import { getStore } from "@netlify/blobs";
import { createHmac, createHash, randomBytes, timingSafeEqual } from "node:crypto";

// Réglages (Netlify > Project configuration > Environment variables)
const ADMIN_KEY = process.env.ADMIN_KEY || "";
const SECRET = process.env.QR_SECRET || ADMIN_KEY;
const DOMAINS = (process.env.ALLOWED_DOMAINS || "um6p.ma").split(",").map((d) => d.trim().toLowerCase()).filter(Boolean);
const CAMPUS_LAT = Number(process.env.CAMPUS_LAT || 32.2185);
const CAMPUS_LNG = Number(process.env.CAMPUS_LNG || -7.9370);
const RADIUS_M = Number(process.env.CAMPUS_RADIUS_M || 1500);
const SLOT_MS = 10_000; // le QR change toutes les 10 secondes
const SLOTS_OK = 3; // un QR reste valable 30 secondes au plus
const PASS_MS = 4 * 60_000; // après le scan, 4 minutes pour remplir le formulaire

const json = (data, status = 200) =>
  new Response(JSON.stringify(data), { status, headers: { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" } });
const fail = (message, status = 400) => json({ ok: false, error: message }, status);
const store = () => getStore({ name: "pointage", consistency: "strong" });

function sameText(a, b) {
  const x = Buffer.from(String(a)), y = Buffer.from(String(b));
  return x.length === y.length && timingSafeEqual(x, y);
}
const isAdmin = (key) => ADMIN_KEY.length >= 6 && sameText(key || "", ADMIN_KEY);
const passFor = (id, device, ts) => createHmac("sha256", SECRET).update(`pass:${id}:${device}:${ts}`).digest("base64url").slice(0, 22);
const tokenFor = (id, slot) => createHmac("sha256", SECRET).update(`${id}:${slot}`).digest("base64url").slice(0, 16);
const clean = (s, n = 120) => String(s ?? "").replace(/\s+/g, " ").trim().slice(0, n);
const slug = (s) => clean(s, 80).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
const isDate = (d) => /^\d{4}-\d{2}-\d{2}$/.test(d || "");
const isTime = (t) => /^\d{2}:\d{2}$/.test(t || "");
const COURSE_ID = /^(\d{4}-\d{2}-\d{2})\.([A-Za-z0-9_-]{6,16})$/;

function mondayOf(date) {
  const d = new Date(date + "T12:00:00Z");
  d.setUTCDate(d.getUTCDate() - ((d.getUTCDay() + 6) % 7));
  return d.toISOString().slice(0, 10);
}
function distanceM(lat1, lng1, lat2, lng2) {
  const R = 6371000, r = (d) => (d * Math.PI) / 180;
  const a = Math.sin(r(lat2 - lat1) / 2) ** 2 + Math.cos(r(lat1)) * Math.cos(r(lat2)) * Math.sin(r(lng2 - lng1) / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}
async function listJSON(prefix) {
  const s = store();
  const { blobs } = await s.list({ prefix });
  const out = await Promise.all(blobs.map((b) => s.get(b.key, { type: "json" })));
  return out.filter(Boolean);
}

// ---------- Séances : événements (sessions/) et cours du planning (weeks/) ----------
async function getWeek(monday) {
  return (await store().get(`weeks/${monday}`, { type: "json" })) || { monday, courses: [] };
}
async function getSession(id) {
  id = clean(id, 40);
  const m = id.match(COURSE_ID);
  if (!m) return await store().get(`sessions/${id}`, { type: "json" });
  const wk = await getWeek(m[1]);
  const c = wk.courses.find((x) => x.id === id);
  if (!c) return null;
  return { ...c, name: c.cours, type: "Cours", open: c.open !== false, course: true };
}
async function setOpen(id, open) {
  const m = id.match(COURSE_ID);
  if (!m) {
    const ses = await store().get(`sessions/${id}`, { type: "json" });
    if (!ses) return null;
    ses.open = open;
    await store().setJSON(`sessions/${id}`, ses);
    return ses;
  }
  const wk = await getWeek(m[1]);
  const c = wk.courses.find((x) => x.id === id);
  if (!c) return null;
  c.open = open;
  await store().setJSON(`weeks/${m[1]}`, wk);
  return c;
}
function courseId(c) {
  const h = createHash("sha1").update([c.promotion, c.date, c.debut, c.groupe || "", c.cours].join("|")).digest("base64url").slice(0, 8);
  return `${mondayOf(c.date)}.${h}`;
}
function normCourse(x) {
  const c = {
    date: clean(x.date, 10), debut: clean(x.debut, 5), fin: clean(x.fin, 5), promotion: clean(x.promotion, 60),
    cours: clean(x.cours, 200), groupe: clean(x.groupe, 30), salle: clean(x.salle, 60), enseignant: clean(x.enseignant, 120),
    emailEns: clean(x.emailEns, 200).toLowerCase(),
  };
  if (!isDate(c.date) || !isTime(c.debut) || !isTime(c.fin) || !c.promotion || !c.cours) return null;
  return c;
}

export default async (req) => {
  if (!ADMIN_KEY) return fail("Configuration manquante : ajoutez la variable ADMIN_KEY dans Netlify.", 500);
  const url = new URL(req.url);
  const route = url.pathname.replace(/^\/api\/?/, "").replace(/\/$/, "");
  const body = req.method === "POST" ? await req.json().catch(() => ({})) : {};
  const key = body.key || req.headers.get("x-admin-key") || "";
  const s = store();

  try {
    // ---------- Étudiant ----------
    if (route === "claim" && req.method === "POST") {
      const id = clean(body.id, 40), device = clean(body.device, 64);
      const ses = await getSession(id);
      if (!ses) return fail("Séance introuvable.", 404);
      if (!ses.open) return fail("Le pointage de cette séance est fermé.");
      if (device.length < 10) return fail("Appareil non reconnu. Rechargez la page.");
      const slot = Math.floor(Date.now() / SLOT_MS);
      let fresh = false;
      for (let i = 0; i < SLOTS_OK; i++) if (sameText(body.t || "", tokenFor(id, slot - i))) fresh = true;
      if (!fresh) return fail("Ce QR code a expiré. Scannez le QR code affiché en ce moment dans la salle.", 410);
      const ts = Date.now();
      return json({ ok: true, pass: passFor(id, device, ts), ts, session: { name: ses.name, type: ses.type, date: ses.date, promotion: ses.promotion || "", debut: ses.debut || "", fin: ses.fin || "" } });
    }

    if (route === "checkin" && req.method === "POST") {
      const id = clean(body.id, 40);
      const ses = await getSession(id);
      if (!ses) return fail("Séance introuvable.", 404);
      if (!ses.open) return fail("Le pointage de cette séance est fermé.");
      const device = clean(body.device, 64);
      const ts = Number(body.ts);
      if (!Number.isFinite(ts) || Date.now() - ts > PASS_MS || ts > Date.now() + 5000 || !sameText(body.pass || "", passFor(id, device, ts)))
        return fail("Le délai de pointage est dépassé. Scannez à nouveau le QR code affiché dans la salle.", 410);

      const name = clean(body.name, 80);
      const email = clean(body.email, 120).toLowerCase();
      if (name.length < 3) return fail("Indiquez votre nom et prénom.");
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return fail("Adresse email invalide.");
      const domain = email.split("@")[1];
      if (DOMAINS.length && !DOMAINS.includes(domain)) return fail(`Utilisez votre adresse email ${DOMAINS.map((d) => "@" + d).join(" ou ")}.`);

      const lat = Number(body.lat), lng = Number(body.lng), acc = Math.max(0, Number(body.acc) || 0);
      if (!Number.isFinite(lat) || !Number.isFinite(lng)) return fail("Activez la localisation de votre téléphone pour pointer.");
      const dist = distanceM(lat, lng, CAMPUS_LAT, CAMPUS_LNG);
      if (dist - Math.min(acc, 300) > RADIUS_M) return fail("Vous devez être sur le campus UM6P de Benguerir pour pointer.", 403);

      const record = { session: id, name, email, at: new Date().toISOString(), distance_m: Math.round(dist), device };
      const byEmail = await s.setJSON(`checkins/${id}/${email}`, record, { onlyIfNew: true });
      if (!byEmail.modified) {
        const prev = await s.get(`checkins/${id}/${email}`, { type: "json" });
        return json({ ok: true, already: true, at: prev?.at, name: prev?.name || name });
      }
      const byDevice = await s.setJSON(`devices/${id}/${device}`, { email }, { onlyIfNew: true });
      if (!byDevice.modified) {
        const owner = await s.get(`devices/${id}/${device}`, { type: "json" });
        if (owner?.email !== email) {
          await s.delete(`checkins/${id}/${email}`);
          return fail("Ce téléphone a déjà servi à pointer une autre personne pour cette séance.", 409);
        }
      }
      return json({ ok: true, at: record.at, name });
    }

    // ---------- Administration ----------
    if (!isAdmin(key)) return fail("Code administrateur incorrect.", 401);
    if (route === "admin/check") return json({ ok: true });

    // Événements ponctuels
    if (route === "admin/sessions" && req.method === "GET") {
      const list = await listJSON("sessions/");
      const counts = await Promise.all(list.map(async (x) => (await s.list({ prefix: `checkins/${x.id}/` })).blobs.length));
      list.forEach((x, i) => (x.count = counts[i]));
      list.sort((a, b) => (b.date + b.createdAt).localeCompare(a.date + a.createdAt));
      return json({ ok: true, sessions: list });
    }
    if (route === "admin/sessions" && req.method === "POST") {
      const name = clean(body.name, 120), type = body.type === "Cours" ? "Cours" : "Événement", date = clean(body.date, 10);
      if (!name) return fail("Donnez un nom à la séance.");
      if (!isDate(date)) return fail("Date invalide.");
      const id = randomBytes(6).toString("base64url");
      const ses = { id, name, type, date, open: true, createdAt: new Date().toISOString() };
      await s.setJSON(`sessions/${id}`, ses);
      return json({ ok: true, session: ses });
    }
    if (route === "admin/open" && req.method === "POST") {
      const r = await setOpen(clean(body.id, 40), !!body.open);
      return r ? json({ ok: true }) : fail("Séance introuvable.", 404);
    }
    if (route === "admin/delete" && req.method === "POST") {
      const id = clean(body.id, 40);
      for (const p of [`checkins/${id}/`, `devices/${id}/`]) {
        const { blobs } = await s.list({ prefix: p });
        await Promise.all(blobs.map((b) => s.delete(b.key)));
      }
      const m = id.match(COURSE_ID);
      if (m) {
        const wk = await getWeek(m[1]);
        wk.courses = wk.courses.filter((x) => x.id !== id);
        await s.setJSON(`weeks/${m[1]}`, wk);
      } else await s.delete(`sessions/${id}`);
      return json({ ok: true });
    }
    if (route === "admin/remove-checkin" && req.method === "POST") {
      await s.delete(`checkins/${clean(body.id, 40)}/${clean(body.email, 120).toLowerCase()}`);
      return json({ ok: true });
    }
    if (route === "admin/token") {
      const id = clean(url.searchParams.get("id"), 40);
      const ses = await getSession(id);
      if (!ses) return fail("Séance introuvable.", 404);
      const now = Date.now(), slot = Math.floor(now / SLOT_MS);
      const count = (await s.list({ prefix: `checkins/${id}/` })).blobs.length;
      const sub = ses.course ? [ses.promotion, ses.groupe, `${ses.debut}–${ses.fin}`, ses.salle].filter(Boolean).join(" · ") : "";
      return json({ ok: true, token: tokenFor(id, slot), msLeft: SLOT_MS - (now % SLOT_MS), open: ses.open, name: ses.name, type: ses.type, date: ses.date, sub, count });
    }

    // Promotions et listes d'étudiants
    if (route === "admin/promos" && req.method === "GET") {
      const list = await listJSON("promos/");
      return json({ ok: true, promos: list.map((p) => ({ name: p.name, count: p.students.length, updatedAt: p.updatedAt, students: url.searchParams.get("full") ? p.students : undefined })) });
    }
    if (route === "admin/promos" && req.method === "POST") {
      const name = clean(body.name, 60);
      if (!name) return fail("Choisissez une promotion.");
      const seen = new Set();
      const students = (Array.isArray(body.students) ? body.students : []).map((x) => ({
        nom: clean(x.nom, 80), prenom: clean(x.prenom, 80), email: clean(x.email, 120).toLowerCase(), groupe: clean(x.groupe, 10),
      })).filter((x) => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x.email) && !seen.has(x.email) && seen.add(x.email));
      if (!students.length) return fail("Aucun étudiant valide dans la liste.");
      await s.setJSON(`promos/${slug(name)}`, { name, students, updatedAt: new Date().toISOString() });
      return json({ ok: true, count: students.length });
    }

    // Planning
    if (route === "admin/import" && req.method === "POST") {
      const incoming = (Array.isArray(body.courses) ? body.courses : []).map(normCourse).filter(Boolean);
      if (!incoming.length) return fail("Aucun cours valide à importer.");
      const byWeek = new Map();
      const ids0 = new Set();
      for (const c of incoming) {
        c.id = courseId(c);
        if (ids0.has(c.id)) continue;
        ids0.add(c.id);
        const m = mondayOf(c.date);
        if (!byWeek.has(m)) byWeek.set(m, []);
        byWeek.get(m).push(c);
      }
      let created = 0, updated = 0, removed = 0;
      for (const [monday, list] of byWeek) {
        const wk = await getWeek(monday);
        const promos = new Set(list.map((c) => c.promotion));
        const old = new Map(wk.courses.map((c) => [c.id, c]));
        const keep = wk.courses.filter((c) => !promos.has(c.promotion));
        for (const c of list) {
          const prev = old.get(c.id);
          if (prev) {
            updated++;
            // les modifications faites à la main dans l'outil sont prioritaires sur le fichier
            const edited = new Set((prev.history || []).flatMap((h) => h.changes.map((x) => x.field)));
            const merged = { ...c, open: prev.open, history: prev.history };
            for (const f of ["salle", "enseignant", "emailEns", "debut", "fin"]) if (edited.has(f) || !c[f]) merged[f] = prev[f] || c[f];
            keep.push(merged);
          } else { created++; keep.push({ ...c, open: true }); }
        }
        const ids = new Set(list.map((c) => c.id));
        removed += wk.courses.filter((c) => promos.has(c.promotion) && !ids.has(c.id)).length;
        keep.sort((a, b) => (a.date + a.debut + a.promotion).localeCompare(b.date + b.debut + b.promotion));
        await s.setJSON(`weeks/${monday}`, { monday, courses: keep, updatedAt: new Date().toISOString() });
      }
      return json({ ok: true, created, updated, removed, weeks: byWeek.size });
    }
    if (route === "admin/week" && req.method === "GET") {
      const monday = mondayOf(clean(url.searchParams.get("monday"), 10) || new Date().toISOString().slice(0, 10));
      const wk = await getWeek(monday);
      const { blobs } = await s.list({ prefix: `checkins/${monday}.` });
      const counts = {};
      for (const b of blobs) { const id = b.key.split("/")[1]; counts[id] = (counts[id] || 0) + 1; }
      wk.courses.forEach((c) => (c.count = counts[c.id] || 0));
      return json({ ok: true, monday, courses: wk.courses });
    }
    if (route === "admin/course" && req.method === "POST") {
      const id = clean(body.id, 40), m = id.match(COURSE_ID);
      if (!m) return fail("Cours introuvable.", 404);
      const wk = await getWeek(m[1]);
      const c = wk.courses.find((x) => x.id === id);
      if (!c) return fail("Cours introuvable.", 404);
      const before = { salle: c.salle, enseignant: c.enseignant, emailEns: c.emailEns, debut: c.debut, fin: c.fin };
      for (const f of ["salle", "enseignant", "emailEns", "debut", "fin"]) if (f in body) c[f] = f === "emailEns" ? clean(body[f], 200).toLowerCase() : clean(body[f], 120);
      if (!isTime(c.debut) || !isTime(c.fin)) return fail("Horaire invalide (format HH:MM).");
      const changes = Object.keys(before).filter((k) => before[k] !== c[k]).map((k) => ({ field: k, from: before[k] || "", to: c[k] || "" }));
      if (changes.length) {
        c.history = [...(c.history || []), { at: new Date().toISOString(), changes }].slice(-20);
        await s.setJSON(`weeks/${m[1]}`, wk);
      }
      return json({ ok: true, course: c, changes });
    }

    // Présences
    if (route === "admin/attendance") {
      const id = url.searchParams.get("id");
      if (id) {
        const rows = await listJSON(`checkins/${clean(id, 40)}/`);
        rows.sort((a, b) => a.at.localeCompare(b.at));
        const ses = await getSession(id);
        let roster = null;
        if (ses?.course) roster = (await s.get(`promos/${slug(ses.promotion)}`, { type: "json" }))?.students || null;
        return json({ ok: true, rows, roster });
      }
      const sessions = await listJSON("sessions/");
      const all = [];
      for (const x of sessions) for (const r of await listJSON(`checkins/${x.id}/`)) all.push({ ...r, sessionName: x.name, type: x.type, date: x.date });
      return json({ ok: true, rows: all, sessions });
    }

    return fail("Route inconnue.", 404);
  } catch (e) {
    console.error(e);
    return fail("Erreur serveur. Réessayez dans un instant.", 500);
  }
};

export const config = { path: "/api/*" };
