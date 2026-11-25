import { getStore } from "@netlify/blobs";
import { createHmac, randomBytes, timingSafeEqual } from "node:crypto";

// Réglages (modifiables dans Netlify > Project configuration > Environment variables)
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

function distanceM(lat1, lng1, lat2, lng2) {
  const R = 6371000, r = (d) => (d * Math.PI) / 180;
  const a = Math.sin(r(lat2 - lat1) / 2) ** 2 + Math.cos(r(lat1)) * Math.cos(r(lat2)) * Math.sin(r(lng2 - lng1) / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

const clean = (s, n = 120) => String(s ?? "").replace(/\s+/g, " ").trim().slice(0, n);

async function listJSON(prefix) {
  const s = store();
  const { blobs } = await s.list({ prefix });
  const out = await Promise.all(blobs.map((b) => s.get(b.key, { type: "json" })));
  return out.filter(Boolean);
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
    if (route === "session" && req.method === "GET") {
      const ses = await s.get(`sessions/${clean(url.searchParams.get("id"), 40)}`, { type: "json" });
      if (!ses) return fail("Séance introuvable.", 404);
      return json({ ok: true, session: { id: ses.id, name: ses.name, type: ses.type, date: ses.date, open: ses.open } });
    }

    if (route === "claim" && req.method === "POST") {
      const id = clean(body.id, 40), device = clean(body.device, 64);
      const ses = await s.get(`sessions/${id}`, { type: "json" });
      if (!ses) return fail("Séance introuvable.", 404);
      if (!ses.open) return fail("Le pointage de cette séance est fermé.");
      if (device.length < 10) return fail("Appareil non reconnu. Rechargez la page.");
      const slot = Math.floor(Date.now() / SLOT_MS);
      let fresh = false;
      for (let i = 0; i < SLOTS_OK; i++) if (sameText(body.t || "", tokenFor(id, slot - i))) fresh = true;
      if (!fresh) return fail("Ce QR code a expiré. Scannez le QR code affiché en ce moment dans la salle.", 410);
      const ts = Date.now();
      return json({ ok: true, pass: passFor(id, device, ts), ts, session: { name: ses.name, type: ses.type, date: ses.date } });
    }

    if (route === "checkin" && req.method === "POST") {
      const id = clean(body.id, 40);
      const ses = await s.get(`sessions/${id}`, { type: "json" });
      if (!ses) return fail("Séance introuvable.", 404);
      if (!ses.open) return fail("Le pointage de cette séance est fermé.");

      const device0 = clean(body.device, 64);
      const ts = Number(body.ts);
      if (!Number.isFinite(ts) || Date.now() - ts > PASS_MS || ts > Date.now() + 5000 || !sameText(body.pass || "", passFor(id, device0, ts)))
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

      const device = device0;

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
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return fail("Date invalide.");
      const id = randomBytes(6).toString("base64url");
      const ses = { id, name, type, date, open: true, createdAt: new Date().toISOString() };
      await s.setJSON(`sessions/${id}`, ses);
      return json({ ok: true, session: ses });
    }

    if (route === "admin/open" && req.method === "POST") {
      const ses = await s.get(`sessions/${clean(body.id, 40)}`, { type: "json" });
      if (!ses) return fail("Séance introuvable.", 404);
      ses.open = !!body.open;
      await s.setJSON(`sessions/${ses.id}`, ses);
      return json({ ok: true, session: ses });
    }

    if (route === "admin/delete" && req.method === "POST") {
      const id = clean(body.id, 40);
      for (const p of [`checkins/${id}/`, `devices/${id}/`]) {
        const { blobs } = await s.list({ prefix: p });
        await Promise.all(blobs.map((b) => s.delete(b.key)));
      }
      await s.delete(`sessions/${id}`);
      return json({ ok: true });
    }

    if (route === "admin/remove-checkin" && req.method === "POST") {
      await s.delete(`checkins/${clean(body.id, 40)}/${clean(body.email, 120).toLowerCase()}`);
      return json({ ok: true });
    }

    if (route === "admin/token") {
      const id = clean(url.searchParams.get("id"), 40);
      const ses = await s.get(`sessions/${id}`, { type: "json" });
      if (!ses) return fail("Séance introuvable.", 404);
      const now = Date.now(), slot = Math.floor(now / SLOT_MS);
      const count = (await s.list({ prefix: `checkins/${id}/` })).blobs.length;
      return json({ ok: true, token: tokenFor(id, slot), msLeft: SLOT_MS - (now % SLOT_MS), open: ses.open, name: ses.name, type: ses.type, date: ses.date, count });
    }

    if (route === "admin/attendance") {
      const id = url.searchParams.get("id");
      if (id) {
        const rows = await listJSON(`checkins/${clean(id, 40)}/`);
        rows.sort((a, b) => a.at.localeCompare(b.at));
        return json({ ok: true, rows });
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
