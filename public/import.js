// Lecture des plannings Excel FMS et des listes d'étudiants (s'exécute dans le navigateur, avec SheetJS)
(function () {
  const MONTHS_EN = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  const MONTHS_FR = { janv: 1, jan: 1, fevr: 2, fev: 2, mars: 3, avr: 4, mai: 5, juin: 6, juil: 7, aout: 8, sept: 9, oct: 10, nov: 11, dec: 12 };
  const DAYS_EN = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
  const DAYS_FR = ['lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'];
  const TEACHER = /^\s*((dr|pr|prof|professor|mr|mrs|ms|mme|mlle|madame|monsieur)\b\.?|m\.\s)/i;
  const ROOM = /(tc[\s-]*f[\s-]*\d+|t\d-\d-\d+|amphi\w*|salle|room|lab\b|hall|teaching center|amphitheat)/i;
  const GROUPS = /\bG\d(\s*(et|&|,)\s*G\d)*\b/i;
  const NONCOURSE = /^(pause|lunch|break|sport)/i;
  const strip = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  const pad = (n) => String(n).padStart(2, '0');
  const iso = (d) => `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  const yearFor = (m) => (m >= 8 ? 2026 : 2027); // année universitaire 2026-2027 (modifiable)
  function mondayOf(y, m, d) { const t = new Date(Date.UTC(y, m - 1, d)); t.setUTCDate(t.getUTCDate() - ((t.getUTCDay() + 6) % 7)); return t; }
  function addDays(t, n) { const x = new Date(t); x.setUTCDate(x.getUTCDate() + n); return x; }

  function normTime(s) {
    const m = [...String(s ?? '').replace(/[–—]/g, '-').replace(/\s/g, '').matchAll(/(\d{1,2})[h:](\d{2})?/g)];
    if (m.length < 2) return [null, null];
    const f = (x) => `${pad(+x[1])}:${x[2] || '00'}`;
    return [f(m[0]), f(m[1])];
  }

  function grid(ws) {
    const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const val = (r, c) => { const x = ws[XLSX.utils.encode_cell({ r, c })]; return x ? x.v : null; };
    const merged = new Map();
    for (const m of ws['!merges'] || [])
      for (let r = m.s.r; r <= m.e.r; r++) for (let c = m.s.c; c <= m.e.c; c++) merged.set(r + ',' + c, [m.s.r, m.s.c]);
    const cell = (r, c) => { const o = merged.get(r + ',' + c) || [r, c]; return { v: val(o[0], o[1]), origin: o.join(',') }; };
    return { ref, val, cell };
  }

  function splitText(text) {
    const raw = String(text).replace(/[ \t]+/g, ' ');
    let parts = raw.split(/\s*\/\s*|\n+/).map((p) => p.replace(/\s{2,}/g, ' ').replace(/^[\s,-]+|[\s,-]+$/g, '')).filter(Boolean);
    const course = [], teachers = [], rooms = [], groups = [], times = [];
    const exp = [];
    for (const p of parts) {
      const m = p.match(/\s[–-]\s((dr|pr|mme|mr|m\.)\s.*)$/i);
      if (m) exp.push(p.slice(0, m.index).trim(), m[1].trim()); else exp.push(p);
    }
    parts = [];
    for (let p of exp) {
      const g = p.match(GROUPS);
      if (g && g[0] !== p) { groups.push(g[0].toUpperCase().replace(/ ET /g, ' et ')); p = (p.slice(0, g.index) + p.slice(g.index + g[0].length)).replace(/^[\s,-]+|[\s,-]+$/g, ''); }
      if (p) parts.push(p);
    }
    for (const p of parts) {
      const t = normTime(p);
      if (t[0] && /^[\dh:\s\-–]+$/i.test(p)) { times.push(t); continue; }
      if (/^G\d(\s*(et|&|,)\s*G\d)*$/i.test(p)) { groups.push(p.toUpperCase().replace(/ ET /g, ' et ')); continue; }
      if (TEACHER.test(p)) { teachers.push(p); continue; }
      if (ROOM.test(p) && p.length < 45) { rooms.push(p.replace(/^salle\s+/i, 'Salle ')); continue; }
      course.push(p);
    }
    return { cours: course.join(' – '), enseignant: teachers.join(' / '), salle: rooms.join(' / '), groupe: groups.join(' et '), horaire: times[0] || null };
  }

  function parseWeekLabel(s) {
    const m = strip(s).replace(/\./g, ' ').match(/([a-z]{3})[a-z]*\s*(\d{1,2})/);
    if (!m || MONTHS_EN.indexOf(m[1]) < 0) return null;
    const mo = MONTHS_EN.indexOf(m[1]) + 1;
    return mondayOf(yearFor(mo), mo, +m[2]);
  }

  // Format MED / PHARMA : tout le semestre, 2 semaines côte à côte
  function parseSemester(ws, promotion) {
    const { ref, val, cell } = grid(ws);
    const rows = [], seen = new Set();
    for (let r0 = ref.s.r; r0 <= ref.e.r; r0++) for (let c0 = ref.s.c; c0 <= ref.e.c; c0++) {
      const h = val(r0, c0);
      if (!(typeof h === 'string' && h.trim().toLowerCase() === 'time')) continue;
      const days = {};
      for (let c = c0 + 1; c < c0 + 8; c++) { const v = val(r0, c); if (typeof v === 'string' && DAYS_EN.includes(v.trim().toLowerCase())) days[c] = DAYS_EN.indexOf(v.trim().toLowerCase()); }
      let monday = null;
      for (let rr = r0 - 1; rr > Math.max(-1, r0 - 6) && !monday; rr--)
        for (let cc = c0; cc < c0 + 8; cc++) { const v = val(rr, cc); if (typeof v === 'string' && v.trim().toLowerCase() === 'week') monday = parseWeekLabel(val(rr, cc + 1)); }
      if (!monday) continue;
      for (let r = r0 + 1; r < r0 + 12; r++) {
        const tv = val(r, c0);
        if (typeof tv === 'string' && tv.trim().toLowerCase() === 'time') break;
        const [start, end] = tv ? normTime(tv) : [null, null];
        if (!start) continue;
        for (const [col, di] of Object.entries(days)) {
          const { v, origin } = cell(r, +col);
          if (!v || !String(v).trim()) continue;
          const k = origin + '|' + start;
          if (seen.has(k)) continue;
          seen.add(k);
          const info = splitText(v);
          if (!info.cours && !info.enseignant) continue;
          const [s, e] = info.horaire || [start, end];
          rows.push({ date: iso(addDays(monday, di)), debut: s, fin: e, promotion, cours: info.cours, groupe: info.groupe, salle: info.salle, enseignant: info.enseignant });
        }
      }
    }
    return rows;
  }

  // Format Kiné / Soins infirmiers : une semaine, jours en colonnes
  function parseWeekFr(ws, promotion, monday) {
    const { ref, val, cell } = grid(ws);
    let hr = null, defaultRoom = '';
    for (let r = ref.s.r; r <= Math.min(ref.e.r, 9); r++) for (let c = ref.s.c; c <= ref.e.c; c++) {
      const v = val(r, c);
      if (typeof v !== 'string') continue;
      if (v.trim().toLowerCase() === 'horaires' && !hr) hr = [r, c];
      if (!hr && /salle/i.test(v)) { const m = v.match(/TC[\s-]*F[\s-]*\d+/i); if (m) defaultRoom = m[0].toUpperCase().replace(/\s/g, ''); }
    }
    if (!hr) return [];
    const [r0, c0] = hr, daycol = {};
    for (let c = c0 + 1; c <= ref.e.c; c++) { const v = cell(r0, c).v; if (typeof v === 'string' && DAYS_FR.includes(strip(v).trim())) daycol[c] = DAYS_FR.indexOf(strip(v).trim()); }
    const rows = [], seen = new Set();
    for (let r = r0 + 1; r <= ref.e.r; r++) {
      const tv = val(r, c0);
      if (typeof tv === 'string' && /^modules/i.test(tv.trim())) break;
      const [start, end] = tv ? normTime(tv) : [null, null];
      if (!start) continue;
      for (const [col, di] of Object.entries(daycol)) {
        const { v, origin } = cell(r, +col);
        if (!v || !String(v).trim() || seen.has(origin)) continue;
        seen.add(origin);
        if (NONCOURSE.test(String(v).trim())) continue;
        const info = splitText(v);
        if (!info.cours && !info.enseignant) continue;
        const [s, e] = info.horaire || [start, end];
        rows.push({ date: iso(addDays(monday, di)), debut: s, fin: e, promotion, cours: info.cours, groupe: info.groupe, salle: info.salle || defaultRoom, enseignant: info.enseignant });
      }
    }
    return rows;
  }

  // Format « modèle » : une ligne = un cours
  function parseTemplate(ws) {
    const data = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false, dateNF: 'yyyy-mm-dd' });
    const k = (o, ...names) => { for (const n of names) for (const key of Object.keys(o)) if (strip(key).trim() === n) return String(o[key]).trim(); return ''; };
    return data.map((o) => {
      let d = k(o, 'date'); const m = d.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
      if (m) d = `${m[3].length === 2 ? '20' + m[3] : m[3]}-${pad(m[2])}-${pad(m[1])}`;
      const t = (x) => { const mm = x.match(/(\d{1,2})[h:](\d{2})/); return mm ? `${pad(mm[1])}:${mm[2]}` : ''; };
      return { date: d, debut: t(k(o, 'debut')), fin: t(k(o, 'fin')), promotion: k(o, 'promotion'), cours: k(o, 'cours'), groupe: k(o, 'groupe'), salle: k(o, 'salle'), enseignant: k(o, 'enseignant'), emailEns: k(o, 'email enseignant') };
    }).filter((x) => x.date && x.cours);
  }

  function guessPromotion(filename, ws) {
    const f = strip(filename);
    const n = (f.match(/\b([1-5])\s*(st|nd|rd|th|er|ere|eme|e)?\b/) || f.match(/_([1-5])_/) || [])[1];
    const yr = (f.match(/(1st|2nd|3rd|4th|5th)/) || [])[1];
    const num = yr ? '12345'['1st2nd3rd4th5th'.indexOf(yr) / 3] : n;
    if (/\bmed/.test(f)) return num ? `Médecine ${num}` : '';
    if (/pha|pahr/.test(f)) return num ? `Pharmacie ${num}` : '';
    if (/kin/.test(f)) return 'Kinésithérapie 1';
    if (/soins|infirm|lsi/.test(f)) {
      let k = num;
      if (ws) { const { val } = grid(ws); for (let r = 0; r < 3; r++) for (let c = 0; c < 3; c++) { const v = strip(val(r, c)); const m = v.match(/(\d)\s*(ere|eme|e)\s*a/); if (m) k = m[1]; } }
      return k ? `Soins infirmiers ${k}` : '';
    }
    return '';
  }

  function guessMonday(filename, ws) {
    const txt = strip(filename + ' ' + (ws ? Object.values(ws).slice(0, 40).map((c) => (c && c.v) || '').join(' ') : '')).replace(/_/g, ' ');
    const m = txt.match(/(?:du|semaine du)\s+(\d{1,2})\s*(janv|jan|fevr|fev|mars|avr|mai|juin|juil|aout|sept|oct|nov|dec)/);
    if (!m) return '';
    const mo = MONTHS_FR[m[2]];
    return iso(mondayOf(yearFor(mo), mo, +m[1]));
  }

  window.FMSImport = {
    // Renvoie { format, promotion, monday, rows }
    parseWorkbook(wb, filename, opts = {}) {
      const ws = wb.Sheets[wb.SheetNames[0]];
      const { ref, val } = grid(ws);
      let hasTime = false, hasHoraires = false, hasHeader = false;
      for (let r = ref.s.r; r <= Math.min(ref.e.r, 15); r++) for (let c = ref.s.c; c <= Math.min(ref.e.c, 20); c++) {
        const v = strip(val(r, c)).trim();
        if (v === 'time') hasTime = true;
        if (v === 'horaires') hasHoraires = true;
        if (r === 0 && v === 'promotion') hasHeader = true;
      }
      const promotion = opts.promotion || guessPromotion(filename, ws);
      if (hasHeader) return { format: 'Modèle FMS (une ligne par cours)', promotion: '', monday: '', rows: parseTemplate(ws) };
      if (hasTime) return { format: 'Planning du semestre (Médecine / Pharmacie)', promotion, monday: '', rows: promotion ? parseSemester(ws, promotion) : [] };
      if (hasHoraires) {
        const monday = opts.monday || guessMonday(filename, ws);
        const m = monday ? monday.split('-').map(Number) : null;
        return { format: 'Planning de la semaine (Kiné / Soins infirmiers)', promotion, monday, rows: promotion && m ? parseWeekFr(ws, promotion, new Date(Date.UTC(m[0], m[1] - 1, m[2]))) : [] };
      }
      return { format: 'Format non reconnu', promotion, monday: '', rows: [] };
    },

    // Liste d'étudiants : copier-coller Outlook (« Prénom NOM » <email>; ...), une ligne par étudiant, ou Excel (Nom, Prénom, Email)
    parseStudents(input) {
      const out = [], seen = new Set();
      const add = (full, email, nom, prenom, groupe) => {
        email = String(email || '').trim().toLowerCase();
        if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || seen.has(email)) return;
        seen.add(email);
        if (!nom && full) {
          const words = String(full).replace(/\bstudent\b/i, '').trim().split(/\s+/);
          const up = words.filter((w) => w.length > 1 && w === w.toUpperCase() && /[A-Z]/.test(w));
          const low = words.filter((w) => !up.includes(w));
          nom = up.join(' '); prenom = low.join(' ');
          if (!nom) { nom = words.slice(-1).join(' '); prenom = words.slice(0, -1).join(' '); }
        }
        out.push({ nom: (nom || '').trim(), prenom: (prenom || '').trim(), email, groupe: (groupe || '').trim().toUpperCase() });
      };
      if (Array.isArray(input)) {
        for (const o of input) {
          const g = (...names) => { for (const key of Object.keys(o)) if (names.includes(strip(key).trim())) return o[key]; return ''; };
          add('', g('email', 'e-mail', 'mail', 'adresse email'), g('nom', 'name', 'last name'), g('prenom', 'first name'), g('groupe', 'group'));
        }
        return out;
      }
      const text = String(input);
      const re = /"?([^"<>;\n]*?)"?\s*<\s*([^<>\s;]+@[^<>\s;]+)\s*>/g;
      let m, any = false;
      while ((m = re.exec(text))) { any = true; add(m[1], m[2]); }
      if (!any) for (const line of text.split(/\r?\n/)) {
        const em = line.match(/[^\s;,\t<>]+@[^\s;,\t<>]+/);
        if (em) add(line.replace(em[0], '').replace(/[;,\t<>"]/g, ' ').replace(/\s+/g, ' ').trim(), em[0]);
      }
      return out;
    },
  };
})();
