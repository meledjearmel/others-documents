// Génère la présentation commerciale. Usage :
//   NODE_PATH=<node_modules> PPTX_SCRIPTS=<skill pptx/scripts> node build.js && python3 animate.py <deck>
const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');
const { applyTheme } = require(process.env.PPTX_SCRIPTS + '/apply_theme.js');

const OUT = "Proposition_Serveur_Proxmox_WindowsServer.pptx";
const THEME = { name: "Infra Bleu", headFontFace: "Calibri", bodyFontFace: "Calibri", colors: {
  dk1: "0B1B3F", lt1: "FFFFFF", dk2: "0B2A6F", lt2: "EEF3FD",
  accent1: "2F6BFF", accent2: "00B8D4", accent3: "7FA4FF", accent4: "C9D8FF", accent5: "16B37A", accent6: "64748B",
  hlink: "2F6BFF", folHlink: "0B2A6F" } };
const H = THEME.colors;

// ---------- Prix ----------
const SERVEUR = 500000, LICENCES = 200000, INSTALL = 450000, SITE = 365000, POSTES = 15;
const TOTAL = SERVEUR + LICENCES + INSTALL;
const fmt = n => n.toLocaleString('fr-FR').replace(/ | /g, ' ');
const F = n => fmt(n) + ' FCFA';

// ---------- SVG / icônes ----------
const iconMarkup = (name, color, x, y, size) => RDS.renderToStaticMarkup(React.createElement(fa[name]))
  .replace(/currentColor/g, '#' + color).replace(/ (height|width)="1em"/g, '')
  .replace('<svg ', `<svg x="${x}" y="${y}" width="${size}" height="${size}" `);
const png = async (svg, w, h, scale = 2) => 'image/png;base64,' +
  (await sharp(Buffer.from(svg), { density: 72 * scale }).resize(Math.round(w * scale), Math.round(h * scale)).png().toBuffer()).toString('base64');
const icon = (name, color) => png(`<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256" viewBox="0 0 256 256">${iconMarkup(name, color, 0, 0, 256)}</svg>`, 256, 256, 1);

// Tuile isométrique centrée en (cx, cy) au niveau z=0
const C30 = Math.cos(Math.PI / 6);
const P = (cx, cy, u, v, z) => [cx + (u - v) * C30, cy + (u + v) * 0.5 - z];
const poly = pts => 'M' + pts.map(p => p[0].toFixed(1) + ',' + p[1].toFixed(1)).join(' L') + ' Z';
function tile(cx, cy, s, t, o) {
  const h = s / 2, k = o.inset ?? 0.82;
  const top = [[-h, -h], [h, -h], [h, h], [-h, h]].map(([u, v]) => P(cx, cy, u, v, t));
  const right = [P(cx, cy, h, -h, t), P(cx, cy, h, h, t), P(cx, cy, h, h, 0), P(cx, cy, h, -h, 0)];
  const left = [P(cx, cy, -h, h, t), P(cx, cy, h, h, t), P(cx, cy, h, h, 0), P(cx, cy, -h, h, 0)];
  const inner = [[-h, -h], [h, -h], [h, h], [-h, h]].map(([u, v]) => P(cx, cy, u * k, v * k, t));
  const op = o.opacity ?? 1, sw = o.round ?? 8;
  let svg = `<g opacity="${op}" stroke-linejoin="round">`;
  svg += `<path d="${poly(right)}" fill="#${o.right}" stroke="#${o.right}" stroke-width="${sw}"/>`;
  svg += `<path d="${poly(left)}" fill="#${o.left}" stroke="#${o.left}" stroke-width="${sw}"/>`;
  svg += `<path d="${poly(top)}" fill="#${o.top}" stroke="#${o.top}" stroke-width="${sw}"/>`;
  if (o.inner) svg += `<path d="${poly(inner)}" fill="#${o.inner}" stroke="#${o.edge || o.inner}" stroke-width="2"/>`;
  if (o.icon) {
    const is = o.iconSize || s * 0.38, [du, dv] = o.iconAt || [0, 0];
    svg += `<g transform="matrix(${C30},0.5,${-C30},0.5,${cx},${cy - t})">${iconMarkup(o.icon, o.iconColor, du - is / 2, dv - is / 2, is)}</g>`;
  }
  return svg + '</g>';
}
const WHITE = { top: "FFFFFF", left: "DDE6FB", right: "B9CCF6", inner: "F3F7FF", edge: "C9D8FF" };
const BLUE = { top: "2F6BFF", left: "1E55E0", right: "1744B8", inner: "4A7FFF", edge: "6E97FF" };
const CYAN = { top: "00B8D4", left: "0096AD", right: "007A8D", inner: "22C8E0", edge: "5ED8EA" };

// Illustration de couverture : chaîne postes → switch → serveur → Windows → accès distant
async function coverArt() {
  const W = 900, Hh = 720, s = 170, t = 26;
  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${Hh}" viewBox="0 0 ${W} ${Hh}">`;
  const c = i => [130 + i * 150, 590 - i * 86.6];
  for (let off of [-120, 0, 120]) {
    const [x0, y0] = c(-1), [x1, y1] = c(6);
    svg += `<line x1="${x0}" y1="${y0 + off}" x2="${x1}" y2="${y1 + off}" stroke="#C9D8FF" stroke-width="2" stroke-dasharray="6 8"/>`;
  }
  const specs = [
    { st: WHITE, icon: null },
    { st: WHITE, icon: "FaNetworkWired", iconColor: H.accent1 },
    { st: BLUE, icon: "FaServer", iconColor: "FFFFFF" },
    { st: WHITE, icon: "FaWindows", iconColor: H.accent1 },
    { st: CYAN, icon: "FaUserShield", iconColor: "FFFFFF" }];
  for (let i = specs.length - 1; i >= 0; i--) {
    const [cx, cy] = c(i), sp = specs[i];
    svg += tile(cx, cy, s, t, { ...sp.st, icon: sp.icon, iconColor: sp.iconColor, iconSize: i == 2 ? 80 : 62 });
    if (i == 0) {
      const minis = [];
      for (const u of [-48, 0, 48]) for (const v of [-48, 0, 48]) minis.push([u, v]);
      minis.sort((a, b) => (a[0] + a[1]) - (b[0] + b[1]));
      for (const [u, v] of minis) {
        const [mx, my] = P(cx, cy, u, v, t);
        svg += tile(mx, my, 38, 10, { ...BLUE, inner: null, round: 4, icon: "FaDesktop", iconColor: "FFFFFF", iconSize: 18 });
      }
    }
  }
  return png(svg + '</svg>', W, Hh);
}

// Pile de couches (slide architecture). Renvoie l'image et le centre y de chaque couche (px)
async function stackArt(layers) {
  const s = 200, t = 22, gap = 105, W = 420;
  const Hh = 110 + (layers.length - 1) * gap + 100 + 40, cx = W / 2;
  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${Hh}" viewBox="0 0 ${W} ${Hh}">`;
  const ys = [];
  layers.forEach((l, k) => {
    const cy = Hh - 120 - k * gap;
    ys.push(cy - t + 50);
    svg += tile(cx, cy, s, t, { ...(l.style || WHITE), icon: l.icon, iconColor: l.iconColor || H.accent1, iconSize: 44, iconAt: [50, 50] });
  });
  return { data: await png(svg + '</svg>', W, Hh), W, H: Hh, ys };
}

// Petite pile de 3 tuiles
async function miniStack(icons) {
  const W = 420, Hh = 420, cx = 210;
  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${Hh}" viewBox="0 0 ${W} ${Hh}">`;
  icons.forEach((ic, k) => {
    const st = k == icons.length - 1 ? BLUE : WHITE;
    svg += tile(cx, 300 - k * 70, 220, 30, { ...st, icon: ic, iconColor: k == icons.length - 1 ? "FFFFFF" : H.accent1, iconSize: 70 });
  });
  return png(svg + '</svg>', W, Hh);
}

// Motif de tuiles transparentes pour les fonds bleus
async function bluePattern() {
  const W = 1333, Hh = 750;
  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${Hh}" viewBox="0 0 ${W} ${Hh}"><rect width="100%" height="100%" fill="#${H.accent1}"/>`;
  const ghost = { top: "FFFFFF", left: "FFFFFF", right: "FFFFFF", opacity: 0.07, round: 14 };
  svg += tile(1120, 300, 420, 60, ghost) + tile(1240, 640, 260, 40, ghost) + tile(80, 760, 300, 40, ghost);
  for (let x = 0; x < W; x += 40) for (let y = 0; y < Hh; y += 40) svg += `<circle cx="${x}" cy="${y}" r="1.4" fill="#FFFFFF" opacity="0.15"/>`;
  return png(svg + '</svg>', W, Hh, 1.5);
}
async function lightGrid() {
  const W = 1333, Hh = 750;
  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${Hh}" viewBox="0 0 ${W} ${Hh}"><rect width="100%" height="100%" fill="#FFFFFF"/>`;
  for (let x = 20; x < W; x += 40) for (let y = 20; y < Hh; y += 40) svg += `<circle cx="${x}" cy="${y}" r="1.3" fill="#D3DEF6"/>`;
  return png(svg + '</svg>', W, Hh, 1.5);
}

// Courbe « plus l'entreprise grandit… » (coordonnées en pouces ×100)
const bez = (p, t) => { const m = 1 - t; return [0, 1].map(i => m * m * m * p[0][i] + 3 * m * m * t * p[1][i] + 3 * m * t * t * p[2][i] + t * t * t * p[3][i]); };
const CURVE = [[60, 650], [620, 650], [760, 190], [1270, 175]];
async function curveArt() {
  const W = 1333, Hh = 750;
  const pts = []; for (let i = 0; i <= 60; i++) pts.push(bez(CURVE, i / 60));
  const line = pts.map((p, i) => (i ? 'L' : 'M') + p[0].toFixed(1) + ',' + p[1].toFixed(1)).join(' ');
  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${Hh}" viewBox="0 0 ${W} ${Hh}">
  <defs><linearGradient id="g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#${H.accent1}" stop-opacity="0.28"/><stop offset="1" stop-color="#${H.accent1}" stop-opacity="0"/></linearGradient></defs>
  <path d="${line} L1270,650 L60,650 Z" fill="url(#g)"/>`;
  for (let i = 1; i <= 4; i++) svg += `<line x1="60" y1="${650 - i * 110}" x2="1270" y2="${650 - i * 110}" stroke="#D3DEF6" stroke-width="1.5" stroke-dasharray="4 6"/>`;
  svg += `<line x1="60" y1="650" x2="1270" y2="650" stroke="#9FB6E8" stroke-width="2"/><path d="${line}" fill="none" stroke="#${H.accent1}" stroke-width="5"/>`;
  return png(svg + '</svg>', W, Hh, 1.5);
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: "Calibri", bodyFontFace: "Calibri" };
  pres.title = "Proposition – Infrastructure serveur pour PME";
  pres.subject = "Proxmox, Windows Server, Active Directory, GPO, Tailscale";
  const C = pres.SchemeColor;
  const GRID = await lightGrid(), BLUEBG = await bluePattern();

  pres.defineSlideMaster({ title: "Couverture", background: { data: GRID }, objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 2.1, w: 6.4, h: 2.2, fontSize: 44, bold: true, color: C.text1, valign: "bottom", align: "left" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 4.45, w: 5.9, h: 1.0, fontSize: 18, color: C.accent6, valign: "top" }, text: "" } }] });
  pres.defineSlideMaster({ title: "Clair", background: { data: GRID }, slideNumber: { x: 12.3, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: C.accent6 }, objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.4, w: 9, h: 0.35, fontSize: 14, bold: true, color: C.accent1 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.75, w: 12.1, h: 0.8, fontSize: 32, bold: true, color: C.text1, align: "left" }, text: "" } },
    { text: { text: "Proposition commerciale · Infrastructure informatique · Prix HT en FCFA", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent6 } } }] });
  pres.defineSlideMaster({ title: "Bleu", background: { data: BLUEBG }, objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 2.2, w: 7, h: 0.45, fontSize: 18, color: C.accent4 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.7, w: 7.4, h: 2.2, fontSize: 40, bold: true, color: C.background1, valign: "top", align: "left" }, text: "" } }] });

  // Groupes d'animation : objectName "aN_..." → apparition dans l'ordre N (voir animate.py)
  let seq = 0;
  const A = n => ({ objectName: `a${n}_${seq++}` });
  const S = (m, sec) => pres.addSlide({ masterName: m, sectionTitle: sec });
  const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, color: C.text1, ...o });
  const box = (s, x, y, w, h, fill, n, extra = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.14, fill: { color: fill }, line: { color: fill }, ...A(n), ...extra });
  const soft = () => ({ shadow: { type: "outer", blur: 10, offset: 3, angle: 90, opacity: 0.12, color: "0B2A6F" } });
  const card = (s, x, y, w, h, n) => box(s, x, y, w, h, C.background1, n, { line: { color: "DCE5F7", width: 1 }, ...soft() });
  const badge = async (s, x, y, d, name, n, bg = C.accent1, fg = "FFFFFF") => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg }, ...A(n) });
    s.addImage({ data: await icon(name, fg), x: x + d * 0.26, y: y + d * 0.26, w: d * 0.48, h: d * 0.48, ...A(n) });
  };
  const chip = (s, x, y, w, text, n) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.42, rectRadius: 0.21, fill: { color: C.background2 }, line: { color: C.accent4, width: 1 }, ...A(n) });
    T(s, text, { x, y, w, h: 0.42, fontSize: 13, bold: true, color: C.text2, align: "center", valign: "middle", ...A(n) });
  };

  // ===== 1. Couverture =====
  pres.addSection({ title: "Ouverture" });
  let s = S("Couverture", "Ouverture");
  await badge(s, 0.6, 0.6, 0.6, "FaServer", 1);
  T(s, "PROPOSITION COMMERCIALE", { x: 1.35, y: 0.6, w: 5, h: 0.6, fontSize: 14, bold: true, color: C.accent1, charSpacing: 2, valign: "middle", ...A(1) });
  s.addText([{ text: "Reprenez le " }, { text: "contrôle", options: { color: C.accent1 } }, { text: " de votre parc informatique" }], { placeholder: "title" });
  s.addText("Un serveur, une administration centralisée, 15 postes sous contrôle — et votre site web offert.", { placeholder: "body" });
  chip(s, 0.6, 5.75, 1.6, "Proxmox VE", 2); chip(s, 2.35, 5.75, 2.0, "Windows Server", 2); chip(s, 4.5, 5.75, 2.4, "Administration à distance", 2);
  s.addImage({ data: await coverArt(), x: 6.9, y: 0.75, w: 6.3, h: 5.04, ...A(1) });
  s.addNotes("Accroche : aujourd'hui, l'informatique de l'entreprise fonctionne… jusqu'au jour où elle ne fonctionne plus. Cette proposition montre comment, avec un seul serveur, on passe d'un parc de postes isolés à un système géré, sécurisé et administré à distance — comme dans une grande entreprise, au budget d'une PME.");

  // ===== 2. Le constat =====
  pres.addSection({ title: "Le constat" });
  s = S("Clair", "Le constat");
  s.addText("Le constat", { placeholder: "kicker" });
  s.addText([{ text: "Plus l'entreprise grandit, plus le " }, { text: "désordre coûte cher", options: { color: C.accent1 } }], { placeholder: "title" });
  s.addImage({ data: await curveArt(), x: 0, y: 0, w: 13.333, h: 7.5, ...A(1) });
  const pains = [[0.18, "Quelques postes : tout se règle « à la main »"], [0.42, "Les mots de passe circulent, chacun installe ce qu'il veut"], [0.66, "Personne ne sait plus qui a accès à quoi"], [0.9, "Une panne ou une fuite de données peut tout arrêter"]];
  pains.forEach(([t, txt], i) => {
    const [px, py] = bez(CURVE, t).map(v => v / 100);
    s.addShape(pres.shapes.OVAL, { x: px - 0.14, y: py - 0.14, w: 0.28, h: 0.28, fill: { color: i == 3 ? "E5484D" : C.accent1 }, line: { color: "FFFFFF", width: 3 }, ...A(2 + i) });
    T(s, txt, { x: i == 3 ? px - 5.1 : Math.max(0.6, px - 4.55), y: i == 3 ? py - 0.25 : i == 0 ? py - 0.85 : py - 0.6, w: i == 3 ? 4.8 : 4.3, h: 0.5, fontSize: 15, bold: i == 3, color: i == 3 ? "C62828" : C.text1, align: i == 0 ? "left" : "right", valign: i == 3 ? "middle" : "bottom", ...A(2 + i) });
  });
  s.addNotes("Faire réagir le client : « Vous reconnaissez-vous à l'une de ces étapes ? » Avec 15 postes, l'entreprise est exactement au moment où le bricolage devient un risque. Chaque départ d'employé, chaque virus, chaque panne coûte du temps et de l'argent.");

  // ===== 3. Mission =====
  s = S("Bleu", "Le constat");
  s.addText("Notre réponse", { placeholder: "kicker" });
  s.addText("Un parc informatique géré, sécurisé et administrable à distance.", { placeholder: "title" });
  T(s, "Comme dans une grande entreprise — au budget d'une PME.", { x: 0.8, y: 5.0, w: 7.4, h: 0.6, fontSize: 20, italic: true, color: C.accent4, ...A(1) });
  s.addNotes("Pause. Laisser la phrase s'installer. Le message : on ne vend pas du matériel, on vend de la tranquillité et du contrôle.");

  // ===== 4. Architecture =====
  pres.addSection({ title: "La solution" });
  s = S("Clair", "La solution");
  s.addText("La solution", { placeholder: "kicker" });
  s.addText([{ text: "Une infrastructure " }, { text: "en couches", options: { color: C.accent1 } }, { text: ", simple et solide" }], { placeholder: "title" });
  const layers = [
    { name: "Serveur physique", desc: "Une machine dédiée, fiable, dans vos locaux", icon: "FaServer" },
    { name: "Proxmox VE", desc: "Virtualisation gratuite : le serveur peut évoluer", icon: "FaCubes" },
    { name: "Windows Server", desc: "Le standard des entreprises pour gérer un réseau", icon: "FaWindows" },
    { name: "AD · DNS · DHCP", desc: "Comptes, noms et adresses réseau centralisés", icon: "FaSitemap" },
    { name: "Stratégies GPO", desc: "Règles de sécurité appliquées automatiquement", icon: "FaShieldAlt" },
    { name: "15 postes", desc: "Chaque employé, son compte, son poste maîtrisé", icon: "FaDesktop", style: BLUE, iconColor: "FFFFFF" }];
  const st = await stackArt(layers);
  const ih = 5.5, sc = ih / st.H, iw = st.W * sc, ix = (13.333 - iw) / 2, iy = 1.45;
  s.addImage({ data: st.data, x: ix, y: iy, w: iw, h: ih, ...A(1) });
  layers.forEach((l, k) => {
    const y = iy + st.ys[k] * sc;
    T(s, l.name, { x: 0.6, y: y - 0.22, w: 2.6, h: 0.44, fontSize: 16, bold: true, color: k == 5 ? C.accent1 : C.text2, valign: "middle", ...A(7 - k) });
    s.addShape(pres.shapes.LINE, { x: 3.2, y, w: ix + iw * 0.3 - 3.2, h: 0, line: { color: C.accent4, width: 1.25, dashType: "dash" }, ...A(7 - k) });
    T(s, l.desc, { x: ix + iw + 0.2, y: y - 0.25, w: 12.73 - (ix + iw + 0.2), h: 0.5, fontSize: 14, color: C.accent6, valign: "middle", ...A(7 - k) });
  });
  s.addNotes("Expliquer couche par couche, de bas en haut : un serveur, Proxmox qui le rend évolutif, Windows Server qui est le standard mondial pour gérer un réseau d'entreprise, puis Active Directory et les GPO qui donnent le contrôle sur les 15 postes.");

  // ===== 5. Avant / après =====
  s = S("Clair", "La solution");
  s.addText("Ce qui change concrètement", { placeholder: "kicker" });
  s.addText([{ text: "Aujourd'hui vs " }, { text: "avec notre solution", options: { color: C.accent1 } }], { placeholder: "title" });
  const rows = [["Comptes", "Mots de passe partagés ou notés sur un papier", "Un compte personnel et sécurisé par employé"],
    ["Départ d'un employé", "Ses accès restent ouverts", "Compte désactivé en un clic"],
    ["Logiciels", "Chacun installe ce qu'il veut", "Installations contrôlées par GPO"],
    ["Mises à jour", "Faites… quand on y pense", "Appliquées automatiquement"],
    ["Dépannage", "Déplacement sur site, attente", "Intervention immédiate à distance"]];
  box(s, 3.3, 1.75, 4.5, 0.5, "E2E8F0", 1);
  T(s, "Aujourd'hui", { x: 3.3, y: 1.75, w: 4.5, h: 0.5, fontSize: 15, bold: true, color: C.accent6, align: "center", valign: "middle", ...A(1) });
  box(s, 8.0, 1.75, 4.7, 0.5, C.accent1, 1);
  T(s, "Avec notre solution", { x: 8.0, y: 1.75, w: 4.7, h: 0.5, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", ...A(1) });
  const xIc = await icon("FaTimesCircle", "94A3B8"), okIc = await icon("FaCheckCircle", H.accent1);
  rows.forEach((r, i) => {
    const y = 2.45 + i * 0.85;
    T(s, r[0], { x: 0.6, y, w: 2.5, h: 0.7, fontSize: 16, bold: true, color: C.text2, valign: "middle", ...A(2 + i) });
    box(s, 3.3, y, 4.5, 0.7, "F1F5F9", 2 + i);
    s.addImage({ data: xIc, x: 3.5, y: y + 0.2, w: 0.3, h: 0.3, ...A(2 + i) });
    T(s, r[1], { x: 3.95, y, w: 3.75, h: 0.7, fontSize: 14, color: C.accent6, valign: "middle", ...A(2 + i) });
    card(s, 8.0, y, 4.7, 0.7, 2 + i);
    s.addImage({ data: okIc, x: 8.2, y: y + 0.2, w: 0.3, h: 0.3, ...A(2 + i) });
    T(s, r[2], { x: 8.65, y, w: 3.95, h: 0.7, fontSize: 14, bold: true, color: C.text1, valign: "middle", ...A(2 + i) });
  });
  s.addNotes("C'est la slide qui vend. Insister sur le départ d'un employé : aujourd'hui, un ancien salarié peut encore accéder aux données. Demain : un clic et c'est fini.");

  // ===== 6. Administration à distance =====
  s = S("Clair", "La solution");
  s.addText("Administration à distance", { placeholder: "kicker" });
  s.addText([{ text: "Un administrateur, " }, { text: "où qu'il soit", options: { color: C.accent1 } }], { placeholder: "title" });
  await badge(s, 0.9, 2.5, 1.5, "FaLaptopCode", 1);
  T(s, "Administrateur", { x: 0.4, y: 4.15, w: 2.5, h: 0.4, fontSize: 16, bold: true, color: C.text2, align: "center", ...A(1) });
  T(s, "Depuis le bureau, la maison ou en déplacement", { x: 0.4, y: 4.55, w: 2.5, h: 0.7, fontSize: 13, color: C.accent6, align: "center", ...A(1) });
  s.addShape(pres.shapes.LINE, { x: 2.6, y: 3.25, w: 2.6, h: 0, line: { color: C.accent2, width: 3, dashType: "dash" }, ...A(2) });
  await badge(s, 3.55, 2.9, 0.7, "FaLock", 2, C.accent2);
  T(s, "Tunnel chiffré Tailscale", { x: 2.6, y: 3.75, w: 2.6, h: 0.4, fontSize: 13, bold: true, color: C.accent2, align: "center", ...A(2) });
  s.addImage({ data: await miniStack(["FaServer", "FaWindows"]), x: 5.1, y: 1.75, w: 2.9, h: 2.9, ...A(3) });
  T(s, "Serveur de l'entreprise", { x: 5.0, y: 4.55, w: 3.1, h: 0.4, fontSize: 16, bold: true, color: C.text2, align: "center", ...A(3) });
  s.addShape(pres.shapes.LINE, { x: 8.0, y: 3.25, w: 0.9, h: 0, line: { color: C.accent1, width: 3, endArrowType: "triangle" }, ...A(4) });
  const desk = await icon("FaDesktop", H.accent1);
  for (let i = 0; i < 15; i++) {
    const x = 9.1 + (i % 5) * 0.72, y = 2.1 + Math.floor(i / 5) * 0.8;
    box(s, x, y, 0.6, 0.6, C.background2, 4);
    s.addImage({ data: desk, x: x + 0.15, y: y + 0.15, w: 0.3, h: 0.3, ...A(4) });
  }
  T(s, "15 postes des employés", { x: 9.1, y: 4.55, w: 3.5, h: 0.4, fontSize: 16, bold: true, color: C.text2, align: "center", ...A(4) });
  [["0", "port ouvert sur Internet"], ["24h/24", "accès pour l'administrateur"], ["Minutes", "au lieu d'un déplacement"]].forEach((t, i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 5.45, 3.9, 1.25, 5);
    T(s, t[0], { x: x + 0.3, y: 5.55, w: 3.4, h: 0.6, fontSize: 28, bold: true, color: C.accent1, ...A(5) });
    T(s, t[1], { x: x + 0.3, y: 6.15, w: 3.4, h: 0.4, fontSize: 14, color: C.accent6, ...A(5) });
  });
  s.addNotes("Tailscale crée un tunnel chiffré : l'administrateur gère tout le parc à distance, sans exposer le serveur sur Internet. Résultat : un problème réglé en quelques minutes, sans attendre un déplacement.");

  // ===== 7. Pourquoi ces choix =====
  s = S("Clair", "La solution");
  s.addText("Pourquoi ces choix", { placeholder: "kicker" });
  s.addText([{ text: "Des technologies " }, { text: "éprouvées", options: { color: C.accent1 } }, { text: ", sans abonnement caché" }], { placeholder: "title" });
  s.addShape(pres.shapes.OVAL, { x: 0.9, y: 1.8, w: 4.8, h: 4.8, fill: { color: C.background2 }, line: { color: C.background2 }, ...A(1) });
  s.addImage({ data: await miniStack(["FaServer", "FaCubes", "FaWindows"]), x: 1.0, y: 1.9, w: 4.6, h: 4.6, ...A(1) });
  const why = [["Évolutif", "Proxmox permet d'ajouter d'autres serveurs virtuels plus tard, sans racheter de matériel."],
    ["Standard du marché", "Windows Server et Active Directory : la référence pour gérer les postes d'une entreprise."],
    ["Économique", "Proxmox VE et Tailscale sont gratuits : aucun abonnement mensuel à prévoir."]];
  for (let i = 0; i < 3; i++) {
    const y = 1.95 + i * 1.6;
    await badge(s, 6.5, y, 0.6, "FaCheck", 2 + i, C.accent1);
    T(s, why[i][0], { x: 7.35, y: y - 0.05, w: 5.3, h: 0.5, fontSize: 20, bold: true, color: C.text2, ...A(2 + i) });
    T(s, why[i][1], { x: 7.35, y: y + 0.45, w: 5.3, h: 0.8, fontSize: 15, color: C.accent6, valign: "top", ...A(2 + i) });
  }
  s.addNotes("Rassurer : ce ne sont pas des solutions exotiques. Et surtout, pas d'abonnement : une fois installé, le système appartient à l'entreprise.");

  // ===== 8. Déploiement =====
  pres.addSection({ title: "Investissement" });
  s = S("Clair", "Investissement");
  s.addText("Le déploiement", { placeholder: "kicker" });
  s.addText([{ text: "Nous nous occupons de " }, { text: "tout", options: { color: C.accent1 } }], { placeholder: "title" });
  const steps = [["FaServer", "Serveur", "Installation du serveur et de Proxmox VE"], ["FaWindows", "Windows Server", "Active Directory, DNS et DHCP"],
    ["FaShieldAlt", "Sécurité", "GPO sur mesure : mots de passe, mises à jour, restrictions"], ["FaDesktop", "15 postes", "Intégration de chaque poste au domaine"],
    ["FaUserShield", "Accès distant", "Tailscale configuré et prise en main"]];
  s.addShape(pres.shapes.LINE, { x: 1.4, y: 2.45, w: 10.5, h: 0, line: { color: C.accent4, width: 2, dashType: "dash" }, ...A(1) });
  for (let i = 0; i < 5; i++) {
    const x = 0.6 + i * 2.45;
    await badge(s, x + 0.45, 2.0, 0.9, steps[i][0], 2 + i, i == 4 ? C.accent2 : C.accent1);
    card(s, x, 3.2, 2.25, 2.3, 2 + i);
    T(s, "Étape " + (i + 1), { x: x + 0.2, y: 3.35, w: 1.9, h: 0.3, fontSize: 12, bold: true, color: C.accent1, ...A(2 + i) });
    T(s, steps[i][1], { x: x + 0.2, y: 3.65, w: 1.9, h: 0.45, fontSize: 17, bold: true, color: C.text2, ...A(2 + i) });
    T(s, steps[i][2], { x: x + 0.2, y: 4.15, w: 1.9, h: 1.2, fontSize: 14, color: C.accent6, valign: "top", ...A(2 + i) });
  }
  box(s, 0.6, 5.85, 12.1, 0.85, C.text2, 7);
  T(s, [{ text: "Installation, configuration et déploiement complet : ", options: { color: C.accent4 } }, { text: F(INSTALL), options: { bold: true, color: C.background1 } }],
    { x: 0.9, y: 5.85, w: 11.5, h: 0.85, fontSize: 18, valign: "middle", ...A(7) });
  s.addNotes("Le client n'a rien à gérer : du serveur jusqu'au dernier poste, tout est installé, configuré et testé. Le forfait couvre les 15 postes.");

  // ===== 9. Investissement =====
  s = S("Clair", "Investissement");
  s.addText("Votre investissement", { placeholder: "kicker" });
  s.addText("Un investissement unique", { placeholder: "title" });
  T(s, [{ text: "15 postes administrés comme dans une grande entreprise, pour ", options: {} }, { text: "≈ " + F(Math.round(TOTAL / POSTES / 100) * 100) + " par poste.", options: { bold: true, color: C.accent1 } }],
    { x: 0.6, y: 1.9, w: 6.6, h: 1.8, fontSize: 26, valign: "top", ...A(1) });
  T(s, "Paiement unique · aucun abonnement obligatoire", { x: 0.6, y: 3.75, w: 6.6, h: 0.4, fontSize: 15, italic: true, color: C.accent6, ...A(1) });
  const circ = (x, y, d, fill, n, big, label, color, lc) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill }, ...A(n) });
    T(s, [{ text: big, options: { fontSize: d > 3 ? 40 : d > 2 ? 24 : 20, bold: true, breakLine: true } }, { text: label, options: { fontSize: d > 3 ? 16 : 13, color: lc } }],
      { x: x + 0.15, y: y + d * 0.2, w: d - 0.3, h: d * 0.6, align: "center", valign: "middle", color, ...A(n) });
  };
  circ(8.4, 1.5, 4.2, C.accent1, 2, fmt(TOTAL), "FCFA HT · investissement total", C.background1, C.accent4);
  circ(5.2, 4.3, 2.6, C.accent3, 3, fmt(SERVEUR), "Serveur", C.background1, C.background1);
  circ(7.6, 4.75, 2.2, C.text2, 4, fmt(INSTALL), "Installation et 15 postes", C.background1, C.accent4);
  circ(10.9, 5.3, 1.7, C.background2, 5, fmt(LICENCES), "Licences ×2", C.text2, C.accent6);
  s.addNotes("Annoncer le total puis ramener au coût par poste : environ 77 000 FCFA par poste, une seule fois. À comparer au coût d'une seule journée d'arrêt ou d'une perte de données.");

  // ===== 10. Maintenance =====
  s = S("Clair", "Investissement");
  s.addText("Après le déploiement", { placeholder: "kicker" });
  s.addText([{ text: "Nous restons " }, { text: "à vos côtés", options: { color: C.accent1 } }], { placeholder: "title" });
  card(s, 0.6, 1.8, 5.9, 4.6, 1);
  await badge(s, 0.95, 2.15, 0.9, "FaGift", 1, C.accent1);
  T(s, "3 mois", { x: 0.95, y: 3.25, w: 5.2, h: 1.0, fontSize: 60, bold: true, color: C.accent1, ...A(1) });
  T(s, "de maintenance offerts", { x: 0.95, y: 4.25, w: 5.2, h: 0.5, fontSize: 22, bold: true, color: C.text2, ...A(1) });
  T(s, "Dès la mise en service : surveillance, corrections et ajustements des GPO, sans frais.", { x: 0.95, y: 4.85, w: 5.2, h: 1.2, fontSize: 15, color: C.accent6, valign: "top", ...A(1) });
  s.addShape(pres.shapes.LINE, { x: 6.7, y: 4.1, w: 0.6, h: 0, line: { color: C.accent4, width: 2, endArrowType: "triangle" }, ...A(2) });
  box(s, 7.5, 1.8, 5.2, 4.6, C.text2, 2, soft());
  await badge(s, 7.85, 2.15, 0.9, "FaHeadset", 2, C.accent2);
  T(s, "Ensuite, en option", { x: 7.85, y: 3.25, w: 4.6, h: 0.4, fontSize: 16, color: C.accent4, ...A(2) });
  T(s, "à partir de 100 000 FCFA", { x: 7.85, y: 3.65, w: 4.6, h: 0.6, fontSize: 28, bold: true, color: C.background1, ...A(2) });
  T(s, "par trimestre · formule à définir ensemble", { x: 7.85, y: 4.25, w: 4.6, h: 0.4, fontSize: 15, color: C.accent4, ...A(2) });
  T(s, [{ text: "Mises à jour et supervision du serveur", options: { bullet: true, breakLine: true } }, { text: "Gestion des comptes et des GPO", options: { bullet: true, breakLine: true } }, { text: "Assistance à distance", options: { bullet: true } }],
    { x: 7.85, y: 4.8, w: 4.6, h: 1.4, fontSize: 15, color: C.background1, paraSpaceAfter: 6, valign: "top", ...A(3) });
  s.addNotes("Rassurer : le client n'est pas laissé seul. Trois mois offerts pour qu'il constate la valeur, puis une maintenance trimestrielle à partir de 100 000 FCFA, à discuter selon ses besoins.");

  // ===== 11. Site offert =====
  s = S("Clair", "Investissement");
  s.addText("Et en cadeau", { placeholder: "kicker" });
  s.addText([{ text: "Votre site web vitrine " }, { text: "offert", options: { color: C.accent1 } }], { placeholder: "title" });
  box(s, 0.6, 1.75, 6.2, 4.9, "DCE5F7", 1, soft());
  s.addShape(pres.shapes.RECTANGLE, { x: 0.75, y: 2.2, w: 5.9, h: 4.3, fill: { color: C.background1 }, line: { color: C.background1 }, ...A(1) });
  ["E5484D", "F5B700", "16B37A"].forEach((c, i) => s.addShape(pres.shapes.OVAL, { x: 0.9 + i * 0.28, y: 1.9, w: 0.16, h: 0.16, fill: { color: c }, line: { color: c }, ...A(1) }));
  s.addShape(pres.shapes.RECTANGLE, { x: 0.75, y: 2.2, w: 5.9, h: 1.9, fill: { color: C.accent1 }, line: { color: C.accent1 }, ...A(1) });
  T(s, "Votre Entreprise", { x: 1.05, y: 2.45, w: 5.3, h: 0.6, fontSize: 26, bold: true, color: C.background1, ...A(1) });
  T(s, "Votre savoir-faire, visible partout", { x: 1.05, y: 3.05, w: 5.3, h: 0.4, fontSize: 14, color: C.accent4, ...A(1) });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.05, y: 3.55, w: 1.7, h: 0.38, rectRadius: 0.19, fill: { color: C.background1 }, line: { color: C.background1 }, ...A(1) });
  T(s, "Nous contacter", { x: 1.05, y: 3.55, w: 1.7, h: 0.38, fontSize: 12, bold: true, color: C.accent1, align: "center", valign: "middle", ...A(1) });
  for (let i = 0; i < 3; i++) {
    const x = 1.05 + i * 1.82;
    box(s, x, 4.35, 1.6, 1.9, C.background2, 1);
    await badge(s, x + 0.5, 4.6, 0.6, ["FaBriefcase", "FaStar", "FaEnvelope"][i], 1);
    T(s, ["Services", "Réalisations", "Contact"][i], { x, y: 5.4, w: 1.6, h: 0.45, fontSize: 14, bold: true, color: C.text2, align: "center", ...A(1) });
  }
  T(s, [{ text: "Valeur " + F(SITE), options: { strike: "sngStrike" } }], { x: 7.3, y: 1.8, w: 5.4, h: 0.45, fontSize: 18, color: C.accent6, ...A(2) });
  T(s, "OFFERT", { x: 7.3, y: 2.2, w: 5.4, h: 0.9, fontSize: 54, bold: true, color: C.accent1, ...A(2) });
  const feats = [["FaMobileAlt", "Adapté au mobile, à la tablette et à l'ordinateur"], ["FaFileAlt", "Accueil, À propos, Services, Réalisations, Contact"], ["FaWhatsapp", "Formulaire de contact, Google Maps et WhatsApp"], ["FaSearch", "Visible sur Google, sécurisé en HTTPS"]];
  for (let i = 0; i < feats.length; i++) {
    const y = 3.35 + i * 0.62;
    await badge(s, 7.3, y, 0.48, feats[i][0], 3 + i, C.background2, H.accent1);
    T(s, feats[i][1], { x: 7.95, y, w: 4.8, h: 0.48, fontSize: 15, valign: "middle", ...A(3 + i) });
  }
  box(s, 7.3, 5.95, 5.4, 0.75, C.text2, 7);
  T(s, [{ text: "Valeur totale " + F(TOTAL + SITE) + "  →  ", options: { color: C.accent4 } }, { text: "vous payez " + F(TOTAL), options: { bold: true, color: C.background1 } }],
    { x: 7.5, y: 5.95, w: 5.1, h: 0.75, fontSize: 15, valign: "middle", ...A(7) });
  s.addNotes("Le geste commercial : le site vitrine, d'une valeur de 365 000 FCFA, est offert. L'entreprise gagne à la fois un outil interne solide et une vitrine pour attirer de nouveaux clients.");

  // ===== Option sécurité =====
  pres.addSection({ title: "Option sécurité" });
  s = S("Bleu", "Option sécurité");
  s.addText("Option · proposition séparée", { placeholder: "kicker" });
  s.addText("Et si un ransomware frappait demain ?", { placeholder: "title" });
  T(s, "Un poste infecté peut chiffrer les fichiers partagés et paralyser toute l'activité. Ajoutons une protection professionnelle.", { x: 0.8, y: 4.4, w: 7.2, h: 1.2, fontSize: 20, color: C.accent4, valign: "top", ...A(1) });
  await badge(s, 9.6, 2.3, 2.6, "FaShieldVirus", 2, C.background1, H.accent1);
  s.addNotes("Transition : l'infrastructure organise le parc, mais il faut aussi le protéger. Cette option est indépendante de la proposition principale : le client peut la prendre maintenant ou plus tard.");

  s = S("Clair", "Option sécurité");
  s.addText("Option sécurité · comment ça fonctionne", { placeholder: "kicker" });
  s.addText([{ text: "Une console, " }, { text: "tous les appareils protégés", options: { color: C.accent1 } }], { placeholder: "title" });
  await badge(s, 0.7, 3.3, 1.2, "FaUserCog", 1, C.accent2);
  T(s, "Administrateur", { x: 0.3, y: 4.6, w: 2.0, h: 0.4, fontSize: 15, bold: true, color: C.text2, align: "center", ...A(1) });
  s.addShape(pres.shapes.LINE, { x: 2.0, y: 3.9, w: 0.8, h: 0, line: { color: C.accent2, width: 3, dashType: "dash" }, ...A(2) });
  box(s, 2.9, 3.1, 2.3, 1.6, C.text2, 2, soft());
  T(s, [{ text: "Console de gestion", options: { bold: true, fontSize: 16, breakLine: true } }, { text: "dans le cloud (ou sur site)", options: { fontSize: 13 } }], { x: 3.05, y: 3.1, w: 2.0, h: 1.6, color: C.background1, align: "center", valign: "middle", ...A(2) });
  s.addShape(pres.shapes.LINE, { x: 5.3, y: 3.9, w: 0.9, h: 0, line: { color: C.accent1, width: 3, beginArrowType: "triangle", endArrowType: "triangle" }, ...A(3) });
  s.addShape(pres.shapes.OVAL, { x: 6.4, y: 1.75, w: 4.3, h: 4.3, fill: { color: C.background2 }, line: { color: C.accent4, width: 1.5, dashType: "dash" }, ...A(3) });
  await badge(s, 8.1, 3.45, 0.9, "FaUsers", 3);
  const devs = ["FaWindows", "FaApple", "FaLinux", "FaServer", "FaAndroid", "FaMobileAlt"];
  for (let i = 0; i < 6; i++) {
    const ang = -Math.PI / 2 + i * Math.PI / 3, cx = 8.55 + 1.65 * Math.cos(ang), cy = 3.9 + 1.65 * Math.sin(ang);
    await badge(s, cx - 0.35, cy - 0.35, 0.7, devs[i], 4, C.background1, H.accent1);
  }
  T(s, "Windows · macOS · Linux · serveurs · Android · iOS", { x: 6.0, y: 6.15, w: 5.1, h: 0.4, fontSize: 14, bold: true, color: C.text2, align: "center", ...A(4) });
  ["Protection des terminaux", "Contrôles Web, appareils, applications", "Analyse des causes", "Campagnes de phishing simulées"].forEach((t, i) => {
    box(s, 11.0, 1.9 + i * 1.05, 1.75, 0.85, C.accent1, 5 + i);
    T(s, t, { x: 11.05, y: 1.9 + i * 1.05, w: 1.65, h: 0.85, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle", ...A(5 + i) });
  });
  s.addNotes("L'administrateur pilote la sécurité de tous les appareils depuis une seule console, à distance, comme pour le reste de l'infrastructure. Les postes, le serveur et les téléphones sont couverts.");

  s = S("Clair", "Option sécurité");
  s.addText("Option sécurité", { placeholder: "kicker" });
  s.addText([{ text: "Kaspersky Next " }, { text: "EDR Foundations", options: { color: C.accent1 } }], { placeholder: "title" });
  const sec = [["FaLock", "Protège vos données", "Anti-ransomware et anti-malware éprouvés"],
    ["FaBug", "Bloque les malwares", "Connus et inconnus : fichiers, Web et e-mails"],
    ["FaSyncAlt", "Applications à jour", "Détection automatique des logiciels obsolètes"],
    ["FaUserLock", "Contrôle des usages", "Applications, Web et périphériques encadrés"],
    ["FaProjectDiagram", "Attaques avancées", "Détectées, éliminées, avec analyse des causes"],
    ["FaBell", "Alertes", "Activités suspectes et failles signalées"]];
  for (let i = 0; i < 6; i++) {
    const x = 0.6 + (i % 2) * 3.85, y = 1.75 + Math.floor(i / 2) * 1.6;
    card(s, x, y, 3.65, 1.4, 1 + Math.floor(i / 2));
    await badge(s, x + 0.25, y + 0.35, 0.7, sec[i][0], 1 + Math.floor(i / 2));
    T(s, sec[i][1], { x: x + 1.15, y: y + 0.2, w: 2.35, h: 0.4, fontSize: 15, bold: true, color: C.text2, ...A(1 + Math.floor(i / 2)) });
    T(s, sec[i][2], { x: x + 1.15, y: y + 0.6, w: 2.35, h: 0.7, fontSize: 13, color: C.accent6, valign: "top", ...A(1 + Math.floor(i / 2)) });
  }
  box(s, 8.5, 1.8, 4.2, 4.65, C.text2, 5, soft());
  await badge(s, 8.8, 2.1, 0.9, "FaQuestion", 5, C.accent2);
  T(s, "Combien vous coûterait une journée d'arrêt ?", { x: 8.8, y: 3.15, w: 3.6, h: 1.0, fontSize: 22, bold: true, color: C.background1, valign: "top", ...A(5) });
  T(s, "Fichiers chiffrés, clients en attente, données perdues… La vraie question n'est pas le prix de la protection, mais celui d'une attaque.", { x: 8.8, y: 4.2, w: 3.6, h: 1.3, fontSize: 14, color: C.accent4, valign: "top", ...A(6) });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.8, y: 5.65, w: 3.6, h: 0.5, rectRadius: 0.25, fill: { color: C.accent2 }, line: { color: C.accent2 }, ...A(7) });
  T(s, "Parlons-en · démonstration sur demande", { x: 8.8, y: 5.65, w: 3.6, h: 0.5, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", ...A(7) });
  T(s, "Option indépendante de l'infrastructure · offre adaptée à la taille de votre parc", { x: 0.6, y: 6.6, w: 12.1, h: 0.3, fontSize: 11, italic: true, color: C.accent6, ...A(7) });
  s.addNotes("Ne pas annoncer de prix : poser la question du coût d'une journée d'arrêt et laisser le client y réfléchir. Proposer une démonstration ou un devis séparé s'il montre de l'intérêt.");

  // ===== 11. Appel à l'action =====
  pres.addSection({ title: "Conclusion" });
  s = S("Bleu", "Conclusion");
  s.addText("Prochaine étape", { placeholder: "kicker" });
  s.addText("Prêt à reprendre le contrôle ?", { placeholder: "title" });
  T(s, "Votre infrastructure opérationnelle en 1 à 2 semaines.", { x: 0.8, y: 4.1, w: 7.0, h: 0.6, fontSize: 20, color: C.accent4, ...A(1) });
  const next = [["FaHandshake", "Validation de la proposition"], ["FaTruck", "Livraison et installation du serveur"], ["FaRocket", "Mise en service et site en ligne"]];
  for (let i = 0; i < 3; i++) {
    const y = 1.6 + i * 1.55;
    box(s, 8.6, y, 4.1, 1.3, C.background1, 2 + i, soft());
    await badge(s, 8.85, y + 0.25, 0.8, next[i][0], 2 + i, C.accent1);
    T(s, next[i][1], { x: 9.85, y, w: 2.7, h: 1.3, fontSize: 16, bold: true, color: C.text2, valign: "middle", ...A(2 + i) });
  }
  s.addNotes("Conclure par une question fermée : « On planifie l'installation pour quelle semaine ? »");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log(TOTAL);
})();
