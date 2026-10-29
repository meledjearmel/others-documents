const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');
const { applyTheme } = require(process.env.PPTX_SCRIPTS + '/apply_theme.js');

const THEME = { name: "Infra PME", headFontFace: "Cambria", bodyFontFace: "Calibri", colors: { dk1: "1B1F24", lt1: "FFFFFF", dk2: "0F3D3E", lt2: "EEF4F3", accent1: "0F7C6E", accent2: "E57A2E", accent3: "3A6B8C", accent4: "8FB8AE", accent5: "C9D6D3", accent6: "5C6670", hlink: "0F7C6E", folHlink: "3A6B8C" } };
const HEX = THEME.colors;
const f = n => n.toLocaleString('fr-FR').replace(/ | /g, ' ') + ' FCFA';
const icon = async (name, color) => {
  const svg = RDS.renderToStaticMarkup(React.createElement(fa[name], { color: '#' + color, size: 256 }));
  return 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
};

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: "Cambria", bodyFontFace: "Calibri" };
  pres.title = "Proposition infrastructure serveur PME";
  const C = pres.SchemeColor;

  pres.defineSlideMaster({ title: "Sombre", background: { color: C.text2 }, objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.6, w: 7.2, h: 2.2, fontSize: 44, bold: true, color: C.background1, fontFace: "Cambria", valign: "bottom" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.0, w: 7.2, h: 1.6, fontSize: 20, color: C.accent4, valign: "top" }, text: "" } }] });
  pres.defineSlideMaster({ title: "Contenu", background: { color: C.background1 }, slideNumber: { x: 12.4, y: 7.0, fontSize: 10, color: C.accent6 }, objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 0.9, fontSize: 32, bold: true, color: C.text2, fontFace: "Cambria" }, text: "" } },
    { text: { text: "Prix indicatifs HT en FCFA – octobre 2026", options: { x: 0.6, y: 7.0, w: 6, h: 0.3, fontSize: 10, color: C.accent6 } } }] });

  const S = (m, sec) => pres.addSlide({ masterName: m, sectionTitle: sec });
  const card = (s, x, y, w, h, fill, shadow) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill || C.background2 }, line: { color: fill || C.background2 }, ...(shadow ? { shadow: { type: "outer", blur: 8, offset: 2, angle: 90, opacity: 0.18, color: "000000" } } : {}) });
  const badge = async (s, x, y, d, name, bg = C.accent1, fg = HEX.lt1) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
    s.addImage({ data: await icon(name, fg), x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5 });
  };
  const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, color: C.text1, ...o });

  // 1. Couverture
  pres.addSection({ title: "Introduction" });
  let s = S("Sombre", "Introduction");
  s.addText("Votre infrastructure informatique clé en main", { placeholder: "title" });
  s.addText("Serveur Proxmox + Windows Server pour administrer les postes de votre entreprise — et votre site web vitrine offert", { placeholder: "body" });
  const cov = [["FaServer", "Serveur Proxmox"], ["FaWindows", "Windows Server"], ["FaUsersCog", "AD · DNS · DHCP · GPO"], ["FaGlobe", "Site vitrine offert"]];
  for (let i = 0; i < cov.length; i++) {
    const y = 1.2 + i * 1.3;
    card(s, 8.8, y, 3.9, 1.05, C.background1);
    await badge(s, 9.0, y + 0.18, 0.7, cov[i][0], i == 3 ? C.accent2 : C.accent1);
    T(s, cov[i][1], { x: 9.9, y: y + 0.1, w: 2.7, h: 0.85, fontSize: 16, bold: true, color: C.text2, valign: "middle" });
  }

  // 2. Proposition
  s = S("Contenu", "Introduction");
  s.addText("Ce que nous vous proposons", { placeholder: "title" });
  const prop = [["FaServer", "Un serveur virtualisé", "Proxmox VE (gratuit) héberge Windows Server et laisse de la place pour évoluer"],
    ["FaUsers", "Des comptes centralisés", "Active Directory : un compte par employé, connexion sur n'importe quel poste"],
    ["FaShieldAlt", "Des postes maîtrisés (GPO)", "Mots de passe, mises à jour, restrictions et paramètres appliqués automatiquement"],
    ["FaNetworkWired", "Une administration à distance", "L'administrateur gère tout depuis n'importe où, en sécurité, via Tailscale"]];
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + (i % 2) * 6.15, y = 1.6 + Math.floor(i / 2) * 2.6;
    card(s, x, y, 5.95, 2.3);
    await badge(s, x + 0.35, y + 0.4, 0.8, prop[i][0]);
    T(s, prop[i][1], { x: x + 1.4, y: y + 0.4, w: 4.3, h: 0.6, fontSize: 20, bold: true, color: C.text2 });
    T(s, prop[i][2], { x: x + 1.4, y: y + 1.05, w: 4.3, h: 1.0, fontSize: 15, valign: "top" });
  }

  // 3. Architecture
  pres.addSection({ title: "Architecture" });
  s = S("Contenu", "Architecture");
  s.addText("Architecture : simple et évolutive", { placeholder: "title" });
  card(s, 0.6, 1.5, 8.2, 4.4, C.text2);
  await badge(s, 0.9, 1.75, 0.6, "FaServer", C.accent1);
  T(s, "Serveur physique — Proxmox VE", { x: 1.7, y: 1.75, w: 6.8, h: 0.6, fontSize: 20, bold: true, color: C.background1, valign: "middle" });
  card(s, 0.9, 2.6, 7.6, 3.0, C.background2);
  await badge(s, 1.15, 2.8, 0.5, "FaWindows", C.accent3);
  T(s, "Machine virtuelle · Windows Server 2025", { x: 1.85, y: 2.8, w: 6.4, h: 0.5, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
  const roles = [["Active Directory", "Comptes et groupes"], ["DNS", "Résolution des noms"], ["DHCP", "Adresses IP des postes"], ["GPO", "Règles sur les postes"]];
  roles.forEach((t, i) => {
    const x = 1.15 + i * 1.82;
    card(s, x, 3.55, 1.7, 1.8, C.background1);
    T(s, t[0], { x: x + 0.1, y: 3.7, w: 1.5, h: 0.6, fontSize: 15, bold: true, color: C.accent1, align: "center", valign: "middle" });
    T(s, t[1], { x: x + 0.1, y: 4.35, w: 1.5, h: 0.8, fontSize: 13, align: "center", valign: "top" });
  });
  card(s, 9.2, 1.5, 3.5, 4.4, C.accent2);
  await badge(s, 9.5, 1.8, 0.8, "FaUserShield", C.background1, HEX.accent2);
  T(s, [{ text: "Administrateur à distance", options: { bold: true, fontSize: 18, breakLine: true } }, { text: "Tunnel chiffré Tailscale vers le serveur : aucun port ouvert sur Internet", options: { fontSize: 15 } }], { x: 9.5, y: 2.85, w: 2.9, h: 2.8, color: C.background1, valign: "top", paraSpaceAfter: 10 });
  await badge(s, 0.6, 6.15, 0.5, "FaDesktop", C.accent4, HEX.dk2);
  T(s, "Postes des employés joints au domaine (10 à 25)", { x: 1.3, y: 6.15, w: 11, h: 0.5, fontSize: 16, italic: true, color: C.accent6, valign: "middle" });

  // 4. Matériel & licence
  pres.addSection({ title: "Chiffrage" });
  const SERVEUR = 2200000, WIN = 700000;
  const svc = [["Installation Proxmox VE, RAID, réseau", 150000], ["Windows Server : Active Directory, DNS, DHCP", 250000], ["Stratégies de groupe (GPO) + intégration des postes au domaine", 200000], ["Accès distant administrateur via Tailscale", 50000], ["Documentation + prise en main administrateur", 100000]];
  const SVC = svc.reduce((p, c) => p + c[1], 0), TOT = SERVEUR + WIN + SVC;
  s = S("Contenu", "Chiffrage");
  s.addText("Matériel et licence", { placeholder: "title" });
  const ml = [["FaServer", "Serveur tour", "Dell T160 / HPE ML30\nXeon · 32 Go RAM · 2×960 Go SSD en RAID1", SERVEUR, C.accent1],
    ["FaWindows", "Windows Server 2025", "Édition Standard 16 cœurs\nLicence perpétuelle, 1 serveur", WIN, C.accent3]];
  for (let i = 0; i < 2; i++) {
    const x = 0.6 + i * 4.15;
    card(s, x, 1.5, 3.95, 4.6, C.background2);
    await badge(s, x + 0.35, 1.8, 0.9, ml[i][0], ml[i][4]);
    T(s, ml[i][1], { x: x + 0.35, y: 2.9, w: 3.3, h: 0.5, fontSize: 20, bold: true, color: C.text2 });
    T(s, ml[i][2], { x: x + 0.35, y: 3.45, w: 3.3, h: 1.2, fontSize: 14, valign: "top" });
    T(s, f(ml[i][3]), { x: x + 0.35, y: 5.2, w: 3.3, h: 0.6, fontSize: 24, bold: true, color: ml[i][4] });
  }
  card(s, 8.9, 1.5, 3.8, 4.6, C.text2);
  T(s, "Gratuits", { x: 9.2, y: 1.8, w: 3.2, h: 0.6, fontSize: 22, bold: true, color: C.background1 });
  const free = [["FaCubes", "Proxmox VE", "sans abonnement"], ["FaLink", "Tailscale", "offre gratuite"]];
  for (let i = 0; i < 2; i++) {
    const y = 2.6 + i * 1.2;
    await badge(s, 9.2, y, 0.7, free[i][0], C.accent4, HEX.dk2);
    T(s, [{ text: free[i][1], options: { bold: true, fontSize: 16, breakLine: true } }, { text: free[i][2], options: { fontSize: 14 } }], { x: 10.1, y, w: 2.4, h: 0.8, color: C.background1, valign: "middle" });
  }
  T(s, "0 FCFA", { x: 9.2, y: 5.2, w: 3.2, h: 0.6, fontSize: 24, bold: true, color: C.accent4 });
  T(s, "Onduleur recommandé (≈ 300 000 FCFA, non inclus). Sauvegarde et stockage : hors périmètre, à définir.", { x: 0.6, y: 6.3, w: 12.1, h: 0.4, fontSize: 12, italic: true, color: C.accent6 });

  // 5. Installation
  s = S("Contenu", "Chiffrage");
  s.addText("Installation et mise en service", { placeholder: "title" });
  for (let i = 0; i < svc.length; i++) {
    const y = 1.5 + i * 0.85;
    card(s, 0.6, y, 8.4, 0.7, C.background2);
    s.addShape(pres.shapes.OVAL, { x: 0.75, y: y + 0.12, w: 0.46, h: 0.46, fill: { color: C.accent1 }, line: { color: C.accent1 } });
    T(s, String(i + 1), { x: 0.75, y: y + 0.12, w: 0.46, h: 0.46, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
    T(s, svc[i][0], { x: 1.4, y, w: 5.6, h: 0.7, fontSize: 15, valign: "middle" });
    T(s, f(svc[i][1]), { x: 6.9, y, w: 1.9, h: 0.7, fontSize: 15, bold: true, color: C.text2, align: "right", valign: "middle" });
  }
  card(s, 9.4, 1.5, 3.3, 4.1, C.accent1);
  await badge(s, 9.7, 1.8, 0.8, "FaTools", C.background1, HEX.accent1);
  T(s, "Forfait installation", { x: 9.7, y: 2.9, w: 2.8, h: 0.5, fontSize: 16, color: C.background1 });
  T(s, f(SVC), { x: 9.7, y: 3.4, w: 2.8, h: 0.7, fontSize: 26, bold: true, color: C.background1 });
  T(s, "Mise en service en 1 à 2 semaines", { x: 9.7, y: 4.3, w: 2.8, h: 0.9, fontSize: 14, color: C.background1, valign: "top" });

  // 6. Récapitulatif
  s = S("Contenu", "Chiffrage");
  s.addText("Récapitulatif de l'investissement", { placeholder: "title" });
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Répartition", labels: ["Serveur", "Licence Windows", "Installation"], values: [SERVEUR, WIN, SVC] }], { x: 0.6, y: 1.4, w: 5.4, h: 5.3, holeSize: 55, chartColors: [HEX.accent1, HEX.accent3, HEX.accent2], showLegend: true, legendPos: "b", legendFontFace: "+mn-lt", legendFontSize: 13, showPercent: true, showValue: false, dataLabelColor: "FFFFFF", dataLabelFontSize: 13, dataLabelFontFace: "+mn-lt" });
  [["Serveur", SERVEUR], ["Licence Windows Server", WIN], ["Installation", SVC]].forEach((t, i) => {
    T(s, t[0], { x: 6.8, y: 1.6 + i * 0.75, w: 3.2, h: 0.6, fontSize: 18, valign: "middle" });
    T(s, f(t[1]), { x: 9.6, y: 1.6 + i * 0.75, w: 3.1, h: 0.6, fontSize: 18, align: "right", valign: "middle" });
  });
  card(s, 6.8, 4.1, 5.9, 2.2, C.text2, true);
  T(s, "Total HT", { x: 7.1, y: 4.3, w: 5.3, h: 0.5, fontSize: 16, color: C.accent4 });
  T(s, f(TOT), { x: 7.1, y: 4.8, w: 5.3, h: 0.9, fontSize: 40, bold: true, color: C.background1 });
  T(s, "TVA 18 % en sus : " + f(Math.round(TOT * 1.18)) + " TTC", { x: 7.1, y: 5.7, w: 5.3, h: 0.4, fontSize: 13, color: C.accent4 });

  // 7. Récurrents
  s = S("Contenu", "Chiffrage");
  s.addText("Après l'installation", { placeholder: "title" });
  const rec = [["FaCheckCircle", "0", "FCFA / an", "Aucune licence annuelle : Proxmox VE et Tailscale sont gratuits, Windows Server est perpétuel"],
    ["FaHeadset", "75 000", "FCFA / mois · option", "Maintenance : supervision, mises à jour, assistance à distance"],
    ["FaGlobe", "≈ 30 000", "FCFA / an", "Nom de domaine et hébergement du site vitrine"]];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.6, 3.9, 4.6, C.background2);
    await badge(s, x + 0.3, 1.9, 0.8, rec[i][0], i == 0 ? C.accent1 : i == 1 ? C.accent3 : C.accent2);
    T(s, rec[i][1], { x: x + 0.3, y: 2.9, w: 3.3, h: 1.0, fontSize: 44, bold: true, color: C.text2 });
    T(s, rec[i][2], { x: x + 0.3, y: 3.9, w: 3.3, h: 0.5, fontSize: 15, bold: true, color: C.accent2 });
    T(s, rec[i][3], { x: x + 0.3, y: 4.5, w: 3.3, h: 1.5, fontSize: 15, valign: "top" });
  }

  // 8. Site offert
  pres.addSection({ title: "Bonus" });
  s = S("Contenu", "Bonus");
  s.addText("En cadeau : votre site web vitrine", { placeholder: "title" });
  // maquette navigateur
  card(s, 0.6, 1.5, 6.0, 4.9, C.accent5, true);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.75, y: 1.95, w: 5.7, h: 4.3, fill: { color: C.background1 }, line: { color: C.background1 } });
  [C.accent2, C.accent4, C.accent1].forEach((c, i) => s.addShape(pres.shapes.OVAL, { x: 0.85 + i * 0.3, y: 1.65, w: 0.18, h: 0.18, fill: { color: c }, line: { color: c } }));
  s.addShape(pres.shapes.RECTANGLE, { x: 0.75, y: 1.95, w: 5.7, h: 1.9, fill: { color: C.text2 }, line: { color: C.text2 } });
  T(s, "Votre Entreprise", { x: 1.05, y: 2.2, w: 5, h: 0.6, fontSize: 24, bold: true, color: C.background1, fontFace: "Cambria" });
  T(s, "Votre slogan ici", { x: 1.05, y: 2.8, w: 5, h: 0.4, fontSize: 14, color: C.accent4 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1.05, y: 3.3, w: 1.6, h: 0.35, rectRadius: 0.08, fill: { color: C.accent2 }, line: { color: C.accent2 } });
  T(s, "Nous contacter", { x: 1.05, y: 3.3, w: 1.6, h: 0.35, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle" });
  for (let i = 0; i < 3; i++) {
    const x = 1.05 + i * 1.75;
    card(s, x, 4.1, 1.55, 1.9, C.background2);
    await badge(s, x + 0.5, 4.3, 0.55, ["FaBriefcase", "FaStar", "FaEnvelope"][i]);
    T(s, ["Services", "Réalisations", "Contact"][i], { x, y: 5.0, w: 1.55, h: 0.5, fontSize: 13, bold: true, color: C.text2, align: "center" });
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.1, y: 1.5, w: 2.4, h: 0.6, rectRadius: 0.1, fill: { color: C.accent2 }, line: { color: C.accent2 } });
  T(s, "OFFERT", { x: 7.1, y: 1.5, w: 2.4, h: 0.6, fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle" });
  T(s, "Valeur 500 000 FCFA", { x: 9.7, y: 1.5, w: 3.0, h: 0.6, fontSize: 16, italic: true, color: C.accent6, valign: "middle" });
  const feats = [["FaMobileAlt", "Adapté mobile, tablette et ordinateur"], ["FaFileAlt", "5 pages : Accueil, À propos, Services, Réalisations, Contact"], ["FaWhatsapp", "Formulaire de contact, Google Maps, WhatsApp"], ["FaSearch", "Référencement de base et HTTPS"], ["FaAt", "E-mails professionnels à votre nom"]];
  for (let i = 0; i < feats.length; i++) {
    const y = 2.45 + i * 0.8;
    await badge(s, 7.1, y, 0.55, feats[i][0], C.background2, HEX.accent1);
    T(s, feats[i][1], { x: 7.85, y, w: 4.85, h: 0.55, fontSize: 15, valign: "middle" });
  }

  // 9. Prochaines étapes
  pres.addSection({ title: "Conclusion" });
  s = S("Sombre", "Conclusion");
  s.addText("Prochaines étapes", { placeholder: "title" });
  s.addText("Votre infrastructure opérationnelle en 1 à 2 semaines", { placeholder: "body" });
  const steps = [["FaClipboardCheck", "Valider le nombre de postes et les GPO souhaitées"], ["FaFileSignature", "Signature du devis et commande du serveur"], ["FaRocket", "Installation, mise en service et remise du site vitrine"]];
  for (let i = 0; i < 3; i++) {
    const y = 1.2 + i * 1.75;
    card(s, 8.6, y, 4.1, 1.5, C.background1);
    await badge(s, 8.85, y + 0.35, 0.8, steps[i][0], C.accent2);
    T(s, steps[i][1], { x: 9.85, y: y + 0.15, w: 2.7, h: 1.2, fontSize: 15, bold: true, color: C.text2, valign: "middle" });
  }

  await pres.writeFile({ fileName: "Proposition_Serveur_Proxmox_WindowsServer.pptx" });
  await applyTheme("Proposition_Serveur_Proxmox_WindowsServer.pptx", THEME);
  console.log(TOT);
})();
