const pptxgen = require('pptxgenjs');
const { applyTheme } = require('/root/.claude/skills/synced/2d22e13e-03a7-4f53-876d-c8368813d997_77b68b09-cce3-4263-bfaa-f3c76c1f97e3/pptx/scripts/apply_theme.js');
const THEME={name:"Infra PME",headFontFace:"Cambria",bodyFontFace:"Calibri",colors:{dk1:"1B1F24",lt1:"FFFFFF",dk2:"0F3D3E",lt2:"EEF4F3",accent1:"0F7C6E",accent2:"E57A2E",accent3:"3A6B8C",accent4:"8FB8AE",accent5:"C9D6D3",accent6:"5C6670",hlink:"0F7C6E",folHlink:"3A6B8C"}};
const pres=new pptxgen(); pres.layout="LAYOUT_WIDE"; pres.theme={headFontFace:"Cambria",bodyFontFace:"Calibri"};
pres.title="Proposition infrastructure serveur PME";
const C=pres.SchemeColor;
const f=n=>n.toLocaleString('fr-FR').replace(/ | /g,' ')+' FCFA';
pres.defineSlideMaster({title:"Sombre",background:{color:C.text2},objects:[{placeholder:{options:{name:"title",type:"title",x:0.8,y:2.3,w:11.7,h:1.6,fontSize:44,bold:true,color:C.background1,fontFace:"Cambria"},text:""}},{placeholder:{options:{name:"body",type:"body",x:0.8,y:4.0,w:11.7,h:1.4,fontSize:20,color:C.accent4},text:""}}]});
pres.defineSlideMaster({title:"Contenu",background:{color:C.background1},slideNumber:{x:12.4,y:7.0,fontSize:10,color:C.accent6},objects:[{placeholder:{options:{name:"title",type:"title",x:0.6,y:0.4,w:12.1,h:0.9,fontSize:32,bold:true,color:C.text2,fontFace:"Cambria"},text:""}},{text:{text:"Prix indicatifs HT en FCFA – octobre 2026",options:{x:0.6,y:7.0,w:6,h:0.3,fontSize:10,color:C.accent6}}}]});
const S=(m,sec)=>pres.addSlide({masterName:m,sectionTitle:sec});
const card=(s,x,y,w,h,fill)=>s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y,w,h,rectRadius:0.12,fill:{color:fill||C.background2},line:{color:fill||C.background2}});
const num=(s,x,y,n)=>{s.addShape(pres.shapes.OVAL,{x,y,w:0.6,h:0.6,fill:{color:C.accent1},line:{color:C.accent1}});s.addText(String(n),{x,y,w:0.6,h:0.6,align:"center",valign:"middle",bold:true,fontSize:18,color:C.background1,isTextBox:true,margin:0});};

pres.addSection({title:"Introduction"});
let s=S("Sombre","Introduction");
s.addText("Votre infrastructure informatique clé en main",{placeholder:"title"});
s.addText("Serveur Proxmox + Windows Server pour une petite entreprise de 10 à 25 utilisateurs — et votre site web vitrine offert",{placeholder:"body"});

s=S("Contenu","Introduction");
s.addText("Ce que nous vous proposons",{placeholder:"title"});
[["Un serveur unique, virtualisé","Proxmox VE héberge tous vos services sur une seule machine robuste"],["Une administration centralisée","Comptes, droits, postes et imprimantes gérés depuis Windows Server"],["Vos données protégées","Sauvegardes automatiques quotidiennes, onduleur, RAID"],["Votre vitrine sur le web","Un site professionnel offert pour présenter l'entreprise"]].forEach((t,i)=>{const x=0.6+(i%2)*6.15,y=1.6+Math.floor(i/2)*2.6;card(s,x,y,5.95,2.3);num(s,x+0.35,y+0.4,i+1);s.addText(t[0],{x:x+1.2,y:y+0.35,w:4.5,h:0.7,fontSize:20,bold:true,color:C.text2,isTextBox:true,margin:0});s.addText(t[1],{x:x+1.2,y:y+1.05,w:4.5,h:1.0,fontSize:15,color:C.text1,isTextBox:true,margin:0,valign:"top"});});

pres.addSection({title:"Architecture"});
s=S("Contenu","Architecture");
s.addText("Architecture : un hôte, quatre services",{placeholder:"title"});
card(s,0.6,1.5,12.1,0.8,C.text2);
s.addText("Hôte physique — Proxmox VE (hyperviseur)",{x:0.6,y:1.5,w:12.1,h:0.8,align:"center",valign:"middle",fontSize:20,bold:true,color:C.background1,isTextBox:true});
[["VM 1 · Windows Server","Contrôleur de domaine\nActive Directory, DNS, DHCP, GPO"],["VM 2 · Windows Server","Serveur de fichiers\nPartages, droits, quotas, clichés"],["VM 3 · Windows Server","Bureau à distance (RDS)\nApplications métier, télétravail"],["Proxmox Backup Server","Sauvegarde\nCopies quotidiennes vers NAS"]].forEach((t,i)=>{const x=0.6+i*3.08;card(s,x,2.6,2.9,2.6);s.addText(t[0],{x:x+0.2,y:2.8,w:2.5,h:0.6,fontSize:14,bold:true,color:C.accent1,isTextBox:true,margin:0});s.addText(t[1],{x:x+0.2,y:3.45,w:2.5,h:1.6,fontSize:15,color:C.text1,isTextBox:true,margin:0,valign:"top"});});
s.addText("Postes clients (10–25) · NAS de sauvegarde · Onduleur · Switch réseau",{x:0.6,y:5.5,w:12.1,h:0.6,align:"center",fontSize:16,italic:true,color:C.accent6,isTextBox:true});

pres.addSection({title:"Chiffrage"});
const mat=[["Serveur tour (Dell T360 / HPE ML110) – Xeon 8 cœurs, 64 Go RAM, 2×960 Go SSD RAID1 + 2×4 To",3500000],["NAS de sauvegarde 2 baies + 2×4 To",650000],["Onduleur 1500 VA",450000],["Switch 24 ports gigabit + câblage / petit rack",350000]];
const lic=[["Proxmox VE – souscription Community (1 an, 1 socket)",75000],["Windows Server 2025 Standard 16 cœurs ×2 (couvre 4 VM)",1400000],["CAL utilisateur Windows Server ×20 (≈30 000 l'unité)",600000],["CAL RDS utilisateur ×10 (≈85 000 l'unité)",850000]];
const svc=[["Installation Proxmox, RAID, réseau",250000],["Déploiement AD/DNS/DHCP, GPO, intégration postes",400000],["Serveur de fichiers + RDS",300000],["Sauvegarde PBS + NAS, tests de restauration",150000],["Documentation + formation administrateur (1 jour)",100000]];
const sum=a=>a.reduce((p,c)=>p+c[1],0);
function tab(title,rows,tot){s=S("Contenu","Chiffrage");s.addText(title,{placeholder:"title"});
 const r=rows.map(x=>[{text:x[0],options:{fontSize:15,color:C.text1}},{text:f(x[1]),options:{fontSize:15,align:"right",color:C.text1}}]);
 r.push([{text:"Sous-total",options:{bold:true,fontSize:16,color:C.background1,fill:{color:C.accent1}}},{text:f(tot),options:{bold:true,fontSize:16,align:"right",color:C.background1,fill:{color:C.accent1}}}]);
 s.addTable([[{text:"Poste",options:{bold:true,fontSize:15,color:C.text2,fill:{color:C.background2}}},{text:"Prix HT",options:{bold:true,fontSize:15,align:"right",color:C.text2,fill:{color:C.background2}}}],...r],{x:0.6,y:1.5,w:12.1,colW:[9.1,3.0],rowH:0.6,border:{type:"solid",pt:0.5,color:"C9D6D3"},valign:"middle"});return s;}
tab("Matériel",mat,sum(mat));
s=tab("Licences logicielles",lic,sum(lic));
s.addText("Windows Server Standard autorise 2 VM par licence 16 cœurs ; les CAL sont obligatoires pour chaque utilisateur accédant au domaine.",{x:0.6,y:5.6,w:12.1,h:0.7,fontSize:13,italic:true,color:C.accent6,isTextBox:true});
tab("Installation et mise en service",svc,sum(svc));

const tot=sum(mat)+sum(lic)+sum(svc);
s=S("Contenu","Chiffrage");
s.addText("Récapitulatif de l'investissement",{placeholder:"title"});
s.addChart(pres.charts.DOUGHNUT,[{name:"Répartition",labels:["Matériel","Licences","Installation"],values:[sum(mat),sum(lic),sum(svc)]}],{x:0.6,y:1.5,w:5.6,h:5.2,holeSize:55,chartColors:["0F7C6E","E57A2E","3A6B8C"],showLegend:true,legendPos:"b",legendFontFace:"+mn-lt",legendFontSize:13,showPercent:true,showValue:false,dataLabelColor:"FFFFFF",dataLabelFontSize:13,dataLabelFontFace:"+mn-lt"});
[["Matériel",sum(mat)],["Licences",sum(lic)],["Installation",sum(svc)]].forEach((t,i)=>{s.addText(t[0],{x:6.8,y:1.7+i*0.75,w:3,h:0.6,fontSize:18,color:C.text1,isTextBox:true,margin:0});s.addText(f(t[1]),{x:9.6,y:1.7+i*0.75,w:3.1,h:0.6,fontSize:18,align:"right",color:C.text1,isTextBox:true,margin:0});});
card(s,6.8,4.2,5.9,2.1,C.text2);
s.addText("Total HT",{x:7.1,y:4.35,w:5.3,h:0.5,fontSize:16,color:C.accent4,isTextBox:true,margin:0});
s.addText(f(tot),{x:7.1,y:4.85,w:5.3,h:0.9,fontSize:40,bold:true,color:C.background1,isTextBox:true,margin:0});
s.addText("TVA 18 % en sus : "+f(Math.round(tot*1.18))+" TTC",{x:7.1,y:5.75,w:5.3,h:0.4,fontSize:13,color:C.accent4,isTextBox:true,margin:0});

s=S("Contenu","Chiffrage");
s.addText("Coûts récurrents (optionnels)",{placeholder:"title"});
[["75 000","FCFA / an","Souscription Proxmox (mises à jour stables)"],["75 000","FCFA / mois","Contrat de maintenance : supervision, mises à jour, assistance"],["≈ 30 000","FCFA / an","Nom de domaine + hébergement du site web"]].forEach((t,i)=>{const x=0.6+i*4.1;card(s,x,1.7,3.9,4.2);s.addText(t[0],{x:x+0.3,y:2.0,w:3.3,h:1.1,fontSize:44,bold:true,color:C.accent1,isTextBox:true,margin:0});s.addText(t[1],{x:x+0.3,y:3.1,w:3.3,h:0.5,fontSize:16,color:C.accent2,bold:true,isTextBox:true,margin:0});s.addText(t[2],{x:x+0.3,y:3.8,w:3.3,h:1.6,fontSize:15,color:C.text1,isTextBox:true,margin:0,valign:"top"});});

pres.addSection({title:"Bonus"});
s=S("Contenu","Bonus");
s.addText("En cadeau : votre site web vitrine",{placeholder:"title"});
card(s,0.6,1.5,4.6,5.0,C.accent2);
s.addText([{text:"OFFERT",options:{fontSize:44,bold:true,breakLine:true}},{text:"Valeur : 500 000 FCFA",options:{fontSize:18}}],{x:0.9,y:2.5,w:4.0,h:2.6,color:C.background1,isTextBox:true,valign:"middle"});
s.addText([{text:"Site responsive (mobile, tablette, ordinateur)",options:{bullet:true,breakLine:true}},{text:"5 pages : Accueil, À propos, Services, Réalisations, Contact",options:{bullet:true,breakLine:true}},{text:"Formulaire de contact, carte Google Maps, WhatsApp",options:{bullet:true,breakLine:true}},{text:"Référencement de base (SEO) et certificat HTTPS",options:{bullet:true,breakLine:true}},{text:"Adresses e-mail professionnelles @votre-entreprise.ci",options:{bullet:true}}],{x:5.6,y:1.7,w:7.1,h:4.6,fontSize:17,color:C.text1,paraSpaceAfter:12,isTextBox:true,valign:"top"});

pres.addSection({title:"Conclusion"});
s=S("Sombre","Conclusion");
s.addText("Prochaines étapes",{placeholder:"title"});
s.addText("1. Visite technique et validation du nombre d'utilisateurs   2. Signature du devis   3. Livraison et installation en 2 à 3 semaines",{placeholder:"body"});
(async()=>{await pres.writeFile({fileName:"Proposition_Serveur_Proxmox_WindowsServer.pptx"});await applyTheme("Proposition_Serveur_Proxmox_WindowsServer.pptx",THEME);console.log(tot);})();
