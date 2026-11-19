const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require(path.join(__dirname, "../.claude/skills/pptx/scripts/apply_theme.js"));

const THEME = {
  name: "Olaplex",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1C1C1C", lt1: "FFFFFF", dk2: "3A3A3A", lt2: "F3EEE6",
    accent1: "B8946A", accent2: "1C1C1C", accent3: "E6D9C4", accent4: "6B6B6B",
    accent5: "8C6A44", accent6: "D9C7A7", hlink: "8C6A44", folHlink: "6B6B6B",
  },
};

const IMG = path.join(__dirname, "olaplex.png");
const AUTHOR = "Andrea Habáňová";

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.author = AUTHOR;
  pres.title = "Produktový list – Olaplex No.4 Bond Maintenance šampon";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "Titulni",
    background: { color: C.text1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.55, w: 5.4, h: 1.5,
        fontSize: 36, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "Obsah",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.35, w: 9, h: 0.7,
        fontSize: 30, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { text: { text: `Produktový list · Olaplex No.4 · ${AUTHOR}`, options: {
        x: 0.5, y: 5.2, w: 7, h: 0.3, fontSize: 10, color: C.accent4, margin: 0 } } },
    ],
    slideNumber: { x: 9.0, y: 5.2, w: 0.5, h: 0.3, fontSize: 10, color: "6B6B6B", align: "right" },
  });

  // 1 — Titulní snímek
  pres.addSection({ title: "Úvod" });
  let s = pres.addSlide({ masterName: "Titulni", sectionTitle: "Úvod" });
  s.addText("PRODUKTOVÝ LIST", { x: 0.6, y: 1.0, w: 5, h: 0.4, fontSize: 14, bold: true,
    color: C.accent1, charSpacing: 4, margin: 0, isTextBox: true });
  s.addText("Olaplex No.4 Bond Maintenance šampon", { placeholder: "title" });
  s.addText("250 ml · šampon pro obnovu a posílení vlasů", { x: 0.6, y: 3.15, w: 5.4, h: 0.4,
    fontSize: 16, color: C.accent6, margin: 0, isTextBox: true });
  s.addText([
    { text: "Vypracovala: ", options: { color: C.accent6 } },
    { text: AUTHOR, options: { bold: true, color: C.background1 } },
  ], { x: 0.6, y: 4.5, w: 5.4, h: 0.4, fontSize: 16, margin: 0, isTextBox: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.55, y: 0.5, w: 2.95, h: 4.625,
    fill: { color: C.background1 }, rectRadius: 0.15, objectName: "Ramecek obrazku" });
  s.addImage({ path: IMG, x: 6.92, y: 0.75, w: 2.21, h: 3.8, altText: "Šampon Olaplex No.4 Bond Maintenance 250 ml" });

  // 2 — Výrobce, typ výrobku + specifikace
  pres.addSection({ title: "Produkt" });
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Produkt" });
  s.addText("Výrobce a typ výrobku", { placeholder: "title" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.3, w: 2.6, h: 3.7,
    fill: { color: C.background2 }, rectRadius: 0.15, objectName: "Ramecek obrazku" });
  s.addImage({ path: IMG, x: 0.74, y: 1.45, w: 2.12, h: 3.4, altText: "Šampon Olaplex No.4" });

  const label = { fontSize: 13, bold: true, color: C.accent5, margin: 0, charSpacing: 1 };
  s.addText("VÝROBCE", { x: 3.5, y: 1.3, w: 2, h: 0.3, ...label, isTextBox: true });
  s.addText("OLAPLEX, Inc. (USA)", { x: 3.5, y: 1.6, w: 6, h: 0.4, fontSize: 18, bold: true,
    color: C.text1, margin: 0, isTextBox: true });
  s.addText("TYP VÝROBKU", { x: 3.5, y: 2.2, w: 2, h: 0.3, ...label, isTextBox: true });
  s.addText("Šampon pro péči o vlasy – 250 ml", { x: 3.5, y: 2.5, w: 6, h: 0.4, fontSize: 18,
    bold: true, color: C.text1, margin: 0, isTextBox: true });
  const spec = [
    "pro všechny typy vlasů",
    "obnovuje, posiluje a hydratuje vlasy",
    "vhodný pro poškozené a barvené vlasy",
    "bez SLS a SLES, šetrný pro každodenní mytí",
    "patentovaná technologie obnovy vazeb ve vlasu",
  ];
  s.addText(spec.map((t, i) => ({ text: t, options: { bullet: { code: "2713" }, breakLine: i < spec.length - 1 } })),
    { x: 3.5, y: 3.05, w: 6, h: 1.9, fontSize: 15, color: C.text2, paraSpaceAfter: 6,
      valign: "top", margin: 0, isTextBox: true });

  // 3 — Použití + popis
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Produkt" });
  s.addText("Použití a popis produktu", { placeholder: "title" });
  const steps = [
    ["1", "Naneste", "dostatečné množství šamponu na mokré vlasy."],
    ["2", "Napěňte", "a jemně vmasírujte do vlasů i pokožky hlavy."],
    ["3", "Opláchněte", "důkladně vodou, v případě potřeby opakujte."],
  ];
  steps.forEach(([n, head, txt], i) => {
    const x = 0.5 + i * 3.1;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.3, w: 2.8, h: 1.75,
      fill: { color: C.background2 }, rectRadius: 0.12, objectName: `Krok ${n}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: 1.5, w: 0.5, h: 0.5, fill: { color: C.text1 },
      objectName: `Cislo ${n}` });
    s.addText(n, { x: x + 0.25, y: 1.5, w: 0.5, h: 0.5, fontSize: 16, bold: true, align: "center",
      valign: "middle", color: C.background1, margin: 0, isTextBox: true });
    s.addText(head, { x: x + 0.9, y: 1.5, w: 1.75, h: 0.5, fontSize: 17, bold: true, valign: "middle",
      color: C.text1, margin: 0, isTextBox: true });
    s.addText(txt, { x: x + 0.25, y: 2.15, w: 2.35, h: 0.8, fontSize: 14, color: C.text2,
      valign: "top", margin: 0, isTextBox: true });
  });
  s.addText("POPIS PRODUKTU", { x: 0.5, y: 3.3, w: 4, h: 0.3, ...label, isTextBox: true });
  s.addText("Vysoce hydratační šampon, který obnovuje, posiluje a vyživuje vlasy. Zanechává je " +
    "lesklé, hebké a snadno upravitelné, chrání je před lámáním a roztřepenými konečky. Je šetrný, " +
    "a proto ho lze používat každý den. Nejlepších výsledků dosáhnete v kombinaci s kondicionérem " +
    "Olaplex No.5. Pokud se šampon dostane do očí, důkladně je vypláchněte čistou vodou.",
    { x: 0.5, y: 3.65, w: 9, h: 1.4, fontSize: 14, color: C.text2, valign: "top", margin: 0,
      isTextBox: true });

  // 4 — Složení
  pres.addSection({ title: "Složení" });
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Složení" });
  s.addText("Složení výrobku", { placeholder: "title" });
  const inci = "Water (Aqua), Sodium Lauroyl Methyl Isethionate, Cocamidopropyl Hydroxysultaine, " +
    "Potassium Cocoyl Glycinate, Disodium Cocoyl Glutamate, Sodium Lauroyl Sarcosinate, Potassium Cocoate, " +
    "Decyl Glucoside, Glycereth-26, Bis-Aminopropyl Diglycol Dimaleate, Propanediol, Euterpe Oleracea Fruit " +
    "Extract, Punica Granatum Extract, Rosmarinus Officinalis (Rosemary) Leaf Extract, Musa Sapientum (Banana) " +
    "Fruit Extract, Origanum Vulgare Leaf Extract, Morinda Citrifolia Fruit Extract, Arctium Lappa Root Extract, " +
    "Fragrance (Parfum), Cocamidopropylamine Oxide, Disodium Laureth Sulfosuccinate, Panthenol, Glycol " +
    "Distearate, Sodium Cocoyl Glutamate, Methyl Gluceth-20, Acrylates Copolymer, Prunus Armeniaca (Apricot) " +
    "Kernel Oil, Pseudozyma Epicola/Argania Spinosa Kernel Oil Ferment Filtrate, Pseudozyma Epicola/Camellia " +
    "Sinensis Seed Oil Ferment Extract Filtrate, PEG-120 Methyl Glucose Dioleate, Amodimethicone, Sodium Lauryl " +
    "Sulfoacetate, Guar Hydroxypropyltrimonium Chloride, Polyquaternium-10, Divinyldimethicone/Dimethicone " +
    "Copolymer, Polyquaternium-11, Trisodium Ethylenediamine Disuccinate, C11-15 Pareth-7, Laureth-9, Glycerin, " +
    "Trideceth-12, Hexyl Cinnamal, Limonene, C12-13 Pareth-23, C12-13 Pareth-3, Hydrolyzed Vegetable Protein " +
    "PG-Propyl Silanetriol, Acetic Acid, Citral, Benzoic Acid, Tetrasodium EDTA, Pentasodium Triphosphate, " +
    "Sodium Polyacrylate, Pentaerythrityl Tetra-Di-T-Butyl Hydroxyhydrocinnamate, Disodium EDTA, Sodium " +
    "Benzoate, Potassium Sorbate, Phenoxyethanol, Chlorphenesin, Ethylhexylglycerin, Quaternium-95, PEG-8, " +
    "PEG-8/SMDI Copolymer, Palmitoyl Myristyl Serinate, Citric Acid, Sodium Hyaluronate, Tocopherol, Biotin, " +
    "Helianthus Annuus (Sunflower) Seed Extract, Helianthus Annuus (Sunflower) Seed Oil.";
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.25, w: 9, h: 3.8,
    fill: { color: C.background2 }, rectRadius: 0.12, objectName: "Pozadi slozeni" });
  s.addText(inci, { x: 0.75, y: 1.4, w: 8.5, h: 3.5, fontSize: 11.5, color: C.text2,
    valign: "middle", margin: 0, isTextBox: true, lineSpacingMultiple: 1.05 });

  // 5 — Hlavní účinné látky
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Složení" });
  s.addText("Hlavní účinné látky a jejich význam", { placeholder: "title" });
  const actives = [
    ["Bis-Aminopropyl Diglycol Dimaleate", "patentovaná technologie Olaplex – obnovuje porušené vazby ve vlasu a posiluje ho"],
    ["Sodium Lauroyl Methyl Isethionate, Cocamidopropyl Hydroxysultaine", "jemné tenzidy z kokosu – šetrně myjí vlasy i pokožku"],
    ["Panthenol, Biotin, Sodium Hyaluronate", "hydratace, pružnost a lesk vlasů"],
    ["Apricot Kernel Oil, Argania Spinosa Ferment", "rostlinné oleje – vyživují a zjemňují vlasy"],
    ["Polyquaternium-10, Guar Hydroxypropyltrimonium Chloride", "kondicionační látky – snadné rozčesávání, antistatický účinek"],
    ["Sodium Benzoate, Phenoxyethanol, Hexyl Cinnamal, Limonene", "konzervační prostředky a vonné přísady"],
  ];
  actives.forEach(([name, desc], i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.05, y = 1.25 + row * 1.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.85, h: 1.75,
      fill: { color: C.background2 }, rectRadius: 0.12, objectName: `Latka ${i + 1}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.2, w: 0.4, h: 0.4, fill: { color: C.text1 },
      objectName: `Cislo latky ${i + 1}` });
    s.addText(String(i + 1), { x: x + 0.2, y: y + 0.2, w: 0.4, h: 0.4, fontSize: 13, bold: true,
      align: "center", valign: "middle", color: C.background1, margin: 0, isTextBox: true });
    s.addText(name, { x: x + 0.75, y: y + 0.12, w: 1.95, h: 0.6, fontSize: 11, bold: true,
      color: C.accent5, valign: "middle", margin: 0, isTextBox: true });
    s.addText(desc, { x: x + 0.2, y: y + 0.8, w: 2.5, h: 0.85, fontSize: 12, color: C.text1,
      valign: "top", margin: 0, isTextBox: true });
  });

  // 6 — Zdroj informací
  pres.addSection({ title: "Zdroje" });
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Zdroje" });
  s.addText("Zdroj informací", { placeholder: "title" });
  const src = [
    ["Heureka.cz – Olaplex No.4 Bond Maintenance šampon 250 ml", "https://sampony.heureka.cz"],
    ["OLAPLEX – oficiální stránky výrobce", "https://olaplex.com"],
    ["OLAPLEX – seznam složení (INCI)", "https://cdn.shopify.com/s/files/1/1081/3948/files/OLAPLEX_inci.pdf"],
    ["Obal výrobku – Olaplex No.4 Bond Maintenance Shampoo", null],
  ];
  src.forEach(([t, url], i) => {
    const y = 1.35 + i * 0.9;
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: y + 0.05, w: 0.45, h: 0.45, fill: { color: C.accent1 },
      objectName: `Zdroj ${i + 1}` });
    s.addText(String(i + 1), { x: 0.5, y: y + 0.05, w: 0.45, h: 0.45, fontSize: 14, bold: true,
      align: "center", valign: "middle", color: C.background1, margin: 0, isTextBox: true });
    const runs = [{ text: t, options: { bold: true, color: C.text1, breakLine: !!url } }];
    if (url) runs.push({ text: url, options: { color: C.accent5, fontSize: 12, hyperlink: { url } } });
    s.addText(runs, { x: 1.2, y, w: 8.3, h: 0.6, fontSize: 15, valign: "middle", margin: 0, isTextBox: true });
  });

  const out = path.join(__dirname, "Produktovy_list_Olaplex_No4.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("ok", out);
})();
