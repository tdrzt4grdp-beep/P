const pptxgen = require("pptxgenjs");
const path = require("path");
const { applyTheme } = require(path.join(__dirname, "../.claude/skills/pptx/scripts/apply_theme.js"));

const THEME = {
  name: "Olaplex",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1C1C1C", lt1: "FFFFFF", dk2: "3A3A3A", lt2: "F2EFEA",
    accent1: "B8946A", accent2: "1C1C1C", accent3: "E6D9C4", accent4: "6B6B6B",
    accent5: "8C6A44", accent6: "D9C7A7", hlink: "8C6A44", folHlink: "6B6B6B",
  },
};

const F = (n) => path.join(__dirname, "foto", n + ".jpg");
const AUTHOR = "Andrea Habáňová";

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.author = AUTHOR;
  pres.title = "Produktový list – Olaplex No.4 Bond Maintenance šampon";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  const footer = (x, w) => ({ text: { text: `Olaplex No.4 · ${AUTHOR}`, options: {
    x, y: 5.2, w, h: 0.25, fontSize: 10, color: C.accent4, margin: 0 } } });

  // Fotka vpravo přes celou výšku, text vlevo
  pres.defineSlideMaster({
    title: "Fotka vpravo",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.4, w: 3.6, h: 0.9,
        fontSize: 28, bold: true, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "" } },
      footer(0.5, 3.6),
    ],
  });
  // Fotka vlevo, text vpravo
  pres.defineSlideMaster({
    title: "Fotka vlevo",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 5.95, y: 0.4, w: 3.55, h: 0.9,
        fontSize: 28, bold: true, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "" } },
      footer(5.95, 3.55),
    ],
  });
  // Obsahový snímek bez fotky
  pres.defineSlideMaster({
    title: "Obsah",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.35, w: 9, h: 0.7,
        fontSize: 28, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
      footer(0.5, 6),
    ],
  });

  const label = (s, text, x, y, w) => s.addText(text, { x, y, w, h: 0.28, fontSize: 12, bold: true,
    color: C.accent5, charSpacing: 2, margin: 0, isTextBox: true });
  const photoR = (s, n, alt) => s.addImage({ path: F(n), x: 4.375, y: 0, w: 5.625, h: 5.625, altText: alt });
  const photoL = (s, n, alt) => s.addImage({ path: F(n), x: 0, y: 0, w: 5.625, h: 5.625, altText: alt });
  const numCircle = (s, n, x, y, d, fill, size) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, objectName: `Kruh ${n}` });
    s.addText(String(n), { x, y, w: d, h: d, fontSize: size, bold: true, align: "center", valign: "middle",
      color: C.background1, margin: 0, isTextBox: true });
  };

  // 1 — Titulní snímek
  pres.addSection({ title: "Úvod" });
  let s = pres.addSlide({ masterName: "Fotka vpravo", sectionTitle: "Úvod" });
  photoR(s, "modelka", "Žena s dlouhými lesklými vlasy");
  s.addText("PRODUKTOVÝ LIST", { x: 0.5, y: 1.1, w: 3.6, h: 0.3, fontSize: 13, bold: true,
    color: C.accent5, charSpacing: 3, margin: 0, isTextBox: true });
  s.addText("Olaplex No.4", { x: 0.5, y: 1.5, w: 3.7, h: 0.75, fontSize: 36,
    bold: true, color: C.text1, fontFace: "Arial", valign: "top", margin: 0, isTextBox: true });
  s.addText("Bond Maintenance šampon 250 ml", { x: 0.5, y: 2.3, w: 3.6, h: 0.8, fontSize: 18,
    color: C.text2, valign: "top", margin: 0, isTextBox: true });
  s.addText([
    { text: "Vypracovala: ", options: { color: C.accent4 } },
    { text: AUTHOR, options: { bold: true, color: C.text1 } },
  ], { x: 0.5, y: 4.3, w: 3.7, h: 0.35, fontSize: 15, margin: 0, isTextBox: true });

  // 2 — Název, výrobce, typ + specifikace
  pres.addSection({ title: "Výrobek" });
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Výrobek" });
  s.addText("O výrobku", { placeholder: "title" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.25, w: 2.6, h: 3.8,
    fill: { color: "F2EFEA" }, rectRadius: 0.15, objectName: "Pozadi fotky" });
  s.addImage({ path: F("lahev"), x: 0.75, y: 1.4, w: 2.1, h: 3.5, sizing: { type: "contain", w: 2.1, h: 3.5 },
    altText: "Lahev šamponu Olaplex No.4, 250 ml" });
  label(s, "NÁZEV VÝROBKU", 3.5, 1.25, 3);
  s.addText("Olaplex No.4 Bond Maintenance Shampoo, 250 ml", { x: 3.5, y: 1.53, w: 6, h: 0.35,
    fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true });
  label(s, "VÝROBCE", 3.5, 2.05, 3);
  s.addText("OLAPLEX, Inc. – Kalifornie, USA", { x: 3.5, y: 2.33, w: 6, h: 0.35, fontSize: 16,
    bold: true, color: C.text1, margin: 0, isTextBox: true });
  label(s, "TYP VÝROBKU + SPECIFIKACE", 3.5, 2.85, 4);
  const spec = [
    "šampon pro obnovu a posílení vlasů",
    "pro všechny typy vlasů, i barvené a poškozené",
    "hydratuje a dodává vlasům lesk",
    "šetrný – vhodný pro každodenní mytí",
    "bez SLS a SLES",
  ];
  s.addText(spec.map((t, i) => ({ text: t, options: { bullet: { code: "2713" }, breakLine: i < spec.length - 1 } })),
    { x: 3.5, y: 3.18, w: 6, h: 1.85, fontSize: 15, color: C.text2, paraSpaceAfter: 5, valign: "top",
      margin: 0, isTextBox: true });

  // 3 — Použití
  s = pres.addSlide({ masterName: "Fotka vpravo", sectionTitle: "Výrobek" });
  s.addText("Použití", { placeholder: "title" });
  photoR(s, "pena", "Lahev šamponu Olaplex No.4 obalená pěnou");
  const steps = [
    "Naneste dostatečné množství šamponu na mokré vlasy.",
    "Napěňte a jemně vmasírujte do vlasů a pokožky hlavy.",
    "Důkladně opláchněte vodou, případně postup opakujte.",
  ];
  steps.forEach((t, i) => {
    const y = 1.25 + i * 0.95;
    numCircle(s, i + 1, 0.5, y, 0.45, C.text1, 14);
    s.addText(t, { x: 1.15, y: y - 0.05, w: 2.95, h: 0.8, fontSize: 14, color: C.text2, valign: "top",
      margin: 0, isTextBox: true });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.1, w: 3.6, h: 0.85,
    fill: { color: "F2EFEA" }, rectRadius: 0.1, objectName: "Tip" });
  s.addText([{ text: "Tip: ", options: { bold: true, color: C.accent5 } },
    { text: "Když se šampon dostane do očí, vypláchněte je čistou vodou." }],
    { x: 0.65, y: 4.15, w: 3.3, h: 0.75, fontSize: 12, color: C.text2, valign: "middle", margin: 0, isTextBox: true });

  // 4 — Popis a výsledky
  s = pres.addSlide({ masterName: "Fotka vlevo", sectionTitle: "Výrobek" });
  s.addText("Co šampon umí", { placeholder: "title" });
  photoL(s, "pred_po", "Vlasy před a po použití šamponu Olaplex No.4");
  s.addText("Hydratační šampon, který obnovuje a posiluje vlasy. Po umytí jsou vlasy lesklé, hebké " +
    "a snadno se rozčesávají. Chrání je před lámáním a roztřepenými konečky.",
    { x: 5.95, y: 1.3, w: 3.55, h: 1.75, fontSize: 14, color: C.text2, valign: "top", margin: 0, isTextBox: true });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.95, y: 3.25, w: 3.55, h: 1.65,
    fill: { color: "F2EFEA" }, rectRadius: 0.1, objectName: "Vysledek" });
  s.addText("Před / po", { x: 6.15, y: 3.4, w: 3.2, h: 0.35, fontSize: 16, bold: true, color: C.text1,
    margin: 0, isTextBox: true });
  s.addText("Fotka ukazuje vlasy před umytím a po umytí šamponem No.4 a kondicionérem No.5.",
    { x: 6.15, y: 3.8, w: 3.2, h: 0.95, fontSize: 13, color: C.text2, valign: "top", margin: 0, isTextBox: true });

  // 5 — Složení
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
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.2, w: 6.3, h: 3.85,
    fill: { color: "F2EFEA" }, rectRadius: 0.12, objectName: "Pozadi slozeni" });
  s.addText(inci, { x: 0.65, y: 1.3, w: 6.0, h: 3.65, fontSize: 10, color: C.text2, valign: "middle",
    margin: 0, isTextBox: true });
  s.addImage({ path: F("textura"), x: 7.1, y: 1.2, w: 2.4, h: 2.4, rounding: false, altText: "Textura šamponu" });
  s.addText("Textura šamponu – hustý krémový gel, který dobře pění.", { x: 7.1, y: 3.75, w: 2.4, h: 0.8,
    fontSize: 12, italic: true, color: C.accent4, valign: "top", margin: 0, isTextBox: true });

  // 6 — Hlavní účinné látky
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Složení" });
  s.addText("Hlavní účinné látky a jejich význam", { placeholder: "title" });
  const actives = [
    ["Bis-Aminopropyl Diglycol Dimaleate", "patentovaná látka Olaplex – spojuje porušené vazby ve vlasu a posiluje ho"],
    ["Sodium Lauroyl Methyl Isethionate, Cocamidopropyl Hydroxysultaine", "jemné mycí látky (tenzidy) z kokosu – šetrně myjí"],
    ["Panthenol, Biotin, Sodium Hyaluronate", "hydratace, pružnost a lesk vlasů"],
    ["Apricot Kernel Oil, Argania Spinosa Ferment", "rostlinné oleje – vyživují a zjemňují vlasy"],
    ["Polyquaternium-10, Guar Hydroxypropyltrimonium Chloride", "kondicionační látky – snadné rozčesávání, antistatický účinek"],
    ["Sodium Benzoate, Phenoxyethanol, Hexyl Cinnamal, Limonene", "konzervační prostředky a vonné přísady"],
  ];
  actives.forEach(([name, desc], i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.05, y = 1.2 + row * 1.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.85, h: 1.8,
      fill: { color: "F2EFEA" }, rectRadius: 0.12, objectName: `Latka ${i + 1}` });
    numCircle(s, i + 1, x + 0.2, y + 0.2, 0.4, C.text1, 13);
    s.addText(name, { x: x + 0.75, y: y + 0.12, w: 1.95, h: 0.6, fontSize: 11, bold: true,
      color: C.accent5, valign: "middle", margin: 0, isTextBox: true });
    s.addText(desc, { x: x + 0.2, y: y + 0.8, w: 2.5, h: 0.9, fontSize: 12, color: C.text1,
      valign: "top", margin: 0, isTextBox: true });
  });

  // 7 — Doporučená péče
  s = pres.addSlide({ masterName: "Fotka vlevo", sectionTitle: "Složení" });
  s.addText("Doporučená péče", { placeholder: "title" });
  photoL(s, "rutina", "Produkty Olaplex No.3, No.4 a No.5");
  const routine = [
    ["No.3", "kúra před mytím – opravuje a zjemňuje vlasy"],
    ["No.4", "šampon – umyje, posílí a hydratuje"],
    ["No.5", "kondicionér – vyživí a dodá lesk"],
  ];
  routine.forEach(([n, t], i) => {
    const y = 1.35 + i * 1.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.95, y, w: 0.85, h: 0.75, fill: { color: i === 1 ? C.text1 : C.accent1 },
      rectRadius: 0.1, objectName: `Produkt ${n}` });
    s.addText(n, { x: 5.95, y, w: 0.85, h: 0.75, fontSize: 14, bold: true, align: "center", valign: "middle",
      color: C.background1, margin: 0, isTextBox: true });
    s.addText(t, { x: 6.95, y, w: 2.55, h: 0.75, fontSize: 13, color: C.text2, valign: "middle", margin: 0, isTextBox: true });
  });
  s.addText("Šampon No.4 funguje nejlépe společně s kondicionérem No.5.", { x: 5.95, y: 4.5, w: 3.55, h: 0.5,
    fontSize: 12, italic: true, color: C.accent4, valign: "top", margin: 0, isTextBox: true });

  // 8 — Zdroje
  pres.addSection({ title: "Závěr" });
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Závěr" });
  s.addText("Zdroj informací", { placeholder: "title" });
  const src = [
    ["Heureka.cz – Olaplex No.4 Bond Maintenance šampon 250 ml", "https://sampony.heureka.cz"],
    ["OLAPLEX – oficiální stránky výrobce", "https://olaplex.com"],
    ["OLAPLEX – seznam složení (INCI)", "https://cdn.shopify.com/s/files/1/1081/3948/files/OLAPLEX_inci.pdf"],
    ["Lookfantastic – fotografie výrobku", "https://www.lookfantastic.com/p/olaplex-no.4-strengthening-shampoo-250ml/17924637/"],
    ["Obal výrobku", null],
  ];
  src.forEach(([t, url], i) => {
    const y = 1.25 + i * 0.72;
    numCircle(s, i + 1, 0.5, y + 0.05, 0.4, C.accent1, 13);
    const runs = [{ text: t, options: { bold: true, color: C.text1, breakLine: !!url } }];
    if (url) runs.push({ text: url, options: { color: C.accent5, fontSize: 11, hyperlink: { url } } });
    s.addText(runs, { x: 1.1, y, w: 6.1, h: 0.55, fontSize: 14, valign: "middle", margin: 0, isTextBox: true });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.4, y: 1.25, w: 2.1, h: 3.4, fill: { color: C.text1 },
    rectRadius: 0.15, objectName: "Podekovani" });
  s.addText("Děkuji za pozornost", { x: 7.55, y: 1.45, w: 1.8, h: 1.9, fontSize: 22, bold: true, fontFace: "Arial",
    color: C.background1, valign: "top", margin: 0, isTextBox: true });
  s.addText(AUTHOR, { x: 7.55, y: 3.9, w: 1.8, h: 0.5, fontSize: 13, color: C.accent6, valign: "bottom",
    margin: 0, isTextBox: true });

  const out = path.join(__dirname, "Produktovy_list_Olaplex_No4.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("ok", out);
})();
