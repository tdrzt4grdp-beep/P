// Svatební účes – prezentace (Andrea Habáňová, 3.K)
// Všechny barvy jsou zadané pevně (hex), aby se zobrazily stejně v PowerPointu, Keynote i na telefonu.
const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const gi = require("react-icons/gi");
const fa = require("react-icons/fa6");
const pi = require("react-icons/pi");

const AUTHOR = "Andrea Habáňová";
const CLASS = "3.K";
const COL = {
  wine: "6E2F40", rose: "A8576A", petal: "EBCFD6", blush: "F8EEF0", ink: "33262B",
  muted: "75666B", white: "FFFFFF", gold: "C9A66B",
};
const HEAD = "Georgia", BODY = "Arial"; // písma, která má Windows, Mac i iPhone
// Režim sestavení: "visual" = jen grafika (z ní se udělá obrázek pozadí), "final" = obrázek pozadí + text
const MODE = process.argv[2] || "final";
const BG = (n) => path.join(__dirname, "pozadi", `slide-${n}.png`);
const P = (n) => path.join(__dirname, "foto", "orez", n + ".jpg");

const meta = JSON.parse(fs.readFileSync(path.join(__dirname, "foto", "zdroje_fotek.json"), "utf8"));
const credit = (n) => {
  const m = meta[n]; const a = m.attr || "";
  const by = (a.match(/ by (.+?) is (licensed|marked)/) || [])[1];
  return `Foto: ${by || "rawpixel"}, ${m.lic === "cc0" ? "CC0" : "CC " + m.lic.toUpperCase()}`;
};

async function icon(Comp, color = "#FFFFFF") {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const ic = {
    dress: await icon(fa.FaPersonDress), hair: await icon(gi.GiHairStrands), spark: await icon(pi.PiSparkleFill),
    comb: await icon(gi.GiComb), list: await icon(fa.FaListOl), flower: await icon(pi.PiFlowerFill),
    heart: await icon(fa.FaHeart), check: await icon(fa.FaCheck), eye: await icon(fa.FaEye),
    ring: await icon(gi.GiDiamondRing, "#C9A66B"), spray: await icon(gi.GiSpray),
  };

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
  pres.author = AUTHOR;
  pres.title = "Svatební účes – romantický nízký drdol";
  pres.theme = { headFontFace: HEAD, bodyFontFace: BODY };

  pres.defineSlideMaster({ title: "Tmavy", background: { color: COL.wine }, objects: [] });
  let slideNo = 0;
  const newSlide = (opts) => {
    const sl = pres.addSlide(opts); slideNo += 1;
    if (MODE === "visual") { sl.addText = () => sl; sl.addNotes = () => sl; }
    else {
      sl.addImage({ path: BG(slideNo), x: 0, y: 0, w: 10, h: 5.625, altText: "Pozadí snímku" });
      sl.addShape = () => sl; sl.addImage = () => sl;
      if (opts.masterName === "Obsah") {
        sl.addText(`${AUTHOR} · ${CLASS} · Kadeřnická technologie – Slavnostní účesy`, { x: 0.5, y: 5.27, w: 7, h: 0.22,
          fontSize: 9, color: COL.muted, fontFace: BODY, margin: 0, isTextBox: true });
        sl.addText(String(slideNo), { x: 9.0, y: 5.27, w: 0.5, h: 0.22, fontSize: 9, color: COL.muted,
          fontFace: BODY, align: "right", margin: 0, isTextBox: true });
      }
    }
    return sl;
  };
  pres.defineSlideMaster({
    title: "Obsah",
    background: { color: COL.blush },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 1.2, y: 0.42, w: 8.3, h: 0.55,
        fontSize: 28, bold: true, fontFace: HEAD, color: COL.wine, valign: "middle", align: "left", margin: 0 }, text: "" } },
    ],
  });

  const shadow = () => ({ type: "outer", color: "6E2F40", blur: 8, offset: 2, angle: 90, opacity: 0.15 });
  const card = (s, x, y, w, h, name, fill = COL.white) => s.addShape(pres.shapes.ROUNDED_RECTANGLE,
    { x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 }, rectRadius: 0.1, shadow: shadow(), objectName: name });
  const txt = (s, text, o) => s.addText(text, { margin: 0, isTextBox: true, fontFace: BODY, color: COL.ink, valign: "top", ...o });
  // Fotka v bílém rámečku se stínem + popisek autora
  const framed = (s, n, x, y, w, h, alt, opts = {}) => {
    const b = opts.border ?? 0.07;
    s.addShape(pres.shapes.RECTANGLE, { x: x - b, y: y - b, w: w + 2 * b, h: h + 2 * b, fill: { color: COL.white },
      line: { color: COL.white, width: 0 }, shadow: shadow(), rotate: opts.rotate || 0, objectName: "Ramecek " + n });
    s.addImage({ path: P(n), x, y, w, h, rotate: opts.rotate || 0, altText: alt });
    if (!opts.noCredit) txt(s, credit(opts.src || n), { x: x - b, y: y + h + b + 0.03, w: w + 2 * b, h: 0.17,
      fontSize: 8, color: opts.creditColor || COL.muted });
  };
  const header = (s, icData, section, title) => {
    s.addShape(pres.shapes.OVAL, { x: 0.5, y: 0.38, w: 0.55, h: 0.55, fill: { color: COL.rose },
      line: { color: COL.rose, width: 0 }, objectName: "Ikona sekce" });
    s.addImage({ data: icData, x: 0.63, y: 0.51, w: 0.29, h: 0.29, altText: "" });
    txt(s, section, { x: 1.2, y: 0.18, w: 6, h: 0.24, fontSize: 10, bold: true, color: COL.rose, charSpacing: 3 });
    s.addText(title, { placeholder: "title" });
  };
  const num = (s, n, x, y, d, fill = COL.rose) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill, width: 0 }, objectName: `Cislo ${n}` });
    txt(s, String(n), { x, y, w: d, h: d, fontSize: 12, bold: true, color: COL.white, align: "center", valign: "middle" });
  };
  const badge = (s, icData, x, y, d) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: COL.rose }, line: { color: COL.rose, width: 0 }, objectName: "Ikona" });
    s.addImage({ data: icData, x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5, altText: "" });
  };
  const labelRows = (rows, size = 13) => rows.flatMap(([b, t], i) => [
    { text: b + ": ", options: { bold: true, color: COL.wine } },
    { text: t, options: { color: COL.ink, breakLine: i < rows.length - 1 } },
  ]);
  const ticks = (items, color = COL.ink) => items.map((t, i) => ({ text: t,
    options: { bullet: { code: "2713", indent: 16 }, color, breakLine: i < items.length - 1 } }));

  // 1 ─ TITULNÍ SNÍMEK
  pres.addSection({ title: "Úvod" });
  let s = newSlide({ masterName: "Tmavy", sectionTitle: "Úvod" });
  s.addImage({ data: ic.ring, x: 0.6, y: 0.65, w: 0.42, h: 0.42, altText: "" });
  txt(s, "SAMOSTATNÁ PRÁCE · SLAVNOSTNÍ ÚČESY", { x: 1.15, y: 0.74, w: 4.6, h: 0.3, fontSize: 10, bold: true,
    color: COL.petal, charSpacing: 2 });
  txt(s, "Svatební účes", { x: 0.6, y: 1.35, w: 4.8, h: 0.9, fontSize: 44, bold: true, fontFace: HEAD, color: COL.white });
  txt(s, "Romantický nízký drdol s copánkem a květinovou sponou", { x: 0.6, y: 2.3, w: 4.4, h: 0.85, fontSize: 18,
    italic: true, color: COL.petal });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 3.85, w: 3.3, h: 0.95, fill: { color: "843B4F" },
    line: { color: "843B4F", width: 0 }, rectRadius: 0.1, objectName: "Jmeno" });
  txt(s, [{ text: AUTHOR, options: { bold: true, color: COL.white, breakLine: true } },
    { text: "Třída " + CLASS, options: { color: COL.petal } }],
    { x: 0.8, y: 3.97, w: 3.0, h: 0.75, fontSize: 16, valign: "middle" });
  framed(s, "kytice", 6.55, 0.55, 2.9, 3.63, "Nevěsta v bílých šatech s kyticí", { border: 0.1, rotate: 3, noCredit: true });
  framed(s, "navrh_zezadu", 5.35, 2.55, 1.75, 2.22, "Nízký drdol s květinovými vlásenkami", { border: 0.08, rotate: -5, noCredit: true });
  txt(s, `${credit("kytice")} · ${credit("navrh_zezadu")}`, { x: 5.3, y: 5.3, w: 4.4, h: 0.18, fontSize: 8,
    color: COL.petal, align: "right" });
  s.addNotes("Dobrý den, jmenuji se Andrea Habáňová ze třídy 3.K. Vystupuji jako kadeřnice a představím vám " +
    "svůj návrh svatebního účesu – romantický nízký drdol s copánkem a květinovou sponou.");

  // 2 ─ PŘÍLEŽITOST A ZÁKAZNICE
  pres.addSection({ title: "Zákaznice a vlasy" });
  s = newSlide({ masterName: "Obsah", sectionTitle: "Zákaznice a vlasy" });
  header(s, ic.dress, "01 · PŘÍLEŽITOST", "Příležitost a zákaznice");
  card(s, 0.5, 1.25, 4.75, 3.8, "Zakaznice");
  txt(s, "Nevěsta Kristýna, 27 let", { x: 0.75, y: 1.42, w: 4.3, h: 0.4, fontSize: 18, bold: true, fontFace: HEAD, color: COL.wine });
  txt(s, labelRows([
    ["Kam jde", "letní svatba na zámku, obřad venku v zahradě, hostina a tanec do noci"],
    ["Styl", "romantický, jemný a přirozený"],
    ["Šaty", "bílé saténové, krajka, knoflíčky na zádech, dlouhý závoj"],
    ["Chce", "elegantní účes, odhalený krk a záda, jemné květiny, vydrží celý den"],
  ]), { x: 0.75, y: 2.0, w: 4.3, h: 2.95, fontSize: 14, paraSpaceAfter: 12 });
  framed(s, "saty", 5.65, 1.3, 1.75, 3.25, "Záda svatebních šatů s knoflíčky");
  framed(s, "titul", 7.7, 1.3, 1.75, 3.25, "Nevěsta s vyčesanými vlasy a kyticí");
  s.addNotes("Zákaznice je nevěsta Kristýna. Má letní svatbu na zámku. Šaty jsou bílé saténové s knoflíčky " +
    "na zádech, proto chce vlasy vyčesané – aby byla vidět záda i krk.");

  // 3 ─ CHARAKTERISTIKA VLASŮ
  s = newSlide({ masterName: "Obsah", sectionTitle: "Zákaznice a vlasy" });
  header(s, ic.hair, "02 · VLASY", "Charakteristika vlasů");
  framed(s, "vlasy", 0.6, 1.3, 3.6, 3.35, "Dlouhé husté vlnité vlasy zezadu");
  const tiles = [["DÉLKA", "dlouhé", "pod lopatky"], ["HUSTOTA", "husté", "hodně vlasů"],
    ["SÍLA", "střední", "středně silný vlas"], ["STRUKTURA", "vlnité", "přirozeně mírně"]];
  tiles.forEach(([l, v, d], i) => {
    const x = 4.65 + (i % 2) * 2.45, y = 1.25 + Math.floor(i / 2) * 1.12;
    card(s, x, y, 2.3, 0.98, "Vlastnost " + l);
    txt(s, l, { x: x + 0.18, y: y + 0.1, w: 2.0, h: 0.2, fontSize: 9, bold: true, color: COL.rose, charSpacing: 2 });
    txt(s, v, { x: x + 0.18, y: y + 0.3, w: 2.0, h: 0.38, fontSize: 20, bold: true, fontFace: HEAD, color: COL.wine });
    txt(s, d, { x: x + 0.18, y: y + 0.68, w: 2.0, h: 0.22, fontSize: 11, color: COL.muted });
  });
  card(s, 4.65, 3.55, 4.75, 1.5, "Vhodnost", COL.petal);
  txt(s, "Jsou vhodné pro můj návrh? Ano!", { x: 4.85, y: 3.65, w: 4.4, h: 0.3, fontSize: 14, bold: true, color: COL.wine });
  txt(s, ticks(["délka stačí na drdol i copánek", "hustota dodá drdolu objem",
    "vlnitost pomáhá držet tvar", "jsou těžké → použiji více vlásenek"]),
  { x: 4.85, y: 3.98, w: 4.4, h: 1.0, fontSize: 12 });
  s.addNotes("Vlasy jsou dlouhé, husté a mírně vlnité. To je pro drdol ideální – mám dost vlasů na objem " +
    "i copánek. Jsou ale těžké, takže je musím pořádně upevnit vlásenkami.");

  // 4 ─ NÁVRH ÚČESU
  pres.addSection({ title: "Návrh a postup" });
  s = newSlide({ masterName: "Obsah", sectionTitle: "Návrh a postup" });
  header(s, ic.spark, "03 · MŮJ NÁVRH", "Návrh účesu");
  framed(s, "navrh_zezadu", 0.6, 1.3, 2.45, 3.1, "Nízký drdol s bílými květinovými vlásenkami – zezadu");
  framed(s, "navrh_zboku", 3.35, 1.3, 2.45, 3.1, "Nízký drdol s copánkem a květinovou sponou – z boku");
  txt(s, [{ text: "Tvar: ", options: { bold: true, color: COL.wine } },
    { text: "nízký drdol ze stočených vln, v zátylku mírně vpravo" }],
  { x: 0.55, y: 4.75, w: 5.3, h: 0.4, fontSize: 12, italic: true });
  const views = [["Zepředu", "pěšinka na stranu, jemný objem na temeni, 2 vlnité pramínky u obličeje"],
    ["Z boku", "copánek od spánku dozadu do drdolu, drdol je vidět za uchem"],
    ["Zezadu", "drdol ze smyček, květinová spona s perličkami, pod ním závoj"]];
  views.forEach(([t, d], i) => {
    const y = 1.25 + i * 1.3;
    card(s, 6.2, y, 3.3, 1.15, t);
    badge(s, ic.eye, 6.35, y + 0.15, 0.4);
    txt(s, t, { x: 6.88, y: y + 0.17, w: 2.5, h: 0.35, fontSize: 14, bold: true, color: COL.wine, valign: "middle" });
    txt(s, d, { x: 6.35, y: y + 0.6, w: 3.05, h: 0.52, fontSize: 11 });
  });
  s.addNotes("Navrhla jsem nízký drdol posazený v zátylku trochu na stranu. Zepředu je hladký s pěšinkou na " +
    "stranu a dvěma pramínky u obličeje. Z boku vede copánek do drdolu. Zezadu je drdol ze stočených vln " +
    "a v něm květinová spona.");

  // 5 ─ PŘÍPRAVA, POMŮCKY A PŘÍPRAVKY
  s = newSlide({ masterName: "Obsah", sectionTitle: "Návrh a postup" });
  header(s, ic.comb, "04 · PŘÍPRAVA", "Příprava, pomůcky a přípravky");
  framed(s, "kulma", 0.6, 1.3, 2.2, 1.5, "Kadeřnice natáčí vlasy kulmou");
  framed(s, "vlasenky", 0.6, 3.3, 2.2, 1.5, "Vlásenky");
  card(s, 3.15, 1.25, 3.25, 3.8, "Pomucky");
  badge(s, ic.comb, 3.3, 1.38, 0.4);
  txt(s, "Pomůcky", { x: 3.8, y: 1.38, w: 2.4, h: 0.4, fontSize: 15, bold: true, color: COL.wine, valign: "middle" });
  txt(s, labelRows([["Fén + kulatý kartáč", "vysušení s objemem"], ["Kulma 25 mm", "natočení vln"],
    ["Tupírovací hřeben", "objem na temeni"], ["Klipsy", "rozdělení do sekcí"], ["Gumičky", "copánek a culík"],
    ["Vlásenky a vidličky", "upevnění drdolu"]]), { x: 3.35, y: 1.95, w: 2.95, h: 3.0, fontSize: 13, paraSpaceAfter: 9 });
  card(s, 6.6, 1.25, 2.9, 2.45, "Pripravky");
  badge(s, ic.spray, 6.75, 1.38, 0.4);
  txt(s, "Přípravky", { x: 7.25, y: 1.38, w: 2.1, h: 0.4, fontSize: 15, bold: true, color: COL.wine, valign: "middle" });
  txt(s, labelRows([["Objemová pěna", "při fénování"], ["Tepelná ochrana", "před kulmou"],
    ["Texturizační sprej", "lepší držení"], ["Lak – silná fixace", "konečné zpevnění"]]),
  { x: 6.8, y: 1.95, w: 2.6, h: 1.7, fontSize: 12, paraSpaceAfter: 7 });
  card(s, 6.6, 3.9, 2.9, 1.15, "Priprava", COL.petal);
  txt(s, [{ text: "Příprava vlasů: ", options: { bold: true, color: COL.wine } },
    { text: "umýt den předem (lépe drží), vyfénovat s pěnou a rozčesat." }],
  { x: 6.78, y: 4.0, w: 2.6, h: 0.98, fontSize: 12, valign: "middle" });
  s.addNotes("Vlasy nechám umýt den předem, protože čerstvě umyté vlasy jsou kluzké a hůř drží. Použiji fén, " +
    "kulmu, tupírovací hřeben, klipsy, gumičky a hodně vlásenek. Z přípravků pěnu, tepelnou ochranu, " +
    "texturizační sprej a lak.");

  // 6 ─ PRACOVNÍ POSTUP
  s = newSlide({ masterName: "Obsah", sectionTitle: "Návrh a postup" });
  header(s, ic.list, "05 · POSTUP", "Pracovní postup");
  framed(s, "cop", 0.6, 1.3, 2.35, 3.35, "Zapletený copánek ve vlasech");
  card(s, 3.3, 1.25, 6.2, 3.8, "Kroky");
  const steps = [
    "Rozčešu vlasy, nanesu tepelnou ochranu a rozdělím je klipsy do sekcí.",
    "Kulmou natočím vlny, nechám vychladnout a nastříkám texturizační sprej.",
    "Vytupíruji temeno, uhladím, udělám pěšinku na stranu a oddělím 2 pramínky u obličeje.",
    "Od spánku zapletu copánek směrem dozadu k zátylku.",
    "Zbytek vlasů i s copánkem stáhnu gumičkou do nízkého culíku vpravo v zátylku.",
    "Vlny z culíku stáčím do smyček kolem gumičky a připínám vlásenkami – vznikne drdol.",
    "Vložím ozdoby a závoj, upravím pramínky u obličeje a vše zafixuji lakem.",
  ];
  steps.forEach((t, i) => {
    const y = 1.4 + i * 0.51;
    num(s, i + 1, 3.5, y + 0.04, 0.34);
    txt(s, t, { x: 4.0, y, w: 5.35, h: 0.44, fontSize: 12, valign: "middle" });
  });
  s.addNotes("Postupuji od přípravy přes tvarování až po fixaci: vlny kulmou, objem na temeni, copánek, " +
    "nízký culík, z něj drdol ze smyček a nakonec ozdoby, závoj a lak.");

  // 7 ─ DOPLŇKY
  pres.addSection({ title: "Doplňky a závěr" });
  s = newSlide({ masterName: "Obsah", sectionTitle: "Doplňky a závěr" });
  header(s, ic.flower, "06 · DOPLŇKY", "Doplňky a jejich výběr");
  const acc = [
    ["spona", "Květinová spona", ["krémové květy ladí se šaty", "jemná a lehká, netlačí", "připnu ji nad drdol vpravo"], "Květinová spona do vlasů"],
    ["kvetiny", "Perličkové vlásenky", ["zdobí drdol na více místech", "perličky jako knoflíčky šatů", "zapíchnu je mezi smyčky drdolu"], "Bílé květinové vlásenky"],
    ["zavoj", "Dlouhý závoj", ["připnu hřebínkem pod drdol", "nezakryje účes, na večer sundám", "jemný tyl bez výšivky – nepřebije účes"], "Nevěsta se závojem"],
  ];
  acc.forEach(([img, t, pts, alt], i) => {
    const x = 0.5 + i * 3.075;
    card(s, x, 1.25, 2.85, 3.8, t);
    s.addImage({ path: P(img), x: x + 0.12, y: 1.37, w: 2.61, h: 1.42, altText: alt });
    txt(s, t, { x: x + 0.2, y: 2.92, w: 2.45, h: 0.35, fontSize: 15, bold: true, fontFace: HEAD, color: COL.wine });
    txt(s, ticks(pts), { x: x + 0.2, y: 3.35, w: 2.5, h: 1.35, fontSize: 13, paraSpaceAfter: 6 });
    txt(s, credit(img), { x: x + 0.2, y: 4.78, w: 2.5, h: 0.17, fontSize: 8, color: COL.muted });
  });
  s.addNotes("Doplňky jsou jemné, v krémové a bílé barvě, aby ladily se šaty. Perličky připomínají knoflíčky " +
    "na zádech šatů. Závoj připnu pod drdol, takže účes zůstane vidět.");

  // 8 ─ ZÁVĚREČNÉ ZHODNOCENÍ
  s = newSlide({ masterName: "Obsah", sectionTitle: "Doplňky a závěr" });
  header(s, ic.check, "07 · ZÁVĚR", "Závěrečné zhodnocení");
  card(s, 0.5, 1.25, 5.6, 1.45, "Proc je vhodny", COL.petal);
  txt(s, "Proč je návrh vhodný", { x: 0.7, y: 1.35, w: 5.2, h: 0.3, fontSize: 14, bold: true, color: COL.wine });
  txt(s, ticks(["romantický a elegantní – sedí k nevěstě i šatům", "odhalí krk a záda šatů s knoflíčky",
    "využije délku a hustotu vlasů zákaznice"]), { x: 0.7, y: 1.7, w: 5.2, h: 0.95, fontSize: 12 });
  const ev = [[ic.check, "Drží", "pevný culík jako základ, vlásenky křížem, lak"],
    [ic.heart, "Pohodlný", "nízký drdol netlačí, závoj jde sundat"],
    [ic.eye, "Upravený", "kontrola zrcadlem ze všech stran, svědkyně má vlásenky a lak"]];
  ev.forEach(([icD, t, d], i) => {
    const x = 0.5 + i * 1.9;
    card(s, x, 2.9, 1.7, 2.15, t);
    badge(s, icD, x + 0.15, 3.03, 0.42);
    txt(s, t, { x: x + 0.15, y: 3.55, w: 1.45, h: 0.32, fontSize: 14, bold: true, color: COL.wine });
    txt(s, d, { x: x + 0.15, y: 3.9, w: 1.45, h: 1.1, fontSize: 12 });
  });
  framed(s, "vysledek", 6.5, 1.3, 2.95, 3.35, "Hotový nízký drdol s květinovými vlásenkami", { src: "navrh_zezadu" });
  s.addNotes("Účes je vhodný, protože je romantický, ladí se šaty a ukáže záda. Aby držel celý den, " +
    "dělám pevný základ z culíku a vlásenky dávám křížem. Je pohodlný, nic netahá. Před svatbou " +
    "uděláme zkoušku účesu a svědkyni dám vlásenky a lak na opravy.");

  // 9 ─ DĚKUJI ZA POZORNOST
  s = newSlide({ masterName: "Tmavy", sectionTitle: "Doplňky a závěr" });
  s.addImage({ data: ic.ring, x: 0.6, y: 1.0, w: 0.45, h: 0.45, altText: "" });
  txt(s, "Děkuji za pozornost", { x: 0.6, y: 1.65, w: 5, h: 1.5, fontSize: 44, bold: true, fontFace: HEAD, color: COL.white });
  txt(s, "Máte nějaké otázky?", { x: 0.6, y: 3.2, w: 4.6, h: 0.45, fontSize: 18, italic: true, color: COL.petal });
  txt(s, [{ text: AUTHOR, options: { bold: true, color: COL.white, breakLine: true } },
    { text: "Třída " + CLASS, options: { color: COL.petal } }], { x: 0.6, y: 4.25, w: 4.6, h: 0.7, fontSize: 15 });
  framed(s, "dekuji", 6.3, 0.6, 3.0, 3.75, "Nízký drdol s květinovou sponou z boku",
    { border: 0.1, rotate: 3, src: "navrh_zboku", creditColor: COL.petal });
  s.addNotes("Děkuji za pozornost. Ráda odpovím na vaše otázky.");

  const out = MODE === "visual" ? path.join(__dirname, "pozadi", "graficka_vrstva.pptx")
    : path.join(__dirname, "Svatebni_uces_Habanova.pptx");
  await pres.writeFile({ fileName: out });
  console.log("ok", out);
})();
