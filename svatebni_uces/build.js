const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const { applyTheme } = require(path.join(__dirname, "../.claude/skills/pptx/scripts/apply_theme.js"));

const THEME = {
  name: "Svatebni ucesy",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "33262B", lt1: "FFFFFF", dk2: "5A4A50", lt2: "F8EEF0",
    accent1: "A8576A", accent2: "D9A7B2", accent3: "7E8F7A", accent4: "8A7C80",
    accent5: "7C3A4A", accent6: "EED6DC", hlink: "7C3A4A", folHlink: "8A7C80",
  },
};
const AUTHOR = "Andrea Habáňová";
const P = (n) => path.join(__dirname, "foto", "orez", n + ".jpg");

// Autoři fotek pro popisky (licence CC)
const meta = JSON.parse(fs.readFileSync(path.join(__dirname, "foto", "zdroje_fotek.json"), "utf8"));
const credit = (n) => {
  const m = meta[n]; const a = m.attr || "";
  const by = (a.match(/ by (.+?) is (licensed|marked)/) || [])[1];
  const lic = m.lic === "cc0" ? "CC0" : "CC " + m.lic.toUpperCase();
  return `Foto: ${by || "rawpixel"}, ${lic}`;
};

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.author = AUTHOR;
  pres.title = "Svatební účes – romantický nízký drdol";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  pres.defineSlideMaster({
    title: "Titulni",
    background: { color: C.background2 },
    objects: [],
  });
  pres.defineSlideMaster({
    title: "Obsah",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9, h: 0.7,
        fontSize: 30, bold: true, color: C.accent5, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { text: { text: `${AUTHOR} · Kadeřnická technologie – Slavnostní účesy`, options: {
        x: 0.5, y: 5.25, w: 7, h: 0.25, fontSize: 10, color: C.accent4, margin: 0 } } },
    ],
    slideNumber: { x: 9.0, y: 5.25, w: 0.5, h: 0.25, fontSize: 10, color: "8A7C80", align: "right" },
  });

  const photo = (s, n, x, y, w, h, alt, src) => {
    s.addImage({ path: P(n), x, y, w, h, altText: alt });
    s.addText(credit(src || n), { x, y: y + h + 0.03, w, h: 0.18, fontSize: 8, color: C.accent4,
      margin: 0, isTextBox: true });
  };
  const card = (s, x, y, w, h, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h,
    fill: { color: C.background2 }, rectRadius: 0.1, objectName: name });
  const heading = (s, t, x, y, w) => s.addText(t, { x, y, w, h: 0.32, fontSize: 15, bold: true,
    color: C.accent1, margin: 0, isTextBox: true });
  const bullets = (s, items, x, y, w, h, size = 14) => s.addText(
    items.map((t, i) => {
      const runs = Array.isArray(t) ? t : [t];
      return { text: runs[0], options: { bullet: { indent: 14 }, breakLine: i < items.length - 1 } };
    }),
    { x, y, w, h, fontSize: size, color: C.text1, paraSpaceAfter: 4, valign: "top", margin: 0, isTextBox: true });
  const rich = (s, rows, x, y, w, h, size = 13) => s.addText(
    rows.flatMap(([b, t], i) => [
      { text: b, options: { bold: true, color: C.accent5, bullet: { indent: 14 } } },
      { text: " – " + t, options: { breakLine: i < rows.length - 1 } },
    ]),
    { x, y, w, h, fontSize: size, color: C.text1, paraSpaceAfter: 4, valign: "top", margin: 0, isTextBox: true });
  const circle = (s, n, x, y, d) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: C.accent1 }, objectName: `Krok ${n}` });
    s.addText(String(n), { x, y, w: d, h: d, fontSize: 13, bold: true, align: "center", valign: "middle",
      color: C.background1, margin: 0, isTextBox: true });
  };

  // 1 — TITULNÍ SNÍMEK
  pres.addSection({ title: "Úvod" });
  let s = pres.addSlide({ masterName: "Titulni", sectionTitle: "Úvod" });
  s.addImage({ path: P("kytice"), x: 5.5, y: 0, w: 4.5, h: 5.625, altText: "Nevěsta v bílých šatech s kyticí" });
  s.addText("SAMOSTATNÁ PRÁCE · SVATEBNÍ ÚČES", { x: 0.6, y: 0.9, w: 4.6, h: 0.3, fontSize: 12, bold: true,
    color: C.accent1, charSpacing: 2, margin: 0, isTextBox: true });
  s.addText("Romantický nízký drdol", { x: 0.6, y: 1.35, w: 4.7, h: 1.4, fontSize: 38, bold: true,
    fontFace: "Cambria", color: C.accent5, valign: "top", margin: 0, isTextBox: true });
  s.addText("s copánkem a květinovou sponou pro nevěstu", { x: 0.6, y: 2.8, w: 4.6, h: 0.7, fontSize: 18,
    italic: true, color: C.text2, valign: "top", margin: 0, isTextBox: true });
  s.addText([
    { text: "Jméno: ", options: { color: C.accent4 } }, { text: AUTHOR, options: { bold: true, breakLine: true } },
    { text: "Třída: ", options: { color: C.accent4 } }, { text: "………………", options: { bold: true } },
  ], { x: 0.6, y: 4.1, w: 4.6, h: 0.75, fontSize: 15, color: C.text1, valign: "top", margin: 0, isTextBox: true });
  s.addText(credit("kytice"), { x: 5.55, y: 5.38, w: 4.4, h: 0.2, fontSize: 8, color: C.text2,
    margin: 0, isTextBox: true });
  s.addNotes("Dobrý den, jmenuji se Andrea Habáňová. Vystupuji jako kadeřnice a představím vám svůj návrh " +
    "svatebního účesu – romantický nízký drdol s copánkem a květinovou sponou.");

  // 2 — PŘÍLEŽITOST A ZÁKAZNICE
  pres.addSection({ title: "Zákaznice" });
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Zákaznice" });
  s.addText("Příležitost a zákaznice", { placeholder: "title" });
  rich(s, [
    ["Zákaznice", "nevěsta Kristýna, 27 let"],
    ["Příležitost", "letní svatba na zámku, obřad venku v zahradě, potom hostina a tanec do noci"],
    ["Styl", "romantický, jemný a přirozený, nemá ráda nic přehnaného"],
    ["Šaty", "bílé saténové šaty s knoflíčky na zádech a krajkou, dlouhý závoj"],
    ["Požadovaný vzhled", "elegantní účes, odhalený krk a záda, jemné květiny, účes musí vydržet celý den"],
  ], 0.5, 1.2, 4.6, 3.9, 14);
  photo(s, "saty", 5.4, 1.2, 2.0, 3.7, "Záda svatebních šatů s knoflíčky");
  photo(s, "titul", 7.5, 1.2, 2.0, 3.7, "Nevěsta s vyčesanými vlasy a kyticí");
  s.addNotes("Moje zákaznice je nevěsta Kristýna. Má letní svatbu na zámku. Šaty jsou bílé, saténové, " +
    "s knoflíčky na zádech – proto chce mít vlasy vyčesané, aby byla vidět záda i krk.");

  // 3 — CHARAKTERISTIKA VLASŮ
  pres.addSection({ title: "Návrh" });
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Návrh" });
  s.addText("Charakteristika vlasů", { placeholder: "title" });
  photo(s, "vlasy", 0.5, 1.2, 4.0, 3.7, "Dlouhé husté vlnité vlasy zezadu");
  rich(s, [
    ["Délka", "dlouhé – pod lopatky"],
    ["Hustota", "husté"],
    ["Síla vlasu", "středně silné"],
    ["Struktura", "přirozeně mírně vlnité"],
    ["Stav", "zdravé, bez roztřepených konečků"],
  ], 4.9, 1.2, 4.6, 1.85, 14);
  card(s, 4.9, 3.15, 4.6, 1.9, "Vhodnost vlasu");
  heading(s, "Jsou vlasy vhodné? Ano!", 5.1, 3.25, 4.2);
  bullets(s, [
    "délka stačí na drdol i copánek",
    "hustota dodá drdolu objem",
    "vlnitost pomáhá účesu držet tvar",
    "vlasy jsou těžké – použiji víc vlásenek",
  ], 5.1, 3.6, 4.2, 1.4, 13);
  s.addNotes("Vlasy jsou dlouhé, husté a mírně vlnité. To je pro drdol ideální, protože mám dost vlasů " +
    "na objem i na copánek. Jediná nevýhoda je, že jsou těžké, takže je musím pořádně upevnit.");

  // 4 — NÁVRH ÚČESU
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Návrh" });
  s.addText("Návrh účesu", { placeholder: "title" });
  photo(s, "navrh_zezadu", 0.5, 1.2, 2.6, 3.3, "Nízký drdol s bílými květinovými vlásenkami zezadu");
  photo(s, "navrh_zboku", 3.3, 1.2, 2.6, 3.3, "Nízký drdol s copánkem a květinovou sponou z boku");
  s.addText("Tvar: nízký drdol ze stočených vln, umístěný v zátylku mírně vpravo", { x: 0.5, y: 4.7, w: 5.4,
    h: 0.45, fontSize: 12, italic: true, color: C.text2, valign: "top", margin: 0, isTextBox: true });
  const views = [
    ["Zepředu", "pěšinka na stranu, mírný objem na temeni, 2 vlnité pramínky u obličeje"],
    ["Z boku", "copánek od spánku dozadu k drdolu, drdol je vidět za uchem"],
    ["Zezadu", "drdol ze smyček, květinová spona s perličkami, pod ním závoj"],
  ];
  views.forEach(([t, d], i) => {
    const y = 1.2 + i * 1.3;
    card(s, 6.2, y, 3.3, 1.15, t);
    s.addText(t, { x: 6.38, y: y + 0.1, w: 3.0, h: 0.3, fontSize: 14, bold: true, color: C.accent1,
      margin: 0, isTextBox: true });
    s.addText(d, { x: 6.38, y: y + 0.42, w: 3.0, h: 0.68, fontSize: 12, color: C.text1, valign: "top",
      margin: 0, isTextBox: true });
  });
  s.addNotes("Můj návrh je nízký drdol, který je posazený v zátylku trochu na stranu. Zepředu je hladký " +
    "s pěšinkou na stranu a dvěma pramínky u obličeje. Z boku vede copánek do drdolu. Zezadu je drdol " +
    "ze stočených vln a v něm květinová spona.");

  // 5 — PŘÍPRAVA, POMŮCKY A PŘÍPRAVKY
  pres.addSection({ title: "Postup" });
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Postup" });
  s.addText("Příprava, pomůcky a přípravky", { placeholder: "title" });
  photo(s, "kulma", 0.5, 1.2, 2.4, 1.6, "Kadeřnice natáčí vlasy kulmou");
  photo(s, "vlasenky", 0.5, 3.15, 2.4, 1.6, "Vlásenky");
  heading(s, "Pomůcky", 3.2, 1.2, 3.2);
  rich(s, [
    ["Fén a kulatý kartáč", "vysušení s objemem"],
    ["Kulma 25 mm", "natočení vln"],
    ["Tupírovací hřeben", "objem na temeni"],
    ["Klipsy", "rozdělení vlasů do sekcí"],
    ["Gumičky", "copánek a culík"],
    ["Vlásenky a vidličky", "upevnění drdolu"],
  ], 3.2, 1.55, 3.2, 3.5, 12);
  heading(s, "Přípravky", 6.65, 1.2, 2.85);
  rich(s, [
    ["Objemová pěna", "při fénování"],
    ["Tepelná ochrana", "před kulmou"],
    ["Texturizační sprej", "vlasy lépe drží"],
    ["Lak – silná fixace", "konečné zpevnění"],
  ], 6.65, 1.55, 2.85, 2.0, 12);
  card(s, 6.65, 3.65, 2.85, 1.4, "Priprava");
  s.addText([{ text: "Příprava: ", options: { bold: true, color: C.accent5 } },
    { text: "vlasy umýt den předem (lépe drží), vyfénovat s pěnou a rozčesat." }],
    { x: 6.8, y: 3.75, w: 2.6, h: 1.2, fontSize: 12, color: C.text1, valign: "top", margin: 0, isTextBox: true });
  s.addNotes("Vlasy nechám umýt den předem, protože čisté vlasy jsou moc kluzké a hůř drží. Použiji fén, " +
    "kulmu, tupírovací hřeben, klipsy, gumičky a hodně vlásenek. Z přípravků pěnu, tepelnou ochranu, " +
    "texturizační sprej a lak.");

  // 6 — PRACOVNÍ POSTUP
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Postup" });
  s.addText("Pracovní postup", { placeholder: "title" });
  photo(s, "cop", 0.5, 1.2, 2.6, 3.7, "Zapletený copánek ve vlasech");
  const steps = [
    "Rozčešu vlasy, nanesu tepelnou ochranu a rozdělím je klipsy do sekcí.",
    "Kulmou natočím vlny, nechám je vychladnout a nastříkám texturizační sprej.",
    "Vytupíruji temeno, uhladím ho a udělám pěšinku na stranu. Oddělím 2 pramínky u obličeje.",
    "Od spánku zapletu copánek směrem dozadu k zátylku.",
    "Zbytek vlasů i s copánkem stáhnu gumičkou do nízkého culíku vpravo v zátylku.",
    "Vlny z culíku stáčím do smyček kolem gumičky a připínám je vlásenkami – vznikne drdol.",
    "Vložím ozdoby a závoj, upravím pramínky u obličeje a vše zafixuji lakem.",
  ];
  steps.forEach((t, i) => {
    const y = 1.2 + i * 0.56;
    circle(s, i + 1, 3.4, y + 0.03, 0.36);
    s.addText(t, { x: 3.9, y, w: 5.6, h: 0.5, fontSize: 12, color: C.text1, valign: "middle", margin: 0,
      isTextBox: true });
  });
  s.addNotes("Postupuji od přípravy přes tvarování až po fixaci. Nejdřív vlny kulmou, potom objem na temeni, " +
    "copánek, nízký culík a z něj drdol ze smyček. Nakonec ozdoby, závoj a lak.");

  // 7 — DOPLŇKY
  pres.addSection({ title: "Doplňky a závěr" });
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Doplňky a závěr" });
  s.addText("Doplňky a jejich výběr", { placeholder: "title" });
  const acc = [
    ["spona", "Květinová spona", ["krémové květy ladí s barvou šatů", "jemná, romantická, není těžká"], "Květinová spona do vlasů"],
    ["kvetiny", "Květinové vlásenky s perličkami", ["doplní drdol na více místech", "perličky jako knoflíčky na šatech"], "Bílé květinové vlásenky"],
    ["zavoj", "Dlouhý závoj", ["připnu ho hřebínkem pod drdol", "nezakryje účes, na večer ho sundám"], "Nevěsta se závojem a vyčesaným účesem"],
  ];
  acc.forEach(([img, t, pts, alt], i) => {
    const x = 0.5 + i * 3.075;
    card(s, x, 1.15, 2.85, 3.95, t);
    s.addImage({ path: P(img), x, y: 1.15, w: 2.85, h: 1.55, altText: alt });
    s.addText(t, { x: x + 0.18, y: 2.8, w: 2.5, h: 0.5, fontSize: 14, bold: true, color: C.accent5,
      valign: "top", margin: 0, isTextBox: true });
    bullets(s, pts, x + 0.18, 3.35, 2.5, 1.25, 12);
    s.addText(credit(img), { x: x + 0.18, y: 4.75, w: 2.5, h: 0.2, fontSize: 8, color: C.accent4,
      margin: 0, isTextBox: true });
  });
  s.addNotes("Vybrala jsem jemné doplňky v krémové a bílé barvě, aby ladily se šaty. Perličky připomínají " +
    "knoflíčky na zádech šatů. Závoj připnu pod drdol, takže účes zůstane vidět.");

  // 8 — ZÁVĚREČNÉ ZHODNOCENÍ
  s = pres.addSlide({ masterName: "Obsah", sectionTitle: "Doplňky a závěr" });
  s.addText("Závěrečné zhodnocení", { placeholder: "title" });
  heading(s, "Proč je návrh vhodný", 0.5, 1.15, 5.4);
  bullets(s, [
    "romantický a elegantní – sedí k nevěstě i šatům",
    "odhalí krk a záda šatů s knoflíčky",
    "využije délku a hustotu vlasů zákaznice",
  ], 0.5, 1.5, 5.4, 1.1, 13);
  const eval3 = [
    ["Drží", "pevný culík jako základ, vlásenky křížem, lak"],
    ["Pohodlný", "nízký drdol netlačí, závoj jde sundat"],
    ["Upravený", "kontrola zrcadlem ze všech stran, svědkyně dostane vlásenky a lak"],
  ];
  eval3.forEach(([t, d], i) => {
    const y = 2.7 + i * 0.8;
    card(s, 0.5, y, 5.4, 0.7, t);
    s.addText(t, { x: 0.68, y, w: 1.3, h: 0.7, fontSize: 14, bold: true, color: C.accent1, valign: "middle",
      margin: 0, isTextBox: true });
    s.addText(d, { x: 2.0, y, w: 3.8, h: 0.7, fontSize: 12, color: C.text1, valign: "middle", margin: 0,
      isTextBox: true });
  });
  photo(s, "vysledek", 6.2, 1.15, 3.3, 3.53, "Hotový nízký drdol s květinovými vlásenkami", "navrh_zezadu");
  s.addText("Děkuji za pozornost", { x: 6.2, y: 4.85, w: 3.3, h: 0.35, fontSize: 16, bold: true, italic: true,
    fontFace: "Cambria", color: C.accent5, align: "center", margin: 0, isTextBox: true });
  s.addNotes("Účes je vhodný, protože je romantický, ladí se šaty a ukáže záda. Aby držel celý den, " +
    "dělám pevný základ z culíku a vlásenky dávám křížem. Je pohodlný, nic netahá. Před svatbou " +
    "uděláme zkoušku účesu a svědkyni dám vlásenky a lak na opravy.");

  const out = path.join(__dirname, "Svatebni_uces_Habanova.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("ok", out);
})();
