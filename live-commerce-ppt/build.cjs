// Build: node build.cjs  (needs pptxgenjs, react, react-dom, react-icons, sharp)
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const SKILL = process.env.PPTX_SKILL_DIR;
const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));

const A = (f) => path.join(__dirname, "assets", f);
const OUT = path.join(__dirname, "Live_Commerce_Elektron_tijorat.pptx");

const THEME = {
  name: "Neon Live",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1A1033", lt1: "FFFFFF", dk2: "140C30", lt2: "F3EFFF",
    accent1: "7C3AED", accent2: "E0338C", accent3: "FFB547",
    accent4: "14B8A6", accent5: "B9A4F5", accent6: "6B6385",
    hlink: "7C3AED", folHlink: "B9A4F5",
  },
};
const HEX = THEME.colors;

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9"; // 10 x 5.625
  pres.title = "Live Commerce: efirdan savatga";
  pres.subject = "Elektron tijorat fani uchun taqdimot";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  const footer = (dark) => [
    { text: { text: "Live Commerce · Elektron tijorat", options: { x: 0.6, y: 5.18, w: 5, h: 0.3, fontSize: 10, color: dark ? C.accent5 : C.accent6, margin: 0 } } },
  ];
  const slideNum = (dark) => ({ x: 9.0, y: 5.18, w: 0.5, h: 0.3, fontSize: 10, color: dark ? C.accent5 : C.accent6, align: "right" });

  pres.defineSlideMaster({
    title: "TITLE_DARK",
    background: { color: C.text2 },
    objects: [],
  });
  pres.defineSlideMaster({
    title: "CONTENT_LIGHT",
    background: { color: C.background1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 8.8, h: 0.75, fontSize: 36, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
      ...footer(false),
    ],
    slideNumber: slideNum(false),
  });
  pres.defineSlideMaster({
    title: "CONTENT_DARK",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 4.3, h: 0.75, fontSize: 36, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
      ...footer(true),
    ],
    slideNumber: slideNum(true),
  });

  const T = (s, text, opts) => s.addText(text, { isTextBox: true, margin: 0, ...opts });

  // ---------- 1. Title ----------
  pres.addSection({ title: "Kirish" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Kirish" });
  s.addImage({ path: A("hero.png"), x: 0, y: 0, w: 10, h: 5.625, objectName: "Hero rasm", altText: "Jonli efirda mahsulot ko'rsatayotgan boshlovchi" });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 5.7, h: 5.625, fill: { color: C.text2, transparency: 12 }, line: { type: "none" }, objectName: "Matn foni" });
  T(s, "ELEKTRON TIJORAT FANI", { x: 0.6, y: 1.05, w: 4.8, h: 0.35, fontSize: 12, bold: true, color: C.accent3, charSpacing: 3, objectName: "Fan nomi" });
  T(s, [
    { text: "Live Commerce:", options: { breakLine: true } },
    { text: "efirdan savatga", options: { color: C.accent2 } },
  ], { x: 0.6, y: 1.5, w: 5.0, h: 1.6, fontSize: 40, bold: true, color: C.background1, fontFace: THEME.headFontFace, valign: "top", objectName: "Sarlavha" });
  T(s, "Jonli efir orqali savdo — elektron tijoratning eng tez o'sayotgan formati", { x: 0.6, y: 3.2, w: 4.7, h: 0.8, fontSize: 16, color: C.background2, valign: "top", objectName: "Kichik sarlavha" });
  T(s, "Taqdimot · 2026", { x: 0.6, y: 4.75, w: 4, h: 0.3, fontSize: 12, color: C.accent5, objectName: "Sana" });
  s.addNotes("Assalomu alaykum! Bugun elektron tijoratning eng qiziqarli va tez o'sayotgan yo'nalishlaridan biri — Live Commerce, ya'ni jonli efir orqali savdo haqida gaplashamiz. Ko'pchiligingiz Instagram yoki TikTok'da kimdir efirda mahsulot ko'rsatib sotayotganini ko'rgansiz — aynan shu hodisa bugun milliardlab dollarlik bozorga aylangan.");

  // ---------- 1b. Intro video (Higgsfield) ----------
  s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Kirish" });
  const fs = require("fs");
  s.addMedia({
    type: "video", path: A("intro.mp4"),
    cover: "data:image/png;base64," + fs.readFileSync(A("intro_poster.png")).toString("base64"),
    x: 0, y: 0, w: 10, h: 5.625, objectName: "Kirish videosi",
  });
  s.addNotes("Kirish videosi (30 soniya, Higgsfield'da yaratilgan). Slayd ochilganda avtomatik boshlanadi. Video savdoning evolyutsiyasini ko'rsatadi: 1980-yillardagi televizor teleshopi → smartfondagi jonli efir → mahsulot ekrandan chiqib xaridor qo'liga tushadi → kelajakdagi AI efirlar shahri. Video tugagach savol bering: 'Siz oxirgi marta jonli efirdan biror narsa sotib olganmisiz?'");

  // ---------- 2. Agenda ----------
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Kirish" });
  s.addText("Bugungi reja", { placeholder: "title" });
  const agenda = [
    "Live Commerce nima?",
    "Qisqa tarix: teleshopdan TikTok Shopgacha",
    "Raqamlar va nega u ishlaydi",
    "Formatlar va AI hostlar",
    "O'zbekiston uchun imkoniyatlar va xavflar",
  ];
  agenda.forEach((t, i) => {
    const y = 1.35 + i * 0.7;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.5, h: 0.5, fill: { color: i === 0 ? C.accent2 : C.accent1 }, line: { type: "none" }, objectName: `Raqam doira ${i + 1}` });
    T(s, String(i + 1), { x: 0.6, y, w: 0.5, h: 0.5, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", objectName: `Raqam ${i + 1}` });
    T(s, t, { x: 1.3, y, w: 4.6, h: 0.5, fontSize: 18, color: C.text1, valign: "middle", objectName: `Reja bandi ${i + 1}` });
  });
  s.addImage({ path: A("phone.png"), x: 6.25, y: 1.25, w: 3.25, h: 3.25, objectName: "Smartfonda jonli efir", altText: "Smartfonda jonli savdo efiri" });
  s.addNotes("Taqdimot besh qismdan iborat: avval tushunchani aniqlaymiz, keyin tarixiga nazar tashlaymiz, raqamlar va psixologiyasini ko'ramiz, formatlar va sun'iy intellekt hostlari haqida gaplashamiz, oxirida esa O'zbekiston uchun amaliy xulosalar chiqaramiz.");

  // ---------- 3. Definition ----------
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Kirish" });
  s.addText("Live Commerce nima?", { placeholder: "title" });
  T(s, [
    { text: "Live Commerce", options: { bold: true, color: C.accent1 } },
    { text: " — sotuvchi yoki boshlovchi mahsulotni jonli videoda ko'rsatadi, tomoshabinlar chatda savol beradi va efirdan chiqmasdan xarid qiladi." },
  ], { x: 0.6, y: 1.2, w: 8.8, h: 0.75, fontSize: 16, color: C.text1, valign: "top", objectName: "Ta'rif" });
  const parts = [
    { ic: fa.FaTv, h: "Teleshop", d: "Mahsulotni ko'rsatib sotish" },
    { ic: fa.FaComments, h: "Ijtimoiy tarmoq", d: "Jonli muloqot va jamoa" },
    { ic: fa.FaShoppingCart, h: "Onlayn do'kon", d: "Bir tugmada to'lov" },
    { ic: fa.FaBroadcastTower, h: "Live Commerce", d: "Ko'r, so'ra, darhol sotib ol", hl: true },
  ];
  const cw = 1.8, gap = 0.5, cx0 = 0.65, cy = 2.3, ch = 2.45;
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    const x = cx0 + i * (cw + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy, w: cw, h: ch, rectRadius: 0.12, fill: { color: p.hl ? C.accent1 : C.background2 }, line: { type: "none" }, objectName: `Karta ${p.h}` });
    s.addShape(pres.shapes.OVAL, { x: x + (cw - 0.8) / 2, y: cy + 0.3, w: 0.8, h: 0.8, fill: { color: p.hl ? C.accent2 : C.accent1 }, line: { type: "none" }, objectName: `Ikonka foni ${i + 1}` });
    s.addImage({ data: await icon(p.ic, "FFFFFF"), x: x + (cw - 0.42) / 2, y: cy + 0.49, w: 0.42, h: 0.42, objectName: `Ikonka ${i + 1}`, altText: p.h });
    T(s, p.h, { x: x + 0.1, y: cy + 1.25, w: cw - 0.2, h: 0.4, fontSize: 16, bold: true, align: "center", color: p.hl ? C.background1 : C.text1, objectName: `Karta sarlavha ${i + 1}` });
    T(s, p.d, { x: x + 0.15, y: cy + 1.65, w: cw - 0.3, h: 0.65, fontSize: 13, align: "center", valign: "top", color: p.hl ? C.background2 : C.accent6, objectName: `Karta matn ${i + 1}` });
    if (i < parts.length - 1) {
      T(s, i === parts.length - 2 ? "=" : "+", { x: x + cw, y: cy + 0.9, w: gap, h: 0.6, fontSize: 28, bold: true, align: "center", valign: "middle", color: C.accent2, objectName: `Belgi ${i + 1}` });
    }
  }
  s.addNotes("Oddiy qilib aytganda, Live Commerce — bu teleshop, ijtimoiy tarmoq va onlayn do'konning birikmasi. Teleshopdan u mahsulotni ko'rsatib sotishni oldi, ijtimoiy tarmoqdan — jonli muloqotni, onlayn do'kondan esa bir tugmada to'lovni. Natijada xaridor mahsulotni ko'radi, savol beradi va shu zahoti sotib oladi.");

  // ---------- 4. Timeline ----------
  pres.addSection({ title: "Tarix va raqamlar" });
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Tarix va raqamlar" });
  s.addText("Teleshopdan TikTok Shopgacha", { placeholder: "title" });
  const tl = [
    { y: "1980-y.", t: "Teleshop: TV orqali buyurtma (HSN, QVC)" },
    { y: "2016", t: "Alibaba Taobao Live'ni ishga tushirdi" },
    { y: "2018–19", t: "Douyin va Kuaishou efir savdosini ommalashtirdi" },
    { y: "2020", t: "Pandemiya: efir savdosi keskin o'sdi" },
    { y: "2023", t: "TikTok Shop AQShda ishga tushdi" },
    { y: "2025+", t: "AI hostlar 24/7 efir qilmoqda" },
  ];
  const lineY = 2.7, tx0 = 0.6, step = 8.8 / tl.length;
  s.addShape(pres.shapes.LINE, { x: tx0 + step / 2, y: lineY, w: step * (tl.length - 1), h: 0, line: { color: C.accent5, width: 2 }, objectName: "Vaqt chizig'i" });
  tl.forEach((e, i) => {
    const cxm = tx0 + step * i + step / 2;
    const last = i === tl.length - 1;
    s.addShape(pres.shapes.OVAL, { x: cxm - 0.14, y: lineY - 0.14, w: 0.28, h: 0.28, fill: { color: last ? C.accent2 : C.accent1 }, line: { color: C.background1, width: 2 }, objectName: `Nuqta ${i + 1}` });
    T(s, e.y, { x: cxm - step / 2 + 0.05, y: lineY - 0.85, w: step - 0.1, h: 0.5, fontSize: 20, bold: true, align: "center", valign: "bottom", color: last ? C.accent2 : C.accent1, fontFace: THEME.headFontFace, objectName: `Yil ${i + 1}` });
    T(s, e.t, { x: cxm - step / 2 + 0.08, y: lineY + 0.35, w: step - 0.16, h: 1.2, fontSize: 13, align: "center", valign: "top", color: C.text1, objectName: `Voqea ${i + 1}` });
  });
  T(s, "Har bosqichda bitta g'oya: mahsulotni jonli ko'rsatish xaridga ishonch beradi", { x: 0.6, y: 4.4, w: 8.8, h: 0.4, fontSize: 14, italic: true, color: C.accent6, objectName: "Xulosa satri" });
  s.addNotes("G'oya yangi emas: 1980-yillarda AQShda HSN va QVC teleshoplari mahsulotni televizorda ko'rsatib sotgan. 2016-yilda Alibaba Taobao Live'ni ishga tushirdi va formatni smartfonga ko'chirdi. Keyin Douyin va Kuaishou uni ommaviy hodisaga aylantirdi, 2020-yildagi pandemiya esa keskin o'sishga sabab bo'ldi. 2023-yilda TikTok Shop AQShda ishga tushdi, bugun esa sun'iy intellekt hostlari tunu kun efir qilmoqda.");

  // ---------- 5. Numbers ----------
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Tarix va raqamlar" });
  s.addText("Raqamlar gapiradi", { placeholder: "title" });
  s.addChart(pres.charts.BAR, [{ name: "Konversiya", labels: ["An'anaviy e-tijorat", "Live Commerce"], values: [3, 30] }], {
    x: 0.5, y: 1.2, w: 5.0, h: 3.65, barDir: "col",
    chartColors: [HEX.accent5, HEX.accent2],
    showTitle: true, title: "Konversiya darajasi, % (taxminiy)", titleFontSize: 14, titleColor: HEX.dk1, titleFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: '0"%"',
    catAxisLabelColor: HEX.accent6, catAxisLabelFontSize: 13, catAxisLabelFontFace: "+mn-lt",
    valAxisHidden: true, valAxisMaxVal: 36, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showLegend: false, barGapWidthPct: 60,
    objectName: "Konversiya grafigi", altText: "An'anaviy e-tijoratda konversiya taxminan 3%, Live Commerce'da 30% gacha",
  });
  const stats = [
    { n: "−40%", t: "gacha kam qaytarish: xaridor mahsulotni oldindan real ko'radi" },
    { n: "4,9 trln ¥", t: "Xitoyda 2023-yilgi efir savdosi hajmi (GMV) — taxminan 680 mlrd $" },
  ];
  stats.forEach((st, i) => {
    const y = 1.3 + i * 1.8;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.9, y, w: 3.6, h: 1.55, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Statistika karta ${i + 1}` });
    T(s, st.n, { x: 6.15, y: y + 0.15, w: 3.2, h: 0.7, fontSize: 36, bold: true, color: C.accent1, fontFace: THEME.headFontFace, objectName: `Statistika raqam ${i + 1}` });
    T(s, st.t, { x: 6.15, y: y + 0.85, w: 3.2, h: 0.6, fontSize: 13, color: C.text1, valign: "top", objectName: `Statistika izoh ${i + 1}` });
  });
  T(s, "Manbalar: McKinsey, “It's showtime! How live commerce is transforming the shopping experience” (2021); Statista (2023, GMV)", { x: 0.6, y: 4.88, w: 8.8, h: 0.28, fontSize: 9, color: C.accent6, objectName: "Manbalar" });
  s.addNotes("McKinsey tadqiqotiga ko'ra, jonli efirdagi konversiya 30 foizga yaqinlashadi — bu oddiy onlayn do'konga nisbatan taxminan o'n barobar yuqori. Qaytarishlar esa 40 foizgacha kamayadi, chunki xaridor mahsulotni oldindan real holatda ko'radi. Statista ma'lumotiga ko'ra, faqat Xitoyning o'zida 2023-yilda efir savdosi hajmi 4,9 trillion yuandan oshgan.");

  // ---------- 6. Why it works ----------
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Tarix va raqamlar" });
  s.addText("Nega u ishlaydi?", { placeholder: "title" });
  s.addImage({ path: A("phone.png"), x: 0.6, y: 1.3, w: 3.5, h: 3.5, objectName: "Efir rasmi", altText: "Telefonda jonli savdo efiri" });
  const why = [
    { ic: fa.FaHourglassHalf, h: "Cheklangan vaqt", d: "“Faqat efir davomida −30%” — tezroq qaror qilinadi" },
    { ic: fa.FaUsers, h: "Ijtimoiy isbot", d: "Minglab tomoshabin, layklar va “hozir sotib olindi” xabarlari" },
    { ic: fa.FaHandshake, h: "Ishonch", d: "Mahsulot tahrirsiz ko'rsatiladi, savolga darhol javob" },
    { ic: fa.FaStar, h: "Shoppertainment", d: "Xarid + shou: o'yinlar, sovg'alar, mehmonlar" },
  ];
  for (let i = 0; i < why.length; i++) {
    const w = why[i];
    const x = 4.5 + (i % 2) * 2.55, y = 1.3 + Math.floor(i / 2) * 1.8;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.4, h: 1.65, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: `Sabab karta ${i + 1}` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.2, w: 0.5, h: 0.5, fill: { color: C.accent2 }, line: { type: "none" }, objectName: `Sabab ikonka foni ${i + 1}` });
    s.addImage({ data: await icon(w.ic, "FFFFFF"), x: x + 0.32, y: y + 0.32, w: 0.26, h: 0.26, objectName: `Sabab ikonka ${i + 1}`, altText: w.h });
    T(s, w.h, { x: x + 0.8, y: y + 0.2, w: 1.5, h: 0.5, fontSize: 14, bold: true, color: C.text1, valign: "middle", objectName: `Sabab sarlavha ${i + 1}` });
    T(s, w.d, { x: x + 0.2, y: y + 0.82, w: 2.05, h: 0.75, fontSize: 12, color: C.accent6, valign: "top", objectName: `Sabab matn ${i + 1}` });
  }
  s.addNotes("Live Commerce'ning muvaffaqiyati psixologiyaga asoslangan. Birinchidan, cheklangan vaqtli aksiya — 'faqat efir davomida' — qarorni tezlashtiradi. Ikkinchidan, ijtimoiy isbot: minglab odam tomosha qilayotganini ko'rsangiz, mahsulotga ishonch ortadi. Uchinchidan, mahsulot tahrirsiz, real vaqtda ko'rsatiladi. To'rtinchidan, bu shunchaki xarid emas, balki shou — inglizcha 'shoppertainment' deyiladi.");

  // ---------- 7. Formats table ----------
  pres.addSection({ title: "Formatlar va kelajak" });
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Formatlar va kelajak" });
  s.addText("Kim efir qiladi?", { placeholder: "title" });
  const hd = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.accent1 }, fontSize: 14 } });
  const cell = (t, o = {}) => ({ text: t, options: { color: C.text1, fontSize: 13, ...o } });
  const rows = [
    [hd("Format"), hd("Kim boshlaydi"), hd("Kuchli tomoni"), hd("Misol")],
    [cell("Brend efiri", { bold: true }), cell("Kompaniya xodimi"), cell("Mahsulotni chuqur biladi"), cell("Yangi kolleksiya taqdimoti")],
    [cell("Influencer efiri", { bold: true }), cell("Blogger"), cell("Katta auditoriya va ishonch"), cell("Li Jiaqi — “lab bo'yog'i qiroli”")],
    [cell("Marketplace efiri", { bold: true }), cell("Platforma sotuvchilari"), cell("Xarid tugmasi efir ichida"), cell("Taobao Live, TikTok Shop")],
    [cell("AI-host efiri", { bold: true, color: C.accent2 }), cell("Virtual boshlovchi"), cell("24/7, arzon, ko'p tilli"), cell("Xitoydagi raqamli hostlar")],
  ];
  rows.forEach((r, i) => { if (i > 0 && i % 2 === 0) r.forEach((c) => (c.options.fill = { color: C.background2 })); });
  s.addTable(rows, {
    x: 0.6, y: 1.35, w: 8.8, colW: [1.9, 2.0, 2.4, 2.5], rowH: 0.68,
    valign: "middle", margin: [0.05, 0.12, 0.05, 0.12],
    border: { type: "solid", pt: 0.75, color: HEX.lt2 },
    objectName: "Formatlar jadvali",
  });
  s.addNotes("Efirni kim olib borishiga qarab to'rt formatni ajratamiz. Brend efirlarini kompaniyaning o'zi olib boradi. Influencer efirlari bloggerlar auditoriyasiga tayanadi — masalan, Xitoyda Li Jiaqi 'lab bo'yog'i qiroli' laqabini olgan. Marketplace efirlarida xarid tugmasi to'g'ridan-to'g'ri efir ichida bo'ladi. Eng yangi format — AI-host efirlari, ularni keyingi slaydda ko'ramiz.");

  // ---------- 8. AI hosts (dark) ----------
  s = pres.addSlide({ masterName: "CONTENT_DARK", sectionTitle: "Formatlar va kelajak" });
  s.addImage({ path: A("ai_host.png"), x: 5.1, y: 0, w: 4.9, h: 5.625, sizing: { type: "cover", w: 4.9, h: 5.625 }, objectName: "AI host rasmi", altText: "Virtual AI boshlovchi gologramma mahsulotni ko'rsatmoqda" });
  s.addText("AI hostlar", { placeholder: "title" });
  T(s, "Efir hech qachon to'xtamaydi", { x: 0.6, y: 1.1, w: 4.3, h: 0.4, fontSize: 18, color: C.accent3, italic: true, objectName: "Kichik sarlavha" });
  const ai = [
    { ic: fa.FaClock, t: "24/7 efir — tunda ham sotuv davom etadi" },
    { ic: fa.FaGlobe, t: "Bir vaqtda o'nlab tillarda gapiradi" },
    { ic: fa.FaCoins, t: "Studiya va boshlovchiga xarajat kamayadi" },
    { ic: fa.FaExclamationTriangle, t: "Lekin: ishonch, deepfake va shaffoflik masalasi", warn: true },
  ];
  for (let i = 0; i < ai.length; i++) {
    const a = ai[i], y = 1.8 + i * 0.78;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.52, h: 0.52, fill: { color: a.warn ? C.accent3 : C.accent1 }, line: { type: "none" }, objectName: `AI ikonka foni ${i + 1}` });
    s.addImage({ data: await icon(a.ic, a.warn ? HEX.dk2 : "FFFFFF"), x: 0.73, y: y + 0.13, w: 0.26, h: 0.26, objectName: `AI ikonka ${i + 1}`, altText: "ikonka" });
    T(s, a.t, { x: 1.3, y, w: 3.6, h: 0.52, fontSize: 15, color: a.warn ? C.accent3 : C.background1, valign: "middle", objectName: `AI band ${i + 1}` });
  }
  s.addNotes("Sun'iy intellekt hostlari — raqamli boshlovchilar — Live Commerce'ning keyingi bosqichi. Ular tunu kun charchamasdan efir qiladi, bir vaqtning o'zida ko'plab tillarda gapiradi va studiya xarajatlarini kamaytiradi. Ammo muammolar ham bor: xaridor kim bilan gaplashayotganini bilishi kerak, deepfake va aldov xavfi esa shaffoflik qoidalarini talab qiladi.");

  // ---------- 9. 5 steps ----------
  pres.addSection({ title: "Amaliyot" });
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Amaliyot" });
  s.addText("Tadbirkor uchun 5 qadam", { placeholder: "title" });
  T(s, "O'zbekistonda Instagram va Telegram orqali sotayotgan minglab kichik bizneslar uchun jonli efir — tayyor auditoriyaga eng arzon kirish yo'li.", { x: 0.6, y: 1.15, w: 8.8, h: 0.6, fontSize: 15, color: C.text1, valign: "top", objectName: "Kirish matni" });
  const steps = [
    { h: "Mahsulot", d: "Ko'rsatib tushuntirsa bo'ladigan tovarni tanlang" },
    { h: "Ssenariy", d: "Ilmoq → namoyish → aksiya → “hozir oling”" },
    { h: "Jihoz", d: "Telefon, chiroq va mikrofon yetarli" },
    { h: "Taklif", d: "Faqat efir uchun chegirma yoki sovg'a" },
    { h: "Tahlil", d: "Tomoshabin → buyurtma konversiyasini o'lchang" },
  ];
  const sw = 1.62, sg = 0.175;
  steps.forEach((st, i) => {
    const x = 0.6 + i * (sw + sg), y = 2.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: sw, h: 2.6, rectRadius: 0.12, fill: { color: i === 4 ? C.accent1 : C.background2 }, line: { type: "none" }, objectName: `Qadam karta ${i + 1}` });
    T(s, String(i + 1).padStart(2, "0"), { x: x + 0.18, y: y + 0.18, w: 1.2, h: 0.6, fontSize: 30, bold: true, color: i === 4 ? C.accent3 : C.accent2, fontFace: THEME.headFontFace, objectName: `Qadam raqam ${i + 1}` });
    T(s, st.h, { x: x + 0.18, y: y + 0.85, w: sw - 0.3, h: 0.4, fontSize: 16, bold: true, color: i === 4 ? C.background1 : C.text1, objectName: `Qadam nomi ${i + 1}` });
    T(s, st.d, { x: x + 0.18, y: y + 1.3, w: sw - 0.32, h: 1.15, fontSize: 13, color: i === 4 ? C.background2 : C.accent6, valign: "top", objectName: `Qadam matn ${i + 1}` });
  });
  s.addNotes("Amaliy qism: O'zbekistonda kichik biznes jonli efirni qanday boshlashi mumkin? Birinchi — namoyish qilsa bo'ladigan mahsulot tanlash. Ikkinchi — ssenariy: e'tiborni tortuvchi boshlanish, namoyish, aksiya va chaqiriq. Uchinchi — qimmat studiya shart emas, telefon va yaxshi yorug'lik yetarli. To'rtinchi — faqat efir uchun maxsus taklif. Beshinchi — natijani o'lchash: nechta tomoshabin buyurtmaga aylandi.");

  // ---------- 10. Risks vs solutions ----------
  s = pres.addSlide({ masterName: "CONTENT_LIGHT", sectionTitle: "Amaliyot" });
  s.addText("Xavflar va yechimlar", { placeholder: "title" });
  const cols = [
    { h: "Xavflar", ic: fa.FaExclamationCircle, col: C.accent2, items: ["Bo'rttirilgan reklama va soxta sharhlar", "Sifat efirdagidan farq qilishi", "Impulsiv xarid → ko'p qaytarish", "To'lov va shaxsiy ma'lumotlar xavfsizligi"] },
    { h: "Yechimlar", ic: fa.FaCheckCircle, col: C.accent4, items: ["Halol demo va mahsulot sertifikatlari", "Aniq va oddiy qaytarish siyosati", "Iste'molchi huquqlari qonunchiligiga rioya", "Ishonchli to'lov tizimlari (Click, Payme)"] },
  ];
  for (let ci = 0; ci < cols.length; ci++) {
    const c = cols[ci], x = 0.6 + ci * 4.55;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.3, w: 4.25, h: 3.1, rectRadius: 0.12, fill: { color: C.background2 }, line: { type: "none" }, objectName: `${c.h} kartasi` });
    s.addImage({ data: await icon(c.ic, ci === 0 ? HEX.accent2 : HEX.accent4), x: x + 0.3, y: 1.55, w: 0.42, h: 0.42, objectName: `${c.h} ikonka`, altText: c.h });
    T(s, c.h, { x: x + 0.85, y: 1.52, w: 3, h: 0.48, fontSize: 20, bold: true, color: c.col, valign: "middle", objectName: `${c.h} sarlavha` });
    T(s, c.items.map((t, k) => ({ text: t, options: { bullet: true, breakLine: k < c.items.length - 1 } })), { x: x + 0.3, y: 2.2, w: 3.75, h: 2.5, fontSize: 15, color: C.text1, valign: "top", paraSpaceAfter: 8, objectName: `${c.h} ro'yxati` });
  }
  s.addNotes("Har qanday kuchli vosita kabi, Live Commerce'ning ham xavflari bor: bo'rttirilgan reklama, efirdagidan farq qiladigan sifat, hissiyotga berilib qilingan xaridlar va to'lov xavfsizligi. Yechimlar ham aniq: halol namoyish, oddiy qaytarish siyosati, iste'molchi huquqlariga rioya va ishonchli to'lov tizimlaridan foydalanish.");

  // ---------- 11. Conclusion ----------
  pres.addSection({ title: "Xulosa" });
  s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Xulosa" });
  s.addImage({ path: A("hero.png"), x: 0, y: 0, w: 10, h: 5.625, objectName: "Fon rasmi", altText: "Jonli efir studiyasi" });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 5.625, fill: { color: C.text2, transparency: 22 }, line: { type: "none" }, objectName: "Qoraytirish qatlami" });
  T(s, "Xulosa", { x: 0.6, y: 0.5, w: 8.8, h: 0.75, fontSize: 40, bold: true, color: C.background1, fontFace: THEME.headFontFace, objectName: "Xulosa sarlavha" });
  const take = [
    "Live Commerce = video + jonli muloqot + bir tugmada xarid",
    "U konversiyani oshiradi va qaytarishlarni kamaytiradi",
    "Kelajak — AI hostlar va mahalliy marketplace'lardagi efirlar",
  ];
  take.forEach((t, i) => {
    const y = 1.55 + i * 0.75;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.5, h: 0.5, fill: { color: C.accent2 }, line: { type: "none" }, objectName: `Xulosa raqam foni ${i + 1}` });
    T(s, String(i + 1), { x: 0.6, y, w: 0.5, h: 0.5, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", objectName: `Xulosa raqam ${i + 1}` });
    T(s, t, { x: 1.3, y, w: 7, h: 0.5, fontSize: 18, color: C.background1, valign: "middle", objectName: `Xulosa band ${i + 1}` });
  });
  T(s, "Rahmat! Savollaringiz bo'lsa, marhamat", { x: 0.6, y: 4.3, w: 8, h: 0.6, fontSize: 24, bold: true, color: C.accent3, fontFace: THEME.headFontFace, objectName: "Rahmat" });
  s.addNotes("Xulosa qilib aytganda: Live Commerce — video, jonli muloqot va bir tugmada xaridning birikmasi. U konversiyani sezilarli oshiradi va qaytarishlarni kamaytiradi. Kelajakda esa sun'iy intellekt hostlari va mahalliy marketplace'lardagi efirlar bu formatni yanada kengaytiradi. E'tiboringiz uchun rahmat, savollaringizga javob berishga tayyorman.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  require("child_process").execFileSync("python3", [path.join(__dirname, "autoplay.py"), OUT, "slide2.xml", "Kirish videosi", "30160"], { stdio: "inherit" });
  console.log("Wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
