const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");
const { imageSize } = require("image-size");

if (!process.env.SKILL) throw new Error("SKILL is not set: path to the pptx skill folder with scripts/apply_theme.js");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const S = __dirname;
const IMG = (n) => path.join(S, "img", n);
const OUT = process.env.OUT || path.join(S, "..", "Презентация_ИНГОССЕМЬЯ.pptx");
const R = JSON.parse(fs.readFileSync(process.env.RES || path.join(S, "..", "..", "05_finance", "results.json"), "utf8"));

// palette
const P = "1E3FD8", D = "0A2A8A", L = "8FA3F0", A = "B8F02A", T = "151B33";
const TINT = "EEF1FD", MUTED = "5A6280", LINE = "D5DBF5", RED = "C0392B";
const F = "Arial";

const THEME = {
  name: "Ingossemya",
  headFontFace: F, bodyFontFace: F,
  colors: { dk1: T, lt1: "FFFFFF", dk2: D, lt2: TINT, accent1: P, accent2: D, accent3: L,
    accent4: A, accent5: MUTED, accent6: RED, hlink: P, folHlink: D },
};

// layout grid, inches (slide is 10 x 5.625)
const X0 = 0.5, CW = 9.0, Y0 = 1.25, YB = 5.0, G = 0.2;
// n equal columns inside [x, x + w] with gap between them
function cols(n, gap = G, x = X0, w = CW) {
  const cw = (w - gap * (n - 1)) / n;
  return Array.from({ length: n }, (_, i) => ({ x: x + i * (cw + gap), w: cw }));
}

// non-breaking spaces inside numbers and between number and unit
function nb(s) {
  if (typeof s !== "string") return s;
  let r = s;
  for (let i = 0; i < 2; i++) r = r.replace(/(\d) (\d{3})(?!\d)/g, "$1 $2");
  r = r.replace(/(\d) (млн|млрд|тыс|руб|мес|семей|семьи|семья|минут|минуты|лет|раз|человек|прогонов|интервью|анкет|из|%)/g, "$1 $2");
  r = r.replace(/(H[1-6]) /g, "$1 ");
  return r;
}
function runs(arr) { return arr.map((x) => typeof x === "string" ? { text: nb(x) } : { text: nb(x.text), options: x.options || {} }); }

const fmt = (v, d = 1) => {
  const neg = v < 0; v = Math.abs(v);
  let s = v.toFixed(d).replace(".", ",");
  let [i, f] = s.split(",");
  i = i.replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  return (neg ? "-" : "") + i + (f ? "," + f : "");
};
const mln = (v, d = 1) => fmt(v / 1e6, d);

// values from results.json
const B = R["Сценарии"]["Базовый"], PS = R["Сценарии"]["Пессимистичный"], OP = R["Сценарии"]["Оптимистичный"];
const UB = R["Юнит-экономика"]["Базовый"], V1 = R["Версия 1 без триажа (для сравнения)"];
const MC = R["Монте-Карло"], MK = R["Рынок"], BUD = R["Бюджет MVP по категориям"];
const VET = R["Льготы ветеранам"], U1 = V1["Юнит-экономика"];
const MIN_B = R["Минуты координатора на семью"]["Базовый"];
const LOAD_B = fmt(UB["Семей на координатора: Базовый"], 0);
const MVP_M = mln(R["Бюджет MVP, руб"], 2);
const PEAK_AFTER = mln(B["Пиковая потребность в деньгах, руб"] - R["Бюджет MVP, руб"]);

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.theme = { headFontFace: F, bodyFontFace: F };
pres.title = "ИНГОССЕМЬЯ";
pres.author = "Команда проекта ИНГОССЕМЬЯ";
pres.subject = "IngoSight 2026, отборочный этап";

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: X0, y: 0.28, w: CW, h: 0.86,
      fontFace: F, fontSize: 21, bold: true, color: D, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 9.05, y: 5.22, w: 0.45, h: 0.25, fontFace: F, fontSize: 8, color: MUTED, align: "right" },
});
pres.defineSlideMaster({ title: "DARK", background: { color: D }, objects: [] });

function addSlide(section, master = "CONTENT") {
  return pres.addSlide({ masterName: master, sectionTitle: section });
}
function title(s, t) { s.addText(nb(t), { placeholder: "title" }); }
function text(s, t, o) {
  const opts = Object.assign({ fontFace: F, fontSize: 12, color: T, margin: 0, valign: "top", isTextBox: true }, o);
  s.addText(Array.isArray(t) ? runs(t) : nb(t), opts);
}
function src(s, t) {
  text(s, "Источники: " + t, { x: X0, y: 5.18, w: 8.5, h: 0.3, fontSize: 7.5, color: MUTED, valign: "bottom" });
}
function box(s, x, y, w, h, fill, line) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08,
    fill: fill ? { color: fill } : { type: "none" }, line: line ? { color: line, width: 1 } : { type: "none" } });
}
// picture scaled to fit inside the box, centered horizontally, aligned to the top
function pic(s, file, x, y, w, h) {
  const { width, height } = imageSize(IMG(file));
  const k = Math.min(w / width, h / height), pw = width * k, ph = height * k;
  const px = x + (w - pw) / 2;
  s.addImage({ path: IMG(file), x: px, y, w: pw, h: ph });
  return { x: px, y, w: pw, h: ph };
}
const ICON_COLOR = { [P]: "blue", [D]: "navy", [A]: "lime", FFFFFF: "white" };
// lucide icon in a filled circle of diameter d
function badge(s, name, x, y, d, fill = TINT, color = P) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
  const k = d * 0.56;
  s.addImage({ path: IMG(`icons/${name}_${ICON_COLOR[color]}.png`), x: x + (d - k) / 2, y: y + (d - k) / 2, w: k, h: k });
}
const TAGS = {
  "факт": { fill: D, color: "FFFFFF", w: 0.42 },
  "расчет": { fill: L, color: T, w: 0.52 },
  "допущение": { fill: A, color: T, w: 0.74 },
};
function tag(s, x, y, kind) {
  const t = TAGS[kind];
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: t.w, h: 0.2, rectRadius: 0.1, fill: { color: t.fill }, line: { type: "none" } });
  text(s, kind, { x, y, w: t.w, h: 0.2, fontSize: 8, color: t.color, align: "center", valign: "middle", bold: true });
}
function chip(s, x, y, w, label, fill = A, color = T) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.26, rectRadius: 0.13, fill: { color: fill }, line: { type: "none" } });
  text(s, label, { x, y, w, h: 0.26, fontSize: 9, bold: true, color, align: "center", valign: "middle" });
}
function bullets(items, o = {}) {
  return items.map((t, j) => ({ text: nb(t), options: Object.assign({ bullet: { indent: 10 }, breakLine: j < items.length - 1, paraSpaceAfter: 4 }, o) }));
}

// ---------------- 1. Title ----------------
pres.addSection({ title: "Титул" });
{
  const s = addSlide("Титул", "DARK");
  chip(s, 0.6, 0.55, 1.55, "IngoSight 2026");
  text(s, "Отборочный этап", { x: 2.3, y: 0.55, w: 3, h: 0.26, fontSize: 11, color: L, valign: "middle" });
  text(s, "ИНГОССЕМЬЯ", { x: 0.6, y: 1.35, w: 5.6, h: 1.0, fontSize: 46, bold: true, color: "FFFFFF", valign: "middle" });
  text(s, "Координатор Ингосстраха организует медпомощь пожилому родителю, а взрослый ребенок в тот же день знает, что сделано", {
    x: 0.6, y: 2.45, w: 5.4, h: 1.2, fontSize: 18, color: "FFFFFF" });
  s.addShape(pres.shapes.LINE, { x: 0.6, y: 3.85, w: 5.4, h: 0, line: { color: P, width: 1 } });
  text(s, [{ text: "Направление ", options: { color: L } }, { text: "«Здоровье и благополучие»", options: { color: "FFFFFF", bold: true } }], {
    x: 0.6, y: 4.05, w: 5.4, h: 0.35, fontSize: 14 });
  text(s, "Конкурс им. В. И. Щербакова, Ингосстрах", { x: 0.6, y: 4.45, w: 5.4, h: 0.3, fontSize: 11, color: L });
  s.addShape(pres.shapes.OVAL, { x: 6.55, y: 0.85, w: 3.3, h: 3.3, fill: { color: P, transparency: 30 }, line: { type: "none" } });
  pic(s, "ph_feed.png", 6.6, 0.3, 3.0, 5.05);
  s.addNotes("Мы команда проекта ИНГОССЕМЬЯ. Взрослые дети организуют медпомощь пожилым родителям на расстоянии, и между «справляюсь сам» и сиделкой сейчас ничего нет. Наш сервис закрывает этот разрыв силами координаторов и медицинской инфраструктуры Ингосстраха. Дальше покажем, что подтверждают данные, как работает сервис и сколько стоит проверка.");
}

// ---------------- 2. Segment ----------------
pres.addSection({ title: "Сегмент и проблема" });
{
  const s = addSlide("Сегмент и проблема");
  title(s, "Покупает работающий взрослый ребенок, который организует медпомощь родителю издалека");
  text(s, "Взрослые 35-60 лет, у которых родитель 70+ с хронической болезнью живет отдельно", { x: X0, y: Y0, w: CW, h: 0.3, fontSize: 12, color: MUTED });
  const Q = [
    ["users", "Пользователь", "Ребенок в чате, родитель по обычному телефону"],
    ["user", "Кто решает", "Взрослый ребенок, иногда с братьями и сестрами"],
    ["wallet", "Кто платит", "Ребенок (B2C) или работодатель через ДМС (B2B2C)"],
    ["zap", "Триггер", "Диагноз, госпитализация, первый упущенный эпизод"],
  ];
  const [left, right] = cols(2, G, X0, 5.8);
  const top = 1.7, th = 1.38;
  Q.forEach(([ic, h, b], i) => {
    const c = i % 2 ? right : left, y = top + Math.floor(i / 2) * (th + G);
    box(s, c.x, y, c.w, th, TINT);
    badge(s, ic, c.x + 0.18, y + 0.18, 0.5, "FFFFFF", P);
    text(s, h, { x: c.x + 0.82, y: y + 0.18, w: c.w - 1.0, h: 0.5, fontSize: 13, bold: true, color: D, valign: "middle" });
    text(s, b, { x: c.x + 0.18, y: y + 0.8, w: c.w - 0.36, h: 0.5, fontSize: 11 });
  });
  const jx = X0 + 5.8 + G, jw = CW - 5.8 - G, jh = th * 2 + G;
  box(s, jx, top, jw, jh, D);
  text(s, "Задача (JTBD)", { x: jx + 0.2, y: top + 0.18, w: jw - 0.4, h: 0.3, fontSize: 12, bold: true, color: A });
  text(s, [
    { text: "Когда ", options: { bold: true, color: A } }, "у родителя что-то со здоровьем, ",
    { text: "я хочу ", options: { bold: true, color: A } }, "быстро организовать помощь, не бросая работу, ",
    { text: "чтобы ", options: { bold: true, color: A } }, "не узнать о проблеме слишком поздно",
  ], { x: jx + 0.2, y: top + 0.6, w: jw - 0.4, h: jh - 0.8, fontSize: 13, color: "FFFFFF" });
  const by = top + jh + 0.18;
  badge(s, "circle-alert", X0, by, 0.4, TINT, P);
  text(s, [{ text: "Сейчас: ", options: { bold: true } }, "звонки каждый день, Госуслуги за родителя, поездки «на всякий случай». Родитель скрывает симптомы, чтобы не беспокоить"],
    { x: X0 + 0.55, y: by, w: CW - 0.55, h: 0.4, fontSize: 10.5, valign: "middle" });
  src(s, "концепция и пояснительная записка команды; признаки сильной проблемы проверяем в интервью 23-25.10");
  s.addNotes("Сегмент мы описываем по поведению, а не по возрасту: это человек, который сам организует медпомощь родителю, хотя живет отдельно. Пользователей двое, ребенок в чате и родитель по телефону, а решает и платит ребенок или его работодатель. Триггер почти всегда событие: диагноз, госпитализация или первый случай, когда о проблеме узнали поздно.");
}

// ---------------- 3. Problem evidence ----------------
{
  const s = addSlide("Сегмент и проблема");
  title(s, "Почти 9 из 10 нуждающихся в уходе остаются на семье: госсистема охватывает 11%");
  const N = [
    ["25,8", "млн людей 65+ в России, 17,7% населения", "факт", "S1"],
    ["19%", "россиян ухаживают за близким раз в неделю и чаще", "факт", "S2"],
    ["54%", "старше трудоспособного возраста хронически болеют", "факт", "S3"],
    ["89%", "нуждающихся вне системы долговременного ухода", "расчет", "S4"],
    [">10%", "ухаживающих сократили работу или уволились", "факт", "S5"],
  ];
  const ch = 2.1;
  cols(5).forEach((c, i) => {
    const dark = i === 3;
    box(s, c.x, Y0, c.w, ch, dark ? D : TINT);
    text(s, N[i][0], { x: c.x + 0.14, y: Y0 + 0.12, w: c.w - 0.28, h: 0.6, fontSize: 28, bold: true, color: dark ? A : P, valign: "middle" });
    text(s, N[i][1], { x: c.x + 0.14, y: Y0 + 0.8, w: c.w - 0.28, h: 0.9, fontSize: 10.5, color: dark ? "FFFFFF" : T });
    tag(s, c.x + 0.14, Y0 + ch - 0.32, N[i][2]);
    text(s, N[i][3], { x: c.x + c.w - 0.5, y: Y0 + ch - 0.32, w: 0.36, h: 0.2, fontSize: 8, color: dark ? L : MUTED, align: "right", valign: "middle" });
  });
  const hy = Y0 + ch + 0.3;
  text(s, "Что не сходится с гипотезой и как мы это учли", { x: X0, y: hy, w: CW, h: 0.3, fontSize: 12.5, bold: true, color: D });
  const C = [
    ["route", "Рынок не растет", "Ухаживают 19% против 21% в 2021. Считаем по сегменту"],
    ["circle-alert", "Нет данных по сегменту", "Долю детей, живущих отдельно, берем 60% и проверяем анкетой"],
    ["ban", "Старые цифры убрали", "Опрос НАФИ 2016 года о тратах на уход не используем"],
  ];
  cols(3).forEach((c, i) => {
    const y = hy + 0.45;
    badge(s, C[i][0], c.x, y, 0.42, TINT, P);
    text(s, C[i][1], { x: c.x + 0.55, y, w: c.w - 0.55, h: 0.42, fontSize: 11, bold: true, color: P, valign: "middle" });
    text(s, C[i][2], { x: c.x, y: y + 0.5, w: c.w, h: 0.5, fontSize: 10 });
  });
  src(s, "S1 Росстат через Демоскоп, 01.01.2025; S2 ВЦИОМ, 18.09.2026, n=1600; S3 слова министра Мурашко, 07.2026, пересказ СМИ; S4 Минтруд через АСИ, 21.07.2026 (100% - 11%); S5 ВШЭ, 2020");
  s.addNotes("Здесь только перепроверенные открытые данные, у каждого числа метка: факт из источника или наш расчет. Главная цифра: система долговременного ухода охватывает 11% нуждающихся, остальное тянут семьи. Мы честно показываем и то, что против нас: рынок ухаживающих не растет, а доли детей, живущих отдельно, в открытых данных нет, поэтому это допущение для анкеты.");
}

// ---------------- 4. Interviews template ----------------
{
  const s = addSlide("Сегмент и проблема");
  title(s, "Интервью и опрос покажут, проходит ли проблема порог H1");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: X0, y: Y0, w: 4.6, h: 0.4, rectRadius: 0.08, fill: { color: A }, line: { type: "none" } });
  text(s, "Заполняется по итогам интервью 23-25.10", { x: X0 + 0.15, y: Y0, w: 4.4, h: 0.4, fontSize: 13, bold: true, color: T, valign: "middle" });
  text(s, "Порог H1: 60%+ тратят 3+ часа в неделю и имели упущенный эпизод за полгода", { x: X0 + 4.8, y: Y0, w: CW - 4.8, h: 0.4, fontSize: 10, color: MUTED, valign: "middle" });
  const [c1, c2, c3] = cols(3);
  const top = Y0 + 0.6, h = YB - top;
  box(s, c1.x, top, c1.w, h, null, LINE);
  badge(s, "users", c1.x + 0.2, top + 0.18, 0.42, TINT, P);
  text(s, "Респонденты", { x: c1.x + 0.75, y: top + 0.18, w: c1.w - 0.9, h: 0.42, fontSize: 12.5, bold: true, color: D, valign: "middle" });
  [["__", "интервью, план 15+"], ["__", "анкет, план 150+"]].forEach(([b, l], i) => {
    const y = top + 0.8 + i * 1.0;
    text(s, b, { x: c1.x + 0.2, y, w: c1.w - 0.4, h: 0.5, fontSize: 30, bold: true, color: L });
    text(s, l, { x: c1.x + 0.2, y: y + 0.52, w: c1.w - 0.4, h: 0.3, fontSize: 10.5 });
  });
  box(s, c2.x, top, c2.w, h, null, LINE);
  badge(s, "list-checks", c2.x + 0.2, top + 0.18, 0.42, TINT, P);
  text(s, "3 повторяющихся паттерна", { x: c2.x + 0.75, y: top + 0.18, w: c2.w - 0.9, h: 0.42, fontSize: 12.5, bold: true, color: D, valign: "middle" });
  [1, 2, 3].forEach((n, i) => {
    const y = top + 0.85 + i * 0.85;
    s.addShape(pres.shapes.OVAL, { x: c2.x + 0.2, y, w: 0.34, h: 0.34, fill: { color: P }, line: { type: "none" } });
    text(s, String(n), { x: c2.x + 0.2, y, w: 0.34, h: 0.34, fontSize: 11, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    text(s, "Паттерн: ___\nУ __ из __ респондентов", { x: c2.x + 0.68, y: y - 0.02, w: c2.w - 0.85, h: 0.6, fontSize: 10, color: MUTED });
  });
  const qh = (h - G) / 2;
  box(s, c3.x, top, c3.w, qh, TINT);
  badge(s, "message-circle", c3.x + 0.2, top + 0.18, 0.42, "FFFFFF", P);
  text(s, "Цитаты", { x: c3.x + 0.75, y: top + 0.18, w: c3.w - 0.9, h: 0.42, fontSize: 12.5, bold: true, color: D, valign: "middle" });
  text(s, "«___»\n«___»", { x: c3.x + 0.2, y: top + 0.75, w: c3.w - 0.4, h: 0.6, fontSize: 11, italic: true, color: MUTED });
  box(s, c3.x, top + qh + G, c3.w, qh, TINT);
  badge(s, "circle-alert", c3.x + 0.2, top + qh + G + 0.18, 0.42, "FFFFFF", P);
  text(s, "Что против гипотезы", { x: c3.x + 0.75, y: top + qh + G + 0.18, w: c3.w - 0.9, h: 0.42, fontSize: 12.5, bold: true, color: RED, valign: "middle" });
  text(s, "___\nЧто меняем: ___", { x: c3.x + 0.2, y: top + qh + G + 0.75, w: c3.w - 0.4, h: 0.6, fontSize: 10.5, color: MUTED });
  src(s, "собственное полевое исследование 23-25.10.2026: интервью и онлайн-анкета; результаты также заменят допущения 60% и 70% в оценке рынка");
  s.addNotes("Это единственный слайд, который мы заполним после полевого исследования 23-25 октября. Мы заранее зафиксировали порог H1 и структуру: сколько людей опросили, три повторяющихся паттерна, цитаты и то, что противоречит гипотезе. Если порог не пройден, мы покажем это здесь прямо и поменяем решение.");
}

// ---------------- 5. Alternatives ----------------
pres.addSection({ title: "Рынок и решение" });
{
  const s = addSlide("Рынок и решение");
  title(s, "Между «справляюсь сам» и сиделкой ничего нет");
  const ALT = [
    ["user", "Сам", "0 руб.", "но ценой времени и отгулов", "Все держится на одном человеке"],
    ["users", "Соседи и родные", "подарки", "и ответные услуги", "Нет регулярности"],
    ["video", "Телемедицина", "разово", "консультация", "Родитель не пользуется приложениями"],
    ["shield-check", "Полис ухода", "8-25 тыс.", "руб. в год", "Платит только после страхового случая"],
    ["heart-handshake", "ИНГОССЕМЬЯ", "2 490", "руб. в месяц", "Организует помощь заранее"],
    ["stethoscope", "ДМС родителю", "43-78 тыс.", "руб. в год", "Оплачивает лечение, но не организует его"],
    ["bed-double", "Сиделка", "30-72 тыс.", "руб. в месяц", "Дорого и избыточно"],
  ];
  const C = cols(7, 0.12);
  s.addShape(pres.shapes.LINE, { x: X0, y: Y0 + 0.12, w: CW, h: 0, line: { color: L, width: 1.5, endArrowType: "triangle" } });
  text(s, "дешевле", { x: X0, y: Y0 - 0.12, w: 1.2, h: 0.2, fontSize: 8.5, color: MUTED });
  text(s, "дороже", { x: X0 + CW - 1.2, y: Y0 - 0.12, w: 1.2, h: 0.2, fontSize: 8.5, color: MUTED, align: "right" });
  const top = Y0 + 0.35, h = 2.7;
  ALT.forEach(([ic, name, price, unit, gap], i) => {
    const c = C[i], hi = i === 4;
    box(s, c.x, top, c.w, h, hi ? D : TINT);
    const d = 0.56;
    badge(s, ic, c.x + (c.w - d) / 2, top + 0.18, d, hi ? A : "FFFFFF", hi ? D : P);
    text(s, name, { x: c.x + 0.04, y: top + 0.85, w: c.w - 0.08, h: 0.42, fontSize: hi ? 9.5 : 10.5, bold: true, color: hi ? "FFFFFF" : D, align: "center", valign: "middle" });
    text(s, price, { x: c.x + 0.06, y: top + 1.3, w: c.w - 0.12, h: 0.36, fontSize: 14, bold: true, color: hi ? A : P, align: "center", valign: "middle" });
    text(s, unit, { x: c.x + 0.06, y: top + 1.66, w: c.w - 0.12, h: 0.36, fontSize: 8.5, color: hi ? L : MUTED, align: "center" });
    s.addShape(pres.shapes.LINE, { x: c.x + 0.2, y: top + 2.08, w: c.w - 0.4, h: 0, line: { color: hi ? P : LINE, width: 1 } });
    text(s, gap, { x: c.x + 0.08, y: top + 2.18, w: c.w - 0.16, h: 0.8, fontSize: 9, color: hi ? "FFFFFF" : T, align: "center" });
  });
  const by = top + h + G;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: X0, y: by, w: CW, h: YB - by, rectRadius: 0.08, fill: { color: A }, line: { type: "none" } });
  text(s, [{ text: "В 12-29 раз дешевле сиделки. ", options: { bold: true } }, "Только ИНГОССЕМЬЯ организует помощь до страхового случая"],
    { x: X0 + 0.2, y: by, w: CW - 0.4, h: YB - by, fontSize: 11.5, color: T, valign: "middle" });
  src(s, "S6 Peterburg2, 14.08.2026; S7 360.ru, 09.2026; S9 ingos.ru, ДМС 43 050-78 120 руб. в год; S8 Коммерсантъ, 02.09.2026. Цена подписки: допущение, кратность: расчет");
  s.addNotes("Мы начинали с карты поведения, а не со списка конкурентов. Сегодня ребенок либо все делает сам, либо нанимает сиделку за 30-72 тысячи в месяц. ДМС и полис ухода платят за лечение или после наступления случая, но никто не организует помощь для самостоятельного родителя. Этот разрыв и закрывает ИНГОССЕМЬЯ.");
}

// ---------------- 6. TAM SAM SOM ----------------
{
  const s = addSlide("Рынок и решение");
  title(s, "Чтобы выйти на " + fmt(MK["SOM на 60 мес, семей (базовый сценарий)"] / 1e3, 1) + " тыс. семей, нужно " + fmt(MK["Доля SOM от SAM, %"], 2) + "% целевого рынка");
  // nested circles share the bottom edge
  const cx = 2.6, bottom = 4.75, RINGS = [[3.5, TINT, LINE], [2.3, L, null], [1.05, D, null]];
  RINGS.forEach(([d, f, l]) =>
    s.addShape(pres.shapes.OVAL, { x: cx - d / 2, y: bottom - d, w: d, h: d, fill: { color: f }, line: l ? { color: l, width: 1 } : { type: "none" } }));
  const ringTop = (d) => bottom - d;
  text(s, [{ text: "TAM ", options: { bold: true, color: D } }, { text: fmt(MK["TAM, семей"] / 1e6, 2) + " млн семей", options: { bold: true } }],
    { x: cx - 1.5, y: ringTop(3.5) + 0.22, w: 3, h: 0.28, fontSize: 12, align: "center" });
  text(s, fmt(MK["TAM, руб в год"] / 1e9, 1) + " млрд руб. в год", { x: cx - 1.5, y: ringTop(3.5) + 0.5, w: 3, h: 0.24, fontSize: 10, align: "center", color: MUTED });
  text(s, [{ text: "SAM ", options: { bold: true, color: D } }, { text: fmt(MK["SAM, семей"] / 1e6, 2) + " млн семей", options: { bold: true } }],
    { x: cx - 1.15, y: ringTop(2.3) + 0.2, w: 2.3, h: 0.28, fontSize: 12, align: "center" });
  text(s, fmt(MK["SAM, руб в год"] / 1e9, 1) + " млрд руб. в год", { x: cx - 1.15, y: ringTop(2.3) + 0.48, w: 2.3, h: 0.24, fontSize: 10, align: "center" });
  text(s, "SOM", { x: cx - 0.525, y: ringTop(1.05) + 0.25, w: 1.05, h: 0.25, fontSize: 10.5, bold: true, color: A, align: "center", valign: "middle" });
  text(s, fmt(MK["SOM на 60 мес, семей (базовый сценарий)"], 0), { x: cx - 0.525, y: ringTop(1.05) + 0.5, w: 1.05, h: 0.28, fontSize: 11, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
  text(s, "SOM: семей через 60 месяцев в базовом сценарии", { x: cx - 2.0, y: bottom + 0.06, w: 4.0, h: 0.2, fontSize: 8.5, color: MUTED, align: "center" });
  // formula
  const fx = 5.2, fw = CW - (fx - X0);
  text(s, "Как считали", { x: fx, y: Y0, w: fw, h: 0.3, fontSize: 13, bold: true, color: D });
  const rows = [
    ["Люди 65+", "25,8 млн", "факт"],
    ["x с хроническим заболеванием", "54%", "факт"],
    ["x живут отдельно от детей", "60%", "допущение"],
    ["x ребенок работает и координирует сам", "70%", "допущение"],
    ["= TAM, семей", "5,85 млн", "расчет"],
    ["x города 250 тыс.+, доход позволяет", "35%", "допущение"],
    ["= SAM, семей", "2,05 млн", "расчет"],
    ["x чек " + fmt(UB["ARPU с НДС, руб/мес"], 0) + " руб. с льготами x 12", fmt(MK["SAM, руб в год"] / 1e9, 1) + " млрд", "расчет"],
  ];
  const rh = 0.36, tagX = fx + fw - 0.74;
  rows.forEach(([a, b, k], i) => {
    const y = Y0 + 0.4 + i * rh, sum = a.startsWith("=");
    if (sum) s.addShape(pres.shapes.RECTANGLE, { x: fx, y, w: fw, h: rh - 0.04, fill: { color: TINT }, line: { type: "none" } });
    text(s, a, { x: fx + 0.08, y, w: 2.7, h: rh - 0.04, fontSize: 10, bold: sum, valign: "middle" });
    text(s, b, { x: tagX - 0.95, y, w: 0.85, h: rh - 0.04, fontSize: 10.5, bold: true, color: sum ? P : T, align: "right", valign: "middle" });
    tag(s, tagX, y + (rh - 0.04 - 0.2) / 2, k);
  });
  text(s, "Две доли пока допущения, поэтому это порядок величины. Даже при 40% вместо 60% TAM 3,9 млн семей",
    { x: fx, y: Y0 + 0.4 + rows.length * rh + 0.12, w: fw, h: 0.45, fontSize: 9.5, color: MUTED });
  src(s, "S1 Росстат через Демоскоп; S3 Мурашко, 07.2026; доли 60%, 70%, 35%: допущения команды, проверяем анкетой и тестом цены; расчет 05_finance/results.json");
  s.addNotes("Размер считаем от людей 65+ через несколько долей, и каждая доля помечена. Две из них пока допущения, поэтому TAM 5,85 миллиона семей это порядок величины. Важнее другое: для базового сценария нам нужно " + fmt(MK["SOM на 60 мес, семей (базовый сценарий)"], 0) + " семей через пять лет, это " + fmt(MK["Доля SOM от SAM, %"], 2) + "% от SAM, так что запас по рынку большой даже при более скромных долях.");
}

// ---------------- 7. Value proposition ----------------
{
  const s = addSlide("Рынок и решение");
  title(s, "Ребенок платит 2 490 руб. в месяц и узнает о проблеме родителя в тот же день");
  const ph = pic(s, "ph_event.png", X0, Y0 - 0.05, 2.1, YB - Y0 + 0.1);
  const fx = ph.x + ph.w + 0.3, fw = 3.6;
  const V = [
    ["Для", "работающих детей, чей пожилой родитель живет отдельно"],
    ["ИНГОССЕМЬЯ", "организует медпомощь родителю и сообщает о каждом шаге"],
    ["Благодаря", "координаторам и врачам Ингосстраха"],
    ["В отличие от", "звонков «на нервах» и сиделки за 30-72 тыс. руб."],
    ["Ребенок", "тратит меньше времени и узнает о проблемах вовремя"],
  ];
  const vh = (YB - Y0 - 0.15 * 4) / 5;
  V.forEach(([h, b], i) => {
    const y = Y0 + i * (vh + 0.15);
    box(s, fx, y, 1.2, vh, i === 1 ? P : D);
    text(s, h, { x: fx, y, w: 1.2, h: vh, fontSize: i === 1 ? 9.5 : 10.5, bold: true, color: i === 1 ? "FFFFFF" : A, align: "center", valign: "middle" });
    text(s, b, { x: fx + 1.32, y, w: fw - 1.32, h: vh, fontSize: 10.5, valign: "middle" });
  });
  const kx = fx + fw + 0.3, kw = X0 + CW - kx;
  const K = [["hand-coins", "в 12-29 раз", "дешевле сиделки"], ["bell", "в тот же день", "узнает о пропуске лекарства"], ["phone", "1 номер", "вместо регистратуры и аптеки"]];
  const kh = (YB - Y0 - 0.4 - G * 2) / 3;
  K.forEach(([ic, b, l], i) => {
    const y = Y0 + i * (kh + G);
    box(s, kx, y, kw, kh, TINT);
    badge(s, ic, kx + 0.15, y + (kh - 0.5) / 2, 0.5, "FFFFFF", P);
    text(s, b, { x: kx + 0.78, y: y + 0.1, w: kw - 0.88, h: 0.4, fontSize: 17, bold: true, color: P, valign: "middle" });
    text(s, l, { x: kx + 0.78, y: y + 0.5, w: kw - 0.88, h: kh - 0.58, fontSize: 9.5 });
  });
  text(s, "Родитель сам решает, что видит ребенок", { x: kx, y: YB - 0.3, w: kw, h: 0.3, fontSize: 10.5, bold: true, color: D, valign: "middle" });
  src(s, "формула ценностного предложения из материалов кейса; цена 2 490 руб.: допущение, проверяем тестом цены; сиделка: S6, S7");
  s.addNotes("Ценностное предложение мы собрали строго по формуле кейса. Слева экран, который ребенок получает в день, когда мама пропустила лекарство. Отличие для пользователя измеримо: меньше часов на организацию, новости в тот же день и один номер вместо трех. Цена 2 490 рублей пока допущение, ее проверим тестом цены и анкетой.");
}

// ---------------- 8. Scenario ----------------
{
  const s = addSlide("Рынок и решение");
  title(s, "Сценарий из 7 шагов опирается на системы, которые у Ингосстраха уже работают");
  pic(s, "scenario.png", X0, Y0 - 0.1, CW, YB - Y0 + 0.1);
  src(s, "сервисный сценарий команды, 07_docs/diagrams/Пользовательский_сценарий; нагрузка " + fmt(MIN_B, 1) + " минуты на семью: расчет финмодели");
  s.addNotes("Сценарий идет от триггера до возврата: ребенок подключает родителя в чат-боте в MAX за 5 минут, координатор знакомится с родителем и берет у него письменное согласие на медданные. Дальше каждый день короткий звонок робота, а живой человек подключается раз в месяц или по тревоге. Нижняя строка показывает, на какие системы Ингосстраха опирается каждый шаг: контакт-центр, телемедицина, клиники ДМС, МедРокет, IngoRead.");
}

// ---------------- 9. Prototype: child app ----------------
{
  const s = addSlide("Рынок и решение");
  title(s, "Ребенок видит ленту событий в мини-приложении MAX и в личном кабинете на компьютере");
  const capH = 0.3, boxH = YB - Y0 - capH - 0.05;
  const d = pic(s, "desk_feed.png", X0, Y0, 5.2, boxH);
  const px = d.x + d.w + 0.15, pw = (X0 + CW - px - 0.1) / 2;
  const p1 = pic(s, "ph_feed.png", px, Y0, pw, boxH);
  const p2 = pic(s, "ph_summary.png", px + pw + 0.1, Y0, pw, boxH);
  const capY = Y0 + Math.max(d.h, p1.h) + 0.05;
  [[d, "Личный кабинет на компьютере"], [p1, "Лента событий"], [p2, "Месячная сводка"]].forEach(([r, t]) =>
    text(s, t, { x: r.x - 0.1, y: capY, w: r.w + 0.2, h: capH, fontSize: 10, bold: true, color: D, align: "center", valign: "middle" }));
  src(s, "кликабельный прототип 06_prototype/Прототип_ИНГОССЕМЬЯ.html, режимы «Телефон» и «Компьютер»; данные семей вымышленные");
  s.addNotes("Это кликабельный прототип. Ребенок получает те же события в мини-приложении MAX на телефоне и в личном кабинете на компьютере: ленту с пометкой «требует внимания» и месячную сводку с предложением ДМС без давления. Родителю приложения не нужны, ему звонят на обычный телефон.");
}

// ---------------- 10. Prototype: coordinator ----------------
{
  const s = addSlide("Рынок и решение");
  title(s, "Координатор ведет " + LOAD_B + " семьи в одном окне: очередь, скрипт звонка и отчет");
  const c = pic(s, "coord.png", X0, Y0, 5.6, YB - Y0);
  const rx = c.x + c.w + 0.3, rw = X0 + CW - rx;
  text(s, LOAD_B + " семьи", { x: rx, y: Y0, w: rw, h: 0.55, fontSize: 30, bold: true, color: P, valign: "middle" });
  text(s, "на координатора на Базовом тарифе, " + fmt(MIN_B, 1) + " мин на семью в месяц", { x: rx, y: Y0 + 0.58, w: rw, h: 0.45, fontSize: 10 });
  tag(s, rx, Y0 + 1.08, "расчет");
  const F3 = [["list-checks", "Очередь по срочности", "красные флаги робота сверху"], ["phone-call", "Скрипт звонка", "6 пунктов, врач по кнопке"], ["send", "Отчет ребенку", "собирается сам, отправка в 1 клик"]];
  F3.forEach(([ic, h, b], i) => {
    const y = Y0 + 1.5 + i * 0.85;
    badge(s, ic, rx, y, 0.5, TINT, P);
    text(s, h, { x: rx + 0.65, y, w: rw - 0.65, h: 0.28, fontSize: 11.5, bold: true, color: D });
    text(s, b, { x: rx + 0.65, y: y + 0.28, w: rw - 0.65, h: 0.4, fontSize: 9.5, color: MUTED });
  });
  src(s, "кликабельный прототип 06_prototype/Прототип_ИНГОССЕМЬЯ.html, режим «Координатор»; нагрузка: расчет финмодели; данные семей вымышленные");
  s.addNotes("Справа рабочее место координатора: очередь по срочности, скрипт звонка и отчет ребенку в один клик. Именно этот экран определяет, сколько семей тянет один координатор, а значит и юнит-экономику: по модели около " + LOAD_B + " семей на человека.");
}

// ---------------- 11. Right to win ----------------
{
  const s = addSlide("Рынок и решение");
  title(s, "Почти все блоки сервиса у Ингосстраха уже есть: стартап не соберет их за 3 месяца и 5 млн руб.");
  const Gx = [
    ["headset", "Медицинский контакт-центр ДМС", "уже записывает к врачам"],
    ["hospital", "Клиники ДМС, ЛМС, «Будь здоров»", "прием, анализы, медсестра"],
    ["video", "Телемедицина", "медицинскую оценку дает врач"],
    ["file-text", "МедРокет и IngoRead", "API клиник и разбор выписок"],
    ["bot", "Роботизация контакт-центра", "основа ежедневного звонка"],
    ["shield-check", "B2B2C и доверие", "клиенты ДМС, ruAAA, с 1947 года"],
  ];
  const C = cols(3), th = 1.15;
  Gx.forEach(([ic, h, b], i) => {
    const c = C[i % 3], y = Y0 + Math.floor(i / 3) * (th + G);
    box(s, c.x, y, c.w, th, TINT);
    badge(s, ic, c.x + 0.18, y + (th - 0.62) / 2, 0.62, "FFFFFF", P);
    text(s, h, { x: c.x + 0.95, y: y + 0.18, w: c.w - 1.1, h: 0.45, fontSize: 11, bold: true, color: D, valign: "middle" });
    text(s, b, { x: c.x + 0.95, y: y + 0.64, w: c.w - 1.1, h: 0.35, fontSize: 9.5, color: MUTED });
  });
  const by = Y0 + 2 * th + 2 * G;
  box(s, X0, by, CW, YB - by, D);
  badge(s, "smartphone", X0 + 0.2, by + (YB - by - 0.5) / 2, 0.5, P, "FFFFFF");
  text(s, [{ text: "Честное ограничение. ", options: { bold: true, color: A } },
    "IngoMobile нет в Google Play и App Store, а иностранные мессенджеры для финорганизаций запрещены (41-ФЗ). Поэтому MVP идет через MAX, СМС и звонки"],
    { x: X0 + 0.85, y: by, w: CW - 1.05, h: YB - by, fontSize: 10.5, color: "FFFFFF", valign: "middle" });
  src(s, "S9 ingos.ru; S12 МедРокет, 11.2025; S11 Эксперт РА, 02.12.2025; S14 RuStore; S15 Habr; клиники и IngoRead по ресерчу команды; CX World Awards 2025 по материалам кейса");
  s.addNotes("Мы называем только те ресурсы, которые реально работают в сценарии. Главный из них: операторы медицинского контакт-центра ДМС уже каждый день организуют медпомощь, мы переупаковываем эту компетенцию для семьи без полиса. Есть и слабые места: IngoMobile недоступен в Google Play и App Store с 31.01.2025, а с 01.06.2025 финорганизациям запрещено информировать клиентов через иностранные мессенджеры, поэтому MVP не зависит от приложения.");
}

// ---------------- 12. Veterans ----------------
{
  const s = addSlide("Рынок и решение");
  title(s, "Особые условия для ветеранов СВО и ветеранов: льготы, навигатор и партнерство с фондом «Защитники Отечества»");
  const ph = pic(s, "ph_tariff.png", X0, Y0 - 0.05, 2.1, YB - Y0 + 0.1);
  const mx = ph.x + ph.w + 0.3, mw = 3.7;
  const short = { "Родители погибших участников СВО": "Родители погибших участников СВО",
    "Участники СВО на службе (оплачивают сервис для родителя)": "Участники СВО на службе",
    "Ветераны боевых действий, включая ветеранов СВО, и инвалиды войны": "Ветераны боевых действий",
    "Ветераны ВОВ и ветераны труда (родитель)": "Родитель ветеран ВОВ или труда" };
  const cats = Object.entries(VET["Категории"]);
  const C2 = cols(2, 0.15, mx, mw), th = 0.85;
  cats.forEach(([c, v], i) => {
    const col = C2[i % 2], y = Y0 + Math.floor(i / 2) * (th + 0.15), first = i === 0;
    box(s, col.x, y, col.w, th, first ? D : TINT);
    text(s, v["Скидка, %"] === 100 ? "Бесплатно" : "-" + v["Скидка, %"] + "%", { x: col.x + 0.14, y: y + 0.08, w: col.w - 0.28, h: 0.4, fontSize: 18, bold: true, color: first ? A : P, valign: "middle" });
    text(s, short[c] || c, { x: col.x + 0.14, y: y + 0.46, w: col.w - 0.28, h: 0.36, fontSize: 9, color: first ? "FFFFFF" : T });
  });
  const ny = Y0 + 2 * th + 0.15 + 0.15;
  text(s, "Не только скидка", { x: mx, y: ny, w: mw, h: 0.28, fontSize: 11.5, bold: true, color: D });
  const NP = [["route", "Навигатор льгот: внеочередная помощь, лекарства, санаторий"], ["brain", "Отдельная линия координаторов и психолог"], ["id-card", "Проверка по удостоверению, копии не храним"]];
  NP.forEach(([ic, t], i) => {
    const y = ny + 0.35 + i * 0.45;
    badge(s, ic, mx, y, 0.38, TINT, P);
    text(s, t, { x: mx + 0.5, y, w: mw - 0.5, h: 0.38, fontSize: 9.5, valign: "middle" });
  });
  const rx = mx + mw + 0.25, rw = X0 + CW - rx;
  box(s, rx, Y0, rw, YB - Y0, D);
  badge(s, "heart-handshake", rx + 0.18, Y0 + 0.18, 0.5, P, "FFFFFF");
  text(s, "Фонд «Защитники Отечества»: партнер, а не конкурент", { x: rx + 0.8, y: Y0 + 0.14, w: rw - 0.95, h: 0.58, fontSize: 10.5, bold: true, color: A, valign: "middle" });
  const kp = [["4 500+", "социальных координаторов фонда", "факт"],
    [fmt(VET["Средняя скидка с подписки B2C, %"], 1) + "%", "средняя скидка, " + mln(B["Льготы ветеранам за 60 мес (недополученная выручка), руб"]) + " млн руб. за 60 мес уже в модели", "расчет"],
    ["40 семей", "в месяц через фонд с 6-го месяца, проверяем H6", "допущение"]];
  kp.forEach(([b, l, k], i) => {
    const y = Y0 + 0.9 + i * 0.95;
    text(s, b, { x: rx + 0.18, y, w: rw - 0.36, h: 0.36, fontSize: 18, bold: true, color: "FFFFFF", valign: "middle" });
    text(s, l, { x: rx + 0.18, y: y + 0.38, w: rw - 0.36, h: 0.45, fontSize: 8.5, color: L });
  });
  src(s, "S28 83-ФЗ от 21.04.2025 (Гарант); S29 ВЗГЛЯД, 01.06.2026; S30 ТВ Звезда, А. Цивилева; доли и скидки: допущения команды; Финансовая_модель_ИНГОССЕМЬЯ.xlsx, лист «Льготы ветеранам»");
  s.addNotes(nb("Мы встроили в продукт особые условия для ветеранов СВО и ветеранов, и это не разовая акция. У семей участников СВО та же задача стоит острее: участник на службе не может каждый день звонить родителю, а у родителей погибших участников рядом вообще нет ребенка, поэтому для них тариф «Базовый» бесплатный. Участникам СВО на службе скидка 50%, ветеранам боевых действий 30%, ветеранам ВОВ и ветеранам труда 10%. Но главное не скидка, а навигатор: координатор помогает получить положенную по закону внеочередную медпомощь и льготные лекарства. Фонд «Защитники Отечества» для нас партнер: его социальные координаторы решают социальные вопросы и передают нам медицинские задачи по родителю. Льготы стоят сервису " + mln(B["Льготы ветеранам за 60 мес (недополученная выручка), руб"]) + " миллиона рублей за пять лет, и эта сумма уже вычтена из всех цифр модели."));
}

// ---------------- 13. Business model ----------------
pres.addSection({ title: "Экономика" });
{
  const s = addSlide("Экономика");
  title(s, "Деньги приходят из подписки, опции к ДМС и комиссий, а экономику решают три параметра");
  const FLOW = [["wallet", "Платит", "ребенок или работодатель"], ["landmark", "Продает", "сервисная компания группы"],
    ["headset", "Делает", "координатор, робот, врачи"], ["user-round", "Получает", "родитель"]];
  const fc = cols(4, 0.4), fh = 0.62;
  FLOW.forEach(([ic, h, b], i) => {
    const c = fc[i];
    badge(s, ic, c.x, Y0, fh, i === 3 ? A : TINT, i === 3 ? D : P);
    text(s, h, { x: c.x + fh + 0.1, y: Y0 + 0.02, w: c.w - fh - 0.1, h: 0.28, fontSize: 11, bold: true, color: D });
    text(s, b, { x: c.x + fh + 0.1, y: Y0 + 0.3, w: c.w - fh - 0.1, h: 0.32, fontSize: 9.5, color: MUTED });
    if (i < 3) s.addShape(pres.shapes.LINE, { x: c.x + c.w + 0.06, y: Y0 + fh / 2, w: 0.28, h: 0, line: { color: L, width: 1.5, endArrowType: "triangle" } });
  });
  const ry = Y0 + fh + 0.3;
  text(s, "Источники дохода", { x: X0, y: ry, w: 3, h: 0.28, fontSize: 12, bold: true, color: D });
  tag(s, X0 + 1.6, ry + 0.04, "допущение");
  const REV = [["2 490 руб./мес", "Базовый тариф"], ["4 990 руб./мес", "Расширенный тариф"], ["1 990 руб./мес", "опция к корпоративному ДМС"], ["10%", "комиссия с услуг партнеров"]];
  const rc = cols(4), rh = 0.85;
  REV.forEach(([b, l], i) => {
    const c = rc[i];
    box(s, c.x, ry + 0.38, c.w, rh, TINT);
    text(s, b, { x: c.x + 0.14, y: ry + 0.46, w: c.w - 0.28, h: 0.38, fontSize: 15, bold: true, color: P, valign: "middle" });
    text(s, l, { x: c.x + 0.14, y: ry + 0.86, w: c.w - 0.28, h: 0.3, fontSize: 9.5 });
  });
  const ky = ry + 0.38 + rh + 0.3;
  const KP = [[LOAD_B + " семьи", "на координатора (" + fmt(MIN_B, 1) + " мин на семью)"],
    [fmt(UB["CAC при полном бюджете (смесь каналов), руб"], 0) + " руб.", "CAC в смеси каналов"], ["25% и 6%", "отток в 1-й месяц и дальше"]];
  const kc = cols(3);
  text(s, "3 ключевых параметра", { x: X0, y: ky - 0.02, w: 3, h: 0.28, fontSize: 12, bold: true, color: D });
  KP.forEach(([b, l], i) => {
    const c = kc[i], y = ky + 0.32;
    box(s, c.x, y, c.w, YB - y, D);
    text(s, b, { x: c.x + 0.16, y, w: 1.55, h: YB - y, fontSize: 16, bold: true, color: A, valign: "middle" });
    text(s, l, { x: c.x + 1.7, y, w: c.w - 1.85, h: YB - y, fontSize: 9, color: "FFFFFF", valign: "middle" });
  });
  src(s, "Финансовая_модель_ИНГОССЕМЬЯ.xlsx, базовый сценарий; НДС 22% по 425-ФЗ учтен в ARPU; подписку продает сервисная компания группы (ст. 6 закона 4015-1), схему подтверждает юрист");
  s.addNotes("Пользуются сервисом ребенок и родитель, платит ребенок или работодатель. Подписку продает отдельная сервисная компания группы, потому что статья 6 закона об организации страхового дела ограничивает страховщика страховой деятельностью; схему подтвердит юрист в первый месяц MVP. Плюс кросс-продажи ДМС и полиса ухода и льготы ветеранам со средней скидкой " + fmt(VET["Средняя скидка с подписки B2C, %"], 1) + "%. Сильнее всего экономику двигают нагрузка на координатора, стоимость привлечения и отток, поэтому все три мы меряем в MVP.");
}

// ---------------- 14. Unit economics ----------------
{
  const s = addSlide("Экономика");
  title(s, "Робот-триаж перевернул юнит-экономику: с минус " + fmt(-U1["Маржинальный доход на семью, руб/мес"], 0) + " до плюс " + fmt(UB["Маржинальный доход на семью, руб/мес"], 0) + " руб. с семьи в месяц");
  const cats = ["Выручка без НДС", "Координатор и накладные", "Прочие переменные", "Маржинальный доход"];
  const k = ["ARPU без НДС, руб/мес", "Координатор, супервизия, накладные на семью, руб/мес", "Прочие переменные на семью, руб/мес", "Маржинальный доход на семью, руб/мес"];
  s.addChart(pres.charts.BAR, [
    { name: "Версия 1: живой звонок раз в неделю", labels: cats, values: k.map((x) => U1[x]) },
    { name: "Сейчас: робот каждый день + триаж", labels: cats, values: k.map((x) => UB[x]) },
  ], { x: X0 - 0.1, y: Y0 - 0.1, w: 5.6, h: 3.6, barDir: "col", barGrouping: "clustered", chartColors: [L, P],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0", dataLabelFontSize: 9, dataLabelFontFace: F, dataLabelColor: T,
    catAxisLabelFontSize: 9, catAxisLabelFontFace: F, catAxisLabelColor: T, valAxisHidden: true,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: true, legendPos: "t", legendFontSize: 9, legendFontFace: F,
    showTitle: true, title: "Руб. на семью в месяц", titleFontSize: 10, titleFontFace: F, titleColor: MUTED });
  text(s, "Маржинальный доход = выручка без НДС с учетом льгот ветеранам + комиссия партнеров - затраты", { x: X0, y: YB - 0.2, w: 5.3, h: 0.2, fontSize: 7.5, color: MUTED });
  const rx = 6.1, rw = X0 + CW - rx;
  text(s, "Что поменяли в продукте", { x: rx, y: Y0, w: rw, h: 0.3, fontSize: 12, bold: true, color: D });
  const ch = [["bot", "Триаж", "робот каждый день, человек по флагу"], ["building-2", "Регион", "контакт-центр, зарплата 70 тыс. руб."], ["wallet", "Цена", "2 490 и 4 990 вместо 1 990 и 3 990"]];
  ch.forEach(([ic, h, b], i) => {
    const y = Y0 + 0.4 + i * 0.58;
    badge(s, ic, rx, y, 0.42, TINT, P);
    text(s, [{ text: h + ": ", options: { bold: true, color: P } }, b], { x: rx + 0.55, y, w: rw - 0.55, h: 0.42, fontSize: 9.5, valign: "middle" });
  });
  const kp = [[fmt(UB["LTV/CAC"], 1), "LTV/CAC (было " + fmt(U1["LTV/CAC"], 2) + ")"], [fmt(UB["Окупаемость CAC, мес"], 1) + " мес", "окупаемость CAC"], [fmt(UB["Маржинальность, %"], 1) + "%", "доля маржи"]];
  const kc = cols(3, 0.1, rx, rw), ky = Y0 + 2.3;
  kp.forEach(([b, l], i) => {
    const c = kc[i];
    box(s, c.x, ky, c.w, YB - ky, TINT);
    text(s, b, { x: c.x + 0.04, y: ky + 0.15, w: c.w - 0.08, h: 0.5, fontSize: 17, bold: true, color: P, align: "center", valign: "middle" });
    text(s, l, { x: c.x + 0.06, y: ky + 0.7, w: c.w - 0.12, h: 0.7, fontSize: 8.5, align: "center" });
  });
  src(s, "Финансовая_модель_ИНГОССЕМЬЯ.xlsx, юнит-экономика базового сценария и версии 1; все значения: расчет модели с учетом льгот ветеранам");
  s.addNotes(nb("Первая версия, где живой координатор звонит каждой семье раз в неделю, теряла " + fmt(-U1["Маржинальный доход на семью, руб/мес"], 0) + " рублей на каждой семье: чем больше клиентов, тем глубже минус. Мы отдали рутину роботу и телемедицине, перенесли координаторов в регион и подняли цену. Теперь семья приносит " + fmt(UB["Маржинальный доход на семью, руб/мес"], 0) + " рублей маржинального дохода с учетом льгот ветеранам, а LTV/CAC " + fmt(UB["LTV/CAC"], 1) + ". Поэтому главная гипотеза MVP именно про триаж."));
}

// ---------------- 15. Financial forecast ----------------
{
  const s = addSlide("Экономика");
  title(s, "Базовый сценарий окупается на " + B["Месяц окупаемости (накопленный > 0)"] + "-м месяце, но за 3 года NPV еще отрицательный");
  pic(s, "Накопленный_результат_по_сценариям.png", X0 - 0.1, Y0 - 0.1, 5.5, YB - Y0 + 0.1);
  const rx = 6.1, rw = X0 + CW - rx;
  const K = [[mln(B["NPV до налога за 36 мес, руб"]) + " млн", "NPV за 36 месяцев, до налога", RED],
    ["+" + mln(B["NPV до налога за 60 мес, руб"]) + " млн", "NPV за 60 месяцев (" + mln(B["NPV после налога за 60 мес, руб"]) + " после налога)", P],
    [B["Месяц операционной безубыточности"] + "-й мес", "операционная безубыточность", P], [mln(B["Пиковая потребность в деньгах, руб"]) + " млн", "пик вложений вместе с MVP", D]];
  const kc = cols(2, 0.1, rx, rw), kh = 1.1;
  K.forEach(([b, l, c], i) => {
    const col = kc[i % 2], y = Y0 + Math.floor(i / 2) * (kh + 0.1);
    box(s, col.x, y, col.w, kh, TINT);
    text(s, b, { x: col.x + 0.1, y: y + 0.08, w: col.w - 0.2, h: 0.45, fontSize: 17, bold: true, color: c, valign: "middle" });
    text(s, l, { x: col.x + 0.1, y: y + 0.55, w: col.w - 0.2, h: 0.5, fontSize: 9 });
  });
  const rows = [
    ["NPV, млн руб.", "36 мес", "60 мес", "Безуб."],
    ["Пессимистичный", mln(PS["NPV до налога за 36 мес, руб"]), mln(PS["NPV до налога за 60 мес, руб"]), "нет"],
    ["Базовый", mln(B["NPV до налога за 36 мес, руб"]), mln(B["NPV до налога за 60 мес, руб"]), B["Месяц операционной безубыточности"] + " мес"],
    ["Оптимистичный", mln(OP["NPV до налога за 36 мес, руб"]), mln(OP["NPV до налога за 60 мес, руб"]), OP["Месяц операционной безубыточности"] + " мес"],
  ].map((r, ri) => r.map((c, ci) => ({ text: nb(String(c)), options: { bold: ri === 0 || ri === 2, fontSize: 9, color: ri === 0 ? "FFFFFF" : T,
    fill: { color: ri === 0 ? D : ri === 2 ? TINT : "FFFFFF" }, align: ci ? "right" : "left" } })));
  const ty = Y0 + 2 * kh + 0.2;
  s.addTable(rows, { x: rx, y: ty, w: rw, colW: [1.15, 0.78, 0.78, rw - 2.71], fontFace: F, rowH: 0.27, margin: [0.02, 0.05, 0.02, 0.05],
    border: { type: "solid", color: LINE, pt: 0.5 }, valign: "middle" });
  text(s, "млн руб., до налога, ставка 22%", { x: rx, y: ty + 4 * 0.27 + 0.06, w: rw, h: 0.2, fontSize: 7.5, color: MUTED });
  src(s, "Финансовая_модель_ИНГОССЕМЬЯ.xlsx, лист «Сводка»; горизонт 60 месяцев; все значения: расчет модели с учетом льгот ветеранам, входные параметры: допущения");
  s.addNotes(nb("Мы не скрываем, что проект не быстрый. В базе операционная безубыточность на " + B["Месяц операционной безубыточности"] + "-м месяце, окупаемость на " + B["Месяц окупаемости (накопленный > 0)"] + "-м, NPV за три года минус " + mln(-B["NPV до налога за 36 мес, руб"]) + " миллиона, за пять лет плюс " + mln(B["NPV до налога за 60 мес, руб"]) + ". Льготы ветеранам уже вычтены. Пессимистичный сценарий уходит в глубокий минус, поэтому решение о росте мы привязываем к результатам MVP, а не к прогнозу."));
}

// ---------------- 16. Risks & sensitivity ----------------
{
  const s = addSlide("Экономика");
  title(s, "NPV сильнее всего зависит от цены и минут координатора: их и проверяем в MVP");
  const [l, r] = cols(2);
  pic(s, "Чувствительность_NPV.png", l.x, Y0 - 0.1, l.w, 2.45);
  const mc = pic(s, "Монте_Карло_NPV.png", r.x, Y0 - 0.1, r.w, 2.0);
  text(s, [{ text: fmt(MC["Вероятность NPV(60 мес) > 0, %"], 1) + "% ", options: { bold: true, color: P } }, "прогонов дают NPV за 60 мес выше нуля, за 36 мес только ",
    { text: fmt(MC["Вероятность NPV(36 мес) > 0, %"], 1) + "%", options: { bold: true, color: RED } }],
    { x: r.x, y: mc.y + mc.h + 0.02, w: r.w, h: 0.4, fontSize: 9.5, valign: "middle" });
  const RK = [["pill", "Медданные (152-ФЗ ст. 10)", "Письменное согласие самого родителя: СМС-код как ПЭП или бумага при визите"],
    ["bell", "Робот пропустит тревогу", "Порог H4: ноль пропусков. Флаг передает семью человеку в тот же день"],
    ["heart-handshake", "Родитель против «контроля»", "Помощь, а не надзор: родитель решает, что видит ребенок (H3)"]];
  const ry = Y0 + 2.5, rh = YB - ry;
  cols(3).forEach((c, i) => {
    box(s, c.x, ry, c.w, rh, TINT);
    badge(s, RK[i][0], c.x + 0.14, ry + 0.14, 0.42, "FFFFFF", P);
    text(s, RK[i][1], { x: c.x + 0.66, y: ry + 0.14, w: c.w - 0.8, h: 0.42, fontSize: 10, bold: true, color: D, valign: "middle" });
    text(s, RK[i][2], { x: c.x + 0.14, y: ry + 0.66, w: c.w - 0.28, h: rh - 0.75, fontSize: 9 });
  });
  src(s, "Финансовая_модель_ИНГОССЕМЬЯ.xlsx, листы «Чувствительность» и «Монте-Карло»: 5 000 прогонов, параметры ±25%; P95 пиковой потребности " + mln(MC["Пиковая потребность P95, руб"]) + " млн руб.; риски: пояснительная записка, раздел 12");
  s.addNotes("Торнадо показывает, что главные рычаги это цена тарифов, минуты координатора на семью и его зарплата. Монте-Карло на 5 000 прогонов дает положительный NPV за пять лет в " + fmt(MC["Вероятность NPV(60 мес) > 0, %"], 1) + "% случаев, а за три года только в " + fmt(MC["Вероятность NPV(36 мес) > 0, %"], 1) + "%. Пиковая потребность в деньгах в 95% прогонов не выше " + mln(MC["Пиковая потребность P95, руб"]) + " миллиона. Внизу три риска, которые мы закрываем дизайном MVP: согласие родителя на медданные, пропуск тревоги роботом и отказ родителя от «контроля».");
}

// ---------------- 17. MVP & budget ----------------
pres.addSection({ title: "MVP и решение" });
{
  const s = addSlide("MVP и решение");
  title(s, "За 12 недель и " + MVP_M + " млн руб. проверяем сервис на 150 семьях");
  text(s, MVP_M + " млн руб.", { x: X0, y: Y0 - 0.05, w: 2.2, h: 0.4, fontSize: 20, bold: true, color: P, valign: "middle" });
  text(s, "из лимита 5 млн,\nзапас " + fmt(R["Запас до лимита 5 млн, руб"], 0) + " руб.", { x: X0 + 2.15, y: Y0 - 0.05, w: 1.9, h: 0.4, fontSize: 9.5, color: MUTED, valign: "middle" });
  const order = Object.entries(BUD).sort((a, b) => a[1] - b[1]);
  s.addChart(pres.charts.BAR, [{ name: "Бюджет", labels: order.map((x) => x[0]), values: order.map((x) => Math.round(x[1] / 1e4) / 100) }],
    { x: X0 - 0.15, y: Y0 + 0.4, w: 4.2, h: 3.15, barDir: "bar", chartColors: [P], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.00",
      dataLabelFontSize: 9, dataLabelFontFace: F, dataLabelColor: T, catAxisLabelFontSize: 9, catAxisLabelFontFace: F, catAxisLabelColor: T,
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 60 });
  text(s, "млн руб. по статьям", { x: X0, y: YB - 0.2, w: 4.0, h: 0.2, fontSize: 8, color: MUTED });
  const M = [["Месяц 1", "Подготовка", ["Юрист: схема и согласия", "Чат-бот в MAX", "3 координатора", "Пилот с фондом"]],
    ["Месяц 2", "50 семей", ["Сотрудники и клиенты ДМС", "Лендинг с предоплатой (H2)", "Хронометраж триажа (H4)"]],
    ["Месяц 3", "150 семей", ["Тест цены", "15 семей ветеранов (H6)", "B2B2C-пилоты (H5)"]]];
  const mc = cols(3, 0.12, 4.75, X0 + CW - 4.75);
  M.forEach(([m, h, items], i) => {
    const c = mc[i];
    box(s, c.x, Y0, c.w, 0.65, i === 0 ? D : P);
    text(s, m, { x: c.x + 0.1, y: Y0 + 0.04, w: c.w - 0.2, h: 0.25, fontSize: 9, bold: true, color: A });
    text(s, h, { x: c.x + 0.1, y: Y0 + 0.29, w: c.w - 0.2, h: 0.32, fontSize: 11, bold: true, color: "FFFFFF", valign: "middle" });
    text(s, bullets(items), { x: c.x, y: Y0 + 0.8, w: c.w, h: 2.2, fontSize: 9.5 });
  });
  const fy = YB - 0.6;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: mc[0].x, y: fy, w: X0 + CW - mc[0].x, h: 0.6, rectRadius: 0.08, fill: { color: A }, line: { type: "none" } });
  badge(s, "calendar-check", mc[0].x + 0.12, fy + 0.1, 0.4, D, "FFFFFF");
  text(s, "Неделя 12: решение по порогам H2-H6", { x: mc[0].x + 0.65, y: fy, w: 3.9, h: 0.6, fontSize: 11.5, bold: true, color: T, valign: "middle" });
  src(s, "Финансовая_модель_ИНГОССЕМЬЯ.xlsx, лист «Бюджет MVP»; дорожная карта 07_docs/diagrams/Дорожная_карта_MVP; выручка за MVP " + fmt(R["Выручка за MVP (базовый), руб"], 0) + " руб. в бюджет не засчитана");
  s.addNotes("MVP частично ручной и не требует новой разработки в IngoMobile. Больше половины бюджета это команда: три координатора из действующих операторов ДМС, супервизор и продакт. Первый месяц уходит на юридическую схему, бот, скрипты, курс по льготам ветеранов и договор о пилоте с фондом «Защитники Отечества», во втором запускаем 50 семей, в третьем растем до 150, из них около 15 семей ветеранов, и тестируем цену 2 490 и 4 990 рублей. В конце двенадцатой недели принимаем решение.");
}

// ---------------- 18. Hypotheses ----------------
{
  const s = addSlide("MVP и решение");
  title(s, "Шесть гипотез с порогами: на неделе 12 решаем, растить проект или остановить");
  const H = [
    ["H1", "Проблема частая и дорогая", "Интервью 15+, анкета 150+", "60%+ тратят 3+ часа в неделю и имели упущенный эпизод"],
    ["H2", "Платят за организацию", "Лендинг с предоплатой", "Конверсия в предоплату 3%+"],
    ["H3", "Родитель принимает звонки", "MVP, 50-150 семей", "70%+ согласий, 80%+ звонков с ответом"],
    ["H4", "Триаж держит нагрузку", "Хронометраж в MVP", "До 49 мин на семью (около " + LOAD_B + " семей), оценка 4,5+, ноль пропущенных тревог"],
    ["H5", "Семьи остаются и покупают еще", "Когорты MVP", "Продление 75%+, доп. услуги у 10%+"],
    ["H6", "Семьи ветеранов приходят через фонд", "Пилот с фондом", "10+ семей в месяц, продление 80%+"],
  ];
  const hdr = ["", "Гипотеза", "Как проверяем", "Порог успеха"].map((t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: D }, fontSize: 9.5 } }));
  const rows = H.map((r, ri) => r.map((c, ci) => ({ text: nb(c), options: { bold: ci < 2, fontSize: ci === 0 ? 11 : 9.5,
    color: ci === 0 ? P : T, fill: { color: ri === 3 ? "F4FBDF" : ri % 2 ? "FFFFFF" : TINT } } })));
  s.addTable([hdr, ...rows], { x: X0, y: Y0 - 0.1, w: CW, colW: [0.45, 2.75, 2.1, 3.7], fontFace: F, rowH: [0.28, 0.38, 0.38, 0.38, 0.44, 0.38, 0.38],
    margin: [0.03, 0.06, 0.03, 0.06], border: { type: "solid", color: "FFFFFF", pt: 1 }, valign: "middle" });
  const DEC = [["", "Все пороги H2-H6 пройдены", "Просим около " + PEAK_AFTER + " млн руб. на рост", P, "FFFFFF"],
    ["", "Пройдена часть порогов", "Меняем цену, сценарий или сегмент, меряем снова", L, T],
    ["", "Не прошли H3 или H4", "Останавливаемся. Потеря не больше " + MVP_M + " млн руб.", D, "FFFFFF"]];
  const dy = 4.0;
  cols(3).forEach((c, i) => {
    const [, h, b, f, col] = DEC[i];
    box(s, c.x, dy, c.w, YB - dy, f);
    text(s, h, { x: c.x + 0.15, y: dy + 0.08, w: c.w - 0.3, h: 0.32, fontSize: 10.5, bold: true, color: col === T ? D : A, valign: "middle" });
    text(s, b, { x: c.x + 0.15, y: dy + 0.42, w: c.w - 0.3, h: 0.52, fontSize: 9.5, color: col });
  });
  src(s, "пояснительная записка, раздел 11; правило решения: допущение команды; H4 главная гипотеза, на ней держится юнит-экономика");
  s.addNotes("Каждая гипотеза имеет способ проверки и числовой порог, заданный заранее. H4 выделена: если триаж не удержит 49 минут на семью, экономика возвращается к версии 1. Правило решения простое: все пороги пройдены, просим деньги на рост; часть, перестраиваем и меряем снова; провал H3 или H4, останавливаемся. H6 проверяет канал через фонд «Защитники Отечества»: если он не сработает, льготы остаются, а канал убираем из модели.");
}

// ---------------- 19. Answer ----------------
{
  const s = addSlide("MVP и решение", "DARK");
  text(s, "Ответ жюри", { x: X0, y: 0.35, w: 4, h: 0.3, fontSize: 11, bold: true, color: A });
  text(s, "Доказательств хватает, чтобы вложить " + MVP_M + " млн руб. в проверку. Для масштабирования их пока нет", {
    x: X0, y: 0.65, w: CW, h: 0.85, fontSize: 21, bold: true, color: "FFFFFF" });
  const yes = ["89% нуждающихся без госухода", "Между «сам» и сиделкой нет продукта", "У Ингосстраха есть почти все блоки", "Особые условия для ветеранов СВО и ветеранов", "Потери ограничены: " + MVP_M + " млн руб."];
  const no = ["H1 проверяем интервью 23-25.10", "Две доли рынка пока допущения", "Нет ни одного платежа за 2 490 руб.", "Нагрузка " + LOAD_B + " семей пока расчетная", "NPV за 36 мес " + mln(B["NPV до налога за 36 мес, руб"]) + " млн руб."];
  const [l, r] = cols(2, 0.4);
  [[l, "badge-check", "Что говорит за проверку", A, yes], [r, "circle-alert", "Что пока не доказано", L, no]].forEach(([c, ic, h, hc, arr]) => {
    badge(s, ic, c.x, 1.65, 0.42, P, "FFFFFF");
    text(s, h, { x: c.x + 0.55, y: 1.65, w: c.w - 0.55, h: 0.42, fontSize: 12.5, bold: true, color: hc, valign: "middle" });
    text(s, bullets(arr, { paraSpaceAfter: 5 }), { x: c.x, y: 2.2, w: c.w, h: 1.75, fontSize: 11, color: "FFFFFF" });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: X0, y: 4.15, w: CW, h: 0.72, rectRadius: 0.08, fill: { color: A }, line: { type: "none" } });
  text(s, [{ text: "Рекомендация: ", options: { bold: true } }, "профинансировать MVP и решать о росте на неделе 12 по порогам H2-H6"],
    { x: X0 + 0.2, y: 4.15, w: CW - 0.4, h: 0.72, fontSize: 12, color: T, valign: "middle" });
  text(s, String(pres.slides.length), { x: 9.05, y: 5.22, w: 0.45, h: 0.25, fontSize: 8, color: L, align: "right" });
  s.addNotes("Наш честный ответ: да, чтобы инвестировать в трехмесячную проверку, и нет, чтобы сразу масштабировать. За проверку говорят открытые данные, пустая ниша между «сам» и сиделкой и готовые блоки Ингосстраха. Против пока то, что спрос и готовность платить не подтверждены реальными людьми и деньгами, а NPV за три года положителен только в " + fmt(MC["Вероятность NPV(36 мес) > 0, %"], 1) + "% прогонов. Поэтому просим " + MVP_M + " миллиона с жесткой точкой решения на двенадцатой неделе.");
}

// ---------------- 20. Sources ----------------
pres.addSection({ title: "Приложение" });
{
  const s = addSlide("Приложение");
  title(s, "Приложение: источники, проверены 22.10.2026");
  const L1 = [
    "S1 Демоскоп по данным Росстата, 65+ в РФ: demoscope.ru/weekly/2025/01079/barom01.php",
    "S2 ВЦИОМ, «Куда уходит старость», 2026: wciom.ru/analytical-reviews/analiticheskii-obzor/kuda-ukhodit-starost",
    "S3 ОСН, А. Мурашко о хронических заболеваниях: osnmedia.ru/obshhestvo/murashko-54-pozhilyh-rossiyan-imeyut-hronicheskie-infektsii",
    "S4 АСИ, Минтруд о долговременном уходе: asi.org.ru/news/2026/07/21/v-rossii-dolgovremennym-uhodom-ohvacheno-11-pozhilyh-i-lyudej-s-invalidnostyu-mintrud",
    "S5 ВШЭ, обследование «Человек, семья и общество»: jsps.hse.ru/article/view/17807",
    "S6 Peterburg2, цены на уход, 14.08.2026: peterburg2.ru/news/skolko-stoit-uhod-za-pozhilymi-roditelyami-v-rossii-realnye-summy-2026-goda-248561.html",
    "S7 360.ru, сиделка и пансионат: 360.ru/news/obschestvo/psiholog-kondrashina-sravnila-rashody-na-sidelku-i-pansionat",
    "S8 Коммерсантъ, полисы долговременного ухода, 02.09.2026: kommersant.ru/doc/8922548",
    "S9 Ингосстрах, ДМС и медицинский контакт-центр: ingos.ru/health/dms",
    "S10 Эксперт РА, рынок страхования 2025: raexpert.ru/researches/insurance/ins_2025",
    "S11 Эксперт РА, рейтинг ruAAA: raexpert.ru/releases/2025/dec02b",
  ];
  const L2 = [
    "S12 МедРокет и Ингосстрах: infullbroker.ru/blog/tsifrovoye-dms-ingosstrakh-i-medroket-sozdali-edinuyu-sistemu-integratsii-klinik-i-strakhovschikov",
    "S13 Мини-приложение Ингосстраха в MAX: infullbroker.ru/blog/ingosstrah-zapustil-mini-prilozhenie-v-nacionalnom-messendzhere-max",
    "S14 RuStore, IngoMobile: rustore.ru/catalog/app/ru.ingos.ingomobile",
    "S15 Habr, IngoMobile недоступен в сторах: habr.com/ru/news/878342",
    "S17 Коммерсантъ, обзор ЦБ за 2025 год: kommersant.ru/doc/8438895",
    "S23 Zoon, вызов медсестры на дом: zoon.ru/msk/medical/type/vyzov_medsestry_na_dom",
    "S27 Страница конкурса IngoSight 2026: changellenge.com/tpage/ingosight-konkursnyy-etap",
    "S28 Гарант, 83-ФЗ от 21.04.2025 о статусе ветерана боевых действий: garant.ru/news/1810720",
    "S29 ВЗГЛЯД, фонд «Защитники Отечества», 01.06.2026: vz.ru/news/2026/6/1/1423828.html",
    "S30 ТВ Звезда, А. Цивилева о социальных координаторах: tvzvezda.ru/news/2026841835-XRS6r.html",
    "Законы (сверка в 02_research/Проверка_правовых_норм.docx): 41-ФЗ от 01.04.2025; 152-ФЗ ст. 9, 10; закон 4015-1 ст. 6; 425-ФЗ от 28.11.2025; 5-ФЗ «О ветеранах»",
    "Внутренние файлы: Финансовая_модель_ИНГОССЕМЬЯ.xlsx; Материалы_исследования_открытые_данные.docx; Льготы_ветеранов_СВО_и_ветеранов.docx",
  ];
  const [c1, c2] = cols(2, 0.3);
  [[L1, c1], [L2, c2]].forEach(([arr, c]) => text(s, arr.map((t, j) => ({ text: nb(t), options: { breakLine: j < arr.length - 1, paraSpaceAfter: 2 } })),
    { x: c.x, y: 1.0, w: c.w, h: 4.1, fontSize: 8, color: T }));
  text(s, "Метки в презентации: факт взят из источника и совпал при проверке; расчет получен командой из фактов; допущение проверяется в MVP", { x: X0, y: 5.2, w: 8.5, h: 0.25, fontSize: 7.5, color: MUTED });
  s.addNotes("Здесь все внешние источники с кодами, которые стоят на слайдах. Факты перепроверены 22 октября 2026 года. Полная таблица фактов с вердиктами лежит в материалах исследования.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
