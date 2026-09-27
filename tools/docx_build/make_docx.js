// Вёрстка отчёта о проверке данных в Word.
// Данные готовит tools/report_data.py, здесь только оформление.
const fs = require("fs");
const path = require("path");
const d = require("docx");
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel,
        Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
        Footer, PageNumber, VerticalAlign, PageOrientation } = d;

const ROOT = path.resolve(__dirname, "..", "..");
const data = JSON.parse(fs.readFileSync(path.join(ROOT, "reports", "report_data.json"), "utf8"));

const FONT = "Calibri";
const MM = 56.7;
const sz = pt => pt * 2;
const CONTENT_W = 15250;   // ширина полосы набора, альбомная А4
const INK = "1A1F36", ACCENT = "2B2F77", MUTED = "5A5F7A";
const LINE = "C7CBDC", HEAD_BG = "EEF0F7", ZEBRA = "F7F8FC";

const t = (text, o = {}) => new TextRun({
  text, font: FONT, size: sz(o.pt || 10), bold: !!o.b, italics: !!o.i,
  color: o.color || INK, break: o.break || undefined });

const p = (children, o = {}) => new Paragraph({
  children: Array.isArray(children) ? children : [children],
  alignment: o.align, spacing: { before: o.before || 0, after: o.after === undefined ? 100 : o.after },
  keepNext: o.keepNext });

const border = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border };
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

function cell(children, o = {}) {
  return new TableCell({
    width: { size: o.w, type: WidthType.DXA },
    borders: o.noBorder ? { top: none, bottom: none, left: none, right: none } : borders,
    shading: o.bg ? { type: ShadingType.CLEAR, color: "auto", fill: o.bg } : undefined,
    verticalAlign: VerticalAlign.TOP,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    columnSpan: o.span,
    children: Array.isArray(children) ? children : [children],
  });
}

// многострочный текст в ячейке
function lines(text, o = {}) {
  const parts = String(text || "").split("\n");
  return new Paragraph({
    spacing: { after: 0 },
    children: parts.map((s, i) => t(s, { ...o, break: i ? 1 : undefined })),
  });
}

const H1 = text => new Paragraph({
  children: [t(text, { pt: 15, b: true, color: ACCENT })],
  spacing: { before: 360, after: 140 }, keepNext: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 6 } },
});

const children = [];

// ---------------------------------------------------------------- титул
children.push(p(t("ОТЧЁТ О ПРОВЕРКЕ ИСХОДНЫХ ДАННЫХ", { pt: 20, b: true, color: ACCENT }),
                { align: AlignmentType.CENTER, after: 60 }));
children.push(p(t("Чемпионат Регулярной хоккейной лиги 3х3", { pt: 12, color: MUTED }),
                { align: AlignmentType.CENTER, after: 40 }));
children.push(p(t(`Дата проверки: ${data.date}`, { pt: 10, color: MUTED }),
                { align: AlignmentType.CENTER, after: 300 }));

// три числа
const NUM_W = Math.floor(CONTENT_W / 3);
children.push(new Table({
  columnWidths: [NUM_W, NUM_W, NUM_W],
  width: { size: NUM_W * 3, type: WidthType.DXA },
  rows: [
    new TableRow({ children: [
      cell(p(t(String(data.files), { pt: 26, b: true, color: ACCENT }), { align: AlignmentType.CENTER, after: 0 }), { w: NUM_W, bg: HEAD_BG }),
      cell(p(t(String(data.games), { pt: 26, b: true, color: ACCENT }), { align: AlignmentType.CENTER, after: 0 }), { w: NUM_W, bg: HEAD_BG }),
      cell(p(t(String(data.findings), { pt: 26, b: true, color: "B03030" }), { align: AlignmentType.CENTER, after: 0 }), { w: NUM_W, bg: HEAD_BG }),
    ]}),
    new TableRow({ children: [
      cell(p(t("файлов проверено", { pt: 9, color: MUTED }), { align: AlignmentType.CENTER, after: 0 }), { w: NUM_W }),
      cell(p(t("матчей проверено", { pt: 9, color: MUTED }), { align: AlignmentType.CENTER, after: 0 }), { w: NUM_W }),
      cell(p(t("найдено несостыковок", { pt: 9, color: MUTED }), { align: AlignmentType.CENTER, after: 0 }), { w: NUM_W }),
    ]}),
  ],
}));
children.push(p(t("Данные проверены без внесения изменений. Ниже перечислены места, которые требуют проверки.",
                  { pt: 10, i: true, color: MUTED }), { before: 200, align: AlignmentType.CENTER }));

// ---------------------------------------------------------------- главное
children.push(H1("Главное"));
children.push(p(t("Находки сгруппированы по смыслу. Начинать стоит сверху — там самые массовые группы.",
                  { pt: 10, color: MUTED })));

const TW = [2600, 1200, CONTENT_W - 3800];
const themeRows = [new TableRow({ tableHeader: true, children: [
  cell(p(t("Тема", { b: true, pt: 10 }), { after: 0 }), { w: TW[0], bg: HEAD_BG }),
  cell(p(t("Находок", { b: true, pt: 10 }), { align: AlignmentType.CENTER, after: 0 }), { w: TW[1], bg: HEAD_BG }),
  cell(p(t("Что это значит", { b: true, pt: 10 }), { after: 0 }), { w: TW[2], bg: HEAD_BG }),
]})];
data.themes.forEach((th, i) => {
  const bg = i % 2 ? ZEBRA : undefined;
  themeRows.push(new TableRow({ children: [
    cell(p(t(th.title, { b: true, pt: 10 }), { after: 0 }), { w: TW[0], bg }),
    cell(p(t(String(th.n), { b: true, pt: 12, color: ACCENT }), { align: AlignmentType.CENTER, after: 0 }), { w: TW[1], bg }),
    cell([p(t(th.detail, { pt: 9, color: MUTED }), { after: th.comment ? 60 : 0 }),
          ...(th.comment ? [p(t(th.comment, { pt: 10 }), { after: 0 })] : [])], { w: TW[2], bg }),
  ]}));
});
children.push(new Table({ columnWidths: TW, width: { size: TW.reduce((a, b) => a + b), type: WidthType.DXA }, rows: themeRows }));

// ---------------------------------------------------------------- сводная
children.push(H1("Сколько матчей проверено"));
const SW = [7500, 4000, CONTENT_W - 11500];
const sumRows = [new TableRow({ tableHeader: true, children: [
  cell(p(t("Соревнование", { b: true, pt: 10 }), { after: 0 }), { w: SW[0], bg: HEAD_BG }),
  cell(p(t("Месяц", { b: true, pt: 10 }), { after: 0 }), { w: SW[1], bg: HEAD_BG }),
  cell(p(t("Игр", { b: true, pt: 10 }), { align: AlignmentType.RIGHT, after: 0 }), { w: SW[2], bg: HEAD_BG }),
]})];
data.summary.forEach((r, i) => {
  const bg = i % 2 ? ZEBRA : undefined;
  sumRows.push(new TableRow({ children: [
    cell(p(t(r.comp, { pt: 10 }), { after: 0 }), { w: SW[0], bg }),
    cell(p(t(r.month, { pt: 10 }), { after: 0 }), { w: SW[1], bg }),
    cell(p(t(String(r.games), { pt: 10 }), { align: AlignmentType.RIGHT, after: 0 }), { w: SW[2], bg }),
  ]}));
});
sumRows.push(new TableRow({ children: [
  cell(p(t("ИТОГО", { b: true, pt: 10 }), { after: 0 }), { w: SW[0], bg: HEAD_BG }),
  cell(p(t("", { pt: 10 }), { after: 0 }), { w: SW[1], bg: HEAD_BG }),
  cell(p(t(String(data.games), { b: true, pt: 10 }), { align: AlignmentType.RIGHT, after: 0 }), { w: SW[2], bg: HEAD_BG }),
]}));
children.push(new Table({ columnWidths: SW, width: { size: SW.reduce((a, b) => a + b), type: WidthType.DXA }, rows: sumRows }));

// ---------------------------------------------------------------- по файлам
children.push(H1("Подробно по каждому файлу"));
children.push(p([t("Последняя колонка — для ответа организаторов. ", { pt: 10, b: true }),
                 t("Напротив каждой строки впишите, как исправляем ошибку.", { pt: 10 })],
                { after: 160 }));

const FW = [800, 2500, 2200, 4400, CONTENT_W - 9900];
data.blocks.forEach(b => {
  children.push(p(t(`${b.title} — ${b.period}`, { pt: 12, b: true }), { before: 300, after: 40, keepNext: true }));
  children.push(p(t(`${b.games} матчей · команд ${b.teams} · игроков ${b.players} · `
                    + `побед по буллитам ${b.shootouts} · файл ${b.file}`,
                    { pt: 9, color: MUTED }), { after: 100, keepNext: true }));
  if (!b.findings.length) {
    children.push(p(t("Ошибок не найдено.", { pt: 10, i: true, color: "2E7D32" })));
    return;
  }
  const rows = [new TableRow({ tableHeader: true, children: [
    cell(p(t("Стр.", { b: true, pt: 9 }), { after: 0 }), { w: FW[0], bg: HEAD_BG }),
    cell(p(t("Матч", { b: true, pt: 9 }), { after: 0 }), { w: FW[1], bg: HEAD_BG }),
    cell(p(t("Счёт", { b: true, pt: 9 }), { after: 0 }), { w: FW[2], bg: HEAD_BG }),
    cell(p(t("Что не так", { b: true, pt: 9 }), { after: 0 }), { w: FW[3], bg: HEAD_BG }),
    cell([p(t("Решение организатора", { b: true, pt: 9 }), { after: 20 }),
          p(t("как исправляем", { pt: 7, i: true, color: MUTED }), { after: 0 })],
         { w: FW[4], bg: HEAD_BG }),
  ]})];
  b.findings.forEach((f, i) => {
    const bg = i % 2 ? ZEBRA : undefined;
    rows.push(new TableRow({ children: [
      cell(p(t(f.row || "—", { pt: 9, b: true }), { after: 0 }), { w: FW[0], bg }),
      cell(lines(f.match || f.what, { pt: 9 }), { w: FW[1], bg }),
      cell(lines(f.score || "—", { pt: 9, color: MUTED }), { w: FW[2], bg }),
      cell(p(t(f.text, { pt: 9 }), { after: 0 }), { w: FW[3], bg }),
      cell(p(t("", { pt: 9 }), { after: 0 }), { w: FW[4] }),
    ]}));
  });
  children.push(new Table({ columnWidths: FW, width: { size: FW.reduce((a, b2) => a + b2), type: WidthType.DXA }, rows }));
});

// ---------------------------------------------------------------- документ
const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: sz(10), color: INK } } } },
  sections: [{
    properties: { page: {
      // размеры книжные, ориентация альбомная — docx меняет их местами сам
      size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
      margin: { top: 15 * MM, bottom: 12 * MM, left: 14 * MM, right: 14 * MM },
    }},
    footers: { default: new Footer({ children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: sz(9), color: MUTED })],
    })]})},
    children,
  }],
});

const out = path.join(ROOT, "reports", "Otchet_o_proverke.docx");
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(out, buf);
  console.log("готово:", out, buf.length, "байт");
});
