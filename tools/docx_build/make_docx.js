// Вёрстка отчёта о проверке данных в Word.
// Данные готовит tools/report_data.py, здесь только оформление.
const fs = require("fs");
const path = require("path");
const d = require("docx");
const { Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow,
        TableCell, WidthType, BorderStyle, ShadingType, Footer, PageNumber,
        VerticalAlign, PageOrientation, PageBreak, VerticalMergeType } = d;

const ROOT = path.resolve(__dirname, "..", "..");
const data = JSON.parse(fs.readFileSync(path.join(ROOT, "reports", "report_data.json"), "utf8"));

const FONT = "Calibri";
const MM = 56.7;
const CONTENT_W = 15250;            // полоса набора, альбомная А4, поля 14 мм
const sz = pt => pt * 2;
const INK = "1A1F36", ACCENT = "2B2F77", MUTED = "5A5F7A";
const LINE = "C7CBDC", HEAD_BG = "EEF0F7", ZEBRA = "F7F8FC";

const t = (text, o = {}) => new TextRun({
  text, font: FONT, size: sz(o.pt || 10), bold: !!o.b, italics: !!o.i,
  color: o.color || INK, break: o.break || undefined });

const p = (children, o = {}) => new Paragraph({
  children: Array.isArray(children) ? children : [children],
  alignment: o.align, spacing: { before: o.before || 0, after: o.after === undefined ? 100 : o.after },
  keepNext: o.keepNext, pageBreakBefore: o.pageBreak });

const border = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border };

const cell = (children, o = {}) => new TableCell({
  width: { size: o.w, type: WidthType.DXA }, borders,
  shading: o.bg ? { type: ShadingType.CLEAR, color: "auto", fill: o.bg } : undefined,
  verticalAlign: VerticalAlign.TOP,
  margins: { top: 60, bottom: 60, left: 90, right: 90 },
  children: Array.isArray(children) ? children : [children] });

const lines = (text, o = {}) => new Paragraph({
  spacing: { after: 0 },
  children: String(text || "").split("\n").map((s, i) => t(s, { ...o, break: i ? 1 : undefined })) });

const H1 = (text, o = {}) => new Paragraph({
  children: [t(text, { pt: 15, b: true, color: ACCENT })],
  spacing: { before: o.pageBreak ? 0 : 360, after: 140 }, keepNext: true,
  pageBreakBefore: !!o.pageBreak,
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 6 } } });

const table = (widths, rows) => new Table({
  columnWidths: widths, width: { size: widths.reduce((a, b) => a + b), type: WidthType.DXA }, rows });

const headRow = (widths, titles) => new TableRow({
  tableHeader: true,
  children: titles.map((x, i) => cell(p(t(x, { b: true, pt: 9 }), { after: 0 }),
                                      { w: widths[i], bg: HEAD_BG })) });

const children = [];

// ---------------------------------------------------------------- титул
children.push(p(t("ОТЧЁТ О ПРОВЕРКЕ ИСХОДНЫХ ДАННЫХ", { pt: 20, b: true, color: ACCENT }),
                { align: AlignmentType.CENTER, after: 60 }));
children.push(p(t("Чемпионат Регулярной хоккейной лиги 3х3", { pt: 12, color: MUTED }),
                { align: AlignmentType.CENTER, after: 40 }));
children.push(p(t(`Дата проверки: ${data.date}`, { pt: 10, color: MUTED }),
                { align: AlignmentType.CENTER, after: 80 }));
children.push(p(t("Где исправление следует из самих данных, оно предложено в последнем столбце "
                  + "таблицы — подтвердите его или отклоните.", { pt: 9, color: MUTED }),
                { align: AlignmentType.CENTER, after: 240 }));

// ---------------------------------------------------------------- вопросы
if (data.questions.length) {
  children.push(H1("Вопросы ко всем дивизионам"));
  const QW = [3800, CONTENT_W - 3800 - 4600, 4600];
  data.questions.forEach(q => {
    children.push(p(t(q.ask, { pt: 11, b: true }), { before: 240, after: 100, keepNext: true }));
    const rows = [headRow(QW, ["Соревнование, месяц", "Строки", "Ответ организатора"])];
    q.items.forEach((it, i) => {
      const bg = i % 2 ? ZEBRA : undefined;
      rows.push(new TableRow({ children: [
        cell(lines(`${it.comp}\n${it.period}`, { pt: 9 }), { w: QW[0], bg }),
        cell(p(t(it.rows || "—", { pt: 9 }), { after: 0 }), { w: QW[1], bg }),
        new TableCell({
          width: { size: QW[2], type: WidthType.DXA }, borders,
          verticalMerge: i === 0 ? VerticalMergeType.RESTART : VerticalMergeType.CONTINUE,
          margins: { top: 60, bottom: 60, left: 90, right: 90 },
          children: [p(t("", { pt: 9 }), { after: 0 })] }),
      ]}));
    });
    children.push(table(QW, rows));
  });
}

// ---------------------------------------------------------------- дивизионы
const FW = [900, 2600, 2400, 4400, CONTENT_W - 10300];
data.divisions.forEach(div => {
  children.push(H1(div.comp, { pageBreak: true }));
  children.push(p(t("В последнем столбце впишите, как исправляем. Если там уже есть вопрос — "
                    + "ответьте на него.", { pt: 9, color: MUTED }), { after: 160 }));

  div.months.forEach(mon => {
    children.push(p(t(mon.label.charAt(0).toUpperCase() + mon.label.slice(1),
                      { pt: 11, b: true }), { before: 200, after: 80, keepNext: true }));
    if (!mon.findings.length) {
      children.push(p(t("Ошибок не найдено.", { pt: 10, i: true, color: "2E7D32" })));
    } else {
      const rows = [headRow(FW, ["Стр.", "Матч", "Счёт", "Что не так", "Как исправляем"])];
      mon.findings.forEach((f, i) => {
        const bg = i % 2 ? ZEBRA : undefined;
        rows.push(new TableRow({ children: [
          cell(p(t(f.row || "—", { pt: 9, b: true }), { after: 0 }), { w: FW[0], bg }),
          cell(lines(f.match || f.what, { pt: 9 }), { w: FW[1], bg }),
          cell(lines(f.score || "—", { pt: 9, color: MUTED }), { w: FW[2], bg }),
          cell(p(t(f.text, { pt: 9 }), { after: 0 }), { w: FW[3], bg }),
          cell(p(t(f.fix || "", { pt: 9, i: true, color: ACCENT }), { after: 0 }), { w: FW[4] }),
        ]}));
      });
      children.push(table(FW, rows));
    }

  });
});

// ---------------------------------------------------------------- количество игр
children.push(H1("Количество игр", { pageBreak: true }));
{
  const first = 6000;
  const rest = Math.floor((CONTENT_W - first) / (data.months.length + 1));
  const GW = [first, ...data.months.map(() => rest), CONTENT_W - first - rest * data.months.length];
  const cap = s => s.charAt(0).toUpperCase() + s.slice(1);
  const num = (v, o = {}) => p(t(String(v), { pt: 10, ...o }), { align: AlignmentType.RIGHT, after: 0 });
  const rows = [headRow(GW, ["Соревнование", ...data.months.map(cap), "Всего"])];
  data.games_table.forEach((r, i) => {
    const bg = i % 2 ? ZEBRA : undefined;
    rows.push(new TableRow({ children: [
      cell(p(t(r.comp, { pt: 10 }), { after: 0 }), { w: GW[0], bg }),
      ...data.months.map((m, j) => cell(num(r.by_month[m] || 0), { w: GW[j + 1], bg })),
      cell(num(r.total, { b: true }), { w: GW[GW.length - 1], bg }),
    ]}));
  });
  rows.push(new TableRow({ children: [
    cell(p(t("Итого", { pt: 10, b: true }), { after: 0 }), { w: GW[0], bg: HEAD_BG }),
    ...data.months.map((m, j) => cell(num(data.games_table.reduce((a, r) => a + (r.by_month[m] || 0), 0), { b: true }),
                                      { w: GW[j + 1], bg: HEAD_BG })),
    cell(num(data.games, { b: true }), { w: GW[GW.length - 1], bg: HEAD_BG }),
  ]}));
  children.push(table(GW, rows));
}

// ---------------------------------------------------------------- документ
const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: sz(10), color: INK } } } },
  sections: [{
    properties: { page: {
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
