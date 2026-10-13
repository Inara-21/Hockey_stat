// Реестр внесённых исправлений по ответам проверяющих — отдельный Word.
// Данные готовит tools/report_data.py (поле applied).
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell,
        WidthType, BorderStyle, ShadingType, Footer, PageNumber, VerticalAlign,
        PageOrientation } = require("docx");

const ROOT = path.resolve(__dirname, "..", "..");
const data = JSON.parse(fs.readFileSync(path.join(ROOT, "reports", "report_data.json"), "utf8"));

const FONT = "Calibri", MM = 56.7, CONTENT_W = 15250, sz = pt => pt * 2;
const INK = "1A1F36", ACCENT = "2B2F77", MUTED = "5A5F7A", LINE = "C7CBDC", HEAD_BG = "EEF0F7", ZEBRA = "F7F8FC";
const t = (text, o = {}) => new TextRun({ text, font: FONT, size: sz(o.pt || 10), bold: !!o.b,
  italics: !!o.i, color: o.color || INK, break: o.break || undefined });
const p = (children, o = {}) => new Paragraph({ children: Array.isArray(children) ? children : [children],
  alignment: o.align, spacing: { before: o.before || 0, after: o.after === undefined ? 100 : o.after },
  keepNext: o.keepNext, pageBreakBefore: o.pageBreak });
const border = { style: BorderStyle.SINGLE, size: 4, color: LINE };
const borders = { top: border, bottom: border, left: border, right: border };
const cell = (children, o = {}) => new TableCell({ width: { size: o.w, type: WidthType.DXA }, borders,
  shading: o.bg ? { type: ShadingType.CLEAR, color: "auto", fill: o.bg } : undefined,
  verticalAlign: VerticalAlign.TOP, margins: { top: 60, bottom: 60, left: 90, right: 90 },
  children: Array.isArray(children) ? children : [children] });
const lines = (text, o = {}) => new Paragraph({ spacing: { after: 0 },
  children: String(text || "").split("\n").map((s, i) => t(s, { ...o, break: i ? 1 : undefined })) });
const H1 = (text, o = {}) => new Paragraph({ children: [t(text, { pt: 15, b: true, color: ACCENT })],
  spacing: { before: o.pageBreak ? 0 : 360, after: 140 }, keepNext: true, pageBreakBefore: !!o.pageBreak,
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 6 } } });
const headRow = (w, titles) => new TableRow({ tableHeader: true,
  children: titles.map((x, i) => cell(p(t(x, { b: true, pt: 9 }), { after: 0 }), { w: w[i], bg: HEAD_BG })) });

const cap = s => s.charAt(0).toUpperCase() + s.slice(1);
const children = [];
children.push(p(t("ВНЕСЁННЫЕ ИСПРАВЛЕНИЯ", { pt: 20, b: true, color: ACCENT }), { align: AlignmentType.CENTER, after: 60 }));
children.push(p(t("Чемпионат Регулярной хоккейной лиги 3х3", { pt: 12, color: MUTED }), { align: AlignmentType.CENTER, after: 40 }));
data.corrections_source.forEach(src => children.push(
  p(t(`Основание: ${src}`, { pt: 9, color: MUTED }), { align: AlignmentType.CENTER, after: 40 })));
children.push(p(t("Исходные файлы не изменялись: исправления применяются поверх них.",
  { pt: 9, color: MUTED }), { align: AlignmentType.CENTER, after: 240 }));

const FW = [900, 2600, 4200, 4200, CONTENT_W - 11900];
data.applied.forEach((div, di) => {
  children.push(H1(div.comp, { pageBreak: di > 0 }));
  div.months.forEach(mon => {
    children.push(p(t(cap(mon.label), { pt: 11, b: true }), { before: 200, after: 80, keepNext: true }));
    const rows = [headRow(FW, ["Стр.", "Матч", "Было", "Стало", "Основание"])];
    mon.items.forEach((it, i) => {
      const bg = i % 2 ? ZEBRA : undefined;
      rows.push(new TableRow({ children: [
        cell(p(t(it.row, { pt: 9, b: true }), { after: 0 }), { w: FW[0], bg }),
        cell(lines(it.match, { pt: 9 }), { w: FW[1], bg }),
        cell(p(t(it.before, { pt: 9, color: MUTED }), { after: 0 }), { w: FW[2], bg }),
        cell(p(t(it.after, { pt: 9, b: true }), { after: 0 }), { w: FW[3], bg }),
        cell(p(t(it.note, { pt: 9 }), { after: 0 }), { w: FW[4], bg }),
      ]}));
    });
    children.push(new Table({ columnWidths: FW, width: { size: CONTENT_W, type: WidthType.DXA }, rows }));
  });
});

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: sz(10), color: INK } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
      margin: { top: 15 * MM, bottom: 12 * MM, left: 14 * MM, right: 14 * MM } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: sz(9), color: MUTED })] })] }) },
    children }],
});
const out = path.join(ROOT, "reports", "Vnesennye_ispravleniya.docx");
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(out, buf); console.log("готово:", out, buf.length, "байт"); });
