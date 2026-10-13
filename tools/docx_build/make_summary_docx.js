// Сводный отчёт по исправленным файлам. Данные — tools/summary_report.py.
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell,
        WidthType, BorderStyle, ShadingType, Footer, PageNumber, VerticalAlign,
        PageOrientation } = require("docx");

const ROOT = path.resolve(__dirname, "..", "..");
const data = JSON.parse(fs.readFileSync(path.join(ROOT, "reports", "svodny_data.json"), "utf8"));

const FONT = "Calibri", MM = 56.7, CONTENT_W = 15250, sz = pt => pt * 2;
const INK = "1A1F36", ACCENT = "2B2F77", MUTED = "5A5F7A", LINE = "C7CBDC", HEAD_BG = "EEF0F7",
      ZEBRA = "F7F8FC", GREEN = "2E7D32", MARK_BG = "FFF2A8";
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
  spacing: { before: o.pageBreak ? 0 : 320, after: 140 }, keepNext: true, pageBreakBefore: !!o.pageBreak,
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 6 } } });
const H2 = text => p(t(text, { pt: 12, b: true }), { before: 220, after: 80, keepNext: true });
const headRow = (w, titles) => new TableRow({ tableHeader: true,
  children: titles.map((x, i) => cell(p(t(x, { b: true, pt: 9 }), { after: 0 }), { w: w[i], bg: HEAD_BG })) });
const table = (w, rows) => new Table({ columnWidths: w, width: { size: w.reduce((a, b) => a + b), type: WidthType.DXA }, rows });
const cap = s => s.charAt(0).toUpperCase() + s.slice(1);
const status = f => f.clean ? (f.answered ? "Правки проверяющих внесены — ошибок нет" : "Ошибок нет")
                            : (f.answered ? "Правки проверяющих внесены, остались ошибки"
                                          : "Ответов проверяющих пока нет — ошибки в разделе дивизиона");

const children = [];
// ------------------------------------------------------------ титул
children.push(p(t("СВОДНЫЙ ОТЧЁТ ОБ ОШИБКАХ", { pt: 20, b: true, color: ACCENT }), { align: AlignmentType.CENTER, after: 60 }));
children.push(p(t(`Чемпионат Регулярной хоккейной лиги 3х3 · ${data.months.join(", ")}`, { pt: 12, color: MUTED }), { align: AlignmentType.CENTER, after: 40 }));
if (data.corrections_date)
  children.push(p(t(`С последними правками от ${data.corrections_date}`, { pt: 12, b: true, color: ACCENT }), { align: AlignmentType.CENTER, after: 40 }));
children.push(p(t(`Дата: ${data.date}`, { pt: 10, color: MUTED }), { align: AlignmentType.CENTER, after: 200 }));

// ------------------------------------------------------------ итог по файлам
children.push(H1("Итог по файлам"));
const clean = data.files.filter(f => f.clean);
if (clean.length) {
  children.push(p(t("В этих файлах ошибок нет:", { pt: 10, b: true, color: GREEN }), { after: 60 }));
  clean.forEach(f => children.push(p(t(`${f.file} — ${f.comp}, ${f.month}`, { pt: 10 }), { after: 30 })));
}
const FW0 = [3600, 5400, 2400, CONTENT_W - 11400];
const frows = [headRow(FW0, ["Файл", "Соревнование", "Месяц", "Итог"])];
data.files.forEach((f, i) => frows.push(new TableRow({ children: [
  cell(p(t(f.file, { pt: 9 }), { after: 0 }), { w: FW0[0], bg: i % 2 ? ZEBRA : undefined }),
  cell(p(t(f.comp, { pt: 9 }), { after: 0 }), { w: FW0[1], bg: i % 2 ? ZEBRA : undefined }),
  cell(p(t(f.month, { pt: 9 }), { after: 0 }), { w: FW0[2], bg: i % 2 ? ZEBRA : undefined }),
  cell(p(t(status(f), { pt: 9, b: true, color: f.clean ? GREEN : "B45309" }), { after: 0 }), { w: FW0[3], bg: i % 2 ? ZEBRA : undefined }),
]})));
children.push(p([], { after: 80 }));
children.push(table(FW0, frows));

// ------------------------------------------------------------ примечания
children.push(H1("Примечания"));
{
  const done = data.divisions_all.filter(x => x.answered).map(x => x.comp);
  const wait = data.divisions_all.filter(x => !x.answered).map(x => x.comp);
  const parts = [];
  if (done.length) parts.push(`Ответы проверяющих внесены${data.corrections_date ? " " + data.corrections_date : ""} по: ${done.join(", ")}. `
    + "По ним в отчёте только то, что осталось.");
  if (wait.length) parts.push(`Ответов пока нет по: ${wait.join(", ")} — указаны все найденные ошибки.`);
  children.push(p([t("1. ", { b: true }), t(parts.join(" "))], { after: 120 }));
}
children.push(p([t("2. Ошибки в составах систематические. ", { b: true }),
  t("Похоже, составы команд копировались из матча в матч, поэтому одна и та же ошибка повторяется во "
    + "всех матчах команды. Исправлять каждый матч не нужно — достаточно прислать верные составы команд "
    + "с игровыми номерами: они приведены в разделах дивизионов в таблицах Составы команд, ошибочные "
    + "строки выделены. Исправления во все матчи мы внесём сами.")], { after: 120 }));

// ------------------------------------------------------------ дивизионы
const RW = [2600, 7200, CONTENT_W - 9800];
const FW = [900, 2600, 2400, 4400, CONTENT_W - 10300];
data.divisions.forEach(div => {
  children.push(H1(div.comp, { pageBreak: true }));
  if (div.teams.length) {
    children.push(H2("Составы команд"));
    children.push(p(t("Выделены игроки, записанные сразу за несколько команд (выступать можно только за одну), "
      + "а также игроки без номера и с повторяющимся номером. Впишите верный состав команды — мы внесём "
      + "его во все матчи.", { pt: 9, color: MUTED }), { after: 100 }));
    const rows = [headRow(RW, ["Команда", "Состав в файле", "Верный состав (ФИО и номер)"])];
    div.teams.forEach((tm, i) => {
      const paras = [];
      tm.rosters.forEach((r, k) => {
        paras.push(p(t(r.label, { pt: 9, i: true, color: MUTED }), { before: k ? 100 : 0, after: 20 }));
        r.players.forEach(pl => paras.push(new Paragraph({ spacing: { after: 0 },
          shading: pl.flag ? { type: ShadingType.CLEAR, color: "auto", fill: MARK_BG } : undefined,
          children: [t(pl.text, { pt: 9, b: pl.flag })].concat(pl.flag ? [t(`  — ${pl.note}`, { pt: 9, color: "B45309" })] : []) })));
      });
      rows.push(new TableRow({ children: [
        cell(p(t(tm.team, { pt: 10, b: true }), { after: 0 }), { w: RW[0], bg: i % 2 ? ZEBRA : undefined }),
        cell(paras, { w: RW[1], bg: i % 2 ? ZEBRA : undefined }),
        cell(p(t("", { pt: 9 }), { after: 0 }), { w: RW[2] }),
      ]}));
    });
    children.push(table(RW, rows));
  }
  if (div.others.some(o => o.findings.length)) {
    children.push(H2(div.teams.length ? "Другие ошибки" : "Ошибки"));
    div.others.forEach(o => {
      children.push(p(t(cap(o.month), { pt: 10, b: true }), { before: 120, after: 60, keepNext: true }));
      if (!o.findings.length) {
        children.push(p(t("Ошибок нет.", { pt: 10, i: true, color: GREEN })));
        return;
      }
      const rows = [headRow(FW, ["Стр.", "Матч", "Счёт", "Что не так", "Как исправляем"])];
      o.findings.forEach((f, i) => {
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
    });
  }
});

// ------------------------------------------------------------ количество игр
children.push(H1("Сверка количества игр", { pageBreak: true }));
children.push(p(t("Просим сверить количество игр по каждому дивизиону за каждый месяц и общую сумму, "
  + "чтобы убедиться, что ни одна игра не потеряна.", { pt: 10 }), { after: 120 }));
const first = 5200, last = 2600;
const mid = Math.floor((CONTENT_W - first - last) / (data.months.length + 1));
const GW = [first, ...data.months.map(() => mid), CONTENT_W - first - last - mid * data.months.length, last];
const num = (v, o = {}) => p(t(String(v), { pt: 10, ...o }), { align: AlignmentType.RIGHT, after: 0 });
const grows = [headRow(GW, ["Соревнование", ...data.months.map(cap), "Всего", "Сходится?"])];
data.games_table.forEach((r, i) => {
  const bg = i % 2 ? ZEBRA : undefined;
  grows.push(new TableRow({ children: [
    cell(p(t(r.comp, { pt: 10 }), { after: 0 }), { w: GW[0], bg }),
    ...data.months.map((m, j) => cell(num(r.cells[m] || "—"), { w: GW[j + 1], bg })),
    cell(num(r.total, { b: true }), { w: GW[GW.length - 2], bg }),
    cell(p(t("", { pt: 10 }), { after: 0 }), { w: GW[GW.length - 1] }),
  ]}));
});
grows.push(new TableRow({ children: [
  cell(p(t("Итого", { pt: 10, b: true }), { after: 0 }), { w: GW[0], bg: HEAD_BG }),
  ...data.months.map((m, j) => cell(num(data.games_total[m], { b: true }), { w: GW[j + 1], bg: HEAD_BG })),
  cell(num(data.games_grand, { b: true }), { w: GW[GW.length - 2], bg: HEAD_BG }),
  cell(p(t("", { pt: 10 }), { after: 0 }), { w: GW[GW.length - 1], bg: HEAD_BG }),
]}));
children.push(table(GW, grows));
if (data.misplaced_total)
  children.push(p(t(`В итог не включены матчи под вопросом: ${data.misplaced_total} — см. раздел Конференция ЗАПАД (2 Дивизион).`,
    { pt: 9, color: MUTED }), { before: 80 }));

// ------------------------------------------------------------ сентябрь
children.push(H1("Ждём данные за сентябрь"));
children.push(p(t("Не хватает матчей за сентябрь 2026 по всем дивизионам. Ждём файлы — после получения "
  + "проверим их так же.", { pt: 11, b: true }), { after: 80 }));
data.comps.forEach(c => children.push(p(t(`${c} — сентябрь 2026: нет данных`, { pt: 10 }), { after: 30 })));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: sz(10), color: INK } } } },
  sections: [{ properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
      margin: { top: 15 * MM, bottom: 12 * MM, left: 14 * MM, right: 14 * MM } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: sz(9), color: MUTED })] })] }) },
    children }],
});
const out = path.join(ROOT, "reports", "Svodny_otchet.docx");
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(out, buf); console.log("готово:", out, buf.length, "байт"); });
