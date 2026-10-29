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
  italics: !!o.i, color: o.color || INK, break: o.break || undefined, highlight: o.hl ? "yellow" : undefined });
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

const children = [];
// ------------------------------------------------------------ титул
// В отчёте только то, что нужно исправить: файлы и месяцы без ошибок, примечания
// и общие вопросы сюда не выводятся.
const errMonths = data.months.filter(m => data.divisions.some(d => d.others.some(o => o.month === m && o.findings.length)
  || d.teams.length));
children.push(p(t("ОТЧЁТ ОБ ОШИБКАХ", { pt: 20, b: true, color: ACCENT }), { align: AlignmentType.CENTER, after: 60 }));
children.push(p(t(`Чемпионат Регулярной хоккейной лиги 3х3 · ${(errMonths.length ? errMonths : data.months).join(", ")}`,
  { pt: 12, color: MUTED }), { align: AlignmentType.CENTER, after: 40 }));
if (data.corrections_date)
  children.push(p(t(`С последними правками от ${data.corrections_date}`, { pt: 12, b: true, color: ACCENT }), { align: AlignmentType.CENTER, after: 200 }));
children.push(p(t(data.divisions.length ? "Ниже только то, что нужно исправить. Ответ впишите в последний столбец."
  : "Ошибок не осталось.", { pt: 11 }), { after: 120 }));

// ------------------------------------------------------------ дивизионы
const RW = [2600, 7200, CONTENT_W - 9800];
const FW = [900, 2600, 2400, 4400, CONTENT_W - 10300];
data.divisions.forEach(div => {
  children.push(H1(div.comp, { pageBreak: true }));
  if (div.teams.length) {
    children.push(H2("Составы команд"));
    children.push(p(t("Выделены игроки, записанные за разные команды в один игровой день (за разные команды "
      + "можно, но не в один день), а также игроки без номера и с повторяющимся номером. Впишите верный "
      + "состав команды — мы внесём его во все матчи.", { pt: 9, color: MUTED }), { after: 100 }));
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
    if (div.teams.length) children.push(H2("Другие ошибки"));
    div.others.filter(o => o.findings.length).forEach(o => {
      const origin = (data.origins || {})[o.file];
      children.push(p(t(cap(o.month) + (origin ? ` · лист ${origin.sheet}` : ""), { pt: 10, b: true }),
        { before: 120, after: 60, keepNext: true }));
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
// «Сходится?» — выделяем, как в присланном отчёте
grows[0] = new TableRow({ tableHeader: true, children: ["Соревнование", ...data.months.map(cap), "Всего", "Сходится?"].map((x, i) =>
  cell(p(t(x, { b: true, pt: 9, hl: x === "Сходится?" }), { after: 0 }), { w: GW[i], bg: HEAD_BG })) });
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

// ------------------------------------------------------------ недостающие месяцы
if (data.missing.length) {
  children.push(new Paragraph({ children: [t("Ждём данные", { pt: 15, b: true, color: ACCENT, hl: true })],
    spacing: { before: 320, after: 140 }, keepNext: true,
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 6 } } }));
  children.push(p(t("Не хватает матчей за эти месяцы. Ждём файлы — после получения проверим их так же.",
    { pt: 11, b: true }), { after: 80 }));
  data.missing.forEach(x => children.push(p(t(`${x.comp} — ${x.month}: нет данных`, { pt: 10 }), { after: 30 })));
}

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
