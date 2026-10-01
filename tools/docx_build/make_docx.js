// Вёрстка отчёта о проверке данных в Word.
// Данные готовит tools/report_data.py, здесь только оформление.
const fs = require("fs");
const path = require("path");
const d = require("docx");
const { Document, Packer, Paragraph, TextRun, AlignmentType, Table, TableRow,
        TableCell, WidthType, BorderStyle, ShadingType, Footer, PageNumber,
        VerticalAlign, PageOrientation, PageBreak } = d;

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
children.push(p(t(`Дата проверки: ${data.date}   ·   проверено матчей: ${data.games}   ·   `
                  + `найдено ошибок: ${data.errors}`,
                  { pt: 10, color: MUTED }),
                { align: AlignmentType.CENTER, after: 80 }));
children.push(p(t("Исходные файлы не изменялись. Если исправление ошибки однозначно следует "
                  + "из самих данных, оно предложено в последнем столбце таблицы с вопросом, "
                  + "подтверждаете ли вы его.", { pt: 9, color: MUTED }),
                { align: AlignmentType.CENTER, after: 240 }));

// ---------------------------------------------------------------- вопросы
if (data.questions.length) {
  children.push(H1("Вопросы ко всем дивизионам"));
  children.push(p(t("Это не ошибки. Данные записаны так во многих матчах, и мы не знаем, "
                    + "допускается ли это правилами соревнования. Просим ответить по каждому "
                    + "пункту.", { pt: 10 }), { after: 160 }));

  const QW = [3800, CONTENT_W - 3800 - 4600, 4600];
  data.questions.forEach(q => {
    children.push(p(t(q.ask, { pt: 11, b: true }), { before: 240, after: 100, keepNext: true }));
    const rows = [headRow(QW, ["Соревнование, месяц", "Что найдено", "Ответ организатора"])];
    q.items.forEach((it, i) => {
      const bg = i % 2 ? ZEBRA : undefined;
      rows.push(new TableRow({ children: [
        cell(lines(`${it.comp}\n${it.period}`, { pt: 9 }), { w: QW[0], bg }),
        cell(p(t(it.text, { pt: 9 }), { after: 0 }), { w: QW[1], bg }),
        cell(p(t("", { pt: 9 }), { after: 0 }), { w: QW[2] }),
      ]}));
    });
    children.push(table(QW, rows));
  });
}

// ---------------------------------------------------------------- сводка
children.push(H1("Сводка: сколько проверено и сколько найдено"));
children.push(p(t("Сверьте количество матчей со своими данными. Если расходится — значит, "
                  + "в исходной таблице не все игры или есть лишние.", { pt: 10, color: MUTED }),
                { after: 140 }));

const SW = [6000, 3400, 3000, CONTENT_W - 12400];
const sumRows = [headRow(SW, ["Соревнование", "Месяц", "Матчей", "Ошибок"])];
data.summary.forEach((r, i) => {
  const bg = i % 2 ? ZEBRA : undefined;
  sumRows.push(new TableRow({ children: [
    cell(p(t(r.comp, { pt: 10 }), { after: 0 }), { w: SW[0], bg }),
    cell(p(t(r.month, { pt: 10 }), { after: 0 }), { w: SW[1], bg }),
    cell(p(t(String(r.games), { pt: 10 }), { align: AlignmentType.RIGHT, after: 0 }), { w: SW[2], bg }),
    cell(p(t(String(r.errors), { pt: 10 }), { align: AlignmentType.RIGHT, after: 0 }), { w: SW[3], bg }),
  ]}));
});
sumRows.push(new TableRow({ children: [
  cell(p(t("ИТОГО", { b: true, pt: 10 }), { after: 0 }), { w: SW[0], bg: HEAD_BG }),
  cell(p(t("", { pt: 10 }), { after: 0 }), { w: SW[1], bg: HEAD_BG }),
  cell(p(t(String(data.games), { b: true, pt: 10 }), { align: AlignmentType.RIGHT, after: 0 }), { w: SW[2], bg: HEAD_BG }),
  cell(p(t(String(data.errors), { b: true, pt: 10 }), { align: AlignmentType.RIGHT, after: 0 }), { w: SW[3], bg: HEAD_BG }),
]}));
children.push(table(SW, sumRows));

// ---------------------------------------------------------------- дивизионы
const FW = [900, 2600, 2400, 4400, CONTENT_W - 10300];
data.divisions.forEach(div => {
  children.push(H1(div.comp, { pageBreak: true }));
  children.push(p([t(`Проверено матчей: ${div.games}.   Найдено ошибок: ${div.errors}.`
                     + (div.proposed ? `   Из них с готовым исправлением: ${div.proposed}.` : ""),
                     { pt: 11, b: true })], { after: 60 }));
  children.push(p(t("Последняя колонка — для вашего ответа: напротив каждой строки впишите, "
                    + "как исправляем ошибку. Где там уже предложено исправление, напишите, "
                    + "подтверждаете вы его или нет.", { pt: 9, color: MUTED }), { after: 160 }));

  div.months.forEach(mon => {
    children.push(p(t(`${mon.label} — матчей ${mon.games}, ошибок ${mon.errors}`,
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
