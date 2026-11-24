/**
 * Сборка файла «Сведения об авторах» для конференции КТИ (филиал) ВолгГТУ.
 *
 * Запуск:  node Сборка/build-authors.js "Текст/Сведения об авторах - Камышин 2026.md" "Готовое/Сведения об авторах.docx"
 *
 * Оргкомитет требует сведения на каждого автора в отдельности, но в одном
 * файле, по образцу из информационного сообщения: таблица «поле — значение».
 * Источник — пары «Ключ: значение», авторы разделены строкой «---».
 */
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, TableLayoutType,
} = require('docx');

const FONT = 'Times New Roman';
const SIZE = 24;                       // полуразмеры = 12 пт
const TEXT_WIDTH = 9355;               // поля набора А4 при полях 3/1,5 см
const LABEL_WIDTH = 3400;

const SOURCE = process.argv[2];
const TARGET = process.argv[3];

const HEADER = [
  'Всероссийская заочная научно-практическая конференция',
  '«Инновационные технологии в обучении и производстве»',
  '26–27 ноября 2026 года',
  'Камышин, Россия',
];

const text = (value, bold) => new Paragraph({
  spacing: { line: 240, before: 20, after: 20 },
  children: [new TextRun({ text: value, font: FONT, size: SIZE, bold: !!bold })],
});

const cell = (value, width, bold) => new TableCell({
  width: { size: width, type: WidthType.DXA },
  margins: { top: 40, bottom: 40, left: 80, right: 80 },
  children: [text(value, bold)],
});

function tableOf(block) {
  const rows = [];
  for (const line of block.split('\n')) {
    const at = line.indexOf(':');
    if (at <= 0) continue;
    rows.push(new TableRow({
      children: [
        cell(line.slice(0, at).trim(), LABEL_WIDTH, true),
        cell(line.slice(at + 1).trim(), TEXT_WIDTH - LABEL_WIDTH, false),
      ],
    }));
  }
  return new Table({
    width: { size: TEXT_WIDTH, type: WidthType.DXA },
    columnWidths: [LABEL_WIDTH, TEXT_WIDTH - LABEL_WIDTH],
    layout: TableLayoutType.FIXED,
    rows,
  });
}

const blocks = fs.readFileSync(SOURCE, 'utf8').split(/^---$/m)
  .map((block) => block.trim()).filter(Boolean);

const children = [];
for (const line of HEADER) {
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 0, after: 0 },
    children: [new TextRun({ text: line, font: FONT, size: SIZE, bold: true })],
  }));
}
for (const block of blocks) {
  children.push(text(''), tableOf(block));
}

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, bottom: 1134, left: 1701, right: 850 },
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(TARGET, buffer);
  const name = TARGET.split(/[\\/]/).pop();
  console.log(`собрано: ${name} (${blocks.length} автора)`);
});
