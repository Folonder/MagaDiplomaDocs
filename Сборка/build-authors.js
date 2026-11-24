/**
 * Сборка файлов «Сведения об авторе» для конференции КТИ (филиал) ВолгГТУ.
 *
 * Запуск:  node Сборка/build-authors.js "Текст/Сведения об авторах - Камышин 2026.md" "Готовое"
 *
 * Сведения подаются на каждого автора ОТДЕЛЬНЫМ файлом, поэтому на каждый
 * блок источника выпускается свой .docx, названный по фамилии автора.
 * Оформление повторяет принятый образец: А4, поля 2,0/2,0/2,5/2,5,
 * Times New Roman 12, шапка по центру и таблица «поле — значение».
 * Источник — пары «Ключ: значение», авторы разделены строкой «---».
 */
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, TableLayoutType, convertMillimetersToTwip,
} = require('docx');

const FONT = 'Times New Roman';
const SIZE = 24;                                      // полуразмеры = 12 пт
const PAGE = { width: convertMillimetersToTwip(210), height: convertMillimetersToTwip(297) };
const MARGIN = {
  top: convertMillimetersToTwip(20), bottom: convertMillimetersToTwip(20),
  left: convertMillimetersToTwip(25), right: convertMillimetersToTwip(25),
};
const TEXT_WIDTH = PAGE.width - MARGIN.left - MARGIN.right;
const LABEL_WIDTH = convertMillimetersToTwip(55);

const HEADER = [
  'XX Всероссийской заочной научно-практической конференции',
  '«Инновационные технологии в обучении и производстве»',
  '26-27 ноября 2026 года',
  'Камышин, Россия',
];

const SOURCE = process.argv[2];
const OUT_DIR = process.argv[3];

const centered = (value, bold) => new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { line: 240, before: 0, after: 0 },
  children: [new TextRun({ text: value, font: FONT, size: SIZE, bold: !!bold })],
});

const cell = (value, width) => new TableCell({
  width: { size: width, type: WidthType.DXA },
  margins: { top: 40, bottom: 40, left: 80, right: 80 },
  children: [new Paragraph({
    spacing: { line: 240, before: 20, after: 20 },
    children: [new TextRun({ text: value, font: FONT, size: SIZE })],
  })],
});

/** Разбор блока в пары «поле — значение» с сохранением порядка. */
function fieldsOf(block) {
  const fields = [];
  for (const line of block.split('\n')) {
    const at = line.indexOf(':');
    if (at > 0) fields.push([line.slice(0, at).trim(), line.slice(at + 1).trim()]);
  }
  return fields;
}

/** «Ролдугин Олег Денисович» -> «Ролдугин О.Д.» для имени файла. */
function shortName(fullName) {
  const [family, first, middle] = fullName.split(/\s+/);
  const initials = [first, middle].filter(Boolean).map((part) => `${part[0]}.`).join('');
  return initials ? `${family} ${initials}` : family;
}

const blocks = fs.readFileSync(SOURCE, 'utf8').split(/^---$/m)
  .map((block) => block.trim()).filter(Boolean);

for (const block of blocks) {
  const fields = fieldsOf(block);
  const author = (fields.find(([label]) => label.startsWith('Ф.И.О.')) || [])[1] || 'Автор';

  const doc = new Document({
    sections: [{
      properties: { page: { size: PAGE, margin: MARGIN } },
      children: [
        centered('СВЕДЕНИЯ ОБ АВТОРЕ', true),
        ...HEADER.map((line) => centered(line)),
        centered(''),
        new Table({
          width: { size: TEXT_WIDTH, type: WidthType.DXA },
          columnWidths: [LABEL_WIDTH, TEXT_WIDTH - LABEL_WIDTH],
          layout: TableLayoutType.FIXED,
          rows: fields.map(([label, value]) => new TableRow({
            children: [cell(label, LABEL_WIDTH), cell(value, TEXT_WIDTH - LABEL_WIDTH)],
          })),
        }),
      ],
    }],
  });

  // Инициалы кончаются точкой, поэтому фамилия берётся в скобки:
  // иначе в имени файла получается двойная точка перед расширением.
  const target = path.join(OUT_DIR, `Сведения об авторе (${shortName(author)}).docx`);
  // eslint-disable-next-line no-loop-func
  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(target, buffer);
    console.log(`собрано: ${path.basename(target)}`);
  });
}
