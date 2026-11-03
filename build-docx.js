/**
 * Сборка раздела ПЗ из markdown в .docx по оформлению кафедры ПОАС ВолгГТУ.
 *
 * Запуск:  npm install docx  (один раз, в любой папке с package.json)
 *          node "Документы/build-docx.js" "Документы/Глава 1 - Анализ задачи и методов.md" "Документы/Раздел 1.docx"
 *
 * Поля 3/1,5/2/2 см, Times New Roman 14, интервал 1,5, абзац 1,25 см,
 * по ширине, только чёрный. Заголовки с абзацного отступа, по левому краю,
 * номер входит в текст заголовка. Разделы с новой страницы не начинаются.
 * Номер страницы внизу по центру. Широкая таблица выносится на альбомную
 * страницу отдельным разделом Word.
 */
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Tab,
  AlignmentType, WidthType, TabStopType, PageOrientation, PageNumber,
  Footer, HeadingLevel, BorderStyle,
} = require('docx');

// ── Параметры страницы и текста (DXA) ────────────────────────────────────────
const PAGE = { width: 11906, height: 16838 };
const MARGIN = { top: 1134, bottom: 1134, left: 1701, right: 850, header: 567, footer: 567 };
const TEXT_WIDTH = PAGE.width - MARGIN.left - MARGIN.right;            // 9355
const TEXT_WIDTH_LAND = PAGE.height - MARGIN.left - MARGIN.right;      // 14287
const FONT = 'Times New Roman';
const SIZE = 28;          // полуразмеры = 14 пт
const LINE = 360;         // полуторный интервал
const FIRST_LINE = 709;   // абзацный отступ 1,25 см
const WIDE_TABLE = 'Таблица 1.5';   // не помещается на книжную страницу

const SOURCE = process.argv[2];
const TARGET = process.argv[3];

// ── Разбор markdown в блоки ──────────────────────────────────────────────────
function parse(markdown) {
  const lines = markdown.split('\n');
  const blocks = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];
    const text = line.trim();

    if (!text) { i += 1; continue; }

    if (text.startsWith('### ')) { blocks.push({ kind: 'h3', text: text.slice(4) }); i += 1; continue; }
    if (text.startsWith('## '))  { blocks.push({ kind: 'h2', text: text.slice(3) }); i += 1; continue; }
    if (text.startsWith('# '))   { blocks.push({ kind: 'h1', text: text.slice(2) }); i += 1; continue; }

    if (/^Таблица\s/.test(text)) { blocks.push({ kind: 'tableCaption', text: dashInCaption(text) }); i += 1; continue; }

    if (text.startsWith('|')) {
      const rows = [];
      while (i < lines.length && lines[i].trim().startsWith('|')) {
        const cells = lines[i].trim().slice(1, -1).split('|').map(c => c.trim());
        if (!cells.every(c => /^:?-{2,}:?$/.test(c))) rows.push(cells);   // строка-разделитель
        i += 1;
      }
      blocks.push({ kind: 'table', rows });
      continue;
    }

    if (text.startsWith('[Рисунок')) {
      const whole = text.replace(/^\[/, '').replace(/\]$/, '');
      const [caption, content] = splitFigure(whole);
      blocks.push({ kind: 'figure', caption: dashInCaption(caption), content });
      i += 1;
      continue;
    }

    if (/\(\d+\.\d+\)\s*$/.test(text) && !/^\d+\.\s/.test(text)) {
      const number = text.match(/\((\d+\.\d+)\)\s*$/)[0];
      blocks.push({ kind: 'formula', text: text.slice(0, text.length - number.length).trim(), number });
      i += 1;
      continue;
    }

    if (/^\d+\.\s/.test(text)) {
      blocks.push({ kind: 'item', text: text.replace(/^(\d+)\.\s/, '$1) ') });
      i += 1;
      continue;
    }

    if (text.startsWith('- ')) { blocks.push({ kind: 'bullet', text: text.slice(2) }); i += 1; continue; }

    blocks.push({ kind: 'para', text });
    i += 1;
  }
  return blocks;
}

function splitFigure(whole) {
  const at = whole.indexOf('Содержание:');
  if (at < 0) return [whole.trim(), ''];
  return [whole.slice(0, at).trim().replace(/\.$/, ''), whole.slice(at + 'Содержание:'.length).trim()];
}

/** В подписях методичка требует дефис, а не тире. */
function dashInCaption(text) {
  return text.replace(/^(Таблица|Рисунок)(\s[\d.]+)\s[–—-]\s/, '$1$2 - ');
}

// ── Построение абзацев ───────────────────────────────────────────────────────
const runs = (text, extra = {}) => {
  const out = [];
  for (const piece of text.split(/(\*\*[^*]+\*\*)/)) {
    if (!piece) continue;
    const bold = piece.startsWith('**');
    out.push(new TextRun({
      text: bold ? piece.slice(2, -2) : piece.replace(/`/g, ''),
      bold, font: FONT, size: SIZE, color: '000000', ...extra,
    }));
  }
  return out;
};

const empty = () => new Paragraph({ children: runs('') });

const body = (text, options = {}) => new Paragraph({
  alignment: options.alignment ?? AlignmentType.JUSTIFIED,
  indent: { firstLine: options.firstLine ?? FIRST_LINE },
  spacing: { line: LINE, before: 0, after: 0 },
  children: runs(text),
});

const heading = (text, level) => new Paragraph({
  heading: level,
  alignment: AlignmentType.LEFT,
  indent: { firstLine: FIRST_LINE },
  spacing: { line: LINE, before: 0, after: 0 },
  keepNext: true,
  children: runs(text),
});

const caption = (text, alignment) => new Paragraph({
  alignment,
  indent: { firstLine: 0 },
  spacing: { line: LINE, before: 0, after: 0 },
  keepNext: alignment === AlignmentType.LEFT,
  children: runs(text),
});

const formula = (text, number, width) => new Paragraph({
  indent: { firstLine: 0 },
  spacing: { line: LINE, before: 0, after: 0 },
  tabStops: [
    { type: TabStopType.CENTER, position: Math.round(width / 2) },
    { type: TabStopType.RIGHT, position: width },
  ],
  children: [
    new TextRun({ children: [new Tab()], font: FONT, size: SIZE }),
    ...runs(text),
    new TextRun({ children: [new Tab()], font: FONT, size: SIZE }),
    ...runs(number),
  ],
});

/** Место под ещё не нарисованный рисунок: рамка с описанием и подпись. */
const figurePlaceholder = (block) => {
  const frame = new Paragraph({
    alignment: AlignmentType.CENTER,
    indent: { firstLine: 0 },
    spacing: { line: LINE, before: 0, after: 0 },
    border: {
      top: { style: BorderStyle.DASHED, size: 6, color: '000000' },
      bottom: { style: BorderStyle.DASHED, size: 6, color: '000000' },
      left: { style: BorderStyle.DASHED, size: 6, color: '000000' },
      right: { style: BorderStyle.DASHED, size: 6, color: '000000' },
    },
    children: runs(`[место для рисунка] ${block.content}`),
  });
  return [frame, caption(block.caption, AlignmentType.CENTER)];
};

const table = (rows, width) => {
  const columns = rows[0].length;
  const first = columns > 6 ? Math.round(width * 0.17) : Math.round(width / columns);
  const rest = Math.floor((width - first) / (columns - 1));
  const widths = [first, ...Array(columns - 1).fill(rest)];
  widths[columns - 1] += width - widths.reduce((a, b) => a + b, 0);

  return new Table({
    columnWidths: widths,
    width: { size: width, type: WidthType.DXA },
    rows: rows.map((cells, index) => new TableRow({
      tableHeader: index === 0,
      children: cells.map((cell, column) => new TableCell({
        width: { size: widths[column], type: WidthType.DXA },
        children: [new Paragraph({
          alignment: column === 0 && columns > 2 ? AlignmentType.LEFT : AlignmentType.CENTER,
          indent: { firstLine: 0 },
          spacing: { line: LINE, before: 0, after: 0 },
          children: runs(cell),
        })],
      })),
    })),
  });
};

// ── Сборка потока элементов ──────────────────────────────────────────────────
function render(blocks, width) {
  const out = [];
  blocks.forEach((block, index) => {
    const next = blocks[index + 1];
    const previous = blocks[index - 1];

    if (block.kind.startsWith('h')) {
      if (previous) { out.push(empty(), empty()); }          // 2 интервала перед заголовком
      const level = { h1: HeadingLevel.HEADING_1, h2: HeadingLevel.HEADING_2, h3: HeadingLevel.HEADING_3 }[block.kind];
      out.push(heading(block.text, level));
      out.push(empty());
      if (next && !next.kind.startsWith('h')) out.push(empty());   // 2 интервала до текста
      return;
    }

    if (block.kind === 'tableCaption') { out.push(empty(), caption(block.text, AlignmentType.LEFT)); return; }
    if (block.kind === 'table') { out.push(table(block.rows, width), empty()); return; }
    if (block.kind === 'figure') { out.push(empty(), ...figurePlaceholder(block), empty()); return; }
    if (block.kind === 'formula') { out.push(empty(), formula(block.text, block.number, width), empty()); return; }
    if (block.kind === 'item' || block.kind === 'bullet') { out.push(body(block.text)); return; }
    out.push(body(block.text));
  });
  return out;
}

/** Широкая таблица и её подпись выносятся в альбомный раздел. */
function split(blocks) {
  const parts = [];
  let current = { landscape: false, blocks: [] };
  for (let i = 0; i < blocks.length; i += 1) {
    const block = blocks[i];
    const wide = block.kind === 'tableCaption' && block.text.startsWith(WIDE_TABLE);
    if (wide) {
      parts.push(current);
      current = { landscape: true, blocks: [block, blocks[i + 1]] };
      parts.push(current);
      current = { landscape: false, blocks: [] };
      i += 1;
      continue;
    }
    current.blocks.push(block);
  }
  parts.push(current);
  return parts.filter(part => part.blocks.length);
}

const footer = () => new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    indent: { firstLine: 0 },
    children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SIZE, color: '000000' })],
  })],
});

const markdown = fs.readFileSync(SOURCE, 'utf8');
const chapter = markdown.split('## Источники')[0];        // список источников собирается отдельно
const parts = split(parse(chapter));

const headingStyle = {
  run: { font: FONT, size: SIZE, bold: false, color: '000000' },
  paragraph: { spacing: { before: 0, after: 0, line: LINE }, indent: { firstLine: FIRST_LINE }, alignment: AlignmentType.LEFT, keepNext: true },
};

const document = new Document({
  styles: {
    default: {
      document: {
        run: { font: FONT, size: SIZE, color: '000000', language: { value: 'ru-RU' } },
        paragraph: { spacing: { line: LINE, before: 0, after: 0 }, indent: { firstLine: FIRST_LINE }, alignment: AlignmentType.JUSTIFIED },
      },
      heading1: headingStyle,
      heading2: headingStyle,
      heading3: headingStyle,
    },
  },
  sections: parts.map(part => ({
    properties: {
      page: {
        size: part.landscape
          ? { width: PAGE.width, height: PAGE.height, orientation: PageOrientation.LANDSCAPE }
          : { width: PAGE.width, height: PAGE.height },
        margin: MARGIN,
      },
    },
    footers: { default: footer() },
    children: render(part.blocks, part.landscape ? TEXT_WIDTH_LAND : TEXT_WIDTH),
  })),
});

Packer.toBuffer(document).then(buffer => {
  fs.writeFileSync(TARGET, buffer);
  console.log('собрано:', path.basename(TARGET), (buffer.length / 1024).toFixed(0) + ' КБ');
});
