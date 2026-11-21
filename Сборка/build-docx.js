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
  ImageRun,
  AlignmentType, WidthType, TabStopType, PageOrientation, PageNumber,
  Footer, HeadingLevel, BorderStyle, TableLayoutType,
  Math: OfficeMath, MathRun, MathFraction, MathRadical, MathSubScript, MathSuperScript, MathSum,
} = require('docx');

// ── Параметры страницы и текста (DXA) ────────────────────────────────────────
const PAGE = { width: 11906, height: 16838 };
const MARGIN = { top: 1134, bottom: 1134, left: 1701, right: 850, header: 567, footer: 567 };
const TEXT_WIDTH = PAGE.width - MARGIN.left - MARGIN.right;            // 9355
const TEXT_WIDTH_LAND = PAGE.height - MARGIN.left - MARGIN.right;      // 14287
const FONT = 'Times New Roman';
const SIZE = 28;          // полуразмеры = 14 пт
// ГОСТ 7.32-2017, п. 6.6.7: в таблице допускается кегль меньше, чем в тексте.
// Методичка кафедры о кегле таблиц молчит, а при 14 пт широкие таблицы
// выходят за поле набора.
const SIZE_TABLE = 24;    // 12 пт
const LINE = 360;         // полуторный интервал
const FIRST_LINE = 709;   // абзацный отступ 1,25 см
// Высота одного «интервала» — пустой строки основного стиля: 14 пт при
// полуторном междустрочии. Вокруг заголовков интервалы задаются отступами
// самого заголовка: пустые абзацы Word поглощает на границе страницы, и
// keepNext на них не держит — заголовок отрывается от своего текста.
const GAP = 480;
const WIDE_TABLE = 'Таблица 1.5';   // не помещается на книжную страницу

// Готовые рисунки: номер -> файл и ширина на странице в пикселях при 96 dpi
// (16 см поля набора = 605 px). Высота считается по пропорциям исходника.
const FIGURES = {
  '1.1': { file: 'fig-1-1-phases.png',   width: 530, source: [640, 232] },
  '1.2': { file: 'fig-1-2-existing.png', width: 530, source: [960, 490] },
  '1.3': { file: 'fig-1-3-activity.png', width: 370, source: [560, 470] },
  '1.4': { file: 'fig-1-4-layouts.png',  width: 530, source: [640, 200] },
};
const FIGURE_DIR = path.join(__dirname, '..', '..', 'images', 'chapter1');

const SOURCE = process.argv[2];
const TARGET = process.argv[3];

// ── Разбор markdown в блоки ──────────────────────────────────────────────────
function parse(markdown) {
  const lines = markdown.split('\n');
  const blocks = [];
  let i = 0;
  // В списке литературы номер источника пишется с точкой, а не как «1)»:
  // это нумерация записей, а не перечисление внутри текста
  let bibliography = false;

  while (i < lines.length) {
    const line = lines[i];
    const text = line.trim();

    if (!text) { i += 1; continue; }

    if (/^#{1,3} Список использованных источников/.test(text)) bibliography = true;

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
      blocks.push({ kind: 'item', text: bibliography ? text : text.replace(/^(\d+)\.\s/, '$1) ') });
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

const empty = (keepNext = false) => new Paragraph({ keepNext, children: runs('') });

const body = (text, options = {}) => new Paragraph({
  alignment: options.alignment ?? AlignmentType.JUSTIFIED,
  indent: { firstLine: options.firstLine ?? FIRST_LINE },
  spacing: { line: LINE, before: 0, after: 0 },
  children: runs(text),
});

const heading = (text, level, before, after) => new Paragraph({
  heading: level,
  alignment: AlignmentType.LEFT,
  indent: { firstLine: FIRST_LINE },
  spacing: { line: LINE, before, after },
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

/**
 * Формулы набираются средствами Word (объект OMML), а не строкой текста:
 * методичка требует редактор формул. Разметка каждой формулы задаётся здесь
 * по её номеру, прочие выводятся одной строкой внутри того же объекта.
 */
// Индекс строится ТОЛЬКО вместе со своим основанием: пустое основание Word
// показывает квадратиком-заполнителем, а не пропуском.
const sub = (base, index) => new MathSubScript({
  children: [new MathRun(base)],
  subScript: [new MathRun(index)],
});

function mathOf(text, number) {
  if (number === '(1.2)') {
    return [
      sub('UCT', 'j'),
      new MathRun(' = '),
      sub('X̄', 'j'),
      new MathRun(' + C'),
      new MathRadical({
        children: [new MathFraction({
          numerator: [new MathRun('ln N')],
          denominator: [sub('n', 'j')],
        })],
      }),
    ];
  }
  if (number === '(1.3)') {
    return [
      new MathRun('E(S) = '),
      new MathFraction({
        numerator: [new MathRun('1')],
        denominator: [sub('s', 'max'), new MathRun(' · 100')],
      }),
      new MathSum({
        children: [
          new MathRun('score(S, '), sub('k', 'i'), new MathRun(') · '), sub('w', 'i'),
        ],
        subScript: [new MathRun('i = 1')],
        superScript: [new MathRun('12')],
      }),
    ];
  }
  if (number === '(1.1)') {
    return [
      new MathRun('G = ⟨R, S, '), sub('s', '0'),
      new MathRun(', A, noop, terminal(·), '), sub('legal', 'i'),
      new MathRun('(·), next(·), payoff(·)⟩'),
    ];
  }
  return [new MathRun(text)];
}

const formula = (text, number, width) => new Paragraph({
  indent: { firstLine: 0 },
  spacing: { line: LINE, before: 0, after: 0 },
  tabStops: [
    { type: TabStopType.CENTER, position: Math.round(width / 2) },
    { type: TabStopType.RIGHT, position: width },
  ],
  children: [
    new TextRun({ children: [new Tab()], font: FONT, size: SIZE }),
    new OfficeMath({ children: mathOf(text, number) }),
    new TextRun({ children: [new Tab()], font: FONT, size: SIZE }),
    ...runs(number),
  ],
});

/** Рисунок с подписью под ним по центру; при отсутствии файла — рамка. */
function figure(block) {
  const number = (block.caption.match(/Рисунок\s+([\d.]+)/) || [])[1];
  const spec = FIGURES[number];
  const file = spec && path.join(FIGURE_DIR, spec.file);
  if (!spec || !fs.existsSync(file)) return figurePlaceholder(block);

  const height = Math.round(spec.width * spec.source[1] / spec.source[0]);
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      indent: { firstLine: 0 },
      spacing: { line: LINE, before: 0, after: 0 },
      keepNext: true,
      children: [new ImageRun({
        type: 'png',
        data: fs.readFileSync(file),
        transformation: { width: spec.width, height },
      })],
    }),
    caption(block.caption, AlignmentType.CENTER),
  ];
}

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
  // У широкой таблицы крайние столбцы держат текст, средние — по одному знаку,
  // поэтому равное деление их бы переполнило, а крайние оставило бы пустыми.
  const columns = rows[0].length;
  let widths;
  if (columns > 8) {
    const head = 2200;
    const tail = 1300;
    const middle = Math.floor((width - head - tail) / (columns - 2));
    widths = [head, ...Array(columns - 2).fill(middle), tail];
  } else {
    const first = columns > 4 ? Math.round(width * 0.22) : Math.round(width / columns);
    widths = [first, ...Array(columns - 1).fill(Math.floor((width - first) / (columns - 1)))];
  }
  widths[columns - 1] += width - widths.reduce((a, b) => a + b, 0);

  return new Table({
    columnWidths: widths,
    width: { size: width, type: WidthType.DXA },
    // Без фиксированной раскладки Word подгоняет столбцы под содержимое и
    // выносит таблицу за правое поле
    layout: TableLayoutType.FIXED,
    // У широкой таблицы обычные поля ячеек съедают треть столбца и заголовки
    // вида k10 переносятся на вторую строку
    ...(columns > 8 ? { margins: { top: 0, bottom: 0, left: 28, right: 28 } } : {}),
    rows: rows.map((cells, index) => new TableRow({
      tableHeader: index === 0,
      children: cells.map((cell, column) => new TableCell({
        width: { size: widths[column], type: WidthType.DXA },
        children: [new Paragraph({
          alignment: column === 0 && columns > 2 ? AlignmentType.LEFT : AlignmentType.CENTER,
          indent: { firstLine: 0 },
          spacing: { line: LINE, before: 0, after: 0 },
          children: runs(cell, { size: SIZE_TABLE }),
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
      const level = { h1: HeadingLevel.HEADING_1, h2: HeadingLevel.HEADING_2, h3: HeadingLevel.HEADING_3 }[block.kind];
      // Перед заголовком 2 интервала, но между заголовком и подзаголовком —
      // ровно 1: его отступ снизу печатает предыдущий заголовок
      const before = !previous || previous.kind.startsWith('h') ? 0 : 2 * GAP;
      // После заголовка 2 интервала до текста и 1 до подзаголовка
      const after = next && next.kind.startsWith('h') ? GAP : 2 * GAP;
      out.push(heading(block.text, level, before, after));
      return;
    }

    if (block.kind === 'tableCaption') { out.push(empty(), caption(block.text, AlignmentType.LEFT)); return; }
    if (block.kind === 'table') { out.push(table(block.rows, width), empty()); return; }
    if (block.kind === 'figure') { out.push(empty(), ...figure(block), empty()); return; }
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
const parts = [{ landscape: false, blocks: parse(chapter) }];   // всё книжной ориентацией

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
