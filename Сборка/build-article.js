/**
 * Сборка статьи для конференции КТИ (филиал) ВолгГТУ из markdown в .docx.
 *
 * Запуск:  node Сборка/build-article.js "Текст/Статья - Камышин 2026.md" "Готовое/Ролдугин О.Д.(ВолгГТУ).docx"
 *
 * Требования организаторов: лист А5, поля 2,0 / 1,8 / 1,8 / 1,9 см,
 * Times New Roman 11, одинарный интервал, абзацный отступ 0,5 см,
 * по ширине, с переносами, без колонтитулов и нумерации страниц.
 * Номер секции — в правом верхнем углу.
 */
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, SymbolRun, Tab,
  AlignmentType, TabStopType, convertMillimetersToTwip,
} = require('docx');

// ── Страница и текст (DXA: 1 см = 567) ───────────────────────────────────────
const PAGE = { width: convertMillimetersToTwip(148), height: convertMillimetersToTwip(210) };
const MARGIN = {
  top: convertMillimetersToTwip(20), bottom: convertMillimetersToTwip(19),
  left: convertMillimetersToTwip(18), right: convertMillimetersToTwip(18),
  header: 0, footer: 0,
};
const TEXT_WIDTH = PAGE.width - MARGIN.left - MARGIN.right;
const FONT = 'Times New Roman';
const SIZE = 22;          // полуразмеры = 11 пт
const SIZE_SMALL = 18;    // 9 пт — таблицы и подписи рисунков
const FIRST_LINE = convertMillimetersToTwip(5);   // абзацный отступ 0,5 см

const SOURCE = process.argv[2];
const TARGET = process.argv[3];

// ── Разбор исходника: шапка «КЛЮЧ: значение», затем текст ────────────────────
function parse(markdown) {
  const [header, body] = markdown.split(/^---$/m);
  const meta = {};
  for (const line of header.split('\n')) {
    const at = line.indexOf(':');
    if (at > 0) meta[line.slice(0, at).trim()] = line.slice(at + 1).trim();
  }

  const blocks = [];
  let inReferences = false;
  for (const raw of body.split('\n')) {
    const text = raw.trim();
    if (!text) continue;
    if (text.startsWith('## ')) {
      inReferences = /литератур/i.test(text);
      blocks.push({ kind: 'heading', text: text.slice(3) });
    } else if (text.startsWith('$$')) {
      blocks.push({ kind: 'formula', number: text.slice(2).trim() });
    } else if (inReferences) {
      blocks.push({ kind: 'reference', text });
    } else {
      blocks.push({ kind: 'para', text });
    }
  }
  return { meta, blocks };
}

// ── Формулы ────────────────────────────────────────────────────────────────
// Формулы набираются обычным текстом, а не объектом Word (OMML):
// объект рисуется математическим шрифтом Cambria Math, и там, где его нет
// (старые версии Word), латиница и греческие буквы показываются пустыми
// квадратами. Для обычного текста Word подбирает шрифт сам, поэтому такая
// запись читается в любой версии.

// Знаков теории множеств в Times New Roman нет вовсе: Word подбирает под них другой
// шрифт, и в старых версиях подстановка даёт пустой квадрат. Такие знаки берутся
// из шрифта Symbol — он входит в любую Windows, и именно так их записывал ещё Word 97.
const SYMBOL_FONT = { '⊆': 'F0CD', '∅': 'F0C6' };

// Запись формулы: «_{…}» — подстрочный индекс.
const FORMULAS = {
  '(1)': 'x(c) ⊆ x(p),    Σ w(c) + w_{рез}(p) = w(p)',
  '(2)': 'x(v) ∩ D ≠ ∅   и   w(v) · ρ ≥ s_{min}',
};

// Переменные печатаются курсивом, знаки действий и словесные индексы — прямым.
const VARIABLE = /[A-Za-zρ]/;

const mathRun = (text, opts) => new TextRun({
  text, font: FONT, size: SIZE, italics: !!opts.italics, subScript: !!opts.subScript,
});

function formulaRuns(source) {
  const parts = [];
  let plain = '';
  const flushPlain = () => {
    if (plain) parts.push(mathRun(plain, {}));
    plain = '';
  };

  for (let at = 0; at < source.length; at += 1) {
    if (source.startsWith('_{', at)) {
      const close = source.indexOf('}', at);
      flushPlain();
      parts.push(mathRun(source.slice(at + 2, close), { subScript: true }));
      at = close;
    } else if (SYMBOL_FONT[source[at]]) {
      flushPlain();
      parts.push(new SymbolRun({ char: SYMBOL_FONT[source[at]], symbolfont: 'Symbol', size: SIZE }));
    } else if (VARIABLE.test(source[at])) {
      flushPlain();
      parts.push(mathRun(source[at], { italics: true }));
    } else {
      plain += source[at];
    }
  }
  flushPlain();
  return parts;
}

// Формула по центру, номер прижат к правому полю.
const formula = (number) => new Paragraph({
  spacing: { line: 240, before: 120, after: 120 },
  indent: { firstLine: 0 },
  tabStops: [
    { type: TabStopType.CENTER, position: Math.round(TEXT_WIDTH / 2) },
    { type: TabStopType.RIGHT, position: TEXT_WIDTH },
  ],
  children: [
    new TextRun({ children: [new Tab()], font: FONT, size: SIZE }),
    ...formulaRuns(FORMULAS[number] || number),
    new TextRun({ children: [new Tab(), number], font: FONT, size: SIZE }),
  ],
});

// ── Абзацы шапки ─────────────────────────────────────────────────────────────
/**
 * Номер автора записывается в исходнике как «^1» и печатается надстрочным
 * индексом — и в строке авторов, и при их адресах.
 */
function runs(text, opts) {
  return text.split(/(\^\d)/).filter(Boolean).map((piece) => new TextRun({
    text: piece.startsWith('^') ? piece.slice(1) : piece,
    superScript: piece.startsWith('^'),
    font: FONT, size: opts.size || SIZE,
    bold: !!opts.bold, italics: !!opts.italics,
  }));
}

const line = (text, opts = {}) => new Paragraph({
  alignment: opts.alignment || AlignmentType.CENTER,
  spacing: { line: 240, before: 0, after: 0 },
  indent: { firstLine: 0 },
  children: runs(text, opts),
});

// Пустая строка основного кегля: у абзаца без текста Word берёт кегль по
// умолчанию, поэтому в нём стоит пробел.
const blank = () => new Paragraph({
  spacing: { line: 240, before: 0, after: 0 },
  children: [new TextRun({ text: ' ', font: FONT, size: SIZE })],
});

const body = (text) => new Paragraph({
  alignment: AlignmentType.JUSTIFIED,
  spacing: { line: 240, before: 0, after: 0 },
  indent: { firstLine: FIRST_LINE },
  children: [new TextRun({ text, font: FONT, size: SIZE })],
});

// ── Сборка ───────────────────────────────────────────────────────────────────
const { meta, blocks } = parse(fs.readFileSync(SOURCE, 'utf8'));

const children = [
  line(meta['СЕКЦИЯ'], { alignment: AlignmentType.RIGHT }),
  line(`УДК ${meta['УДК']}`, { alignment: AlignmentType.LEFT }),
  line(meta['НАЗВАНИЕ']),
  blank(),
  line(meta['АВТОРЫ'], { bold: true }),
  // Организация и адреса — одной строкой, как в принятых работах ВолгГТУ.
  line(`${meta['ОРГАНИЗАЦИЯ']}, Email: ${meta['EMAIL']}`, { italics: true }),
  blank(),
];

for (const block of blocks) {
  if (block.kind === 'formula') children.push(formula(block.number));
  // «Список литературы» — по центру обычным начертанием, с пустой строкой вокруг.
  else if (block.kind === 'heading') children.push(blank(), line(block.text), blank());
  else if (block.kind === 'reference') {
    children.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { line: 240, before: 0, after: 0 },
      indent: { firstLine: FIRST_LINE },
      children: [new TextRun({ text: block.text, font: FONT, size: SIZE })],
    }));
  } else children.push(body(block.text));
}

const doc = new Document({
  sections: [{
    properties: { page: { size: PAGE, margin: MARGIN } },
    children,
  }],
});

/**
 * Включает автоматический перенос слов — требование оргкомитета.
 * Библиотека docx такой настройки не поддерживает, поэтому флаг
 * дописывается прямо в word/settings.xml готового пакета.
 */
function enableHyphenation(buffer) {
  const { unzipSync, zipSync, strFromU8, strToU8 } = require('fflate');
  const files = unzipSync(new Uint8Array(buffer));
  const settings = strFromU8(files['word/settings.xml']);
  if (!settings.includes('w:autoHyphenation')) {
    files['word/settings.xml'] = strToU8(
      settings.replace('</w:settings>', '<w:autoHyphenation w:val="true"/></w:settings>'),
    );
  }
  return Buffer.from(zipSync(files));
}

Packer.toBuffer(doc).then((buffer) => {
  const packed = enableHyphenation(buffer);
  fs.writeFileSync(TARGET, packed);
  const name = TARGET.split(/[\\/]/).pop();
  console.log(`собрано: ${name} ${Math.round(packed.length / 1024)} КБ`);
});
