/**
 * Рисунки раздела 1: чёрно-белые схемы (методичка запрещает цвет) плюс
 * перевод готового снимка экрана в оттенки серого.
 *
 * Запуск:  node "Документы/figures-chapter1.js"
 * Результат: images/chapter1/*.png
 *
 * Отрисовка — системный Chrome через playwright-core из mcts-viz/node_modules.
 */
const fs = require('fs');
const path = require('path');
const { chromium } = require('../../mcts-viz/node_modules/playwright-core');

const OUT = path.join(__dirname, '..', '..', 'images', 'chapter1');
const SCALE = 2;                      // плотность растра под печать
const FONT = 'Times New Roman, serif';

// ── Общее дерево для схем: корень, три ребёнка, два внука ────────────────────
const ROOT = [75, 22];
const KIDS = [[35, 66], [75, 66], [115, 66]];
const GRAND = [[58, 110], [95, 110]];
const NEW_NODE = [112, 150];

const edge = (a, b, w = 1) =>
  `<line x1="${a[0]}" y1="${a[1]}" x2="${b[0]}" y2="${b[1]}" stroke="#000" stroke-width="${w}"/>`;

const node = (p, { fill = '#fff', dashed = false, r = 8 } = {}) =>
  `<circle cx="${p[0]}" cy="${p[1]}" r="${r}" fill="${fill}" stroke="#000" stroke-width="1.2"` +
  (dashed ? ' stroke-dasharray="3 2"' : '') + '/>';

const text = (x, y, s, { size = 11, anchor = 'middle', style = '' } = {}) =>
  `<text x="${x}" y="${y}" font-family="${FONT}" font-size="${size}" text-anchor="${anchor}"` +
  ` fill="#000" ${style}>${s}</text>`;

/** Один этап итерации: дерево с подписью этапа под ним. */
function phasePanel(title, extra, { thickPath = [], dim = [] } = {}) {
  const thick = new Set(thickPath.map(JSON.stringify));
  const isThick = (a, b) => thick.has(JSON.stringify([a, b]));
  const edges = [
    [ROOT, KIDS[0]], [ROOT, KIDS[1]], [ROOT, KIDS[2]],
    [KIDS[1], GRAND[0]], [KIDS[1], GRAND[1]],
  ];
  const dimmed = new Set(dim.map(JSON.stringify));
  return `
    <g>
      ${edges.map(([a, b]) => edge(a, b, isThick(a, b) ? 2.6 : 1)).join('')}
      ${extra || ''}
      ${[ROOT, ...KIDS, ...GRAND].map(p =>
        node(p, { fill: dimmed.has(JSON.stringify(p)) ? '#d9d9d9' : '#fff' })).join('')}
      ${text(75, 214, title, { size: 11 })}
    </g>`;
}

const FIG_ITERATION = `
<svg xmlns="http://www.w3.org/2000/svg" width="640" height="232" viewBox="0 0 640 232">
  <rect width="640" height="232" fill="#fff"/>
  <g transform="translate(0,0)">${phasePanel('а) выбор', '', {
    thickPath: [[ROOT, KIDS[1]], [KIDS[1], GRAND[1]]],
  })}</g>
  <g transform="translate(160,0)">${phasePanel('б) расширение',
    edge(GRAND[1], NEW_NODE) + node(NEW_NODE, { dashed: true }))}</g>
  <g transform="translate(320,0)">${phasePanel('в) симуляция',
    edge(GRAND[1], NEW_NODE) + node(NEW_NODE) +
    `<path d="M ${NEW_NODE[0]} ${NEW_NODE[1] + 9} q 10 10 -4 18 q -10 7 1 13"
       fill="none" stroke="#000" stroke-width="1" stroke-dasharray="4 3"/>` +
    `<rect x="${NEW_NODE[0] - 19}" y="184" width="38" height="15" fill="#fff" stroke="#000"/>` +
    text(NEW_NODE[0], 195, 'исход', { size: 9 }))}</g>
  <g transform="translate(480,0)">${phasePanel('г) обратное распространение',
    edge(GRAND[1], NEW_NODE) + node(NEW_NODE) +
    `<path d="M 108 143 L 100 121" stroke="#000" stroke-width="2.2" marker-end="url(#arr)"/>` +
    `<path d="M 92 101 L 80 78" stroke="#000" stroke-width="2.2" marker-end="url(#arr)"/>` +
    `<path d="M 75 57 L 75 34" stroke="#000" stroke-width="2.2" marker-end="url(#arr)"/>` +
    text(116, 136, '+1', { size: 10, anchor: 'start' }) +
    text(98, 92, '+1', { size: 10, anchor: 'start' }) +
    text(81, 46, '+1', { size: 10, anchor: 'start' }))}</g>
  <defs>
    <marker id="arr" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="5" markerHeight="5" orient="auto">
      <path d="M 0 0 L 10 5 L 0 10 z" fill="#000"/>
    </marker>
  </defs>
</svg>`;

// ── Диаграмма деятельности: работа с существующим модулем ───────────────────
const act = (x, y, w, h, label, lines) => {
  const body = (lines || [label]).map((l, i) =>
    text(x + w / 2, y + h / 2 + 4 + (i - (lines ? lines.length - 1 : 0) / 2) * 13, l, { size: 11})).join('');
  return `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="9" fill="#fff" stroke="#000" stroke-width="1.2"/>${body}`;
};

const arrow = (x1, y1, x2, y2) =>
  `<path d="M ${x1} ${y1} L ${x2} ${y2}" stroke="#000" stroke-width="1.2" fill="none" marker-end="url(#arr2)"/>`;

const FIG_ACTIVITY = `
<svg xmlns="http://www.w3.org/2000/svg" width="560" height="470" viewBox="0 0 560 470">
  <rect width="560" height="470" fill="#fff"/>
  <circle cx="280" cy="22" r="9" fill="#000"/>
  ${arrow(280, 31, 280, 48)}
  ${act(180, 48, 200, 36, 'Загрузить снимок дерева')}
  ${arrow(280, 84, 280, 102)}
  ${act(180, 102, 200, 36, 'Отобразить дерево')}
  ${arrow(280, 138, 280, 158)}
  <path d="M 280 158 L 352 196 L 280 234 L 208 196 z" fill="#fff" stroke="#000" stroke-width="1.2"/>
  ${text(280, 200, 'Картина читаема?', { size: 11 })}
  ${arrow(352, 196, 420, 196)}
  ${text(386, 188, 'да', { size: 11 })}
  ${act(420, 178, 120, 36, 'Анализировать')}
  ${arrow(480, 214, 480, 240)}
  <circle cx="480" cy="252" r="10" fill="#fff" stroke="#000" stroke-width="1.2"/>
  <circle cx="480" cy="252" r="6" fill="#000"/>
  ${arrow(280, 234, 280, 256)}
  ${text(292, 250, 'нет', { size: 11, anchor: 'start' })}
  <rect x="96" y="256" width="368" height="150" fill="none" stroke="#000" stroke-width="1" stroke-dasharray="6 4"/>
  ${text(280, 276, 'Ручная настройка детализации', { size: 11, style: 'font-style="italic"' })}
  ${act(150, 288, 260, 36, 'Задать ограничение по глубине')}
  ${arrow(280, 324, 280, 344)}
  ${act(150, 344, 260, 36, 'Задать порог доли посещений')}
  ${arrow(280, 380, 280, 396)}
  <path d="M 280 396 L 120 396 L 120 120 L 180 120" stroke="#000" stroke-width="1.2" fill="none" marker-end="url(#arr2)"/>
  ${text(280, 428, 'Цикл подбора замыкается на исследователя,', { size: 11 })}
  ${text(280, 444, 'а не на систему', { size: 11 })}
  <defs>
    <marker id="arr2" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto">
      <path d="M 0 0 L 10 5 L 0 10 z" fill="#000"/>
    </marker>
  </defs>
</svg>`;

// ── Три пространственно-заполняющих представления одного дерева ─────────────
const SHARES = [0.45, 0.3, 0.25];                 // доли детей корня
const SUB = [[0.6, 0.4], [0.5, 0.5], [1]];        // доли внуков

function icicle(x, y, w, rowH) {
  let out = `<rect x="${x}" y="${y}" width="${w}" height="${rowH}" fill="#bfbfbf" stroke="#000"/>`;
  let cursor = x;
  SHARES.forEach((share, i) => {
    const cw = w * share;
    out += `<rect x="${cursor}" y="${y + rowH + 2}" width="${cw}" height="${rowH}" fill="${i === 0 ? '#8c8c8c' : '#e6e6e6'}" stroke="#000"/>`;
    let inner = cursor;
    SUB[i].forEach((s) => {
      const iw = cw * s;
      out += `<rect x="${inner}" y="${y + 2 * (rowH + 2)}" width="${iw}" height="${rowH}" fill="#f2f2f2" stroke="#000"/>`;
      inner += iw;
    });
    cursor += cw;
  });
  return out;
}

function treemap(x, y, w, h) {
  let out = `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="#fff" stroke="#000"/>`;
  let cursor = x;
  SHARES.forEach((share, i) => {
    const cw = w * share;
    out += `<rect x="${cursor}" y="${y}" width="${cw}" height="${h}" fill="${i === 0 ? '#8c8c8c' : '#e6e6e6'}" stroke="#000"/>`;
    let top = y;
    SUB[i].forEach((s) => {
      const ih = h * s;
      out += `<rect x="${cursor + 4}" y="${top + 4}" width="${cw - 8}" height="${ih - 8}" fill="#f2f2f2" stroke="#000"/>`;
      top += ih;
    });
    cursor += cw;
  });
  return out;
}

function sunburst(cx, cy, r0, r1, r2) {
  const arc = (from, to, inner, outer, fill) => {
    const a0 = from * 2 * Math.PI - Math.PI / 2;
    const a1 = to * 2 * Math.PI - Math.PI / 2;
    const p = (a, r) => [cx + r * Math.cos(a), cy + r * Math.sin(a)];
    const [x0, y0] = p(a0, outer); const [x1, y1] = p(a1, outer);
    const [x2, y2] = p(a1, inner); const [x3, y3] = p(a0, inner);
    const large = to - from > 0.5 ? 1 : 0;
    return `<path d="M ${x0} ${y0} A ${outer} ${outer} 0 ${large} 1 ${x1} ${y1} L ${x2} ${y2} A ${inner} ${inner} 0 ${large} 0 ${x3} ${y3} z" fill="${fill}" stroke="#000"/>`;
  };
  let out = `<circle cx="${cx}" cy="${cy}" r="${r0}" fill="#bfbfbf" stroke="#000"/>`;
  let at = 0;
  SHARES.forEach((share, i) => {
    out += arc(at, at + share, r0, r1, i === 0 ? '#8c8c8c' : '#e6e6e6');
    let inner = at;
    SUB[i].forEach((s) => {
      out += arc(inner, inner + share * s, r1, r2, '#f2f2f2');
      inner += share * s;
    });
    at += share;
  });
  return out;
}

const FIG_LAYOUTS = `
<svg xmlns="http://www.w3.org/2000/svg" width="640" height="200" viewBox="0 0 640 200">
  <rect width="640" height="200" fill="#fff"/>
  ${icicle(20, 40, 180, 30)}
  ${text(110, 180, 'а) ярусная диаграмма', { size: 12 })}
  ${treemap(240, 40, 160, 98)}
  ${text(320, 180, 'б) вложенные прямоугольники', { size: 12 })}
  ${sunburst(540, 89, 18, 38, 56)}
  ${text(540, 180, 'в) диаграмма sunburst', { size: 12 })}
</svg>`;

// ── Отрисовка ───────────────────────────────────────────────────────────────
async function render(page, svg, file, width, height) {
  await page.setViewportSize({ width, height });
  await page.setContent(
    `<body style="margin:0;background:#fff">${svg}</body>`,
    { waitUntil: 'load' });
  await page.screenshot({ path: path.join(OUT, file), clip: { x: 0, y: 0, width, height } });
  console.log('  ', file);
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  const page = await browser.newPage({ deviceScaleFactor: SCALE });

  console.log('рисунки раздела 1:');
  await render(page, FIG_ITERATION, 'fig-1-1-phases.png', 640, 232);
  await render(page, FIG_ACTIVITY, 'fig-1-3-activity.png', 560, 470);
  await render(page, FIG_LAYOUTS, 'fig-1-4-layouts.png', 640, 200);

  // Снимок существующего модуля — в оттенки серого: методичка запрещает цвет
  const source = path.resolve(__dirname, '..', '..', 'images', 'chapter1', 'src-module.png');
  if (fs.existsSync(source)) {
    const data = fs.readFileSync(source).toString('base64');
    await page.setViewportSize({ width: 960, height: 490 });
    await page.setContent(
      `<body style="margin:0"><img src="data:image/png;base64,${data}"
        style="width:960px;display:block;filter:grayscale(1) contrast(1.08)"></body>`,
      { waitUntil: 'load' });
    const box = await (await page.$('img')).boundingBox();
    await page.screenshot({
      path: path.join(OUT, 'fig-1-2-existing.png'),
      clip: { x: 0, y: 0, width: Math.round(box.width), height: Math.round(box.height) },
    });
    console.log('   fig-1-2-existing.png');
  } else {
    console.log('   ПРОПУЩЕН fig-1-2: нет', source);
  }

  await browser.close();
})();
