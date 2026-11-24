"""Разовый перенос раздела из markdown в LaTeX.

После переноса источником становится .tex — скрипт нужен только для первой
конвертации и для повторного прогона, если markdown ещё правился.

    python "Документы/LaTeX/md2tex.py" ВХОД.md ВЫХОД.tex
"""
import io
import re
import sys

# Готовая разметка формул: автоматический перевод здесь невозможен
FORMULAS = {
    '(1.1)': r'G = \langle R,\; S,\; s_0,\; A,\; noop,\; terminal(\cdot),\; '
             r'legal_i(\cdot),\; next(\cdot),\; payoff(\cdot) \rangle',
    '(1.2)': r'UCT_j = \bar{X}_j + C\sqrt{\dfrac{\ln N}{n_j}}',
    '(1.3)': r'E(S) = \dfrac{1}{s_{\max} \cdot 100}\sum_{i=1}^{12} '
             r'score(S,\, k_i) \cdot w_i',
}

# Рисунки: номер -> файл и ширина на полосе набора
FIGURES = {
    '1.1': ('../Рисунки/fig-1-1-phases.png',   '14cm'),
    '1.2': ('../Рисунки/fig-1-2-existing.png', '14cm'),
    '1.3': ('../Рисунки/fig-1-3-activity.png', '9.8cm'),
    '1.4': ('../Рисунки/fig-1-4-layouts.png',  '14cm'),
}

SPECIALS = [('\\', r'\textbackslash{}'), ('&', r'\&'), ('%', r'\%'), ('$', r'\$'),
            ('#', r'\#'), ('_', r'\_'), ('{', r'\{'), ('}', r'\}'),
            ('~', r'\textasciitilde{}'), ('^', r'\textasciicircum{}')]


def escape(text):
    for a, b in SPECIALS:
        text = text.replace(a, b)
    text = re.sub(r'\*\*([^*]+)\*\*', r'\\textbf{\1}', text)
    text = text.replace('`', '')
    return text


def heading_text(line, level):
    """Снимает номер: в LaTeX он проставляется автоматически."""
    body = line[level:].strip()
    return re.sub(r'^\d+(\.\d+)*\s+', '', body)


def table(rows):
    columns = len(rows[0])
    if columns <= 2:
        spec = '>{\\raggedright\\arraybackslash}p{0.3\\linewidth}'
        spec += '>{\\raggedright\\arraybackslash}p{0.62\\linewidth}'
    else:
        spec = '>{\\raggedright\\arraybackslash}X' + 'c' * (columns - 1)
    out = ['\\begin{tabularx}{\\linewidth}{%s}' % spec, '\\hline']
    for index, cells in enumerate(rows):
        out.append(' & '.join(escape(c) for c in cells) + ' \\\\')
        out.append('\\hline')
        if index == 0:
            out.append('\\endhead' if False else '')
    return '\n'.join(l for l in out if l) + '\n\\end{tabularx}'


def convert(markdown):
    lines = markdown.split('\n')
    out = []
    i = 0
    pending_caption = None

    while i < len(lines):
        line = lines[i]
        text = line.strip()

        if not text:
            i += 1
            continue

        if text.startswith('### '):
            out.append('\\subsubsection{%s}' % escape(heading_text(text, 4)))
            i += 1
            continue
        if text.startswith('## '):
            out.append('\\subsection{%s}' % escape(heading_text(text, 3)))
            i += 1
            continue
        if text.startswith('# '):
            out.append('\\section{%s}' % escape(heading_text(text, 2)))
            i += 1
            continue

        if re.match(r'^Таблица\s', text):
            pending_caption = re.sub(r'^Таблица\s+[\d.]+\s+[–—-]\s+', '', text)
            i += 1
            continue

        if text.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                cells = [c.strip() for c in lines[i].strip()[1:-1].split('|')]
                if not all(re.fullmatch(r':?-{2,}:?', c) for c in cells):
                    rows.append(cells)
                i += 1
            out.append('\\begin{gosttable}{}{%s}' % escape(pending_caption or ''))
            out.append(table(rows))
            out.append('\\end{gosttable}')
            pending_caption = None
            continue

        if text.startswith('[Рисунок'):
            whole = text.strip('[]')
            number = (re.match(r'Рисунок\s+([\d.]+)', whole) or [None, ''])[1]
            caption = re.sub(r'^Рисунок\s+[\d.]+\s+[–—-]\s+', '', whole.split('Содержание:')[0]).strip().rstrip('.')
            path, width = FIGURES.get(number, ('', '10cm'))
            out.append('\\gostfigure{%s}{%s}{%s}' % (path, width, escape(caption)))
            i += 1
            continue

        number = re.search(r'\((\d+\.\d+)\)\s*$', text)
        if number and not re.match(r'^\d+\.\s', text):
            key = number.group(0)
            out.append('\\begin{equation}\n%s\n\\end{equation}' % FORMULAS.get(key, escape(text)))
            i += 1
            continue

        if re.match(r'^\d+\.\s', text):
            items = []
            while i < len(lines) and re.match(r'^\d+\.\s', lines[i].strip()):
                items.append(re.sub(r'^\d+\.\s', '', lines[i].strip()))
                i += 1
            out.append('\\begin{enumerate}')
            out += ['  \\item %s' % escape(item) for item in items]
            out.append('\\end{enumerate}')
            continue

        if text.startswith('- '):
            items = []
            while i < len(lines) and lines[i].strip().startswith('- '):
                items.append(lines[i].strip()[2:])
                i += 1
            out.append('\\begin{itemize}')
            out += ['  \\item %s' % escape(item) for item in items]
            out.append('\\end{itemize}')
            continue

        out.append(escape(text))
        i += 1

    return '\n\n'.join(out) + '\n'


source = io.open(sys.argv[1], encoding='utf-8').read()
body, _, sources = source.partition('## Источники')

tex = convert(body)
entries = re.findall(r'(?m)^\d+\. (.+)$', sources)
if entries:
    tex += '\n\\subsection*{Список использованных источников}\n\n'
    tex += '\\begin{enumerate}[label=\\arabic*.]\n'
    tex += '\n'.join('  \\item %s' % escape(e) for e in entries)
    tex += '\n\\end{enumerate}\n'

io.open(sys.argv[2], 'w', encoding='utf-8', newline='\n').write(tex)
print('перенесено:', sys.argv[2], '-', len(tex), 'знаков,', len(entries), 'источников')
