// Конвертер Markdown в DOCX в официальном оформлении.
// Times New Roman, поля 30/15/20/20 мм, текст по ширине с абзацным отступом 1,25 см,
// номера страниц внизу по центру. Шапки таблиц полужирные, без заливки и без повтора на новой странице.
// Запуск: node md2docx.js <src.md> <out.docx>
const fs = require('fs'), path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ImageRun, HeadingLevel,
  LevelFormat, BorderStyle, AlignmentType, Footer, PageNumber, VerticalAlign,
} = require('docx');

const [src, out] = process.argv.slice(2);
if (!src || !out) {
  console.error('Использование: node md2docx.js <src.md> <out.docx>');
  process.exit(1);
}
const dir = path.dirname(src);
const lines = fs.readFileSync(src, 'utf8').split(/\r?\n/);

const FONT = 'Times New Roman';
const SZ = 24; // 12 pt
const SZ_TABLE = 22; // 11 pt
const LINE = 276; // межстрочный интервал 1,15
const INDENT = 709; // абзацный отступ 1,25 см
const PAGE = { width: 11906, height: 16838 };
const MARGIN = { top: 1134, bottom: 1134, left: 1701, right: 850 };
const TW = PAGE.width - MARGIN.left - MARGIN.right; // ширина текста, DXA
const IMG_MAX_W = 600, IMG_MAX_H = 820; // пиксели

function runs(text, base = {}) {
  const size = base.size || SZ;
  const t = text.replace(/\[([^\]]+)\]\([^)]+\)/g, '$1'); // ссылки: оставить только текст
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*\s][^*]*\*|`[^`]+`)/g;
  let last = 0, m;
  const push = (s, extra = {}) => out.push(new TextRun({ text: s, font: FONT, size, ...base, ...extra }));
  while ((m = re.exec(t))) {
    if (m.index > last) push(t.slice(last, m.index));
    const x = m[0];
    if (x.startsWith('**')) push(x.slice(2, -2), { bold: true });
    else if (x.startsWith('`')) push(x.slice(1, -1));
    else push(x.slice(1, -1), { italics: true });
    last = m.index + x.length;
  }
  if (last < t.length) push(t.slice(last));
  return out;
}

const line = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const borders = { top: line, bottom: line, left: line, right: line };

function cellParagraphs(text, isHeader) {
  return text.split(/<br\s*\/?>/i).map(part => new Paragraph({
    alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
    keepNext: isHeader, // шапка не остается одна внизу страницы
    spacing: { before: 0, after: 0, line: 240 },
    children: runs(part.trim(), { size: SZ_TABLE, ...(isHeader ? { bold: true } : {}) }),
  }));
}

// Ширина столбцов: каждому столбцу хватает места на самое длинное слово (чтобы слова не рвались),
// остаток ширины делится пропорционально объему текста в столбце.
const CHAR_DXA = 118; // средняя ширина строчного символа Times New Roman 11 pt
const CAP_DXA = 170; // ширина заглавного символа
const CELL_PAD = 200;
const MAX_COL_SHARE = 0.3;
const URL_MIN_DXA = 1300;
const LONG_ROW_CHARS = 350; // строки длиннее разрешено переносить на следующую страницу

function columnWidths(hdr, body) {
  const cols = hdr.map((_, i) => [hdr, ...body].map(r => (r[i] || '').replace(/\*\*|`|<br\s*\/?>/g, ' ')));
  const charsWidth = w => [...w].reduce((a, ch) => a + (/[A-ZА-ЯЁ0-9]/.test(ch) ? CAP_DXA : CHAR_DXA), 0);
  // ссылки и имена файлов переносятся по символам, поэтому их длина не задает минимум столбца
  const wordWidth = w => /[/:_.]/.test(w) && w.length > 15 ? URL_MIN_DXA : charsWidth(w);
  const longestWord = cells => Math.max(...cells.flatMap(c => c.split(/[\s-]+/)).map(wordWidth));
  const minW = cols.map(c => Math.min(longestWord(c) + CELL_PAD, TW * MAX_COL_SHARE));
  const minSum = minW.reduce((a, b) => a + b, 0);
  if (minSum >= TW) return fit(capWidest(minW));
  const volume = cols.map(c => c.slice(1).reduce((a, s) => a + s.length, 0) / Math.max(1, c.length - 1) + 1);
  const volSum = volume.reduce((a, b) => a + b, 0);
  const desired = volume.map((v, i) => Math.max(minW[i], v / volSum * TW));
  const desiredSum = desired.reduce((a, b) => a + b, 0);
  if (desiredSum <= TW) return fit(desired);
  // не хватает ширины: сверх минимума каждому столбцу достается одинаковая доля его запроса
  const k = (TW - minSum) / (desiredSum - minSum);
  return fit(desired.map((d, i) => minW[i] + (d - minW[i]) * k));
}

// Минимумы не помещаются: сжимаются только самые широкие столбцы до общего уровня,
// узкие (номера, коды, статусы) сохраняют свою ширину.
function capWidest(widths) {
  const sorted = [...widths].sort((a, b) => a - b);
  let used = 0;
  for (let i = 0; i < sorted.length; i++) {
    const level = (TW - used) / (sorted.length - i);
    if (sorted[i] >= level) return widths.map(w => Math.min(w, level));
    used += sorted[i];
  }
  return widths;
}

function fit(widths) {
  const sum = widths.reduce((a, b) => a + b, 0);
  const cw = widths.map(w => Math.floor(w * TW / sum));
  cw[cw.length - 1] += TW - cw.reduce((a, b) => a + b, 0);
  return cw;
}

function table(rows) {
  const split = r => r.trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());
  const hdr = split(rows[0]);
  const body = rows.slice(2).map(split);
  const cw = columnWidths(hdr, body);
  // tableHeader не задается: шапка не повторяется на следующей странице
  const row = (cells, isHeader) => new TableRow({
    cantSplit: cells.join('').length <= LONG_ROW_CHARS,
    children: cw.map((width, i) => new TableCell({
      width: { size: width, type: WidthType.DXA },
      borders,
      verticalAlign: isHeader ? VerticalAlign.CENTER : VerticalAlign.TOP,
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
      children: cellParagraphs(cells[i] || '', isHeader),
    })),
  });
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cw,
    rows: [row(hdr, true), ...body.map(r => row(r, false))],
  });
}

function image(alt, rel) {
  const buf = fs.readFileSync(path.join(dir, rel));
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const k = Math.min(IMG_MAX_W / w, IMG_MAX_H / h);
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120, after: 60 },
      children: [new ImageRun({
        type: 'png', data: buf, transformation: { width: Math.round(w * k), height: Math.round(h * k) },
        altText: { title: alt, description: alt, name: alt },
      })],
    }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: runs(alt, { italics: true }) }),
  ];
}

// Текст со ссылками выравнивается влево: при выравнивании по ширине длинный URL растягивает пробелы
const textAlign = t => /https?:\/\//.test(t) ? AlignmentType.LEFT : AlignmentType.JUSTIFIED;

const HEADINGS = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3];
const children = [];
for (let i = 0; i < lines.length; i++) {
  const l = lines[i];
  if (!l.trim() || /^---+\s*$/.test(l)) continue;
  let m;
  if (l.startsWith('```')) {
    i++;
    while (i < lines.length && !lines[i].startsWith('```')) {
      children.push(new Paragraph({
        indent: { left: INDENT }, spacing: { after: 0, line: 240 },
        children: [new TextRun({ text: lines[i] || ' ', font: FONT, size: SZ_TABLE })],
      }));
      i++;
    }
    children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
  } else if ((m = l.match(/^(#{1,4}) (.*)/))) {
    const level = Math.min(m[1].length, 3);
    children.push(new Paragraph({ heading: HEADINGS[level - 1], children: runs(m[2].replace(/\*\*/g, '')) }));
  } else if (l.startsWith('|')) {
    const rows = [];
    while (i < lines.length && lines[i].startsWith('|')) rows.push(lines[i++]);
    i--;
    children.push(table(rows));
    children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
  } else if ((m = l.match(/^!\[(.*)\]\((.*)\)/))) {
    children.push(...image(m[1], m[2]));
  } else if ((m = l.match(/^- \[( |x)\] (.*)/i))) {
    children.push(new Paragraph({
      indent: { left: INDENT + 360, hanging: 360 }, spacing: { after: 40, line: LINE },
      children: [new TextRun({ text: (m[1] === ' ' ? '□' : '■') + '\t', font: FONT, size: SZ }), ...runs(m[2])],
      tabStops: [{ type: 'left', position: INDENT + 360 }],
    }));
  } else if ((m = l.match(/^[-*] (.*)/))) {
    children.push(new Paragraph({
      numbering: { reference: 'bullet', level: 0 }, alignment: textAlign(m[1]),
      spacing: { after: 40, line: LINE }, children: runs(m[1]),
    }));
  } else if ((m = l.match(/^(\d+)\. (.*)/))) {
    children.push(new Paragraph({
      indent: { left: INDENT + 360, hanging: 360 }, alignment: textAlign(m[2]),
      spacing: { after: 40, line: LINE }, tabStops: [{ type: 'left', position: INDENT + 360 }],
      children: [new TextRun({ text: m[1] + '.\t', font: FONT, size: SZ }), ...runs(m[2])],
    }));
  } else {
    children.push(new Paragraph({
      alignment: textAlign(l), indent: { firstLine: INDENT },
      spacing: { after: 80, line: LINE }, children: runs(l),
    }));
  }
}

const heading = (id, size, opts) => ({
  id, name: id.replace('Heading', 'Heading '), basedOn: 'Normal', next: 'Normal', quickFormat: true,
  run: { font: FONT, size, bold: true, color: '000000', ...(opts.run || {}) },
  paragraph: { keepNext: true, keepLines: true, outlineLevel: opts.outline, ...opts.paragraph },
});

const doc = new Document({
  creator: 'Команда «ИНГОССЕМЬЯ»',
  title: path.basename(out, '.docx').replace(/_/g, ' '),
  styles: {
    default: { document: { run: { font: FONT, size: SZ, color: '000000' } } },
    paragraphStyles: [
      heading('Heading1', 28, { outline: 0, paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 0, after: 240 } } }),
      heading('Heading2', 26, { outline: 1, paragraph: { spacing: { before: 240, after: 120 } } }),
      heading('Heading3', 24, { outline: 2, paragraph: { spacing: { before: 160, after: 80 } } }),
    ],
  },
  numbering: {
    config: [{
      reference: 'bullet',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: INDENT + 360, hanging: 360 } }, run: { font: FONT } },
      }],
    }],
  },
  sections: [{
    properties: { page: { size: PAGE, margin: MARGIN } },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: SZ_TABLE })],
        })],
      }),
    },
    children,
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(out, b); console.log('ok', out); });
