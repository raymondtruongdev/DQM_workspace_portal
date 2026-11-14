#!/usr/bin/env node
/**
 * Kiểm tra quy ước style của dự án (xem AGENTS.md mục "UI & Style").
 *   npm run lint:styles
 * Không cần cài thêm thư viện. Thoát mã 1 nếu có vi phạm.
 */
const fs = require('fs');
const path = require('path');

const SRC = path.join(__dirname, '..', 'src');
const errors = [];

/** Tiền tố class hợp lệ của từng module (module mới: tự động là `<tên>-`). */
const MODULE_PREFIXES = { report: ['report-', 'bcn-'] };
/** Class cũ chưa có tiền tố (phần Weekly placeholder). Dọn khi làm báo cáo tuần – KHÔNG thêm mới. */
const LEGACY_CLASSES = new Set([
  'auto-calc-label', 'empty-tasks-row', 'id-tag', 'id-parent', 'id-leaf', 'kpi-box', 'kpi-label', 'kpi-value',
  'highlight', 'leaf-rate-input-wrap', 'parent-icon', 'parent-qty-dash', 'parent-rate-pill', 'parent-rate-val',
  'qty-input', 'rate-input', 'rate-unit', 'task-name-text', 'task-name-wrap', 'font-parent', 'unit-text'
]);
/**
 * Từ khóa trong tên class của module gợi ý đang viết lại thành phần có sẵn của core.
 * Gặp từ khóa → lỗi, trừ khi có trong REINVENT_ALLOW (kèm lý do).
 */
const REINVENT_KEYWORDS = {
  btn: '.btn (+ btn--primary/secondary/…)', button: '.btn (+ btn--primary/secondary/…)',
  card: '.card hoặc .section-card', banner: '.section-card__banner',
  badge: '.badge (+ badge--info/success/warning/danger)', pill: '.badge hoặc <FilterChips />',
  chip: '<FilterChips />', chips: '<FilterChips />', tab: '<FilterChips />', tabs: '<FilterChips />',
  header: '<PageHeader />', input: '.input / .textarea / .field', textarea: '.textarea',
  search: '<SearchInput />', spinner: '<LoadingState />', loading: '<LoadingState />',
  empty: '<EmptyState />', toast: '<ToastAlert />'
};
/** Ngoại lệ có chủ đích (tên trùng từ khóa nhưng là phần nội dung riêng của module). */
const REINVENT_ALLOW = {
  'report-info-card': 'mở rộng .card (padding/gap), luôn đi kèm class card',
  'report-template-upgrade-card': 'mở rộng .card (viền cảnh báo), luôn đi kèm class card',
  'report-tile-header': 'dòng đầu bên trong ô báo cáo (không phải header trang)',
  'report-photos-header': 'dòng đầu bên trong hộp upload ảnh',
  'report-template-upgrade-header': 'dòng đầu bên trong thẻ nâng cấp template',
  'report-task-header-bar': 'thanh tiêu đề bảng tiến độ (Weekly placeholder)',
  'report-row-lock-badge': 'ô vuông icon khoá trong bảng, không phải badge chữ'
};
/** Rule được phép dùng màu hex (CSS bản in A4 / xem trước phải cố định theo giấy). */
const HEX_ALLOWED_SELECTOR = /\.bcn-/;
/** File TSX được phép chứa hex (CSS in A4 trong chuỗi, hình thương hiệu module). */
const HEX_ALLOWED_TSX = [
  /ReportPrintPreviewModal\.tsx$/,
  /reportPrintCss\.ts$/, // CSS in A4 dùng chung cho cửa sổ in và file HTML độc lập
  /reportHtml\.ts$/, // khung giấy + thanh công cụ của file HTML độc lập (không có token vì chạy ngoài portal)
  /moduleRegistry\.tsx$/,
  /modules[\\/][^\\/]+[\\/]index\.ts$/ // định nghĩa icon thương hiệu của module
];

function walk(dir, out = []) {
  for (const f of fs.readdirSync(dir)) {
    const p = path.join(dir, f);
    if (fs.statSync(p).isDirectory()) walk(p, out);
    else out.push(p);
  }
  return out;
}
const rel = (p) => path.relative(path.join(__dirname, '..'), p);
const fail = (file, msg) => errors.push(`${rel(file)}: ${msg}`);

/** Tách rule top-level (và rule trong @media): trả [{selector, body}] */
function parseRules(css) {
  const rules = [];
  const clean = css.replace(/\/\*[\s\S]*?\*\//g, '');
  (function scan(text) {
    let i = 0;
    while (i < text.length) {
      const j = text.indexOf('{', i);
      if (j < 0) break;
      const head = text.slice(i, j).trim();
      let d = 1, k = j + 1;
      while (d && k < text.length) {
        if (text[k] === '{') d++;
        else if (text[k] === '}') d--;
        k++;
      }
      const body = text.slice(j + 1, k - 1);
      if (head.startsWith('@media') || head.startsWith('@supports')) scan(body);
      else if (!head.startsWith('@')) rules.push({ selector: head, body });
      i = k;
    }
  })(clean);
  return rules;
}

const files = walk(SRC);
const cssFiles = files.filter((f) => f.endsWith('.css'));
const tsFiles = files.filter((f) => /\.(ts|tsx)$/.test(f));

for (const file of cssFiles) {
  const text = fs.readFileSync(file, 'utf8');
  const isTokens = file.endsWith(path.join('styles', 'tokens.css'));
  const inCore = rel(file).startsWith(path.join('src', 'core'));
  const moduleMatch = rel(file).match(/^src[\\/]modules[\\/]([^\\/]+)[\\/]/);

  if (/!important/.test(text.replace(/\/\*[\s\S]*?\*\//g, ''))) fail(file, 'không dùng !important');

  for (const { selector, body } of parseRules(text)) {
    // 1) hex chỉ trong tokens.css (trừ rule in A4)
    if (!isTokens && /#[0-9a-fA-F]{3,8}\b/.test(body) && !HEX_ALLOWED_SELECTOR.test(selector)) {
      fail(file, `màu hex trong "${selector.slice(0, 50)}" – dùng token var(--color-*)`);
    }
    // 2) core không chứa tên module
    if (inCore && /\.(report|tasks)-/.test(selector)) {
      fail(file, `core không được chứa tên module: "${selector.slice(0, 60)}"`);
    }
    // 3) class của module phải có tiền tố module
    if (moduleMatch) {
      const name = moduleMatch[1];
      const prefixes = MODULE_PREFIXES[name] || [`${name}-`];
      const classes = selector.match(/\.([a-zA-Z_][\w-]*)/g) || [];
      for (const c of classes.map((x) => x.slice(1))) {
        const isShared = /^(btn|badge|input|textarea|field|card|section-card|chip|chips|search|page-header|state|icon|spin|table-responsive|portal-table|text-|font-|is-)/.test(c);
        if (!prefixes.some((p) => c.startsWith(p)) && !LEGACY_CLASSES.has(c) && !isShared) {
          fail(file, `class "${c}" thiếu tiền tố module (${prefixes.join(' hoặc ')})`);
        }
        if (!isShared && !LEGACY_CLASSES.has(c) && !c.startsWith('bcn-') && !REINVENT_ALLOW[c]) {
          const hit = c.split('-').find((seg) => REINVENT_KEYWORDS[seg]);
          if (hit) {
            fail(file, `class "${c}" có vẻ viết lại thành phần có sẵn → dùng ${REINVENT_KEYWORDS[hit]} (hoặc thêm vào REINVENT_ALLOW kèm lý do)`);
          }
        }
      }
    }
  }
}

for (const file of tsFiles) {
  const text = fs.readFileSync(file, 'utf8');
  if (rel(file).startsWith(path.join('src', 'core')) && /from ['"](\.\.\/)+modules\//.test(text)) {
    fail(file, 'core không được import từ modules/');
  }
  if (!HEX_ALLOWED_TSX.some((r) => r.test(file)) && /#[0-9a-fA-F]{6}\b/.test(text)) {
    fail(file, 'màu hex trong TSX – dùng token CSS');
  }
  const code = text.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
  if (/<svg[\s>]/.test(code) && !HEX_ALLOWED_TSX.some((r) => r.test(file)) && !/style-lint: allow-svg/.test(text)) {
    fail(file, '<svg> viết tay – dùng <Icon name="..." />');
  }
}

// ---- Tài liệu phải khớp code: mọi thành phần/biến thể của core đều có trong AGENTS.md ----
(function checkDocs() {
  const agentsPath = path.join(__dirname, '..', 'AGENTS.md');
  const docs = fs.readFileSync(agentsPath, 'utf8');
  const need = (name, kind) => {
    if (!docs.includes(name)) errors.push(`AGENTS.md: thiếu ${kind} "${name}" trong bảng thành phần dùng chung (mục 8.3)`);
  };
  // a) React component xuất ra từ core/components
  for (const f of files.filter((x) => /src[\\/]core[\\/]components[\\/][^\\/]+\.tsx$/.test(x))) {
    const text = fs.readFileSync(f, 'utf8');
    for (const m of text.matchAll(/export const ([A-Z]\w*)\s*:\s*React\.FC/g)) need(m[1], 'component');
  }
  // b) biến thể (block--variant) và block có element (block__el) trong core/styles/components
  const blocks = new Set();
  for (const f of cssFiles.filter((x) => /src[\\/]core[\\/]styles[\\/]components[\\/]/.test(x))) {
    const css = fs.readFileSync(f, 'utf8').replace(/\/\*[\s\S]*?\*\//g, '');
    for (const m of css.matchAll(/\.([a-z][a-z0-9-]*?)--([a-z][a-z0-9-]*)/g)) {
      need(`${m[1]}--${m[2]}`, 'biến thể');
      blocks.add(m[1]);
    }
    for (const m of css.matchAll(/\.([a-z][a-z0-9-]*?)__[a-z]/g)) blocks.add(m[1]);
  }
  blocks.forEach((b) => need(b, 'thành phần'));
})();

if (errors.length) {
  console.error(`\n✖ ${errors.length} vi phạm quy ước style:\n`);
  errors.forEach((e) => console.error('  - ' + e));
  console.error('\nXem AGENTS.md → "UI & Style".\n');
  process.exit(1);
}
console.log(`✔ Style OK (${cssFiles.length} file CSS, ${tsFiles.length} file TS/TSX)`);
