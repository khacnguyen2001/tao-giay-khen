/* =====================================================================
 * Tạo giấy khen hàng loạt — chạy hoàn toàn trên trình duyệt.
 * Không có dữ liệu nào được gửi ra ngoài máy người dùng.
 * ===================================================================== */
'use strict';

/* ---------------- Hằng số ---------------- */
const FONT_FAMILY = 'PlayfairGK';            // tên nội bộ, tránh nhầm với font cài sẵn trên máy
const FONT_LABEL = 'Playfair Display Bold Italic';
const DEFAULT_COLOR = '#5e1a36';
const DEFAULT_SIZE = 114;                    // px trên ảnh gốc
const DEFAULT_X = 1019;                      // tâm ngang của tên (px trên ảnh gốc)
const DEFAULT_Y = 688;                       // giữa chiều cao chữ hoa (px trên ảnh gốc)
const CAP_HEIGHT_EM = 0.708;                 // chiều cao chữ hoa (bảng OS/2 của font)
const X_HEIGHT_EM = 0.514;                   // chiều cao chữ thường
const ZIP_NAME = 'giay_khen_hoc_sinh.zip';
const TEST_COUNT = 3;
const EXCEL_EXT = ['xlsx', 'xls'];
const IMAGE_EXT = ['png', 'jpg', 'jpeg', 'webp'];
const INVALID_FILENAME = /[\/\\:*?"<>|\u0000-\u001f\u007f]/;
const SAFARI_CANVAS_LIMIT = 16777216;        // 4096 × 4096

const $ = (id) => document.getElementById(id);
const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));

const state = {
  fontReady: false,
  capHeightEm: CAP_HEIGHT_EM,
  frame: null,          // { img, w, h, url, fileName, size }
  ref: null,            // như trên
  workbook: null,
  excelFile: null,      // File đã chọn trên website (đọc lại mỗi lần bấm “Đọc danh sách”)
  excelSeq: 0,
  excelLoading: false,
  sheetMatrix: null,    // { maxC, rows: [{ excelRow, cells }] }
  students: [],         // { idx, row, code, name, errors: [], warnings: [], fit }
  cfg: null,
  detected: null,
  zoom: 1,
  fit: true,
  viewMode: 'frame',
  previewIdx: -1,       // -1 = dùng tên mẫu tự nhập
  job: null,            // { cancelled } khi đang tạo hàng loạt
  testing: false,
  testDone: false,
  testUrls: [],
  output: null,         // { zipBlob, count }
  exportCanvas: null,
};

/* ---------------- Tiện ích ---------------- */
function fontString(size) {
  return `italic 700 ${size}px "${FONT_FAMILY}"`;
}

function extOf(name) {
  const m = /\.([^.]+)$/.exec(name || '');
  return m ? m[1].toLowerCase() : '';
}

function formatBytes(n) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// Bỏ dấu tiếng Việt để so khớp tiêu đề cột
function normKey(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/gi, 'd').toLowerCase().replace(/\s+/g, ' ').trim();
}

function cleanText(s) {
  return String(s || '').normalize('NFC').replace(/[ ​﻿]/g, ' ').replace(/\s+/g, ' ').trim();
}

function base64ToArrayBuffer(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}

const yieldToUI = () => new Promise((r) => setTimeout(r, 0));

function toast(msg, type = 'info', ms = 4500) {
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = msg;
  $('toastWrap').appendChild(el);
  setTimeout(() => el.remove(), ms);
}

function showGenMessage(type, html) {
  const el = $('genMessage');
  if (!html) { el.classList.add('hidden'); return; }
  el.className = `alert alert-${type}`;
  el.innerHTML = html;
}

function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.rel = 'noopener';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

function canvasToBlob(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('Trình duyệt không tạo được ảnh PNG (ảnh có thể quá lớn).'))), 'image/png');
  });
}

async function imageSizeOf(blob) {
  if (window.createImageBitmap) {
    const bmp = await createImageBitmap(blob);
    const size = { w: bmp.width, h: bmp.height };
    if (bmp.close) bmp.close();
    return size;
  }
  const url = URL.createObjectURL(blob);
  try {
    const img = await loadImage(url);
    return { w: img.naturalWidth, h: img.naturalHeight };
  } finally {
    URL.revokeObjectURL(url);
  }
}

function loadImage(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('Không giải mã được ảnh'));
    img.src = url;
  });
}

/* ---------------- Font ---------------- */
async function loadFont() {
  const pill = $('fontStatus');
  try {
    const source = window.PLAYFAIR_BOLD_ITALIC_BASE64
      ? base64ToArrayBuffer(window.PLAYFAIR_BOLD_ITALIC_BASE64)
      : 'url("fonts/PlayfairDisplay-BoldItalic-NoLiga.ttf")';
    const face = new FontFace(FONT_FAMILY, source, { style: 'italic', weight: '700' });
    await face.load();
    document.fonts.add(face);
    await document.fonts.ready;

    // Xác nhận font thật sự được dùng (khác font dự phòng)
    const ctx = document.createElement('canvas').getContext('2d');
    const probe = 'Nguyễn Thị Ánh Tuyết';
    ctx.font = fontString(100);
    const wFont = ctx.measureText(probe).width;
    ctx.font = 'italic 700 100px serif';
    const wSerif = ctx.measureText(probe).width;
    if (!document.fonts.check(fontString(DEFAULT_SIZE), probe) || Math.abs(wFont - wSerif) < 0.5) {
      throw new Error('Font chưa được trình duyệt áp dụng');
    }
    ctx.font = fontString(1000);
    const capH = ctx.measureText('H').actualBoundingBoxAscent;
    if (capH > 0) state.capHeightEm = capH / 1000;

    state.fontReady = true;
    pill.className = 'pill pill-ok';
    pill.textContent = `✓ Font ${FONT_LABEL} đã sẵn sàng`;
    computeFitChecks();
    renderStudentTable();
    updateStats();
    requestPreview();
    if (state.frame && state.ref) runDetection(false);
  } catch (err) {
    console.error(err);
    state.fontReady = false;
    pill.className = 'pill pill-err';
    pill.textContent = '✕ Không tải được font';
    toast(`Không tải được font ${FONT_LABEL}: ${err.message}. Kiểm tra thư mục fonts/ (xem README).`, 'err', 10000);
  }
  updateButtons();
}

/* ---------------- Cấu hình ---------------- */
// Vị trí mặc định; nếu ảnh khung quá nhỏ để chứa điểm này thì dùng tâm ảnh
function defaultPosition(w, h) {
  if (DEFAULT_X < w && DEFAULT_Y < h) return { x: DEFAULT_X, y: DEFAULT_Y };
  return { x: Math.round(w / 2), y: Math.round(h / 2) };
}

function defaultCfg() {
  const f = state.frame;
  const pos = f ? defaultPosition(f.w, f.h) : { x: DEFAULT_X, y: DEFAULT_Y };
  return {
    size: DEFAULT_SIZE,
    color: DEFAULT_COLOR,
    align: 'center',
    vAlign: 'middle',
    autoShrink: false,
    minSize: 30,
    x: pos.x,
    y: pos.y,
    maxWidth: f ? Math.round(f.w * 0.7) : 1400,
  };
}

function syncCfgInputs() {
  const c = state.cfg;
  const setVal = (id, v) => {
    const el = $(id);
    if (document.activeElement !== el) el.value = v;
  };
  setVal('cfgSize', c.size);
  $('cfgSizeRange').value = clamp(c.size, 10, 200);
  $('cfgColor').value = c.color;
  setVal('cfgColorText', c.color);
  setVal('cfgX', c.x);
  setVal('cfgY', c.y);
  setVal('cfgMaxWidth', c.maxWidth);
  setVal('cfgMinSize', c.minSize);
  $('cfgAutoShrink').checked = c.autoShrink;
  document.querySelectorAll('input[name=align]').forEach((r) => { r.checked = r.value === c.align; });
  document.querySelectorAll('input[name=valign]').forEach((r) => { r.checked = r.value === c.vAlign; });

  const warn = $('colorWarn');
  if (c.color.toLowerCase() !== DEFAULT_COLOR) {
    warn.textContent = `⚠ Màu đang dùng ${c.color} khác màu mặc định ${DEFAULT_COLOR}.`;
    warn.classList.remove('hidden');
  } else {
    warn.classList.add('hidden');
  }
  const sizeNote = $('sizeNote');
  sizeNote.classList.toggle('warn', c.size !== DEFAULT_SIZE);
  sizeNote.textContent = c.size === DEFAULT_SIZE
    ? `Mặc định ${DEFAULT_SIZE} px — tính theo pixel trên ảnh gốc (không phải point).`
    : `⚠ Đang dùng ${c.size} px, khác cỡ mặc định ${DEFAULT_SIZE} px (pixel trên ảnh gốc).`;
}

let fitTimer = 0;
function onCfgChanged() {
  syncCfgInputs();
  requestPreview();
  clearTimeout(fitTimer);
  fitTimer = setTimeout(() => {
    computeFitChecks();
    renderStudentTable();
    updateStats();
  }, 150);
  invalidateResults('Cấu hình tên đã thay đổi');
}

function bindCfgInputs() {
  const num = (id, key, { min = -Infinity, max = Infinity, int = true } = {}) => {
    const el = $(id);
    el.addEventListener('input', () => {
      const v = int ? Math.round(Number(el.value)) : Number(el.value);
      if (el.value === '' || !Number.isFinite(v) || v < min || v > max) return;
      state.cfg[key] = v;
      onCfgChanged();
    });
    el.addEventListener('change', () => { el.value = state.cfg[key]; });
  };
  num('cfgSize', 'size', { min: 8, max: 500, int: false });
  num('cfgX', 'x');
  num('cfgY', 'y');
  num('cfgMaxWidth', 'maxWidth', { min: 10 });
  num('cfgMinSize', 'minSize', { min: 8, max: 500 });

  $('cfgSizeRange').addEventListener('input', (e) => {
    state.cfg.size = Number(e.target.value);
    onCfgChanged();
  });

  const setColor = (v) => {
    v = v.toLowerCase();
    if (v === '#000000') {
      toast(`Không dùng màu đen #000000 cho tên học sinh. Đã đặt lại màu ${DEFAULT_COLOR}.`, 'warn');
      v = DEFAULT_COLOR;
    }
    state.cfg.color = v;
    onCfgChanged();
  };
  $('cfgColor').addEventListener('input', (e) => setColor(e.target.value));
  $('cfgColorText').addEventListener('change', (e) => {
    let v = e.target.value.trim();
    if (!v.startsWith('#')) v = `#${v}`;
    if (!/^#[0-9a-f]{6}$/i.test(v)) {
      toast('Mã màu không hợp lệ. Dùng dạng #RRGGBB, ví dụ #5e1a36.', 'err');
      e.target.value = state.cfg.color;
      return;
    }
    setColor(v);
    e.target.value = state.cfg.color;
  });

  $('cfgAutoShrink').addEventListener('change', (e) => {
    state.cfg.autoShrink = e.target.checked;
    if (e.target.checked) toast('Tên vượt độ rộng vùng tên sẽ được thu nhỏ. Mỗi trường hợp được ghi rõ trong bảng danh sách.', 'warn', 6000);
    onCfgChanged();
  });

  document.querySelectorAll('input[name=align]').forEach((r) => r.addEventListener('change', () => {
    const prev = state.cfg.align;
    const next = r.value;
    // Giữ nguyên vị trí hiển thị của tên đang xem trước khi đổi kiểu căn
    const L = previewLayout();
    if (L) {
      const left = L.left;
      const width = L.right - L.left;
      if (next === 'left') state.cfg.x = Math.round(left);
      else if (next === 'right') state.cfg.x = Math.round(left + width);
      else state.cfg.x = Math.round(left + width / 2);
    }
    state.cfg.align = next;
    if (prev !== next) onCfgChanged();
  }));
  document.querySelectorAll('input[name=valign]').forEach((r) => r.addEventListener('change', () => {
    const half = (state.capHeightEm * state.cfg.size) / 2;
    if (state.cfg.vAlign === 'middle' && r.value === 'baseline') state.cfg.y = Math.round(state.cfg.y + half);
    if (state.cfg.vAlign === 'baseline' && r.value === 'middle') state.cfg.y = Math.round(state.cfg.y - half);
    state.cfg.vAlign = r.value;
    onCfgChanged();
  }));

  $('btnCenterX').addEventListener('click', () => {
    if (!state.frame) return toast('Chưa có ảnh khung.', 'warn');
    state.cfg.align = 'center';
    state.cfg.x = Math.round(state.frame.w / 2);
    onCfgChanged();
  });
  $('btnCenterY').addEventListener('click', () => {
    if (!state.frame) return toast('Chưa có ảnh khung.', 'warn');
    state.cfg.vAlign = 'middle';
    state.cfg.y = Math.round(state.frame.h / 2);
    onCfgChanged();
  });
  $('btnReset').addEventListener('click', () => {
    state.cfg = defaultCfg();
    onCfgChanged();
    toast(`Đã khôi phục mặc định: ${DEFAULT_SIZE} px, ${DEFAULT_COLOR}, căn giữa, X ${state.cfg.x}, Y ${state.cfg.y}.`, 'ok');
  });
  $('btnDetect').addEventListener('click', () => runDetection(true));
  $('btnApplySuggest').addEventListener('click', () => {
    const d = state.detected;
    if (!d) return;
    state.cfg.size = d.size;
    if (state.cfg.vAlign === 'middle') state.cfg.y = Math.round(d.baseline - (state.capHeightEm * d.size) / 2);
    else state.cfg.y = Math.round(d.baseline);
    onCfgChanged();
    toast(`Đã đổi cỡ chữ thành ${d.size} px theo ảnh mẫu (mặc định là ${DEFAULT_SIZE} px).`, 'warn', 6000);
  });
}

/* ---------------- Dàn trang & vẽ tên ---------------- */
// Tính vị trí thực của tên (dùng chung cho xem trước và xuất ảnh)
function layoutName(ctx, name, cfg) {
  ctx.textAlign = cfg.align;
  ctx.textBaseline = 'alphabetic';
  let size = cfg.size;
  ctx.font = fontString(size);
  let m = ctx.measureText(name);
  let width = m.actualBoundingBoxLeft + m.actualBoundingBoxRight;
  let shrunk = false;
  if (cfg.autoShrink && width > cfg.maxWidth && width > 0) {
    const s = Math.max(cfg.minSize, Math.floor(((cfg.size * cfg.maxWidth) / width) * 10) / 10);
    if (s < size) {
      size = s;
      shrunk = true;
      ctx.font = fontString(size);
      m = ctx.measureText(name);
      width = m.actualBoundingBoxLeft + m.actualBoundingBoxRight;
    }
  }
  const baseline = cfg.vAlign === 'middle' ? cfg.y + (state.capHeightEm * size) / 2 : cfg.y;
  return {
    size,
    shrunk,
    width,
    baseline,
    left: cfg.x - m.actualBoundingBoxLeft,
    right: cfg.x + m.actualBoundingBoxRight,
    top: baseline - m.actualBoundingBoxAscent,
    bottom: baseline + m.actualBoundingBoxDescent,
  };
}

// Vẽ một giấy khen hoàn chỉnh lên canvas có kích thước đúng bằng ảnh gốc.
// Mỗi lần vẽ đều xóa sạch và vẽ lại nền → các ảnh độc lập, không chồng tên.
function renderCertificate(canvas, name, cfg, beforeText) {
  const { img, w, h } = state.frame;
  if (canvas.width !== w) canvas.width = w;
  if (canvas.height !== h) canvas.height = h;
  const ctx = canvas.getContext('2d');
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.globalAlpha = 1;
  ctx.globalCompositeOperation = 'source-over';
  ctx.shadowColor = 'transparent';
  ctx.shadowBlur = 0;
  ctx.clearRect(0, 0, w, h);
  ctx.drawImage(img, 0, 0, w, h);
  if (beforeText) beforeText(ctx);
  if (!name) return null;
  const L = layoutName(ctx, name, cfg);
  ctx.save();
  ctx.font = fontString(L.size);
  ctx.fillStyle = cfg.color;
  ctx.textAlign = cfg.align;
  ctx.textBaseline = 'alphabetic';
  ctx.fillText(name, cfg.x, L.baseline);
  ctx.restore();
  return L;
}

function getExportCanvas() {
  if (!state.exportCanvas) state.exportCanvas = document.createElement('canvas');
  return state.exportCanvas;
}

async function renderStudentPNG(student) {
  const canvas = getExportCanvas();
  renderCertificate(canvas, student.name, state.cfg);
  return canvasToBlob(canvas);
}

/* ---------------- Kiểm tra độ dài tên ---------------- */
const measureCtx = document.createElement('canvas').getContext('2d');

function checkFit(name) {
  const L = layoutName(measureCtx, name, state.cfg);
  const { w, h } = state.frame;
  return {
    ...L,
    tooLong: L.width > state.cfg.maxWidth + 0.5,
    outside: L.left < 0 || L.right > w || L.top < 0 || L.bottom > h,
  };
}

function computeFitChecks() {
  const ready = state.fontReady && state.frame;
  for (const s of state.students) {
    s.warnings = [];
    s.fit = null;
    if (!ready || s.errors.length) continue;
    const f = checkFit(s.name);
    s.fit = f;
    if (f.shrunk) s.warnings.push(`Đã tự thu nhỏ cỡ chữ ${state.cfg.size} → ${f.size} px`);
    if (f.tooLong) s.warnings.push(`Tên rộng ${Math.round(f.width)} px, vượt vùng tên ${state.cfg.maxWidth} px`);
    if (f.outside) s.warnings.push('Tên bị tràn ra ngoài ảnh — sẽ bị cắt');
  }
}

const validStudents = () => state.students.filter((s) => !s.errors.length);

/* ---------------- Xem trước ---------------- */
let previewRaf = 0;
function requestPreview() {
  if (previewRaf) return;
  previewRaf = requestAnimationFrame(() => {
    previewRaf = 0;
    drawPreview();
  });
}

function currentPreviewName() {
  if (state.previewIdx >= 0 && state.students[state.previewIdx]) return state.students[state.previewIdx].name;
  return cleanText($('previewCustom').value);
}

function previewLayout() {
  if (!state.frame || !state.fontReady) return null;
  const name = currentPreviewName();
  return name ? layoutName(measureCtx, name, state.cfg) : null;
}

function cssPx() {
  const canvas = $('previewCanvas');
  const r = canvas.getBoundingClientRect();
  return r.width ? state.frame.w / r.width : 1;
}

function drawPreview() {
  const canvas = $('previewCanvas');
  if (!state.frame) return;
  const cfg = state.cfg;
  const { w, h } = state.frame;
  const name = state.fontReady ? currentPreviewName() : '';
  const mode = state.ref ? state.viewMode : 'frame';
  let L = null;

  if (mode === 'ref') {
    if (canvas.width !== w) canvas.width = w;
    if (canvas.height !== h) canvas.height = h;
    const ctx = canvas.getContext('2d');
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.globalAlpha = 1;
    ctx.clearRect(0, 0, w, h);
    ctx.drawImage(state.ref.img, 0, 0, w, h);
  } else {
    const opacity = Number($('overlayOpacity').value) / 100;
    L = renderCertificate(canvas, name, cfg, mode === 'overlay' ? (ctx) => {
      ctx.save();
      ctx.globalAlpha = opacity;
      ctx.drawImage(state.ref.img, 0, 0, w, h);
      ctx.restore();
    } : null);
  }

  if ($('showGuides').checked) drawGuides(canvas.getContext('2d'), L, mode);
  updatePreviewInfo(L, name);
}

function drawGuides(ctx, L, mode) {
  const cfg = state.cfg;
  const { w, h } = state.frame;
  const px = cssPx();
  ctx.save();
  ctx.lineWidth = px;

  // Trục giữa ảnh
  ctx.setLineDash([6 * px, 5 * px]);
  ctx.strokeStyle = 'rgba(0, 110, 230, .45)';
  ctx.beginPath();
  ctx.moveTo(w / 2, 0);
  ctx.lineTo(w / 2, h);
  ctx.stroke();

  // Vùng tên tối đa
  const mw = cfg.maxWidth;
  const zx = cfg.align === 'center' ? cfg.x - mw / 2 : cfg.align === 'left' ? cfg.x : cfg.x - mw;
  const zTop = L ? L.top - 6 * px : cfg.y - cfg.size * 0.6;
  const zH = L ? L.bottom - L.top + 12 * px : cfg.size * 1.2;
  ctx.strokeStyle = 'rgba(94, 26, 54, .55)';
  ctx.setLineDash([3 * px, 4 * px]);
  ctx.strokeRect(zx, zTop, mw, zH);

  // Vùng tên dò được trên ảnh mẫu
  if (state.detected && mode !== 'frame') {
    const d = state.detected;
    ctx.strokeStyle = 'rgba(20, 150, 80, .9)';
    ctx.setLineDash([5 * px, 3 * px]);
    ctx.strokeRect(d.x0, d.y0, d.x1 - d.x0 + 1, d.y1 - d.y0 + 1);
  }

  // Khung bao tên hiện tại
  if (L) {
    const bad = L.width > cfg.maxWidth + 0.5 || L.left < 0 || L.right > w || L.top < 0 || L.bottom > h;
    ctx.setLineDash([]);
    ctx.strokeStyle = bad ? 'rgba(210, 30, 20, .95)' : 'rgba(0, 110, 230, .85)';
    ctx.strokeRect(L.left, L.top, L.right - L.left, L.bottom - L.top);
    // Đường chân chữ
    ctx.setLineDash([2 * px, 3 * px]);
    ctx.beginPath();
    ctx.moveTo(L.left - 20 * px, L.baseline);
    ctx.lineTo(L.right + 20 * px, L.baseline);
    ctx.stroke();
  }

  // Điểm neo X, Y
  const r = 9 * px;
  ctx.setLineDash([]);
  ctx.strokeStyle = 'rgba(255, 255, 255, .95)';
  ctx.lineWidth = 3 * px;
  ctx.beginPath();
  ctx.moveTo(cfg.x - r, cfg.y); ctx.lineTo(cfg.x + r, cfg.y);
  ctx.moveTo(cfg.x, cfg.y - r); ctx.lineTo(cfg.x, cfg.y + r);
  ctx.stroke();
  ctx.strokeStyle = 'rgba(220, 30, 60, 1)';
  ctx.lineWidth = 1.5 * px;
  ctx.stroke();
  ctx.restore();
}

function updatePreviewInfo(L, name) {
  const cfg = state.cfg;
  const f = state.frame;
  let info = `X ${cfg.x} · Y ${cfg.y} · cỡ chữ ${L ? L.size : cfg.size} px · màu ${cfg.color}`;
  if (L) info += ` · tên rộng ${Math.round(L.width)} px`;
  $('posInfo').textContent = info;

  const warnEl = $('previewWarn');
  const msgs = [];
  if (!state.fontReady) msgs.push('Font đang được tải — chưa thể hiển thị tên đúng kiểu chữ.');
  if (L) {
    if (L.left < 0 || L.right > f.w || L.top < 0 || L.bottom > f.h) msgs.push(`Tên “${escapeHtml(name)}” bị tràn ra ngoài ảnh. Hãy giảm cỡ chữ hoặc dời vị trí.`);
    else if (L.width > cfg.maxWidth + 0.5) msgs.push(`Tên “${escapeHtml(name)}” rộng ${Math.round(L.width)} px, vượt độ rộng vùng tên ${cfg.maxWidth} px. Hãy giảm cỡ chữ, tăng độ rộng vùng tên hoặc bật “Tự thu nhỏ tên quá dài”.`);
    if (L.shrunk) msgs.push(`Tên này đang được tự thu nhỏ từ ${cfg.size} px xuống ${L.size} px.`);
  }
  const longCount = validStudents().filter((s) => s.fit && (s.fit.tooLong || s.fit.outside)).length;
  if (longCount) msgs.push(`Có <b>${longCount}</b> học sinh có tên quá dài hoặc tràn ảnh — chọn bộ lọc “Chỉ cảnh báo” ở bảng danh sách để xem.`);
  if (msgs.length) {
    warnEl.innerHTML = msgs.join('<br>');
    warnEl.classList.remove('hidden');
  } else {
    warnEl.classList.add('hidden');
  }
}

/* ---------------- Thu phóng ---------------- */
function fitZoom() {
  const st = $('stage');
  const aw = st.clientWidth - 32;
  const ah = st.clientHeight - 32;
  return clamp(Math.min(aw / state.frame.w, ah / state.frame.h), 0.02, 4);
}

function applyZoom() {
  if (!state.frame) return;
  if (state.fit) state.zoom = fitZoom();
  state.zoom = clamp(state.zoom, 0.02, 4);
  const canvas = $('previewCanvas');
  canvas.style.width = `${state.frame.w * state.zoom}px`;
  canvas.style.height = `${state.frame.h * state.zoom}px`;
  $('zoomLabel').textContent = `${Math.round(state.zoom * 100)}%`;
  requestPreview();
}

function setZoom(z) {
  state.fit = false;
  state.zoom = z;
  applyZoom();
}

/* ---------------- Kéo thả vị trí tên ---------------- */
function bindPreviewInteractions() {
  const canvas = $('previewCanvas');
  let drag = null;
  const locked = () => !state.frame || state.job || state.testing || (state.ref && state.viewMode === 'ref');
  const toImage = (e) => {
    const r = canvas.getBoundingClientRect();
    return { x: ((e.clientX - r.left) * state.frame.w) / r.width, y: ((e.clientY - r.top) * state.frame.h) / r.height };
  };

  canvas.addEventListener('pointerdown', (e) => {
    if (locked() || e.button !== 0) return;
    e.preventDefault();
    canvas.setPointerCapture(e.pointerId);
    const p = toImage(e);
    drag = { sx: p.x, sy: p.y, ox: state.cfg.x, oy: state.cfg.y, moved: false };
    canvas.classList.add('dragging');
    canvas.focus({ preventScroll: true });
  });
  canvas.addEventListener('pointermove', (e) => {
    if (!drag) return;
    const p = toImage(e);
    state.cfg.x = Math.round(drag.ox + p.x - drag.sx);
    state.cfg.y = Math.round(drag.oy + p.y - drag.sy);
    drag.moved = true;
    syncCfgInputs();
    requestPreview();
  });
  const end = () => {
    if (!drag) return;
    const moved = drag.moved;
    drag = null;
    canvas.classList.remove('dragging');
    if (moved) onCfgChanged();
  };
  canvas.addEventListener('pointerup', end);
  canvas.addEventListener('pointercancel', end);
  canvas.addEventListener('dblclick', (e) => {
    if (locked()) return;
    const p = toImage(e);
    state.cfg.x = Math.round(p.x);
    state.cfg.y = Math.round(p.y);
    onCfgChanged();
  });
  canvas.addEventListener('keydown', (e) => {
    if (locked()) return;
    const step = e.shiftKey ? 10 : 1;
    const d = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[e.key];
    if (!d) return;
    e.preventDefault();
    state.cfg.x += d[0];
    state.cfg.y += d[1];
    onCfgChanged();
  });

  // Cmd/Ctrl + cuộn (hoặc chụm 2 ngón trên trackpad) để thu phóng
  $('stage').addEventListener('wheel', (e) => {
    if (!state.frame || !(e.ctrlKey || e.metaKey)) return;
    e.preventDefault();
    setZoom(state.zoom * Math.exp(-e.deltaY * 0.01));
  }, { passive: false });

  $('zoomIn').addEventListener('click', () => setZoom(state.zoom * 1.25));
  $('zoomOut').addEventListener('click', () => setZoom(state.zoom / 1.25));
  $('zoom100').addEventListener('click', () => setZoom(1));
  $('zoomFit').addEventListener('click', () => { state.fit = true; applyZoom(); });
  if (window.ResizeObserver) new ResizeObserver(() => { if (state.fit) applyZoom(); }).observe($('stage'));

  $('previewStudent').addEventListener('change', (e) => {
    state.previewIdx = e.target.value === '' ? -1 : Number(e.target.value);
    $('customNameTool').classList.toggle('hidden', state.previewIdx >= 0);
    renderStudentTable();
    requestPreview();
  });
  $('previewCustom').addEventListener('input', requestPreview);
  $('viewMode').addEventListener('change', (e) => {
    state.viewMode = e.target.value;
    $('opacityTool').classList.toggle('hidden', state.viewMode !== 'overlay');
    $('previewCanvas').classList.toggle('locked', state.viewMode === 'ref');
    requestPreview();
  });
  $('overlayOpacity').addEventListener('input', requestPreview);
  $('showGuides').addEventListener('change', requestPreview);
}

function setPreviewStudent(idx) {
  state.previewIdx = idx;
  $('previewStudent').value = idx >= 0 ? String(idx) : '';
  $('customNameTool').classList.toggle('hidden', idx >= 0);
  renderStudentTable();
  requestPreview();
}

/* ---------------- Dò vị trí tên từ ảnh mẫu ---------------- */
// So sánh ảnh mẫu (có tên) với ảnh khung gốc: phần khác nhau chính là tên.
function diffMask(sx, sy, sw, sh, dw, dh) {
  const { frame, ref } = state;
  const make = () => {
    const c = document.createElement('canvas');
    c.width = dw;
    c.height = dh;
    return c.getContext('2d', { willReadFrequently: true });
  };
  const a = make();
  const b = make();
  a.drawImage(frame.img, sx, sy, sw, sh, 0, 0, dw, dh);
  const kx = ref.w / frame.w;
  const ky = ref.h / frame.h;
  b.drawImage(ref.img, sx * kx, sy * ky, sw * kx, sh * ky, 0, 0, dw, dh);
  const da = a.getImageData(0, 0, dw, dh).data;
  const db = b.getImageData(0, 0, dw, dh).data;
  const mask = new Uint8Array(dw * dh);
  for (let i = 0, p = 0; p < mask.length; i += 4, p++) {
    const d = Math.max(Math.abs(da[i] - db[i]), Math.abs(da[i + 1] - db[i + 1]), Math.abs(da[i + 2] - db[i + 2]), Math.abs(da[i + 3] - db[i + 3]));
    if (d > 60) mask[p] = 1;
  }
  return mask;
}

function groupRuns(counts, minCount, gap) {
  const runs = [];
  let cur = null;
  for (let i = 0; i < counts.length; i++) {
    if (counts[i] < minCount) continue;
    if (cur && i - cur.end <= gap) {
      cur.end = i;
      cur.mass += counts[i];
    } else {
      cur = { start: i, end: i, mass: counts[i] };
      runs.push(cur);
    }
  }
  return runs.sort((p, q) => q.mass - p.mass)[0] || null;
}

// Tìm cụm điểm khác biệt lớn nhất (dòng chữ tên)
function findTextBox(mask, w, h) {
  const rows = new Uint32Array(h);
  for (let y = 0; y < h; y++) {
    let n = 0;
    for (let x = 0, o = y * w; x < w; x++) n += mask[o + x];
    rows[y] = n;
  }
  const band = groupRuns(rows, 2, Math.max(2, Math.round(h * 0.012)));
  if (!band) return null;
  const cols = new Uint32Array(w);
  for (let y = band.start; y <= band.end; y++) {
    for (let x = 0, o = y * w; x < w; x++) cols[x] += mask[o + x];
  }
  const seg = groupRuns(cols, 1, Math.max(3, Math.round((band.end - band.start + 1) * 1.2)));
  if (!seg) return null;
  const rowCounts = [];
  let y0 = -1;
  let y1 = -1;
  for (let y = band.start; y <= band.end; y++) {
    let n = 0;
    for (let x = seg.start, o = y * w; x <= seg.end; x++) n += mask[o + x];
    rowCounts.push(n);
    if (n) {
      if (y0 < 0) y0 = y;
      y1 = y;
    }
  }
  if (y0 < 0) return null;
  return { x0: seg.start, x1: seg.end, y0, y1, rows: rowCounts.slice(y0 - band.start, y1 - band.start + 1), mass: seg.mass };
}

function detectNamePosition() {
  const W = state.frame.w;
  const H = state.frame.h;
  // Lượt 1: độ phân giải thấp để tìm vùng khác biệt
  const k = Math.min(1, 1200 / Math.max(W, H));
  const cw = Math.max(1, Math.round(W * k));
  const ch = Math.max(1, Math.round(H * k));
  const coarse = findTextBox(diffMask(0, 0, W, H, cw, ch), cw, ch);
  if (!coarse) return null;
  // Lượt 2: độ phân giải gốc trong vùng đã tìm (mở rộng biên)
  const kx = W / cw;
  const ky = H / ch;
  const pad = Math.ceil((coarse.y1 - coarse.y0 + 1) * ky) + 4;
  const rx = clamp(Math.floor(coarse.x0 * kx) - pad, 0, W - 1);
  const ry = clamp(Math.floor(coarse.y0 * ky) - pad, 0, H - 1);
  const rw = clamp(Math.ceil((coarse.x1 + 1) * kx) + pad, rx + 1, W) - rx;
  const rh = clamp(Math.ceil((coarse.y1 + 1) * ky) + pad, ry + 1, H) - ry;
  const box = findTextBox(diffMask(rx, ry, rw, rh, rw, rh), rw, rh);
  if (!box) return null;

  // Dải chữ thường (x-height): các dòng có mật độ điểm cao nhất → suy ra đường chân chữ
  const rows = box.rows;
  let peak = 0;
  for (let i = 1; i < rows.length; i++) if (rows[i] > rows[peak]) peak = i;
  const thr = rows[peak] * 0.45;
  let top = peak;
  let bottom = peak;
  while (top > 0 && rows[top - 1] >= thr) top--;
  while (bottom < rows.length - 1 && rows[bottom + 1] >= thr) bottom++;
  const xHeight = bottom - top + 1;
  const baseline = ry + box.y0 + bottom + 1;

  return {
    x0: rx + box.x0,
    x1: rx + box.x1,
    y0: ry + box.y0,
    y1: ry + box.y1,
    cx: rx + (box.x0 + box.x1 + 1) / 2,
    baseline,
    xHeight,
    size: Math.round(xHeight / X_HEIGHT_EM),
  };
}

function runDetection(apply) {
  const info = $('detectInfo');
  const btnApply = $('btnApplySuggest');
  btnApply.classList.add('hidden');
  if (!state.frame || !state.ref) {
    info.textContent = 'Tải ảnh khung gốc và ảnh mẫu có tên để tự xác định vị trí tên.';
    return;
  }
  let d = null;
  try {
    d = detectNamePosition();
  } catch (err) {
    console.error(err);
  }
  state.detected = d;
  if (!d) {
    info.innerHTML = '⚠ Không tìm thấy khác biệt rõ ràng giữa ảnh mẫu và ảnh khung gốc. Hãy kiểm tra lại 2 ảnh, hoặc kéo tên trên ảnh để đặt vị trí thủ công.';
    info.className = 'note warn';
    requestPreview();
    return;
  }
  if (apply) {
    state.cfg.align = 'center';
    state.cfg.x = Math.round(d.cx);
    state.cfg.y = state.cfg.vAlign === 'middle'
      ? Math.round(d.baseline - (state.capHeightEm * state.cfg.size) / 2)
      : Math.round(d.baseline);
    onCfgChanged();
  }
  const diffPct = Math.abs(d.size - state.cfg.size) / state.cfg.size;
  info.className = 'note';
  info.innerHTML = `✓ Đã xác định tên trên ảnh mẫu: tâm ngang X ≈ <b>${Math.round(d.cx)}</b>, chân chữ Y ≈ <b>${d.baseline}</b>, rộng ${d.x1 - d.x0 + 1} px.`
    + `<br>Cỡ chữ ước tính trên ảnh mẫu ≈ <b>${d.size} px</b> (đang dùng ${state.cfg.size} px).`
    + (diffPct > 0.08 ? ` Cỡ ước tính khác cỡ đang dùng — bạn có thể giữ nguyên ${state.cfg.size} px hoặc áp dụng cỡ gợi ý.` : '')
    + (apply ? '' : `<br>Đang giữ vị trí mặc định X ${state.cfg.x}, Y ${state.cfg.y}. Bấm “Dùng vị trí dò được” nếu muốn đặt tên theo ảnh mẫu.`)
    + '<br>Chọn chế độ xem “Chồng ảnh mẫu” để đối chiếu.';
  if (diffPct > 0.08) btnApply.classList.remove('hidden');
  if (state.ref.w / state.ref.h - state.frame.w / state.frame.h > 0.01 || state.frame.w / state.frame.h - state.ref.w / state.ref.h > 0.01) {
    info.innerHTML += '<br><span class="warn">⚠ Ảnh mẫu khác tỉ lệ với ảnh khung, vị trí dò được có thể lệch — hãy kiểm tra lại bằng mắt.</span>';
  }
  requestPreview();
}

/* ---------------- Tải file ---------------- */
function setZone(kind, status, text) {
  const map = {
    excel: ['dzExcel', 'excelName', 'pickExcel'],
    frame: ['dzFrame', 'frameName', 'pickFrame'],
    ref: ['dzRef', 'refName', 'pickRef'],
  };
  const [zone, nameEl, btn] = map[kind];
  $(zone).classList.toggle('is-ok', status === 'ok');
  $(zone).classList.toggle('is-error', status === 'error');
  $(nameEl).textContent = text;
  $(btn).textContent = status === 'ok' ? 'Đổi file' : 'Chọn file';
}

function bindDropzone(zoneId, inputId, handler) {
  const zone = $(zoneId);
  const input = $(inputId);
  zone.addEventListener('click', () => {
    if (state.job || state.testing) return toast('Đang tạo ảnh, vui lòng chờ hoặc hủy trước khi đổi file.', 'warn');
    input.click();
  });
  input.addEventListener('click', (e) => e.stopPropagation());
  input.addEventListener('change', () => {
    const file = input.files && input.files[0];
    input.value = '';
    if (file) handler(file);
  });
  zone.addEventListener('dragenter', (e) => { e.preventDefault(); zone.classList.add('drag'); });
  zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('drag'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('drag');
    if (state.job || state.testing) return;
    const files = e.dataTransfer && e.dataTransfer.files;
    if (files && files.length > 1) toast('Chỉ nhận 1 file cho mỗi ô — đã dùng file đầu tiên.', 'warn');
    if (files && files[0]) handler(files[0]);
  });
}

function formatTime(ms) {
  const d = new Date(ms);
  const pad = (n) => String(n).padStart(2, '0');
  return `${pad(d.getHours())}:${pad(d.getMinutes())} ${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
}

function describeFile(file) {
  return `${file.name} · ${formatBytes(file.size)} · sửa lúc ${formatTime(file.lastModified)}`;
}

async function parseExcelFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array', cellText: true, cellDates: true });
  if (!wb.SheetNames.length) throw new Error('file không có trang tính nào');
  return wb;
}

async function handleExcel(file) {
  if (!EXCEL_EXT.includes(extOf(file.name))) {
    setZone('excel', 'error', `“${file.name}” không phải file Excel (.xlsx, .xls)`);
    toast('File danh sách phải có định dạng .xlsx hoặc .xls.', 'err');
    return;
  }
  // Bỏ ngay dữ liệu của file cũ để không thể đọc nhầm trong lúc file mới đang tải
  const seq = ++state.excelSeq;
  state.excelLoading = true;
  state.excelFile = null;
  state.workbook = null;
  state.sheetMatrix = null;
  clearStudents();
  $('listSource').textContent = '';
  setZone('excel', '', `Đang đọc ${file.name}…`);
  updateButtons();
  try {
    const wb = await parseExcelFile(file);
    if (seq !== state.excelSeq) return;   // đã chọn file khác trong lúc đang đọc
    state.workbook = wb;
    state.excelFile = file;
    const sel = $('sheetSelect');
    sel.innerHTML = wb.SheetNames.map((n, i) => `<option value="${i}">${escapeHtml(n)}</option>`).join('');
    const firstNonEmpty = wb.SheetNames.findIndex((n) => wb.Sheets[n] && wb.Sheets[n]['!ref']);
    sel.value = String(Math.max(0, firstNonEmpty));
    setZone('excel', 'ok', `${describeFile(file)} · ${wb.SheetNames.length} trang tính`);
    $('excelOptions').classList.remove('hidden');
    loadSheet(true);
    toast(`Đã tải “${file.name}”. Kiểm tra cột mã và cột tên rồi bấm “Đọc danh sách học sinh”.`, 'ok');
  } catch (err) {
    if (seq !== state.excelSeq) return;
    console.error(err);
    $('excelOptions').classList.add('hidden');
    setZone('excel', 'error', `Không đọc được “${file.name}”: file hỏng, có mật khẩu hoặc không phải Excel.`);
    toast('Không đọc được file Excel.', 'err');
  } finally {
    if (seq === state.excelSeq) state.excelLoading = false;
    updateButtons();
    updateSteps();
  }
}

async function handleImage(kind, file) {
  const label = kind === 'frame' ? 'ảnh khung gốc' : 'ảnh mẫu';
  const typeOk = !file.type || /^image\/(png|jpeg|webp)$/.test(file.type);
  if (!IMAGE_EXT.includes(extOf(file.name)) || !typeOk) {
    setZone(kind, 'error', `“${file.name}” không phải ảnh hợp lệ (.png, .jpg, .webp)`);
    toast(`File ${label} phải là ảnh .png, .jpg hoặc .webp.`, 'err');
    return;
  }
  const url = URL.createObjectURL(file);
  let img;
  try {
    img = await loadImage(url);
    if (img.decode) await img.decode().catch(() => {});
    if (!img.naturalWidth || !img.naturalHeight) throw new Error('ảnh rỗng');
  } catch (err) {
    URL.revokeObjectURL(url);
    setZone(kind, 'error', `Không mở được “${file.name}” — file hỏng hoặc định dạng không hỗ trợ.`);
    toast(`Không mở được ${label}.`, 'err');
    return;
  }
  const prev = state[kind];
  if (prev) URL.revokeObjectURL(prev.url);
  const w = img.naturalWidth;
  const h = img.naturalHeight;
  state[kind] = { img, w, h, url, fileName: file.name, size: file.size };
  $(kind === 'frame' ? 'frameThumb' : 'refThumb').innerHTML = `<img src="${url}" alt="">`;
  setZone(kind, 'ok', `${file.name} · ${w} × ${h} px · ${formatBytes(file.size)}`);

  if (kind === 'frame') {
    const sizeChanged = !prev || prev.w !== w || prev.h !== h;
    if (sizeChanged) {
      state.detected = null;
      state.cfg = { ...state.cfg, ...defaultPosition(w, h), maxWidth: Math.round(w * 0.7) };
      state.fit = true;
    }
    $('stageEmpty').classList.add('hidden');
    $('previewCanvas').classList.remove('hidden');
    if (w * h > SAFARI_CANVAS_LIMIT) toast(`Ảnh khung rất lớn (${w} × ${h} px). Safari có thể không xử lý được — nên dùng Chrome.`, 'warn', 8000);
    invalidateResults('Ảnh khung đã thay đổi');
  }
  if (kind === 'ref' && state.frame && (w !== state.frame.w || h !== state.frame.h)) {
    toast(`Ảnh mẫu (${w} × ${h}) khác kích thước ảnh khung (${state.frame.w} × ${state.frame.h}) — sẽ được co giãn để đối chiếu.`, 'warn', 7000);
  }
  if (kind === 'frame' && state.ref && (state.ref.w !== w || state.ref.h !== h)) {
    toast(`Ảnh mẫu (${state.ref.w} × ${state.ref.h}) khác kích thước ảnh khung (${w} × ${h}) — sẽ được co giãn để đối chiếu.`, 'warn', 7000);
  }
  updateImageInfo();
  // Chỉ hiển thị kết quả dò, giữ nguyên vị trí mặc định (bấm “Dùng vị trí dò được” để áp dụng)
  if (state.frame && state.ref) runDetection(false);
  if (kind === 'frame') onCfgChanged();
  computeFitChecks();
  renderStudentTable();
  updateStats();
  applyZoom();
  updateButtons();
  updateSteps();
}

function updateImageInfo() {
  const parts = [];
  if (state.frame) parts.push(`Ảnh gốc: <b>${state.frame.w} × ${state.frame.h} px</b>`);
  if (state.ref) parts.push(`Ảnh mẫu: ${state.ref.w} × ${state.ref.h} px`);
  $('imgInfo').innerHTML = parts.join(' · ') || 'Chưa có ảnh khung';
}

/* ---------------- Đọc Excel ---------------- */
function cellToText(cell) {
  if (!cell || cell.t === 'z' || cell.t === 'e') return '';
  let s;
  if (cell.t === 'n') {
    // Ưu tiên chuỗi hiển thị (giữ số 0 ở đầu nếu ô định dạng 000…), tránh dạng 1.2E+12
    if (cell.w != null && !/e[+-]?\d/i.test(cell.w)) s = cell.w;
    else s = Number.isInteger(cell.v) ? cell.v.toFixed(0) : String(cell.v);
  } else if (cell.t === 'd') {
    s = cell.w != null ? cell.w : String(cell.v);
  } else if (cell.t === 'b') {
    s = cell.v ? 'TRUE' : 'FALSE';
  } else {
    s = cell.v == null ? '' : String(cell.v);
  }
  return cleanText(s);
}

function readSheetMatrix(ws) {
  const rows = new Map();
  let maxC = -1;
  for (const addr of Object.keys(ws)) {
    if (addr[0] === '!') continue;
    const text = cellToText(ws[addr]);
    if (text === '') continue;
    const { r, c } = XLSX.utils.decode_cell(addr);
    if (!rows.has(r)) rows.set(r, []);
    rows.get(r)[c] = text;
    if (c > maxC) maxC = c;
  }
  const order = [...rows.keys()].sort((a, b) => a - b);   // giữ đúng thứ tự dòng trong Excel
  return { maxC, rows: order.map((r) => ({ excelRow: r + 1, cells: rows.get(r) })) };
}

const HEADER_HINT = /\b(ma|ten|ho|name|id|code|stt|sbd|lop|mshs|mssv)\b/;
const CODE_HEADER = /(^|\s)(ma|id|code|mshs|mssv|msv|sbd)(\s|$)|ma hoc sinh|ma so|so bao danh|student id/;
const NAME_HEADER = /(ten|name)/;

function guessColumns(matrix, hasHeader) {
  let code = -1;
  let name = -1;
  if (hasHeader && matrix.rows[0]) {
    const cells = matrix.rows[0].cells;
    for (let c = 0; c <= matrix.maxC; c++) {
      const n = normKey(cells[c]);
      if (!n) continue;
      if (code < 0 && CODE_HEADER.test(n)) code = c;
      else if (name < 0 && NAME_HEADER.test(n) && !/lop|truong/.test(n)) name = c;
    }
  }
  const data = matrix.rows.slice(hasHeader ? 1 : 0, 40);
  const ratio = (c, test) => {
    const vals = data.map((r) => r.cells[c]).filter(Boolean);
    return vals.length ? vals.filter(test).length / vals.length : 0;
  };
  if (name < 0) {
    let best = 0;
    for (let c = 0; c <= matrix.maxC; c++) {
      if (c === code) continue;
      const r = ratio(c, (v) => /\s/.test(v) && /\p{L}/u.test(v) && !/\d/.test(v));
      if (r > best + 0.01) { best = r; name = c; }
    }
  }
  if (code < 0) {
    let best = 0;
    for (let c = 0; c <= matrix.maxC; c++) {
      if (c === name) continue;
      const r = ratio(c, (v) => !/\s/.test(v) && /\d/.test(v) && /\p{L}/u.test(v));
      if (r > best + 0.01) { best = r; code = c; }
    }
  }
  if (code < 0) code = name === 0 ? 1 : 0;
  if (name < 0 || name === code) name = code === 0 ? 1 : 0;
  return { code, name };
}

function loadSheet(autoHeader) {
  const wb = state.workbook;
  if (!wb) return;
  const ws = wb.Sheets[wb.SheetNames[Number($('sheetSelect').value)]];
  state.sheetMatrix = ws ? readSheetMatrix(ws) : { maxC: -1, rows: [] };
  const m = state.sheetMatrix;
  if (autoHeader) {
    const first = m.rows[0];
    $('hasHeader').checked = !!first && first.cells.some((v) => HEADER_HINT.test(normKey(v)));
  }
  buildColumnOptions();
}

function buildColumnOptions() {
  const m = state.sheetMatrix;
  const hasHeader = $('hasHeader').checked;
  const headers = hasHeader && m.rows[0] ? m.rows[0].cells : [];
  const nCols = Math.max(m.maxC + 1, 2);
  const opts = [];
  for (let c = 0; c < nCols; c++) {
    const letter = XLSX.utils.encode_col(c);
    const h = headers[c] ? ` — ${escapeHtml(headers[c])}` : '';
    opts.push(`<option value="${c}">Cột ${letter}${h}</option>`);
  }
  $('codeCol').innerHTML = opts.join('');
  $('nameCol').innerHTML = opts.join('');
  const g = guessColumns(m, hasHeader);
  $('codeCol').value = String(g.code);
  $('nameCol').value = String(g.name);
  renderSampleRows();
}

function renderSampleRows() {
  const m = state.sheetMatrix;
  if (!m) return;
  const hasHeader = $('hasHeader').checked;
  const cc = Number($('codeCol').value);
  const nc = Number($('nameCol').value);
  const data = m.rows.slice(hasHeader ? 1 : 0);
  const sample = data.slice(0, 3);
  if (!sample.length) {
    $('sampleRows').innerHTML = '<p class="note warn">Trang tính này không có dòng dữ liệu nào.</p>';
    return;
  }
  $('sampleRows').innerHTML = `<table><tr><th>Dòng</th><th>Mã</th><th>Tên</th></tr>${sample.map((r) =>
    `<tr><td>${r.excelRow}</td><td class="mono">${escapeHtml(r.cells[cc] || '—')}</td><td>${escapeHtml(r.cells[nc] || '—')}</td></tr>`).join('')}</table>`
    + `<p class="note">Xem trước 3 / ${data.length} dòng dữ liệu.</p>`;
}

// Đọc lại đúng file đã chọn trên website ngay lúc bấm nút, để luôn lấy nội dung mới nhất của file đó
async function rereadSelectedExcel() {
  const file = state.excelFile;
  const sheetName = state.workbook.SheetNames[Number($('sheetSelect').value)];
  const seq = state.excelSeq;
  let wb;
  try {
    wb = await parseExcelFile(file);
  } catch (err) {
    console.error(err);
    throw new Error(`Không đọc lại được “${file.name}”. File có thể đã bị sửa, đổi tên hoặc di chuyển sau khi chọn — hãy bấm “Đổi file” và chọn lại file này.`);
  }
  if (seq !== state.excelSeq) return false;
  if (!wb.Sheets[sheetName]) throw new Error(`Trang tính “${sheetName}” không còn trong file. Hãy chọn lại file.`);
  state.workbook = wb;
  const cc = $('codeCol').value;
  const nc = $('nameCol').value;
  state.sheetMatrix = readSheetMatrix(wb.Sheets[sheetName]);
  const nOpts = $('codeCol').options.length;
  if (state.sheetMatrix.maxC + 1 > nOpts) {
    buildColumnOptions();
    $('codeCol').value = cc;
    $('nameCol').value = nc;
  }
  renderSampleRows();
  return true;
}

async function readStudentList() {
  if (state.job || state.testing) return;
  if (state.excelLoading) {
    toast('File Excel đang được đọc, vui lòng chờ trong giây lát.', 'warn');
    return;
  }
  if (!state.excelFile || !state.workbook || !state.sheetMatrix) {
    toast('Chưa tải file Excel danh sách học sinh.', 'err');
    return;
  }
  const cc = Number($('codeCol').value);
  const nc = Number($('nameCol').value);
  if (cc === nc) {
    toast('Cột mã học sinh và cột tên học sinh phải khác nhau.', 'err');
    return;
  }
  state.excelLoading = true;
  updateButtons();
  try {
    if (!(await rereadSelectedExcel())) return;
  } catch (err) {
    toast(err.message, 'err', 10000);
    return;
  } finally {
    state.excelLoading = false;
    updateButtons();
  }
  const file = state.excelFile;
  const sheetName = state.workbook.SheetNames[Number($('sheetSelect').value)];
  const rows = state.sheetMatrix.rows.slice($('hasHeader').checked ? 1 : 0);
  const students = rows.map((r, i) => ({
    idx: i,
    row: r.excelRow,
    code: cleanText(r.cells[cc]),
    name: cleanText(r.cells[nc]),
    errors: [],
    warnings: [],
    fit: null,
  }));

  const groups = new Map();
  for (const s of students) {
    if (!s.code) s.errors.push('Thiếu mã học sinh');
    else if (INVALID_FILENAME.test(s.code)) s.errors.push('Mã chứa ký tự không dùng được trong tên file ( / \\ : * ? " < > | )');
    else if (s.code.startsWith('.')) s.errors.push('Mã không được bắt đầu bằng dấu chấm');
    else if (s.code.length > 150) s.errors.push('Mã quá dài để làm tên file');
    if (!s.name) s.errors.push('Thiếu tên học sinh');
    if (s.code) {
      // macOS không phân biệt hoa/thường trong tên file → HS001 và hs001 là trùng
      const key = s.code.normalize('NFC').toLowerCase();
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(s);
    }
  }
  for (const list of groups.values()) {
    if (list.length < 2) continue;
    for (const s of list) {
      const others = list.filter((o) => o !== s).map((o) => o.row).join(', ');
      s.errors.push(`Mã bị trùng với dòng ${others}`);
    }
  }

  state.students = students;
  // Đặt lại bộ lọc/tìm kiếm của lần đọc trước để không che mất danh sách mới
  const nErr = students.filter((s) => s.errors.length).length;
  $('listFilter').value = nErr ? 'error' : 'all';
  $('listSearch').value = '';
  computeFitChecks();
  const firstValid = students.findIndex((s) => !s.errors.length);
  fillPreviewSelect();
  setPreviewStudent(firstValid);
  updateStats();
  invalidateResults(state.testDone || state.output ? 'Danh sách học sinh đã được đọc lại' : '');
  updateButtons();
  updateSteps();

  $('listSource').innerHTML = `Nguồn: <b>${escapeHtml(file.name)}</b> · ${formatBytes(file.size)} · sửa lúc ${formatTime(file.lastModified)} · trang tính “${escapeHtml(sheetName)}” · đọc lúc ${formatTime(Date.now())}`;
  if (!students.length) {
    toast(`Không tìm thấy dòng dữ liệu nào trong “${file.name}”.`, 'warn');
  } else if (nErr) {
    toast(`Đã đọc ${students.length} học sinh từ “${file.name}” — có ${nErr} dòng lỗi sẽ bị bỏ qua. Bảng đang lọc “Chỉ dòng lỗi”, chọn “Tất cả” để xem toàn bộ.`, 'warn', 8000);
  } else {
    toast(`Đã đọc ${students.length} học sinh từ “${file.name}”, không có lỗi dữ liệu.`, 'ok');
  }
}

function clearStudents() {
  state.students = [];
  $('listSource').textContent = '';
  state.previewIdx = -1;
  fillPreviewSelect();
  setPreviewStudent(-1);
  updateStats();
  invalidateResults('');
}

function fillPreviewSelect() {
  const opts = ['<option value="">— Tên mẫu tự nhập —</option>'];
  for (const s of state.students) {
    if (!s.name) continue;
    opts.push(`<option value="${s.idx}">${escapeHtml(s.code || '(thiếu mã)')} — ${escapeHtml(s.name)}</option>`);
  }
  $('previewStudent').innerHTML = opts.join('');
}

/* ---------------- Bảng danh sách ---------------- */
function statusOf(s) {
  if (s.errors.length) return 'error';
  if (s.warnings.length) return 'warn';
  return 'ok';
}

function renderStudentTable() {
  const body = $('studentBody');
  if (!state.students.length) {
    body.innerHTML = '<tr class="empty"><td colspan="6">Chưa có dữ liệu. Tải file Excel rồi bấm “Đọc danh sách học sinh”.</td></tr>';
    return;
  }
  const filter = $('listFilter').value;
  const q = normKey($('listSearch').value);
  const rows = [];
  let shown = 0;
  for (const s of state.students) {
    const st = statusOf(s);
    if (filter !== 'all' && filter !== st) continue;
    if (q && !normKey(`${s.code} ${s.name}`).includes(q)) continue;
    shown++;
    const tag = st === 'error' ? '<span class="tag tag-err">Lỗi — bỏ qua</span>'
      : st === 'warn' ? '<span class="tag tag-warn">Cảnh báo</span>'
        : '<span class="tag tag-ok">Hợp lệ</span>';
    const issues = s.errors.map((e) => `<span class="issue err">• ${escapeHtml(e)}</span>`).join('')
      + s.warnings.map((e) => `<span class="issue warn">• ${escapeHtml(e)}</span>`).join('');
    const canUse = !s.errors.length;
    rows.push(`<tr class="row-${st}${s.idx === state.previewIdx ? ' selected' : ''}" data-idx="${s.idx}">
      <td>${s.idx + 1}</td>
      <td>${s.row}</td>
      <td class="code">${s.code ? `${escapeHtml(s.code)}.png` : '—'}</td>
      <td>${s.name ? escapeHtml(s.name) : '<span class="muted">—</span>'}</td>
      <td>${tag}${issues}</td>
      <td class="actions">
        <button type="button" class="linkbtn" data-act="view" ${s.name ? '' : 'disabled'}>Xem</button>
        <button type="button" class="linkbtn" data-act="png" ${canUse ? '' : 'disabled'}>PNG</button>
      </td></tr>`);
  }
  body.innerHTML = rows.join('') || '<tr class="empty"><td colspan="6">Không có dòng nào khớp bộ lọc.</td></tr>';
  if (shown === 0 && filter !== 'all' && !q) {
    body.innerHTML = '<tr class="empty"><td colspan="6">Không có dòng nào thuộc nhóm này.</td></tr>';
  }
}

function bindTable() {
  $('studentBody').addEventListener('click', (e) => {
    const btn = e.target.closest('button[data-act]');
    if (!btn) return;
    const idx = Number(btn.closest('tr').dataset.idx);
    if (btn.dataset.act === 'view') {
      setPreviewStudent(idx);
      $('secC').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } else {
      downloadOne(idx);
    }
  });
  $('listFilter').addEventListener('change', renderStudentTable);
  $('listSearch').addEventListener('input', renderStudentTable);
}

function updateStats(done) {
  const total = state.students.length;
  const err = state.students.filter((s) => s.errors.length).length;
  const warn = state.students.filter((s) => !s.errors.length && s.warnings.length).length;
  $('statTotal').textContent = total;
  $('statValid').textContent = total - err;
  $('statErr').textContent = err;
  $('statWarn').textContent = warn;
  $('statDone').textContent = done != null ? done : state.output ? state.output.count : 0;
}

/* ---------------- Trạng thái nút & tiến trình ---------------- */
function updateButtons() {
  const busy = !!state.job || state.testing;
  const hasValid = state.students.some((s) => !s.errors.length);
  const canRender = state.fontReady && !!state.frame;
  $('btnReadList').disabled = busy || state.excelLoading || !state.workbook;
  $('btnReadList2').disabled = busy || state.excelLoading || !state.workbook;
  $('btnPreview').disabled = !state.frame;
  $('btnDownloadOne').disabled = busy || !canRender || !hasValid;
  $('btnTest').disabled = busy || !canRender || !hasValid;
  $('btnGenerate').disabled = busy || !(state.testDone && $('testApprove').checked);
  $('btnCancel').disabled = !state.job;
  $('btnZip').disabled = busy || !state.output || !state.output.count;
  $('btnDetect').disabled = busy || !(state.frame && state.ref);
  $('cfgFieldset').disabled = busy;
  $('testApprove').disabled = busy;
  $('previewCanvas').classList.toggle('locked', busy || state.viewMode === 'ref');
  document.querySelectorAll('#excelOptions select, #excelOptions input').forEach((el) => { el.disabled = busy; });
}

function updateSteps() {
  const done = {
    excel: !!state.workbook,
    frame: !!state.frame,
    ref: !!state.ref,
    list: state.students.length > 0,
    test: state.testDone && $('testApprove').checked,
    done: !!state.output,
  };
  document.querySelectorAll('.steps li').forEach((li) => li.classList.toggle('done', !!done[li.dataset.step]));
}

function updateProgress(fraction, text, finished) {
  $('progressBar').style.width = `${Math.round(clamp(fraction, 0, 1) * 1000) / 10}%`;
  $('progressBar').classList.toggle('done', !!finished);
  $('progressText').textContent = text;
}

function clearTestResults() {
  for (const url of state.testUrls) URL.revokeObjectURL(url);
  state.testUrls = [];
  $('testGrid').innerHTML = '';
  $('approveWrap').classList.add('hidden');
  $('testApprove').checked = false;
  state.testDone = false;
}

function releaseOutput() {
  state.output = null;   // trình duyệt tự giải phóng Blob khi không còn tham chiếu
}

function invalidateResults(reason) {
  if (state.job || state.testing) return;
  const had = state.testDone || !!state.output;
  clearTestResults();
  releaseOutput();
  if (had) {
    updateProgress(0, 'Chưa bắt đầu');
    if (reason) showGenMessage('warn', `${escapeHtml(reason)} — kết quả cũ đã được hủy. Vui lòng tạo thử lại trước khi tạo toàn bộ.`);
  }
  updateStats();
  updateButtons();
  updateSteps();
}

/* ---------------- Kiểm tra trước khi tạo ---------------- */
function preflight(forFull) {
  const fail = (msg) => {
    showGenMessage('err', msg);
    toast(msg.replace(/<[^>]+>/g, ''), 'err');
    return false;
  };
  if (!state.fontReady) return fail(`Font ${FONT_LABEL} chưa tải xong — chưa thể xuất ảnh.`);
  if (!state.frame) return fail('Chưa tải ảnh khung giấy khen gốc.');
  if (!state.students.length) return fail('Chưa đọc danh sách học sinh. Tải file Excel và bấm “Đọc danh sách học sinh”.');
  computeFitChecks();
  renderStudentTable();
  updateStats();
  const valid = validStudents();
  if (!valid.length) return fail('Không có học sinh hợp lệ nào để tạo giấy khen.');
  const outside = valid.filter((s) => s.fit && s.fit.outside);
  if (outside.length) {
    return fail(`Có <b>${outside.length}</b> tên bị tràn ra ngoài ảnh (ví dụ: ${escapeHtml(outside[0].name)}). Hãy giảm cỡ chữ hoặc điều chỉnh vị trí trước khi tạo.`);
  }
  if (forFull) {
    if (!state.testDone || !$('testApprove').checked) return fail('Hãy tạo thử và xác nhận đã kiểm tra ảnh thử trước khi tạo toàn bộ.');
    const long = valid.filter((s) => s.fit && s.fit.tooLong);
    if (long.length && !window.confirm(`Có ${long.length} tên vượt độ rộng vùng tên (${state.cfg.maxWidth} px), ví dụ “${long[0].name}”.\n\nVẫn tiếp tục tạo toàn bộ giấy khen?`)) return false;
  }
  if (!state.ref) toast('Chưa tải ảnh mẫu có tên — vị trí tên đang được đặt thủ công.', 'warn');
  return true;
}

/* ---------------- Tạo thử ---------------- */
function pickTestStudents() {
  const valid = validStudents();
  const picks = [];
  const add = (s) => { if (s && !picks.includes(s) && picks.length < TEST_COUNT) picks.push(s); };
  add(valid[0]);
  let longest = null;
  for (const s of valid) if (!longest || s.fit.width > longest.fit.width) longest = s;
  add(longest);
  add(valid[valid.length - 1]);
  for (const s of valid) add(s);
  return picks;
}

async function runTest() {
  if (state.job || state.testing) return;
  if (!preflight(false)) return;
  clearTestResults();
  releaseOutput();
  state.testing = true;
  updateButtons();
  showGenMessage('', '');
  const grid = $('testGrid');
  try {
    const picks = pickTestStudents();
    for (let i = 0; i < picks.length; i++) {
      const s = picks[i];
      updateProgress((i + 1) / picks.length, `Đang tạo ảnh thử ${i + 1}/${picks.length}: ${s.code}.png`);
      const blob = await renderStudentPNG(s);
      const dims = await imageSizeOf(blob);
      const sizeOk = dims.w === state.frame.w && dims.h === state.frame.h;
      const url = URL.createObjectURL(blob);
      state.testUrls.push(url);
      const card = document.createElement('div');
      card.className = 'test-item';
      card.innerHTML = `<a href="${url}" target="_blank" rel="noopener" title="Mở ảnh kích thước thật"><img src="${url}" alt="${escapeHtml(s.name)}"></a>
        <div class="meta"><b class="mono">${escapeHtml(s.code)}.png</b>${escapeHtml(s.name)}<br>
        <span class="${sizeOk ? 'ok' : 'bad'}">${dims.w} × ${dims.h} px ${sizeOk ? '✓ đúng kích thước gốc' : '✕ sai kích thước'}</span> · ${formatBytes(blob.size)}<br>
        <a href="${url}" download="${escapeHtml(s.code)}.png">Tải ảnh này</a></div>`;
      grid.appendChild(card);
      if (!sizeOk) throw new Error('Ảnh thử không đúng kích thước ảnh gốc.');
    }
    state.testDone = true;
    $('approveWrap').classList.remove('hidden');
    updateProgress(1, `Đã tạo ${picks.length} ảnh thử`, true);
    showGenMessage('ok', `Đã tạo ${picks.length} ảnh thử từ ảnh khung gốc. Bấm vào từng ảnh để xem kích thước thật, sau đó đánh dấu <b>“Tôi đã kiểm tra ảnh thử”</b> để mở khóa nút tạo toàn bộ.`);
  } catch (err) {
    console.error(err);
    showGenMessage('err', `Lỗi khi tạo ảnh thử: ${escapeHtml(err.message)}`);
    state.testDone = false;
  } finally {
    state.testing = false;
    updateButtons();
    updateSteps();
  }
}

/* ---------------- Tạo toàn bộ ---------------- */
async function generateAll() {
  if (state.job || state.testing) return;
  if (!preflight(true)) return;
  const list = validStudents();
  const skipped = state.students.length - list.length;
  releaseOutput();
  const job = { cancelled: false };
  state.job = job;
  updateButtons();
  showGenMessage('', '');
  const t0 = performance.now();
  let zip = new JSZip();
  const used = new Set();
  let done = 0;
  const cancelled = () => {
    zip = null;
    updateProgress(0, 'Đã hủy');
    updateStats(0);
    showGenMessage('warn', `Đã hủy quá trình tạo. ${done} ảnh đã tạo trước khi hủy đã được loại bỏ — không có file ZIP nào được tạo.`);
  };
  try {
    for (let i = 0; i < list.length; i++) {
      if (job.cancelled) return cancelled();
      const s = list[i];
      const fileName = `${s.code}.png`;
      const key = fileName.normalize('NFC').toLowerCase();
      if (used.has(key)) throw new Error(`Phát hiện trùng tên file ${fileName} — dừng để tránh ghi đè.`);
      const blob = await renderStudentPNG(s);
      zip.file(fileName, blob, { binary: true });
      used.add(key);
      done++;
      updateProgress((done / list.length) * 0.92, `Đang tạo ảnh ${done}/${list.length}: ${fileName}`);
      updateStats(done);
      await yieldToUI();
    }
    if (job.cancelled) return cancelled();

    const count = Object.values(zip.files).filter((f) => !f.dir).length;
    if (count !== list.length || used.size !== list.length) {
      throw new Error(`Số ảnh trong ZIP (${count}) không khớp số học sinh hợp lệ (${list.length}).`);
    }
    updateProgress(0.92, 'Đang đóng gói ZIP…');
    // STORE: PNG đã được nén không mất dữ liệu, không nén lại để giữ nguyên từng byte ảnh
    const zipBlob = await zip.generateAsync({ type: 'blob', compression: 'STORE', mimeType: 'application/zip' }, (meta) => {
      updateProgress(0.92 + (meta.percent / 100) * 0.08, `Đang đóng gói ZIP… ${Math.round(meta.percent)}%`);
    });
    zip = null;
    if (job.cancelled) return cancelled();

    state.output = { zipBlob, count };
    const secs = ((performance.now() - t0) / 1000).toFixed(1);
    updateProgress(1, `Hoàn tất: ${count} ảnh PNG · ZIP ${formatBytes(zipBlob.size)} · ${secs} giây`, true);
    updateStats(count);
    showGenMessage('ok', `✅ <b>Hoàn tất!</b> Đã tạo ${count}/${list.length} giấy khen cho học sinh hợp lệ`
      + `${skipped ? `, bỏ qua ${skipped} dòng lỗi` : ''}. Bấm <b>“Tải ZIP”</b> để tải <span class="mono">${ZIP_NAME}</span>.`);
    toast(`Hoàn tất! Đã tạo ${count} giấy khen.`, 'ok');
  } catch (err) {
    console.error(err);
    zip = null;
    releaseOutput();
    updateStats(0);
    updateProgress(0, 'Đã dừng do lỗi');
    showGenMessage('err', `Lỗi khi tạo giấy khen: ${escapeHtml(err.message)}`);
  } finally {
    state.job = null;
    $('btnCancel').textContent = '✕ Hủy';
    updateButtons();
    updateSteps();
  }
}

async function downloadOne(idx) {
  if (state.job || state.testing) return;
  if (idx == null) idx = state.previewIdx;
  if (idx < 0 || !state.students[idx]) {
    toast('Hãy chọn một học sinh ở ô “Tên xem trước” (hoặc bấm “PNG” trong bảng danh sách).', 'warn', 6000);
    return;
  }
  const s = state.students[idx];
  if (!state.fontReady) return toast('Font chưa tải xong — chưa thể xuất ảnh.', 'err');
  if (!state.frame) return toast('Chưa tải ảnh khung giấy khen gốc.', 'err');
  if (s.errors.length) return toast(`Dòng ${s.row} bị lỗi: ${s.errors.join('; ')}`, 'err');
  const fit = checkFit(s.name);
  if (fit.outside) return toast('Tên này bị tràn ra ngoài ảnh — hãy điều chỉnh trước khi tải.', 'err');
  if (fit.tooLong) toast('Lưu ý: tên này vượt độ rộng vùng tên.', 'warn');
  try {
    const blob = await renderStudentPNG(s);
    downloadBlob(blob, `${s.code}.png`);
  } catch (err) {
    toast(`Không tạo được ảnh: ${err.message}`, 'err');
  }
}

/* ---------------- Khởi động ---------------- */
function init() {
  state.cfg = defaultCfg();
  syncCfgInputs();
  bindCfgInputs();
  bindPreviewInteractions();
  bindTable();

  bindDropzone('dzExcel', 'fileExcel', handleExcel);
  bindDropzone('dzFrame', 'fileFrame', (f) => handleImage('frame', f));
  bindDropzone('dzRef', 'fileRef', (f) => handleImage('ref', f));
  // Chặn trình duyệt mở file khi thả ra ngoài vùng nhận
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());

  $('sheetSelect').addEventListener('change', () => { clearStudents(); loadSheet(true); });
  $('hasHeader').addEventListener('change', () => { clearStudents(); buildColumnOptions(); });
  $('codeCol').addEventListener('change', () => { clearStudents(); renderSampleRows(); });
  $('nameCol').addEventListener('change', () => { clearStudents(); renderSampleRows(); });
  $('btnReadList').addEventListener('click', readStudentList);
  $('btnReadList2').addEventListener('click', readStudentList);

  $('btnPreview').addEventListener('click', () => {
    if (state.previewIdx < 0) {
      const first = state.students.findIndex((s) => !s.errors.length);
      if (first >= 0) setPreviewStudent(first);
    }
    $('viewMode').value = 'frame';
    $('viewMode').dispatchEvent(new Event('change'));
    $('secC').scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
  $('btnDownloadOne').addEventListener('click', () => downloadOne());
  $('btnTest').addEventListener('click', runTest);
  $('testApprove').addEventListener('change', () => { updateButtons(); updateSteps(); });
  $('btnGenerate').addEventListener('click', generateAll);
  $('btnCancel').addEventListener('click', () => {
    if (!state.job) return;
    state.job.cancelled = true;
    $('btnCancel').textContent = 'Đang hủy…';
    $('btnCancel').disabled = true;
  });
  $('btnZip').addEventListener('click', () => {
    if (!state.output || !state.output.count) return toast('Chưa có ảnh nào được tạo.', 'err');
    downloadBlob(state.output.zipBlob, ZIP_NAME);
  });

  if (!window.XLSX || !window.JSZip) {
    toast('Thiếu thư viện trong thư mục libs/ (xlsx.full.min.js, jszip.min.js). Xem README.', 'err', 15000);
  }
  updateButtons();
  updateSteps();
  loadFont();
}

// Dùng cho kiểm thử tự động / gỡ lỗi trong Console
window.__giayKhen = { state, renderCertificate, detectNamePosition, layoutName };

init();
