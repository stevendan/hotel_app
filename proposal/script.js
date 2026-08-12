"use strict";

// ---- default hotel / sales rep values (from Proposal.docx) ----
const DEFAULTS = {
  pkg_venue: "Ballroom 273m2",
  pkg_setup: "Round table",
  pkg_time: "08:00-17:00",
  pkg_event_type: "Full-day meeting package",
  sales_rep_name: "Vu Thi Hien (Mrs.)",
  sales_rep_title: "Event Sales Executive",
  sales_rep_phone: "0845542994",
  sales_rep_email: "hien.vuthi@movenpick.com"
};

const SAMPLE = `Chào em, bên công ty Paula's Choice muốn đặt phòng hội thảo cả ngày cho khoảng 50 khách,
tổ chức trong 2 ngày 02-03/03/2026. Chị Yến là người liên hệ.
Bên em báo giá gói full-day khoảng 950.000/khách giúp chị nhé.
Cần thuê thêm 1 màn hình LED 3,5m x 6m, giá thuê tầm 12.000.000/ngày.
Báo giá giữ hiệu lực đến 30-01-2026 nhé em.`;

const $ = id => document.getElementById(id);
const PLACE_KEYS = [
  "company_name","contact_name","contact_salutation","event_date",
  "pkg_date","pkg_venue","pkg_setup","pkg_time","pkg_event_type",
  "pkg_qty","pkg_days","pkg_unit_price","pkg_total",
  "addon_desc","addon_qty","addon_days","addon_unit_price","addon_total",
  "subtotal","service_charge","vat","grand_total",
  "sales_rep_name","sales_rep_title","sales_rep_phone","sales_rep_email","valid_until"
];

// ---- number helpers ----
function fmtVND(n) {
  if (!isFinite(n) || n === 0) return "";
  return Math.round(n).toLocaleString("en-US");
}
function toNumber(v) {
  const n = parseFloat(String(v).replace(/[^\d.-]/g, ""));
  return isFinite(n) ? n : 0;
}
// Parse a Vietnamese-style money string: "950.000", "1,2 triệu", "12tr", "500k"
function parseMoney(raw) {
  if (!raw) return 0;
  let s = raw.toLowerCase().trim();
  let mult = 1;
  const trieu = s.match(/([\d.,]+)\s*(triệu|tr\b|tr\.)/);
  const nghin = s.match(/([\d.,]+)\s*(nghìn|nghin|ngàn|ngan|k\b)/);
  if (trieu) { mult = 1e6; s = trieu[1]; }
  else if (nghin) { mult = 1e3; s = nghin[1]; }
  if (mult !== 1) {
    const num = parseFloat(s.replace(/\./g, "").replace(",", "."));
    return isFinite(num) ? Math.round(num * mult) : 0;
  }
  // plain grouped number like 950.000 or 12,000,000
  const digits = s.replace(/[.,\s](?=\d{3}\b)/g, "").replace(/[^\d]/g, "");
  return digits ? parseInt(digits, 10) : 0;
}

// ---- recompute totals ----
function recompute() {
  const qty = toNumber($("pkg_qty").value);
  const days = toNumber($("pkg_days").value);
  const unit = toNumber($("pkg_unit_price").value);
  const pkgTotal = unit * qty * days;

  const aQty = toNumber($("addon_qty").value);
  const aDays = toNumber($("addon_days").value);
  const aUnit = toNumber($("addon_unit_price").value);
  const addonTotal = aUnit * aQty * aDays;

  const subtotal = pkgTotal + addonTotal;
  const service = subtotal * 0.05;
  const vat = (subtotal + service) * 0.08;
  const grand = subtotal + service + vat;

  $("pkg_total").value = fmtVND(pkgTotal);
  $("addon_total").value = fmtVND(addonTotal);
  $("subtotal").value = fmtVND(subtotal);
  $("service_charge").value = fmtVND(service);
  $("vat").value = fmtVND(vat);
  $("grand_total").value = fmtVND(grand);
}
["pkg_qty","pkg_days","pkg_unit_price","addon_qty","addon_days","addon_unit_price"]
  .forEach(id => $(id).addEventListener("input", recompute));

// ---- apply defaults ----
function applyDefaults() {
  Object.keys(DEFAULTS).forEach(k => { if (!$(k).value) $(k).value = DEFAULTS[k]; });
}

// ---- heuristic analysis of the chat ----
function analyze(text) {
  const out = {};
  const t = text.replace(/\r/g, "");

  // contact name + salutation: chị/anh/cô/em/Ms./Mr./Mrs. + Name
  const nameM = t.match(/\b(ch[ịi]|anh|cô|c[ôo]|em|Ms\.?|Mr\.?|Mrs\.?|Miss)\s+([A-ZÀ-Ỹ][\p{L}]+(?:\s+[A-ZÀ-Ỹ][\p{L}]+)?)/u);
  if (nameM) {
    const title = nameM[1].toLowerCase();
    const enTitle = { "chị": "Ms.", "chi": "Ms.", "cô": "Ms.", "co": "Ms.",
                      "anh": "Mr.", "em": "Ms." }[title] || nameM[1];
    out.contact_name = (enTitle.match(/^(Ms|Mr|Mrs|Miss)/i) ? enTitle : "Ms.") + " " + nameM[2];
    const viTitle = /anh/i.test(title) ? "anh" : "chị";
    out.contact_salutation = viTitle + " " + nameM[2];
  }

  // company: after "công ty" / "company"
  let compM = t.match(/(?:c[ôo]ng ty|company)\s*[:\-]?\s*([A-Za-zÀ-Ỹ0-9'’&.\- ]{2,45}?)(?=\s+(?:mu[ôố]n|c[ầa]n|đặt|book|-|,|\.|\n)|$)/i);
  if (compM) out.company_name = compM[1].trim().replace(/\s+/g, " ");

  // number of guests / pax
  const paxM = t.match(/(?:kho[ảa]ng\s*)?(\d{1,4})\s*(?:kh[áa]ch|ngư[ờo]i|pax|guests?|đại bi[ểe]u|delegates?)/i);
  if (paxM) out.pkg_qty = paxM[1];

  // number of days
  const dayM = t.match(/(\d{1,2})\s*(?:ng[àa]y|days?)/i);
  if (dayM) out.pkg_days = dayM[1];

  // event date (range or single): 02-03/03/2026, 02/03/2026, 2-3/3
  const dateM = t.match(/(\d{1,2}(?:\s*[-–]\s*\d{1,2})?[\/\-]\d{1,2}(?:[\/\-]\d{2,4})?)/);
  if (dateM) { out.event_date = dateM[1].replace(/\s+/g, ""); out.pkg_date = out.event_date; }

  // valid until: "hiệu lực đến <date>"
  const validM = t.match(/hi[ệe]u l[ựu]c\s*(?:đ[ếe]n|tới)?\s*(?:ng[àa]y\s*)?(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})/i);
  if (validM) out.valid_until = validM[1];

  // unit price per guest: near "khách" / "pax" / "gói"
  const priceM = t.match(/([\d.,]+\s*(?:tri[ệe]u|tr\b|ngh[ìi]n|ng[àa]n|k\b)?)\s*(?:vnd|vnđ|đ|₫)?\s*\/\s*(?:kh[áa]ch|pax|ngư[ờo]i)/i)
    || t.match(/(?:g[óo]i|đ[ơo]n gi[áa]|gi[áa])\D{0,12}?([\d.,]+\s*(?:tri[ệe]u|tr\b|ngh[ìi]n|ng[àa]n|k\b)?)/i);
  if (priceM) { const p = parseMoney(priceM[1]); if (p) out.pkg_unit_price = p; }

  // add-on (LED / màn hình / equipment) with a price
  const ledM = t.match(/(m[àa]n h[ìi]nh led[^\n,.]*|led[^\n,.]*)/i);
  if (ledM) out.addon_desc = ledM[1].trim().replace(/\s+/g, " ");
  const addonPriceM = t.match(/(?:led|m[àa]n h[ìi]nh|thu[êe])[^\n]*?([\d.,]+\s*(?:tri[ệe]u|tr\b|ngh[ìi]n|ng[àa]n|k\b))\s*(?:vnd|vnđ|đ|₫)?\s*(?:\/\s*ng[àa]y)?/i);
  if (addonPriceM) {
    const ap = parseMoney(addonPriceM[1]);
    if (ap) {
      out.addon_unit_price = ap;
      out.addon_qty = out.addon_qty || "1";
      out.addon_days = out.addon_days || (out.pkg_days || "1");
    }
  }

  return out;
}

function setStatus(msg, type) {
  const el = $("status");
  el.textContent = msg;
  el.className = "status " + type;
}

// ---- buttons ----
$("sampleBtn").addEventListener("click", () => { $("chatInput").value = SAMPLE; });

$("analyzeBtn").addEventListener("click", () => {
  const text = $("chatInput").value.trim();
  if (!text) { setStatus("Vui lòng dán nội dung tin nhắn trước khi phân tích.", "err"); return; }
  const data = analyze(text);
  Object.keys(data).forEach(k => { if ($(k)) $(k).value = data[k]; });
  applyDefaults();
  recompute();
  const found = Object.keys(data).length;
  setStatus("Đã phân tích và điền " + found + " chỉ số. Hãy kiểm tra & chỉnh lại nếu cần, sau đó bấm \"Tạo file báo giá\".", "info");
});

// ---- docx generation ----
function escapeXml(v) {
  return String(v)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

async function loadTemplateBuffer() {
  try {
    const res = await fetch("proposal-template.docx");
    if (res.ok) return await res.arrayBuffer();
  } catch (e) { /* file:// or missing – fall back */ }
  return null;
}

function pickTemplateFile() {
  return new Promise((resolve, reject) => {
    const inp = $("templateFile");
    inp.value = "";
    inp.onchange = () => {
      const f = inp.files[0];
      if (!f) { reject(new Error("Chưa chọn file mẫu.")); return; }
      const r = new FileReader();
      r.onload = () => resolve(r.result);
      r.onerror = () => reject(new Error("Không đọc được file mẫu."));
      r.readAsArrayBuffer(f);
    };
    inp.click();
  });
}

async function generate() {
  setStatus("Đang tạo file báo giá...", "info");
  recompute();

  let buf = await loadTemplateBuffer();
  if (!buf) {
    setStatus("Không tự tải được proposal-template.docx — vui lòng chọn file mẫu.", "info");
    try { buf = await pickTemplateFile(); }
    catch (e) { setStatus(e.message, "err"); return; }
  }

  try {
    const zip = await JSZip.loadAsync(buf);
    const docPath = "word/document.xml";
    let xml = await zip.file(docPath).async("string");

    PLACE_KEYS.forEach(key => {
      const val = escapeXml(($(key) && $(key).value) ? $(key).value : "");
      xml = xml.split("{{" + key + "}}").join(val);
    });

    zip.file(docPath, xml);
    const blob = await zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });

    const company = ($("company_name").value || "KhachHang").replace(/[\\/:*?"<>|]/g, "").trim();
    const dateTag = ($("event_date").value || "").replace(/[\\/:*?"<>|]/g, "-").trim();
    const fname = "Bao gia - " + company + (dateTag ? " - " + dateTag : "") + ".docx";

    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = fname;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);

    setStatus("Đã tạo xong: " + fname, "ok");
  } catch (e) {
    console.error(e);
    setStatus("Lỗi khi tạo file: " + e.message, "err");
  }
}

$("generateBtn").addEventListener("click", generate);

// init defaults on load
applyDefaults();
recompute();
