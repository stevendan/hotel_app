"use strict";

const DEFAULT_TEMPLATE = "contract-template.docx";
const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const MONTHS_EN = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];

const VARIABLES = [
  { name: "excecute_date",                    meaning: "Ngày thực hiện hợp đồng" },
  { name: "contract_number",                  meaning: "Số hợp đồng" },
  { name: "contract_created_date",            meaning: "ngày tạo hợp đồng" },
  { name: "contract_created_month",           meaning: "tháng tạo hợp đồng" },
  { name: "contract_created_month_en",        meaning: "tháng tạo hợp đồng Tiếng Anh" },
  { name: "contract_created_year",            meaning: "Năm tạo hợp đồng" },
  { name: "client_company_name_en",           meaning: "Tên công ty khách hàng tiếng Anh" },
  { name: "client_company_name_vi",           meaning: "Tên công ty khách hàng tiếng Việt" },
  { name: "client_representative_name_en",    meaning: "Tên đầy đủ và danh xưng Tiếng Anh" },
  { name: "client_representative_name_vi",    meaning: "Tên đầy đủ và danh xưng Tiếng Việt" },
  { name: "client_representative_name_en_vi", meaning: "Tên đầy đủ và danh xưng Tiếng Anh-Việt" },
  { name: "client_representative_title_en",   meaning: "Chức danh tiếng Anh" },
  { name: "client_representative_title_vi",   meaning: "Chức danh tiếng Việt" },
  { name: "client_address_en",                meaning: "Địa chỉ tiếng Anh" },
  { name: "client_address_vi",                meaning: "Địa chỉ tiếng Việt" },
  { name: "client_phone",                     meaning: "Số điện thoại công ty khách hàng" },
  { name: "client_mobile_phone",              meaning: "Số di động của khách hàng" },
  { name: "client_email",                     meaning: "Email của khách hàng" },
  { name: "client_VAT_code",                  meaning: "Mã số thuế của khách hàng" },
  { name: "deposit_percent",                  meaning: "Phần trăm đặt cọc" },
  { name: "request_to_settle_before_date",    meaning: "Yêu cầu thanh toán trước ngày" },
  { name: "contract_cancellation_date",       meaning: "Ngày hủy hợp đồng" },
  { name: "returning_agreement_date",         meaning: "Ngày hạn gửi lại hợp đồng cho khách sạn" },
  { name: "returning_agreement_month",        meaning: "Tháng hạn gửi lại hợp đồng cho khách sạn" },
  { name: "returning_agreement_month_en",     meaning: "Tháng tiếng Anh hạn gửi lại hợp đồng cho khách sạn" },
  { name: "returning_agreement_year",         meaning: "Năm hạn gửi lại hợp đồng cho khách sạn" },
  { name: "sale_rep_fullname",                meaning: "Danh xưng và tên sale" },
  { name: "sale_rep_title_en",                meaning: "Chức vụ sale Tiếng Anh" },
  { name: "sale_rep_title_vi",                meaning: "Chức vụ sale Tiếng Việt" },
  { name: "event_setup_type",                 meaning: "Kiểu bố trí sự kiện" },
  { name: "event_start_time",                 meaning: "Giờ bắt đầu sự kiện" },
  { name: "event_end_time",                   meaning: "Giờ kết thúc sự kiện" },
  { name: "meal_setup_type",                  meaning: "Kiểu bố trí bữa ăn" },
  { name: "meal_venue",                       meaning: "Địa điểm dùng bữa" },
  { name: "meal_start_time",                  meaning: "Giờ bắt đầu bữa ăn" },
  { name: "meal_end_time",                    meaning: "Giờ kết thúc bữa ăn" },
  { name: "morning_tea_break_start_time",     meaning: "Giờ bắt đầu tiệc trà buổi sáng" },
  { name: "morning_tea_break_end_time",       meaning: "Giờ kết thúc tiệc trà buổi sáng" },
  { name: "afternoon_tea_break_start_time",   meaning: "Giờ bắt đầu tiệc trà buổi chiều" },
  { name: "afternoon_tea_break_end_time",     meaning: "Giờ kết thúc tiệc trà buổi chiều" },
  { name: "tea_break_venue",                  meaning: "Địa điểm tiệc trà (tea break)" },
  { name: "number_of_persons",                meaning: "Số lượng khách" },
  { name: "unit_price",                       meaning: "Đơn giá" },
  { name: "LED_unit_price",                   meaning: "Đơn giá màn hình LED" }
];

const fields = new Map();
let variablesByName = new Map();
let selectedTemplateBuffer = null;
let criteriaTemplateFile = null;
let extractAbortController = null;

// Thư mục và danh sách file mẫu có sẵn. Tên file theo cấu trúc:
// {venue}_{day-type}_{package-type}_{meal-type}.docx
const TEMPLATE_DIR = "template";
const TEMPLATE_FILES = [
  "Ballroom_fullday_package_lunch.docx",
  "Ballroom_fullday_package_nolunch.docx",
  "Ballroom_fullday_room-rental_nomeal.docx",
  "Ballroom_halfday-afternoon_package_dinner.docx",
  "Ballroom_halfday-afternoon_package_lunch.docx",
  "Ballroom_halfday-afternoon_package_nomeal.docx",
  "Ballroom_halfday-morning_package_lunch.docx",
  "Ballroom_halfday-morning_package_nomeal.docx",
  "Ballroom_halfday_room-rental_nomeal.docx",
  "Ballroom_YEP_package1_dinner.docx",
  "Ballroom_YEP_package1_lunch.docx",
  "Ballroom_YEP_package2_dinner(editting).docx",
  "Ballroom_YEP_package2_lunch.docx",
  "Ballroom_YEP_package3_dinner(editting).docx",
  "Ballroom_YEP_package3_lunch(editting).docx",
  "Lotus-Apricot_fullday_package_lunch.docx",
  "Lotus-Apricot_fullday_package_nomeal.docx",
  "Lotus-Apricot_fullday_room-rental_nomeal.docx",
  "Lotus-Apricot_halfday-afternoon_package_dinner.docx",
  "Lotus-Apricot_halfday-afternoon_package_lunch.docx",
  "Lotus-Apricot_halfday-afternoon_package_nomeal.docx",
  "Lotus-Apricot_halfday-morning_package_lunch.docx",
  "Lotus-Apricot_halfday-morning_package_nomeal.docx",
  "Lotus-Apricot_halfday_room-rental_nomeal.docx",
  "Lotus-Apricot_YEP_package1_dinner.docx",
  "Lotus-Apricot_YEP_package1_lunch.docx",
  "Orchid_fullday_package_lunch.docx",
  "Orchid_fullday_package_nomeal.docx",
  "Orchid_fullday_room-rental_nomeal.docx",
  "Orchid_halfday-afternoon_package_dinner.docx",
  "Orchid_halfday-afternoon_package_lunch.docx",
  "Orchid_halfday-afternoon_package_nomeal.docx",
  "Orchid_halfday-morning_package_lunch.docx",
  "Orchid_halfday-morning_package_nomeal.docx",
  "Orchid_halfday_room-rental_nomeal.docx",
  "Orchid_YEP_package1_dinner.docx",
  "Orchid_YEP_package1_lunch.docx"
];

// Các tiêu chí chọn mẫu và nhãn hiển thị.
const TEMPLATE_CRITERIA = {
  venue: {
    selectId: "criteriaVenue",
    options: [
      { value: "Ballroom", label: "Ballroom" },
      { value: "Lotus",    label: "Lotus" },
      { value: "Apricot",  label: "Apricot" },
      { value: "Orchid",   label: "Orchid" }
    ]
  },
  day: {
    selectId: "criteriaDay",
    options: [
      { value: "fullday",           label: "Cả ngày (fullday)" },
      { value: "halfday",           label: "Nửa ngày (halfday)" },
      { value: "halfday-morning",   label: "Nửa ngày sáng" },
      { value: "halfday-afternoon", label: "Nửa ngày chiều" },
      { value: "YEP",               label: "YEP (tiệc tất niên)" }
    ]
  },
  pkg: {
    selectId: "criteriaPackage",
    options: [
      { value: "package",     label: "Trọn gói (package)" },
      { value: "room-rental", label: "Thuê phòng (room-rental)" },
      { value: "package1",    label: "Gói 1 (package1)" },
      { value: "package2",    label: "Gói 2 (package2)" },
      { value: "package3",    label: "Gói 3 (package3)" }
    ]
  },
  meal: {
    selectId: "criteriaMeal",
    options: [
      { value: "lunch",  label: "Ăn trưa (lunch)" },
      { value: "dinner", label: "Ăn tối (dinner)" },
      { value: "nomeal", label: "Không kèm bữa ăn (nomeal)" }
    ]
  }
};

// Lotus và Apricot dùng chung mẫu có phần đầu là "Lotus-Apricot".
function mapVenueToPrefix(venue) {
  return venue === "Lotus" || venue === "Apricot" ? "Lotus-Apricot" : venue;
}

// Tách tên file mẫu thành các thành phần tiêu chí.
function parseTemplateName(file) {
  const editing = /\(editting\)/i.test(file);
  const base = file.replace(/\.docx$/i, "").replace(/\s*\(editting\)\s*$/i, "");
  const [venue, day, pkg, meal] = base.split("_");
  return { venue, day, pkg, meal, editing, file };
}

const TEMPLATE_RECORDS = TEMPLATE_FILES.map(parseTemplateName);

// "nomeal" khớp cả token "nomeal" lẫn "nolunch" trong tên file.
function mealMatches(recordMeal, wanted) {
  if (recordMeal === wanted) return true;
  return wanted === "nomeal" && (recordMeal === "nomeal" || recordMeal === "nolunch");
}

// Tìm file mẫu phù hợp nhất với bộ tiêu chí đã chọn.
function resolveTemplateFile(criteria) {
  const venue = mapVenueToPrefix(criteria.venue);
  const matches = TEMPLATE_RECORDS.filter(r =>
    r.venue === venue &&
    r.day === criteria.day &&
    r.pkg === criteria.pkg &&
    mealMatches(r.meal, criteria.meal));
  if (!matches.length) return null;
  return matches.find(r => !r.editing) || matches[0];
}

// Các trường luôn hiển thị và lưu ở dạng chữ IN HOA.
const UPPERCASE_FIELDS = new Set(["client_company_name_en", "client_company_name_vi"]);

// Trường thời gian (xuất ra dạng h:mm, ví dụ 9:00) và trường giá tiền.
const TIME_FIELDS = new Set([
  "event_start_time", "event_end_time",
  "meal_start_time", "meal_end_time",
  "morning_tea_break_start_time", "morning_tea_break_end_time",
  "afternoon_tea_break_start_time", "afternoon_tea_break_end_time"
]);
const PRICE_FIELDS = new Set(["unit_price", "LED_unit_price"]);

// "09:00" -> "9:00" (bỏ số 0 ở đầu giờ).
function formatTimeValue(value) {
  const match = value.match(/^(\d{1,2}):(\d{2})/);
  return match ? `${Number(match[1])}:${match[2]}` : value;
}

// Định dạng số tiền có dấu phẩy phân tách hàng nghìn.
function formatPriceInput(input) {
  const digits = input.value.replace(/\D/g, "");
  input.value = digits ? Number(digits).toLocaleString("en-US") : "";
}

// Chuẩn hóa ô giờ về định dạng 24 giờ HH:MM (chấp nhận "9:0", "930", "9.30"...).
function normalizeTimeInput(input) {
  let value = input.value.trim();
  if (!value) { clearFieldInvalid(input); return; }
  let digits = value.replace(/\D/g, "");
  if (digits.length === 3) digits = "0" + digits;
  if (digits.length === 4) value = `${digits.slice(0, 2)}:${digits.slice(2)}`;
  const match = value.match(/^(\d{1,2}):(\d{1,2})$/);
  if (!match) {
    input.classList.add("is-invalid");
    input.setAttribute("aria-invalid", "true");
    return;
  }
  const hour = Math.min(23, Number(match[1]));
  const minute = Math.min(59, Number(match[2]));
  input.value = `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
  clearFieldInvalid(input);
}

const GEMINI_MODEL = "gemini-flash-latest";
const GEMINI_API_URL = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;
const GEMINI_KEY_STORAGE = "hta_gemini_api_key";

// Các trường AI có thể trích xuất từ tin nhắn (bỏ qua trường tự tính và nhân viên sale).
// type "date" nhận giá trị ISO YYYY-MM-DD từ AI rồi hiển thị dd/mm/yyyy.
const EXTRACTABLE_FIELDS = [
  { name: "contract_number",                  type: "text" },
  { name: "excecute_date",                    type: "date" },
  { name: "client_company_name_en",           type: "text" },
  { name: "client_company_name_vi",           type: "text" },
  { name: "client_representative_name_en",    type: "text" },
  { name: "client_representative_name_vi",    type: "text" },
  { name: "client_representative_name_en_vi", type: "text" },
  { name: "client_representative_title_en",   type: "text" },
  { name: "client_representative_title_vi",   type: "text" },
  { name: "client_address_en",                type: "text" },
  { name: "client_address_vi",                type: "text" },
  { name: "client_phone",                     type: "text" },
  { name: "client_mobile_phone",              type: "text" },
  { name: "client_email",                     type: "text" },
  { name: "client_VAT_code",                  type: "text" },
  { name: "deposit_percent",                  type: "text" },
  { name: "request_to_settle_before_date",    type: "date" },
  { name: "contract_cancellation_date",       type: "date" }
];

const $ = id => document.getElementById(id);

function setStatus(message = "", type = "info") {
  const status = $("status");
  status.textContent = message;
  status.className = message ? `status is-visible ${type}` : "status";
}

function toLocalIsoDate(date) {
  const offsetDate = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return offsetDate.toISOString().slice(0, 10);
}

function applyDateParts(sourceId, prefix) {
  const value = $(sourceId).value;
  const parts = value.split("-").map(Number);
  if (parts.length !== 3 || parts.some(Number.isNaN)) return;

  const [year, month, day] = parts;
  setFieldValue(`${prefix}_date`, String(day));
  setFieldValue(`${prefix}_month`, String(month));
  setFieldValue(`${prefix}_month_en`, MONTHS_EN[month - 1] || "");
  setFieldValue(`${prefix}_year`, String(year));
}

function setFieldValue(name, value) {
  const input = fields.get(name);
  if (!input) return;
  input.value = UPPERCASE_FIELDS.has(name) ? value.toUpperCase() : value;
  if (value.trim()) clearFieldInvalid(input);
}

function formatDateDMY(isoDate) {
  const [y, m, d] = isoDate.split('-');
  return d && m && y ? `${d}/${m}/${y}` : '';
}

function prepareDateHelpers() {
  const today = new Date();
  const returnDate = new Date(today);
  returnDate.setDate(returnDate.getDate() + 7);

  const todayIso = toLocalIsoDate(today);
  const returnIso = toLocalIsoDate(returnDate);

  $('contractCreatedFullDate').value = todayIso;
  $('returnAgreementFullDate').value = returnIso;

  const dispCreated = document.getElementById('contractCreatedFullDate_display');
  const dispReturn = document.getElementById('returnAgreementFullDate_display');
  if (dispCreated) dispCreated.value = formatDateDMY(todayIso);
  if (dispReturn) dispReturn.value = formatDateDMY(returnIso);

  $("contractCreatedFullDate").addEventListener("change", () => applyDateParts("contractCreatedFullDate", "contract_created"));
  $("returnAgreementFullDate").addEventListener("change", () => applyDateParts("returnAgreementFullDate", "returning_agreement"));

  applyDateParts("contractCreatedFullDate", "contract_created");
  applyDateParts("returnAgreementFullDate", "returning_agreement");
}

function initFields() {
  fields.clear();
  for (const input of document.querySelectorAll("#formFields input[name], #formFields select[name]")) {
    fields.set(input.name, input);
  }
  variablesByName = new Map(VARIABLES.map(v => [v.name, v]));
}

function getValues() {
  const currency = $("priceCurrency")?.value || "VND";
  const result = {};
  for (const [name, input] of fields) {
    let value = (input.value || "").trim();
    if (TIME_FIELDS.has(name)) value = formatTimeValue(value);
    else if (PRICE_FIELDS.has(name) && value) value = `${value} ${currency}`;
    result[name] = value;
  }
  return result;
}

function parseWordXml(xml) {
  const documentXml = new DOMParser().parseFromString(xml, "application/xml");
  if (documentXml.getElementsByTagName("parsererror").length) {
    throw new Error("Không đọc được nội dung XML trong file mẫu.");
  }
  return documentXml;
}

function paragraphTextNodes(documentXml) {
  return Array.from(documentXml.getElementsByTagNameNS(WORD_NS, "p"))
    .map(paragraph => Array.from(paragraph.getElementsByTagNameNS(WORD_NS, "t")));
}

function collectPlaceholders(xml) {
  const placeholders = new Set();
  paragraphTextNodes(parseWordXml(xml)).forEach(nodes => {
    const text = nodes.map(node => node.textContent || "").join("");
    for (const match of text.matchAll(/\{\{([A-Za-z][A-Za-z0-9_]*)\}\}/g)) placeholders.add(match[1]);
  });
  return placeholders;
}

function findTextLocation(nodes, offset) {
  let cursor = offset;
  for (let index = 0; index < nodes.length; index += 1) {
    const length = (nodes[index].textContent || "").length;
    if (cursor < length) return { index, offset: cursor };
    cursor -= length;
  }
  return null;
}

function setText(node, value) {
  node.textContent = value;
  if (/^\s|\s$/.test(value)) node.setAttributeNS(XML_NS, "xml:space", "preserve");
  else node.removeAttributeNS(XML_NS, "space");
}

// Word có thể tách {{variable}} thành nhiều thẻ <w:t>. Hàm này thay thế theo
// từng paragraph, giữ lại định dạng của vị trí bắt đầu biến và vẫn hợp lệ XML.
function replacePlaceholdersInXml(xml, values) {
  const documentXml = parseWordXml(xml);
  paragraphTextNodes(documentXml).forEach(nodes => {
    const fullText = nodes.map(node => node.textContent || "").join("");
    const matches = Array.from(fullText.matchAll(/\{\{([A-Za-z][A-Za-z0-9_]*)\}\}/g));

    for (let matchIndex = matches.length - 1; matchIndex >= 0; matchIndex -= 1) {
      const match = matches[matchIndex];
      const key = match[1];
      const start = findTextLocation(nodes, match.index);
      const end = findTextLocation(nodes, match.index + match[0].length - 1);
      if (!start || !end) continue;

      const firstText = nodes[start.index].textContent || "";
      const lastText = nodes[end.index].textContent || "";
      const before = firstText.slice(0, start.offset);
      const after = lastText.slice(end.offset + 1);
      const replacement = Object.prototype.hasOwnProperty.call(values, key) ? values[key] : match[0];

      if (start.index === end.index) {
        setText(nodes[start.index], before + replacement + after);
      } else {
        setText(nodes[start.index], before + replacement);
        for (let i = start.index + 1; i < end.index; i += 1) setText(nodes[i], "");
        setText(nodes[end.index], after);
      }
    }
  });
  return new XMLSerializer().serializeToString(documentXml);
}

function wordXmlParts(zip) {
  return zip.file(/^word\/(?:document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i);
}

function embeddedXlsxParts(zip) {
  return zip.file(/^word\/embeddings\/.*\.xlsx$/i);
}

// Escape XML special chars before injecting values into raw XML strings
function escapeXml(str) {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function collectPlaceholdersFromText(text) {
  const placeholders = new Set();
  for (const match of text.matchAll(/\{\{([A-Za-z][A-Za-z0-9_]*)\}\}/g)) placeholders.add(match[1]);
  return placeholders;
}

function replacePlaceholdersInText(text, values) {
  return text.replace(/\{\{([A-Za-z][A-Za-z0-9_]*)\}\}/g, (match, key) =>
    Object.prototype.hasOwnProperty.call(values, key) ? escapeXml(values[key]) : match
  );
}

// Excel XML has no split-run problem, so plain regex replacement is safe here
async function collectPlaceholdersFromEmbeddedExcel(zip) {
  const placeholders = new Set();
  for (const xlsxFile of embeddedXlsxParts(zip)) {
    const inner = await window.JSZip.loadAsync(await xlsxFile.async("arraybuffer"));
    for (const f of inner.file(/^xl\/(sharedStrings|worksheets\/.+)\.xml$/i)) {
      for (const p of collectPlaceholdersFromText(await f.async("string"))) placeholders.add(p);
    }
  }
  return placeholders;
}

async function processEmbeddedExcel(zip, values) {
  for (const xlsxFile of embeddedXlsxParts(zip)) {
    const inner = await window.JSZip.loadAsync(await xlsxFile.async("arraybuffer"));
    for (const f of inner.file(/^xl\/(sharedStrings|worksheets\/.+)\.xml$/i)) {
      inner.file(f.name, replacePlaceholdersInText(await f.async("string"), values));
    }
    zip.file(xlsxFile.name, await inner.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } }));
  }
}

async function getTemplateBuffer() {
  if (selectedTemplateBuffer) return selectedTemplateBuffer;
  if (criteriaTemplateFile) {
    const url = `${TEMPLATE_DIR}/${encodeURIComponent(criteriaTemplateFile)}`;
    const response = await fetch(url, { cache: "no-store" });
    if (!response.ok) throw new Error(`Không tìm thấy ${criteriaTemplateFile}.`);
    return response.arrayBuffer();
  }
  const response = await fetch(DEFAULT_TEMPLATE, { cache: "no-store" });
  if (!response.ok) throw new Error(`Không tìm thấy ${DEFAULT_TEMPLATE}.`);
  return response.arrayBuffer();
}

function readableVariableName(name) {
  return variablesByName.get(name)?.meaning || `{{${name}}}`;
}

function clearFieldInvalid(input) {
  input.classList.remove("is-invalid");
  input.removeAttribute("aria-invalid");
}

function clearInvalidFields() {
  for (const input of fields.values()) {
    if (input.classList.contains("is-invalid")) clearFieldInvalid(input);
  }
}

function markMissingFields(missing) {
  clearInvalidFields();
  let firstInput = null;
  for (const name of missing) {
    const input = fields.get(name);
    if (!input) continue;
    input.classList.add("is-invalid");
    input.setAttribute("aria-invalid", "true");
    const onFill = () => {
      if (input.value.trim()) clearFieldInvalid(input);
    };
    input.addEventListener("input", onFill);
    input.addEventListener("change", onFill);
    if (!firstInput) firstInput = input;
  }
  if (firstInput) {
    firstInput.scrollIntoView({ behavior: "smooth", block: "center" });
    firstInput.focus({ preventScroll: true });
  }
}

// Ngày dd/mm/yyyy -> d.m.yyyy (bỏ số 0 ở đầu) để đặt tên file.
function formatDateForFilename(dmy) {
  const parts = String(dmy).split("/");
  if (parts.length !== 3) return "";
  const [d, m, y] = parts.map(p => p.trim());
  if (!d || !m || !y || [d, m].some(v => Number.isNaN(Number(v)))) return "";
  return `${Number(d)}.${Number(m)}.${y}`;
}

// Rút gọn tên công ty thành tên viết tắt (bỏ các từ chỉ loại hình doanh nghiệp).
function shortClientName(values) {
  const raw = (values.client_company_name_en || values.client_company_name_vi || "").toUpperCase();
  if (!raw.trim()) return "";
  const patterns = [
    /\bCÔNG TY\b/g, /\bCONG TY\b/g,
    /\bTNHH\b/g, /\bTRÁCH NHIỆM HỮU HẠN\b/g,
    /\bCỔ PHẦN\b/g, /\bCO PHAN\b/g,
    /\bMỘT THÀNH VIÊN\b/g, /\bMTV\b/g,
    /\bDOANH NGHIỆP TƯ NHÂN\b/g, /\bDNTN\b/g, /\bTẬP ĐOÀN\b/g,
    /\bJOINT STOCK\b/g, /\bCOMPANY\b/g, /\bLIMITED\b/g,
    /\bCORPORATION\b/g, /\bCORP\b/g, /\bENTERPRISES?\b/g, /\bGROUP\b/g,
    /\bCO\.?,?\b/g, /\bLTD\.?\b/g, /\bJSC\b/g, /\bINC\.?\b/g,
    /\bLLC\b/g, /\bPLC\b/g, /\bPTE\.?\b/g
  ];
  let name = raw;
  for (const p of patterns) name = name.replace(p, " ");
  return name.replace(/[.,]/g, " ").replace(/\s+/g, " ").trim();
}

function download(blob, values) {
  const datePart = formatDateForFilename(values.excecute_date || "");
  const shortName = shortClientName(values);
  const rawName = [datePart, "MVPL", shortName, "Contract"].filter(Boolean).join("_");
  const safeName = rawName.replace(/[\\/:*?"<>|]/g, "").replace(/\s+/g, " ").trim() || "MVPL_Contract";
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${safeName}.docx`;
  document.body.append(link);
  link.click();
  window.setTimeout(() => {
    URL.revokeObjectURL(link.href);
    link.remove();
  }, 1500);
  return link.download;
}

// Kiểm tra phần trăm đặt cọc phải nằm trong khoảng 0-100.
function validateDepositPercent() {
  const input = fields.get("deposit_percent");
  if (!input) return true;
  const raw = input.value.trim();
  if (!raw) { clearFieldInvalid(input); return true; }
  const num = Number(raw);
  if (Number.isFinite(num) && num >= 0 && num <= 100) {
    clearFieldInvalid(input);
    return true;
  }
  input.classList.add("is-invalid");
  input.setAttribute("aria-invalid", "true");
  setStatus("Phần trăm đặt cọc phải nằm trong khoảng từ 0 đến 100.", "error");
  return false;
}

async function generateContract() {
  if (!window.JSZip) {
    setStatus("Chưa tải được thư viện tạo file Word. Hãy kiểm tra kết nối Internet rồi tải lại trang.", "error");
    return;
  }

  const button = $("generateBtn");
  button.disabled = true;
  setStatus("Đang kiểm tra dữ liệu và tạo hợp đồng…", "info");

  try {
    let buffer;
    try {
      buffer = await getTemplateBuffer();
    } catch (error) {
      throw new Error(error?.message || `Không thể mở biểu mẫu. Hãy chọn lại tiêu chí hoặc bấm “Chọn file mẫu thủ công”.`);
    }

    const zip = await window.JSZip.loadAsync(buffer);
    const parts = wordXmlParts(zip);
    if (!parts.length) throw new Error("File đã chọn không có nội dung Word hợp lệ.");

    const xmlParts = await Promise.all(parts.map(async file => ({
      file,
      xml: await file.async("string")
    })));
    const requiredVariables = new Set(xmlParts.flatMap(part => Array.from(collectPlaceholders(part.xml))));
    for (const p of await collectPlaceholdersFromEmbeddedExcel(zip)) requiredVariables.add(p);
    const values = getValues();
    const missing = Array.from(requiredVariables).filter(name => !values[name]);
    if (missing.length) {
      markMissingFields(missing);
      const listed = missing.slice(0, 6).map(readableVariableName).join(", ");
      const more = missing.length > 6 ? ` và ${missing.length - 6} trường khác` : "";
      throw new Error(`Vui lòng điền các trường đang có trong mẫu: ${listed}${more}.`);
    }
    if (!validateDepositPercent()) {
      const depositInput = fields.get("deposit_percent");
      if (depositInput) {
        depositInput.scrollIntoView({ behavior: "smooth", block: "center" });
        depositInput.focus({ preventScroll: true });
      }
      return;
    }
    clearInvalidFields();

    xmlParts.forEach(({ file, xml }) => zip.file(file.name, replacePlaceholdersInXml(xml, values)));
    await processEmbeddedExcel(zip, values);
    const blob = await zip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });
    const filename = download(blob, values);
    setStatus(`Đã tạo xong ${filename}. Hãy mở file và kiểm tra lại nội dung trước khi gửi khách hàng.`, "success");
  } catch (error) {
    console.error(error);
    setStatus(error.message || "Có lỗi xảy ra khi tạo hợp đồng.", "error");
  } finally {
    button.disabled = false;
  }
}

function setSampleDate(name, date) {
  const input = fields.get(name);
  if (!input) return;
  const iso = toLocalIsoDate(date);
  input.value = formatDateDMY(iso);
  clearFieldInvalid(input);
  const native = input.closest(".date-picker-wrap")?.querySelector(".date-native");
  if (native) native.value = iso;
}

function fillSampleData() {
  const sample = {
    contract_number: "HTA-2026-0142",
    client_company_name_en: "Skyline Travel & Events Co., Ltd",
    client_company_name_vi: "Công ty TNHH Du lịch và Sự kiện Skyline",
    client_representative_name_en: "Mr. Nguyen Van An",
    client_representative_name_vi: "Ông Nguyễn Văn An",
    client_representative_name_en_vi: "Mr./Ông Nguyen Van An",
    client_representative_title_en: "General Director",
    client_representative_title_vi: "Tổng Giám đốc",
    client_address_en: "12 Nguyen Hue Street, District 1, Ho Chi Minh City",
    client_address_vi: "12 Nguyễn Huệ, Quận 1, TP. Hồ Chí Minh",
    client_phone: "028 3822 1234",
    client_mobile_phone: "0909 123 456",
    client_email: "an.nguyen@skyline-events.vn",
    client_VAT_code: "0312345678",
    deposit_percent: "50"
  };
  for (const [name, value] of Object.entries(sample)) setFieldValue(name, value);

  const eventSample = {
    event_setup_type: "Banquet",
    meal_setup_type: "Banquet",
    event_start_time: "08:00",
    event_end_time: "17:00",
    meal_start_time: "11:30",
    meal_end_time: "13:30",
    morning_tea_break_start_time: "09:00",
    morning_tea_break_end_time: "09:30",
    afternoon_tea_break_start_time: "15:00",
    afternoon_tea_break_end_time: "15:30",
    tea_break_venue: "Pre-function area",
    number_of_persons: "200"
  };
  for (const [name, value] of Object.entries(eventSample)) setFieldValue(name, value);
  const unitPrice = fields.get("unit_price");
  if (unitPrice) { unitPrice.value = "1000000"; formatPriceInput(unitPrice); }
  const ledPrice = fields.get("LED_unit_price");
  if (ledPrice) { ledPrice.value = "5000000"; formatPriceInput(ledPrice); }

  const today = new Date();
  const addDays = days => {
    const date = new Date(today);
    date.setDate(date.getDate() + days);
    return date;
  };

  setSampleDate("excecute_date", addDays(30));
  setSampleDate("request_to_settle_before_date", addDays(14));
  setSampleDate("contract_cancellation_date", addDays(21));

  const createdIso = toLocalIsoDate(today);
  const returnIso = toLocalIsoDate(addDays(7));
  $("contractCreatedFullDate").value = createdIso;
  $("returnAgreementFullDate").value = returnIso;
  const dispCreated = $("contractCreatedFullDate_display");
  const dispReturn = $("returnAgreementFullDate_display");
  if (dispCreated) dispCreated.value = formatDateDMY(createdIso);
  if (dispReturn) dispReturn.value = formatDateDMY(returnIso);
  applyDateParts("contractCreatedFullDate", "contract_created");
  applyDateParts("returnAgreementFullDate", "returning_agreement");

  clearInvalidFields();
  setStatus("Đã điền dữ liệu mẫu. Hãy kiểm tra và chỉnh sửa trước khi tạo file.", "info");
}

function initDatePickers() {
  document.querySelectorAll('.date-picker-wrap').forEach(wrap => {
    const display = wrap.querySelector('input[type="text"]');
    const native = wrap.querySelector('.date-native');
    if (!display || !native) return;
    display.addEventListener('click', () => { try { native.showPicker(); } catch (_) {} });
    native.addEventListener('change', () => {
      if (!native.value) return;
      display.value = formatDateDMY(native.value);
      display.dispatchEvent(new Event('change', { bubbles: true }));
    });
  });
}

function setLookupHint(message = "", type = "info") {
  const hint = $("lookupVatHint");
  if (!hint) return;
  hint.textContent = message;
  hint.className = message ? `lookup-hint is-${type}` : "lookup-hint";
}

// Tra cứu tên và địa chỉ công ty theo mã số thuế qua API công khai của VietQR
// (không cần API key, dữ liệu tổng hợp từ Cục Thuế).
async function lookupClientByVatCode() {
  const button = $("lookupVatBtn");
  const codeInput = fields.get("client_VAT_code");
  if (!button || !codeInput) return;

  const code = codeInput.value.trim().replace(/[\s-]/g, "");
  if (!/^\d{10}(\d{3})?$/.test(code)) {
    setLookupHint("Mã số thuế phải gồm 10 hoặc 13 chữ số.", "error");
    codeInput.focus();
    return;
  }

  button.disabled = true;
  setLookupHint("Đang tra cứu thông tin doanh nghiệp…", "info");

  try {
    const response = await fetch(`https://api.vietqr.io/v2/business/${code}`, { cache: "no-store" });
    if (!response.ok) throw new Error("network");
    const payload = await response.json();
    const data = payload && payload.data;

    if (payload.code !== "00" || !data || !data.name) {
      setLookupHint("Không tìm thấy doanh nghiệp với mã số thuế này.", "error");
      return;
    }

    if (data.name) setFieldValue("client_company_name_vi", data.name);
    if (data.internationalName) setFieldValue("client_company_name_en", data.internationalName);
    if (data.address) setFieldValue("client_address_vi", data.address);

    const filled = ["Tên công ty tiếng Việt"];
    if (data.internationalName) filled.push("Tên công ty tiếng Anh");
    if (data.address) filled.push("Địa chỉ tiếng Việt");
    setLookupHint(`Đã cập nhật: ${filled.join(", ")}. Vui lòng bổ sung các trường còn lại.`, "success");
  } catch (error) {
    console.error(error);
    setLookupHint("Không kết nối được tới dịch vụ tra cứu. Hãy kiểm tra Internet rồi thử lại.", "error");
  } finally {
    button.disabled = false;
  }
}

function initSaleRepPicker() {
  const group = $("segment_sale_rep_fullname");
  const nameField = $("sale_rep_fullname_field");
  const nameInput = $("field_sale_rep_fullname");
  const titleVi = $("field_sale_rep_title_vi");
  const titleEn = $("field_sale_rep_title_en");
  if (!group || !nameInput || !titleVi || !titleEn) return;
  const radios = group.querySelectorAll('input[type="radio"]');
  const sync = () => {
    const checked = group.querySelector('input[type="radio"]:checked');
    if (!checked) return;
    const isCustom = checked.value === "";
    if (nameField) nameField.hidden = !isCustom;
    titleVi.readOnly = !isCustom;
    titleEn.readOnly = !isCustom;
    if (isCustom) {
      nameInput.value = "";
      titleVi.value = "";
      titleEn.value = "";
      nameInput.focus();
    } else {
      nameInput.value = checked.value;
      titleVi.value = checked.dataset.titleVi || "";
      titleEn.value = checked.dataset.titleEn || "";
    }
  };
  radios.forEach(radio => radio.addEventListener("change", sync));
  sync();
}

function setExtractStatus(message = "", type = "info") {
  const status = $("extractStatus");
  if (!status) return;
  status.textContent = message;
  status.className = message ? `status is-visible ${type}` : "status";
}

function setExtractLoading(active) {
  const loader = $("extractLoader");
  const button = $("extractBtn");
  if (loader) loader.hidden = !active;
  if (button) button.classList.toggle("is-loading", active);
}

// Điền ngày dạng ISO (YYYY-MM-DD) vào cả ô hiển thị dd/mm/yyyy và input date ẩn.
function applyDateFieldIso(name, iso) {
  const input = fields.get(name);
  if (!input || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return;
  input.value = formatDateDMY(iso);
  clearFieldInvalid(input);
  const native = input.closest(".date-picker-wrap")?.querySelector(".date-native");
  if (native) native.value = iso;
}

function buildExtractionPrompt() {
  const lines = EXTRACTABLE_FIELDS.map(field => {
    const meaning = variablesByName.get(field.name)?.meaning || field.name;
    return `- ${field.name}: ${meaning}`;
  });
  return [
    "Bạn là trợ lý trích xuất thông tin hợp đồng sự kiện cho khách sạn.",
    "Người dùng sẽ dán nội dung tin nhắn/email của khách hàng.",
    "Hãy đọc kỹ và trả về DUY NHẤT một JSON object với các khóa sau (chỉ điền khi chắc chắn, không có thì để null):",
    ...lines,
    "",
    "Quy tắc:",
    "- Ngày (các trường *_date) trả về theo định dạng YYYY-MM-DD.",
    "- deposit_percent chỉ là con số (0-100), không kèm ký tự %.",
    "- client_representative_name_en_vi là tên đầy đủ kèm danh xưng dạng song ngữ, ví dụ \"Mr./Ông Nguyen Van An\".",
    "- Mã số thuế (client_VAT_code) chỉ gồm chữ số.",
    "- Không bịa thông tin. Nếu không tìm thấy thì để null.",
    "- Chỉ trả về JSON hợp lệ, không kèm giải thích."
  ].join("\n");
}

function applyExtractedValues(data) {
  let filled = 0;
  for (const field of EXTRACTABLE_FIELDS) {
    const raw = data[field.name];
    if (raw === null || raw === undefined) continue;
    const value = String(raw).trim();
    if (!value) continue;
    if (field.type === "date") applyDateFieldIso(field.name, value);
    else setFieldValue(field.name, value);
    filled += 1;
  }
  return filled;
}

function openExtractDialog() {
  const dialog = $("extractDialog");
  if (!dialog) return;
  const keyInput = $("geminiKey");
  const rememberInput = $("rememberKey");
  const savedKey = localStorage.getItem(GEMINI_KEY_STORAGE);
  if (savedKey && keyInput) {
    keyInput.value = savedKey;
    if (rememberInput) rememberInput.checked = true;
  }
  setExtractStatus("");
  setExtractLoading(false);
  const extractBtn = $("extractBtn");
  if (extractBtn) extractBtn.disabled = false;
  if (typeof dialog.showModal === "function") dialog.showModal();
  else dialog.setAttribute("open", "");
}

function closeExtractDialog() {
  if (extractAbortController) {
    extractAbortController.abort();
    extractAbortController = null;
  }
  setExtractLoading(false);
  const extractBtn = $("extractBtn");
  if (extractBtn) extractBtn.disabled = false;
  const dialog = $("extractDialog");
  if (!dialog) return;
  if (typeof dialog.close === "function") dialog.close();
  else dialog.removeAttribute("open");
}

async function extractInformation() {
  const button = $("extractBtn");
  const keyInput = $("geminiKey");
  const messageInput = $("messageContent");
  if (!button || !keyInput || !messageInput) return;

  const apiKey = keyInput.value.trim();
  const message = messageInput.value.trim();

  if (!apiKey) {
    setExtractStatus("Vui lòng nhập Google Gemini API key.", "error");
    keyInput.focus();
    return;
  }
  if (!message) {
    setExtractStatus("Vui lòng dán nội dung tin nhắn cần trích xuất.", "error");
    messageInput.focus();
    return;
  }

  if ($("rememberKey")?.checked) localStorage.setItem(GEMINI_KEY_STORAGE, apiKey);
  else localStorage.removeItem(GEMINI_KEY_STORAGE);

  button.disabled = true;
  setExtractLoading(true);
  setExtractStatus("");
  extractAbortController = new AbortController();

  try {
    const response = await fetch(GEMINI_API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": apiKey
      },
      body: JSON.stringify({
        system_instruction: { parts: [{ text: buildExtractionPrompt() }] },
        contents: [{ parts: [{ text: message }] }],
        generationConfig: { temperature: 0, responseMimeType: "application/json" }
      }),
      signal: extractAbortController.signal
    });

    if (response.status === 400 || response.status === 401 || response.status === 403) {
      throw new Error("API key không hợp lệ hoặc chưa được cấp quyền.");
    }
    if (response.status === 429) throw new Error("Đã vượt hạn mức sử dụng miễn phí. Hãy thử lại sau.");
    if (response.status === 503) throw new Error("Máy chủ AI đang quá tải (503). Đây là sự cố tạm thời phía Google, không phải lỗi của bạn. Hãy chờ vài giây rồi bấm Trích xuất lại.");
    if (!response.ok) throw new Error(`Dịch vụ AI trả về lỗi (${response.status}).`);

    const payload = await response.json();
    const content = payload.candidates?.[0]?.content?.parts?.map(part => part.text || "").join("");
    if (!content) throw new Error("AI không trả về nội dung. Hãy thử lại.");

    let data;
    try {
      data = JSON.parse(content);
    } catch (_) {
      throw new Error("Không đọc được kết quả từ AI. Hãy thử lại.");
    }

    const filled = applyExtractedValues(data);
    if (!filled) {
      setExtractStatus("Không tìm thấy thông tin phù hợp trong tin nhắn.", "error");
      return;
    }

    clearInvalidFields();
    closeExtractDialog();
    setStatus(`Đã trích xuất và điền ${filled} trường từ tin nhắn. Hãy kiểm tra lại trước khi tạo hợp đồng.`, "success");
  } catch (error) {
    if (error.name === "AbortError") return;
    console.error(error);
    const offline = error instanceof TypeError;
    setExtractStatus(offline ? "Không kết nối được tới dịch vụ AI. Hãy kiểm tra Internet rồi thử lại." : (error.message || "Có lỗi khi trích xuất thông tin."), "error");
  } finally {
    extractAbortController = null;
    button.disabled = false;
    setExtractLoading(false);
  }
}

// Dịch "Địa chỉ tiếng Việt" sang tiếng Anh qua endpoint dịch công khai của Google (không cần API key).
async function translateAddressToEnglish() {
  const viInput = fields.get("client_address_vi");
  if (!viInput) return;
  const text = viInput.value.trim();
  if (!text) return;

  try {
    const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=vi&tl=en&dt=t&q=${encodeURIComponent(text)}`;
    const response = await fetch(url, { cache: "no-store" });
    if (!response.ok) throw new Error("network");
    const data = await response.json();
    const translated = Array.isArray(data?.[0]) ? data[0].map(segment => segment?.[0] || "").join("") : "";
    if (translated.trim()) setFieldValue("client_address_en", translated.trim());
  } catch (error) {
    console.error(error);
  }
}

function getCriteriaSelection() {
  return {
    venue: $(TEMPLATE_CRITERIA.venue.selectId).value,
    day: $(TEMPLATE_CRITERIA.day.selectId).value,
    pkg: $(TEMPLATE_CRITERIA.pkg.selectId).value,
    meal: $(TEMPLATE_CRITERIA.meal.selectId).value
  };
}

// Cập nhật file mẫu và trạng thái dựa trên tiêu chí đang chọn.
function updateResolvedTemplate() {
  const info = $("templateResolved");
  const resolved = resolveTemplateFile(getCriteriaSelection());
  if (resolved) {
    criteriaTemplateFile = resolved.file;
    info.className = "template-resolved is-ok";
    info.textContent = `Mẫu phù hợp: ${resolved.file}` + (resolved.editing ? " (bản đang chỉnh sửa)" : "");
    $("generateBtn").disabled = false;
  } else {
    criteriaTemplateFile = null;
    info.className = "template-resolved is-warn";
    info.textContent = "Không tìm thấy mẫu phù hợp với các tiêu chí đã chọn.";
    $("generateBtn").disabled = true;
  }
}

// Khi đổi tiêu chí thì bỏ file mẫu chọn thủ công (nếu có) để tiêu chí được ưu tiên.
function onCriteriaChange() {
  if (selectedTemplateBuffer) {
    selectedTemplateBuffer = null;
    $("templateName").textContent = "";
  }
  updateMealVenueOption();
  updateResolvedTemplate();
}

// Tùy chọn "theo sảnh" (data-venue-option) của các select địa điểm cập nhật theo loại sảnh đang chọn.
function updateMealVenueOption() {
  const venue = $(TEMPLATE_CRITERIA.venue.selectId)?.value || "";
  for (const opt of document.querySelectorAll('#formFields option[data-venue-option="true"]')) {
    opt.value = venue;
    opt.textContent = venue || "—";
  }
}

// Giờ bắt đầu/kết thúc sự kiện mặc định theo Thời lượng.
const DAY_EVENT_TIME_DEFAULTS = {
  "fullday":           ["08:00", "17:00"],
  "halfday-morning":   ["08:00", "12:00"],
  "halfday-afternoon": ["13:30", "17:30"],
  "YEP":               ["10:00", "14:00"]
};

function applyEventTimeDefaults() {
  const preset = DAY_EVENT_TIME_DEFAULTS[$(TEMPLATE_CRITERIA.day.selectId)?.value];
  if (!preset) return;
  const [start, end] = preset;
  const startInput = fields.get("event_start_time");
  const endInput = fields.get("event_end_time");
  if (startInput) { startInput.value = start; clearFieldInvalid(startInput); }
  if (endInput) { endInput.value = end; clearFieldInvalid(endInput); }
}

// Giờ bắt đầu/kết thúc bữa ăn mặc định theo Bữa ăn.
const MEAL_TIME_DEFAULTS = {
  "lunch":  ["12:00", "13:30"],
  "dinner": ["18:00", "22:00"]
};

function applyMealTimeDefaults() {
  const preset = MEAL_TIME_DEFAULTS[$(TEMPLATE_CRITERIA.meal.selectId)?.value];
  if (!preset) return;
  const [start, end] = preset;
  const startInput = fields.get("meal_start_time");
  const endInput = fields.get("meal_end_time");
  if (startInput) { startInput.value = start; clearFieldInvalid(startInput); }
  if (endInput) { endInput.value = end; clearFieldInvalid(endInput); }
}

// Ẩn nhóm "Bữa ăn" khi không kèm bữa ăn (nomeal).
function updateMealFieldsVisibility() {
  const isNoMeal = $(TEMPLATE_CRITERIA.meal.selectId)?.value === "nomeal";
  const subgroup = fields.get("meal_start_time")?.closest(".subgroup");
  if (subgroup) subgroup.hidden = isNoMeal;
}

// Ẩn nhóm "Tiệc trà" khi gói dịch vụ là thuê phòng (room-rental).
function updateTeaBreakVisibility() {
  const isRoomRental = $(TEMPLATE_CRITERIA.pkg.selectId)?.value === "room-rental";
  const subgroup = fields.get("morning_tea_break_start_time")?.closest(".subgroup");
  if (subgroup) subgroup.hidden = isRoomRental;
}

// Điều chỉnh ô "Bữa ăn" theo gói dịch vụ:
// - room-rental: buộc nomeal và khóa.
// - package1/2/3: cho chọn nhưng bỏ tùy chọn nomeal.
// - còn lại: mở khóa, đủ tùy chọn.
function updateMealCriteriaForPackage() {
  const mealSelect = $(TEMPLATE_CRITERIA.meal.selectId);
  if (!mealSelect) return;
  const pkg = $(TEMPLATE_CRITERIA.pkg.selectId)?.value;
  const isRoomRental = pkg === "room-rental";
  const hideNomeal = pkg === "package1" || pkg === "package2" || pkg === "package3";

  const nomealOption = Array.from(mealSelect.options).find(o => o.value === "nomeal");
  if (nomealOption) {
    nomealOption.hidden = hideNomeal;
    nomealOption.disabled = hideNomeal;
  }
  if (isRoomRental) {
    mealSelect.value = "nomeal";
    mealSelect.disabled = true;
  } else {
    mealSelect.disabled = false;
    if (hideNomeal && mealSelect.value === "nomeal") mealSelect.value = "lunch";
  }
}

// Giới hạn Thời lượng theo gói dịch vụ.
// - package (trọn gói): cả ngày, nửa ngày sáng, nửa ngày chiều.
// - room-rental (thuê phòng): cả ngày, nửa ngày.
// - package1/2/3: chỉ YEP và khóa.
const PACKAGE_ALLOWED_DAYS = {
  "package":     new Set(["fullday", "halfday-morning", "halfday-afternoon"]),
  "room-rental": new Set(["fullday", "halfday"]),
  "package1":    new Set(["YEP"]),
  "package2":    new Set(["YEP"]),
  "package3":    new Set(["YEP"])
};
function updateDayOptionsForPackage() {
  const daySelect = $(TEMPLATE_CRITERIA.day.selectId);
  if (!daySelect) return;
  const pkg = $(TEMPLATE_CRITERIA.pkg.selectId)?.value;
  const allowed = PACKAGE_ALLOWED_DAYS[pkg];
  for (const opt of daySelect.options) {
    const hide = !!allowed && !allowed.has(opt.value);
    opt.hidden = hide;
    opt.disabled = hide;
  }
  if (allowed && !allowed.has(daySelect.value)) daySelect.value = [...allowed][0];
  daySelect.disabled = pkg === "package1" || pkg === "package2" || pkg === "package3";
}

// Khởi tạo các trường chi tiết sự kiện: địa điểm dùng bữa và định dạng giá tiền.
function initEventDetailFields() {
  updateMealVenueOption();
  applyEventTimeDefaults();
  applyMealTimeDefaults();
  updateMealCriteriaForPackage();
  updateDayOptionsForPackage();
  updateMealFieldsVisibility();
  updateTeaBreakVisibility();
  $(TEMPLATE_CRITERIA.day.selectId)?.addEventListener("change", applyEventTimeDefaults);
  $(TEMPLATE_CRITERIA.meal.selectId)?.addEventListener("change", () => {
    applyMealTimeDefaults();
    updateMealFieldsVisibility();
  });
  $(TEMPLATE_CRITERIA.pkg.selectId)?.addEventListener("change", () => {
    updateMealCriteriaForPackage();
    updateDayOptionsForPackage();
    applyEventTimeDefaults();
    applyMealTimeDefaults();
    updateMealFieldsVisibility();
    updateTeaBreakVisibility();
    updateResolvedTemplate();
  });
  for (const name of PRICE_FIELDS) {
    const input = fields.get(name);
    if (input) input.addEventListener("input", () => formatPriceInput(input));
  }
  for (const name of TIME_FIELDS) {
    const input = fields.get(name);
    if (input) input.addEventListener("change", () => normalizeTimeInput(input));
  }
}

function initTemplateSelector() {
  for (const config of Object.values(TEMPLATE_CRITERIA)) {
    const select = $(config.selectId);
    if (!select) continue;
    select.innerHTML = "";
    for (const opt of config.options) {
      const option = document.createElement("option");
      option.value = opt.value;
      option.textContent = opt.label;
      select.append(option);
    }
    select.addEventListener("change", onCriteriaChange);
  }
  updateResolvedTemplate();
}

async function initialize() {
  initFields();
  prepareDateHelpers();
  initDatePickers();
  initSaleRepPicker();
  initTemplateSelector();
  initEventDetailFields();
  $("generateBtn").disabled = false;

  const lookupBtn = $("lookupVatBtn");
  if (lookupBtn) {
    lookupBtn.addEventListener("click", lookupClientByVatCode);
    const vatInput = fields.get("client_VAT_code");
    if (vatInput) {
      vatInput.addEventListener("keydown", event => {
        if (event.key === "Enter") {
          event.preventDefault();
          lookupClientByVatCode();
        }
      });
    }
  }

  $("templateFile").addEventListener("change", async event => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      selectedTemplateBuffer = await file.arrayBuffer();
      $("templateName").textContent = `Đang dùng: ${file.name}`;
      $("generateBtn").disabled = false;
      setStatus("Đã chọn biểu mẫu thủ công. Hệ thống sẽ kiểm tra các biến của mẫu này khi xuất file.", "info");
    } catch (error) {
      selectedTemplateBuffer = null;
      setStatus("Không thể đọc file mẫu đã chọn.", "error");
    }
  });

  $("generateBtn").addEventListener("click", generateContract);

  const sampleBtn = $("sampleDataBtn");
  if (sampleBtn) sampleBtn.addEventListener("click", fillSampleData);

  const depositInput = fields.get("deposit_percent");
  if (depositInput) {
    depositInput.addEventListener("change", validateDepositPercent);
    depositInput.addEventListener("input", () => {
      if (depositInput.classList.contains("is-invalid") && validateDepositPercent()) setStatus("");
    });
  }

  for (const name of UPPERCASE_FIELDS) {
    const input = fields.get(name);
    if (!input) continue;
    input.addEventListener("input", event => {
      if (event.isComposing) return;
      const start = input.selectionStart;
      const end = input.selectionEnd;
      input.value = input.value.toUpperCase();
      try { input.setSelectionRange(start, end); } catch (_) {}
    });
    input.addEventListener("change", () => { input.value = input.value.toUpperCase(); });
  }

  const addressViInput = fields.get("client_address_vi");
  if (addressViInput) addressViInput.addEventListener("change", translateAddressToEnglish);

  const openExtractBtn = $("openExtractBtn");
  if (openExtractBtn) openExtractBtn.addEventListener("click", openExtractDialog);
  $("extractBtn")?.addEventListener("click", extractInformation);
  $("extractCloseBtn")?.addEventListener("click", closeExtractDialog);
  $("extractCancelBtn")?.addEventListener("click", closeExtractDialog);
}

document.addEventListener("DOMContentLoaded", initialize);
