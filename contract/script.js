"use strict";

const DEFAULT_TEMPLATE = "contract-template.docx";
const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const MONTHS_EN = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];

// Dùng khi trình duyệt mở trực tiếp bằng file:// và không thể fetch variable.csv.
const FALLBACK_VARIABLES = [
  ["excecute_date", "Ngày thực hiện hợp đồng"],
  ["contract_number", "Số hợp đồng"],
  ["contract_created_date", "Ngày tạo hợp đồng"],
  ["contract_created_month", "Tháng tạo hợp đồng"],
  ["contract_created_month_en", "Tháng tạo hợp đồng tiếng Anh"],
  ["contract_created_year", "Năm tạo hợp đồng"],
  ["client_company_name_en", "Tên công ty khách hàng tiếng Anh"],
  ["client_company_name_vi", "Tên công ty khách hàng tiếng Việt"],
  ["client_representative_name_en", "Họ tên và danh xưng tiếng Anh"],
  ["client_representative_name_vi", "Họ tên và danh xưng tiếng Việt"],
  ["client_representative_name_en_vi", "Họ tên và danh xưng Anh - Việt"],
  ["client_representative_title_en", "Chức danh tiếng Anh"],
  ["client_representative_title_vi", "Chức danh tiếng Việt"],
  ["client_address_en", "Địa chỉ tiếng Anh"],
  ["client_address_vi", "Địa chỉ tiếng Việt"],
  ["client_office_address_en", "Địa chỉ văn phòng tiếng Anh"],
  ["client_office_address_vi", "Địa chỉ văn phòng tiếng Việt"],
  ["client_phone", "Số điện thoại công ty khách hàng"],
  ["client_VAT_code", "Mã số thuế của khách hàng"],
  ["deposit_percent", "Phần trăm đặt cọc"],
  ["request_to_settle_before_date", "Yêu cầu thanh toán trước ngày"],
  ["contract_cancellation_date", "Ngày hủy hợp đồng"],
  ["returning_agreement_date", "Ngày hạn gửi lại hợp đồng cho khách sạn"],
  ["returning_agreement_month", "Tháng hạn gửi lại hợp đồng cho khách sạn"],
  ["returning_agreement_month_en", "Tháng tiếng Anh hạn gửi lại hợp đồng"],
  ["returning_agreement_year", "Năm hạn gửi lại hợp đồng cho khách sạn"],
  ["sale_rep_fullname", "Danh xưng và tên sales"],
  ["sale_rep_title_en", "Chức vụ sales tiếng Anh"],
  ["sale_rep_title_vi", "Chức vụ sales tiếng Việt"]
].map(([name, meaning]) => ({ name, meaning }));

const FIELD_GROUPS = [
  {
    title: "Thông tin hợp đồng",
    names: [
      "excecute_date", "contract_number", "contract_created_date",
      "contract_created_month", "contract_created_month_en", "contract_created_year"
    ]
  },
  {
    title: "Thông tin khách hàng (Bên B)",
    names: [
      "client_company_name_en", "client_company_name_vi",
      "client_representative_name_en", "client_representative_name_vi",
      "client_representative_name_en_vi", "client_representative_title_en",
      "client_representative_title_vi", "client_address_en", "client_address_vi",
      "client_office_address_en", "client_office_address_vi", "client_phone", "client_VAT_code"
    ]
  },
  {
    title: "Thanh toán và thời hạn",
    names: [
      "deposit_percent", "request_to_settle_before_date", "contract_cancellation_date",
      "returning_agreement_date", "returning_agreement_month",
      "returning_agreement_month_en", "returning_agreement_year"
    ]
  },
  {
    title: "Thông tin nhân viên kinh doanh",
    names: ["sale_rep_fullname", "sale_rep_title_en", "sale_rep_title_vi"]
  }
];

// Các biến đang xuất hiện trong contract-template.docx. Khi chọn mẫu khác,
// hệ thống sẽ đọc lại trực tiếp trong file trước lúc xuất.
const DEFAULT_TEMPLATE_VARIABLES = new Set([
  "client_address_en", "client_address_vi", "client_company_name_en",
  "client_company_name_vi", "client_office_address_vi", "client_phone",
  "client_representative_name_en_vi", "client_representative_name_en",
  "client_representative_name_vi", "client_representative_title_en",
  "client_representative_title_vi", "client_VAT_code", "contract_cancellation_date",
  "contract_created_date", "contract_created_month_en", "contract_created_month",
  "contract_created_year", "contract_number", "excecute_date",
  "returning_agreement_date", "returning_agreement_month_en",
  "returning_agreement_month", "returning_agreement_year"
]);

const fields = new Map();
let variablesByName = new Map();
let selectedTemplateBuffer = null;

const $ = id => document.getElementById(id);

function setStatus(message = "", type = "info") {
  const status = $("status");
  status.textContent = message;
  status.className = message ? `status is-visible ${type}` : "status";
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    if (char === '"') {
      if (quoted && text[i + 1] === '"') {
        cell += '"';
        i += 1;
      } else {
        quoted = !quoted;
      }
    } else if (char === "," && !quoted) {
      row.push(cell.trim());
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && text[i + 1] === "\n") i += 1;
      row.push(cell.trim());
      if (row.some(value => value)) rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }
  row.push(cell.trim());
  if (row.some(value => value)) rows.push(row);
  return rows;
}

function cleanVariableName(value) {
  return String(value || "").replace(/^\uFEFF/, "").trim().replace(/^\{\{|\}\}$/g, "");
}

function readVariablesFromCsv(text) {
  const rows = parseCsv(text);
  return rows
    .slice(1)
    .map(([variable, meaning]) => ({ name: cleanVariableName(variable), meaning: String(meaning || "").trim() }))
    .filter(item => /^[A-Za-z][A-Za-z0-9_]*$/.test(item.name));
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
  if (input) input.value = value;
}

function prepareDateHelpers() {
  const today = new Date();
  const returnDate = new Date(today);
  returnDate.setDate(returnDate.getDate() + 7);

  $("contractCreatedFullDate").value = toLocalIsoDate(today);
  $("returnAgreementFullDate").value = toLocalIsoDate(returnDate);

  $("contractCreatedFullDate").addEventListener("change", () => applyDateParts("contractCreatedFullDate", "contract_created"));
  $("returnAgreementFullDate").addEventListener("change", () => applyDateParts("returnAgreementFullDate", "returning_agreement"));
}

function inputTypeFor(name) {
  if (name === "client_phone") return "tel";
  return "text";
}

function placeholderFor(name) {
  if (name.includes("date") || name.includes("month") || name.includes("year")) return "Ví dụ: 12/08/2026";
  if (name.includes("percent")) return "Ví dụ: 50%";
  return "Nhập thông tin";
}

function createField(variable) {
  const wrapper = document.createElement("div");
  const activeInDefaultTemplate = DEFAULT_TEMPLATE_VARIABLES.has(variable.name);
  wrapper.className = `field${activeInDefaultTemplate ? "" : " is-unused"}`;
  if (!activeInDefaultTemplate) wrapper.title = "Biến này có trong variable.csv nhưng chưa xuất hiện trong contract-template.docx hiện tại.";

  const label = document.createElement("label");
  const inputId = `field_${variable.name}`;
  label.htmlFor = inputId;
  label.append(document.createTextNode(variable.meaning || variable.name));
  if (activeInDefaultTemplate) {
    const required = document.createElement("span");
    required.className = "required-mark";
    required.textContent = "*";
    required.title = "Có trong contract-template.docx";
    label.append(required);
  }
  const code = document.createElement("code");
  code.textContent = `{{${variable.name}}}`;
  label.append(code);

  const input = document.createElement("input");
  input.id = inputId;
  input.name = variable.name;
  input.type = inputTypeFor(variable.name);
  input.placeholder = placeholderFor(variable.name);
  input.autocomplete = "off";

  wrapper.append(label, input);
  fields.set(variable.name, input);
  return wrapper;
}

function renderFields(variables) {
  const container = $("formFields");
  container.replaceChildren();
  fields.clear();
  variablesByName = new Map(variables.map(variable => [variable.name, variable]));
  const rendered = new Set();

  const groups = FIELD_GROUPS.map(group => ({
    ...group,
    variables: group.names.map(name => variablesByName.get(name)).filter(Boolean)
  }));
  const remaining = variables.filter(variable => !FIELD_GROUPS.some(group => group.names.includes(variable.name)));
  if (remaining.length) groups.push({ title: "Thông tin khác", variables: remaining });

  groups.forEach(group => {
    if (!group.variables.length) return;
    const section = document.createElement("section");
    const title = document.createElement("h3");
    title.className = "group-title";
    title.textContent = group.title;
    const grid = document.createElement("div");
    grid.className = "field-grid";

    group.variables.forEach(variable => {
      if (!rendered.has(variable.name)) {
        grid.append(createField(variable));
        rendered.add(variable.name);
      }
    });
    section.append(title, grid);
    container.append(section);
  });

  applyDateParts("contractCreatedFullDate", "contract_created");
  applyDateParts("returnAgreementFullDate", "returning_agreement");
}

function getValues() {
  return Object.fromEntries(Array.from(fields, ([name, input]) => [name, input.value.trim()]));
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

async function getTemplateBuffer() {
  if (selectedTemplateBuffer) return selectedTemplateBuffer;
  const response = await fetch(DEFAULT_TEMPLATE, { cache: "no-store" });
  if (!response.ok) throw new Error(`Không tìm thấy ${DEFAULT_TEMPLATE}.`);
  return response.arrayBuffer();
}

function readableVariableName(name) {
  return variablesByName.get(name)?.meaning || `{{${name}}}`;
}

function download(blob, values) {
  const source = values.contract_number || values.client_company_name_vi || values.client_company_name_en || "Hop dong su kien";
  const safeName = source.replace(/[\\/:*?"<>|]/g, "").replace(/\s+/g, " ").trim() || "Hop dong su kien";
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `Hop dong - ${safeName}.docx`;
  document.body.append(link);
  link.click();
  window.setTimeout(() => {
    URL.revokeObjectURL(link.href);
    link.remove();
  }, 1500);
  return link.download;
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
      throw new Error(`Không thể tự mở ${DEFAULT_TEMPLATE}. Hãy bấm “Chọn biểu mẫu khác” để chọn file mẫu thủ công.`);
    }

    const zip = await window.JSZip.loadAsync(buffer);
    const parts = wordXmlParts(zip);
    if (!parts.length) throw new Error("File đã chọn không có nội dung Word hợp lệ.");

    const xmlParts = await Promise.all(parts.map(async file => ({
      file,
      xml: await file.async("string")
    })));
    const requiredVariables = new Set(xmlParts.flatMap(part => Array.from(collectPlaceholders(part.xml))));
    const values = getValues();
    const missing = Array.from(requiredVariables).filter(name => !values[name]);
    if (missing.length) {
      const listed = missing.slice(0, 6).map(readableVariableName).join(", ");
      const more = missing.length > 6 ? ` và ${missing.length - 6} trường khác` : "";
      throw new Error(`Vui lòng điền các trường đang có trong mẫu: ${listed}${more}.`);
    }

    xmlParts.forEach(({ file, xml }) => zip.file(file.name, replacePlaceholdersInXml(xml, values)));
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

async function loadVariables() {
  try {
    const response = await fetch("variable.csv", { cache: "no-store" });
    if (!response.ok) throw new Error("Không đọc được variable.csv");
    const variables = readVariablesFromCsv(await response.text());
    return variables.length ? variables : FALLBACK_VARIABLES;
  } catch (error) {
    console.warn("Dùng danh sách biến dự phòng:", error);
    return FALLBACK_VARIABLES;
  }
}

async function initialize() {
  prepareDateHelpers();
  renderFields(await loadVariables());
  $("generateBtn").disabled = false;

  $("templateFile").addEventListener("change", async event => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      selectedTemplateBuffer = await file.arrayBuffer();
      $("templateName").textContent = `Đang dùng: ${file.name}`;
      setStatus("Đã chọn biểu mẫu mới. Hệ thống sẽ kiểm tra các biến của mẫu này khi xuất file.", "info");
    } catch (error) {
      selectedTemplateBuffer = null;
      setStatus("Không thể đọc file mẫu đã chọn.", "error");
    }
  });

  $("generateBtn").addEventListener("click", generateContract);
}

document.addEventListener("DOMContentLoaded", initialize);
