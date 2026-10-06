const state = { files: [], tolerance: "0", busy: false };

const dropzone = document.querySelector("#dropzone");
const fileInput = document.querySelector("#file-input");
const fileList = document.querySelector("#file-list");
const errorBox = document.querySelector("#error");
const options = document.querySelector("#options");
const actionBar = document.querySelector("#action-bar");
const resultBox = document.querySelector("#result");
const buildBtn = document.querySelector("#build-btn");
const toleranceInput = document.querySelector("#tolerance");

const COLUMNS = [
  ["policyholder", "نام بیمه‌گذار"],
  ["insurance_no", "شماره بیمه"],
  ["insurance_type", "نوع بیمه"],
  ["month", "ماه"],
  ["national_id", "کد ملی"],
  ["first_name", "نام"],
  ["last_name", "نام خانوادگی"],
  ["full_name", "نام و نام خانوادگی"],
  ["note", "توضیحات"],
];

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    fileInput.click();
  }
});
["dragenter", "dragover"].forEach((name) => {
  dropzone.addEventListener(name, (event) => {
    event.preventDefault();
    dropzone.classList.add("is-over");
  });
});
["dragleave", "drop"].forEach((name) => {
  dropzone.addEventListener(name, (event) => {
    event.preventDefault();
    dropzone.classList.remove("is-over");
  });
});
dropzone.addEventListener("drop", (event) => uploadFiles(event.dataTransfer.files));
fileInput.addEventListener("change", () => {
  uploadFiles(fileInput.files);
  fileInput.value = "";
});
document.querySelector("#demo-btn").addEventListener("click", loadDemo);
buildBtn.addEventListener("click", buildReport);
toleranceInput.addEventListener("input", () => {
  state.tolerance = toleranceInput.value;
});
fileList.addEventListener("input", onFileInput);
fileList.addEventListener("change", onFileChange);
fileList.addEventListener("click", onFileClick);

function faDigits(value) {
  return String(value).replace(/\d/g, (digit) => "۰۱۲۳۴۵۶۷۸۹"[digit]);
}

function money(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return value;
  return new Intl.NumberFormat("fa-IR").format(number);
}

function esc(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function showError(message) {
  errorBox.hidden = !message;
  errorBox.textContent = message || "";
}

async function api(url, options) {
  const response = await fetch(url, options);
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    const detail = data.detail;
    throw new Error(typeof detail === "string" ? detail : "درخواست انجام نشد.");
  }
  return data;
}

async function uploadFiles(fileListLike) {
  const files = [...fileListLike];
  if (!files.length) return;
  showError("");
  const body = new FormData();
  files.forEach((file) => body.append("uploaded_files", file));
  try {
    const data = await api("/api/upload", { method: "POST", body });
    data.files.forEach((file) => state.files.push(adopt(file)));
    render();
  } catch (error) {
    showError(error.message);
  }
}

async function loadDemo() {
  showError("");
  document.querySelector("#demo-btn").disabled = true;
  try {
    const data = await api("/api/demo", { method: "POST" });
    state.files = data.files.map(adopt);
    resultBox.hidden = true;
    render();
  } catch (error) {
    showError(error.message);
  } finally {
    document.querySelector("#demo-btn").disabled = false;
  }
}

function adopt(file) {
  const sheet = file.sheets.find((item) => item.name === file.active_sheet) || file.sheets[0];
  return {
    id: file.id,
    filename: file.filename,
    label: file.label,
    role: file.role || "due",
    factor: "1",
    activeSheet: sheet.name,
    headerRow: sheet.header_row,
    sheets: file.sheets,
    columns: columnsFrom(sheet.suggestions),
    touched: false,
  };
}

function columnsFrom(suggestions) {
  const col = (key) => (suggestions[key] == null ? "" : String(suggestions[key]));
  const due = col("amount_due") || col("amount");
  const paid = suggestions.amount_paid != null && suggestions.amount_paid !== suggestions.amount_due
    ? col("amount_paid")
    : "";
  return {
    policyholder: col("policyholder"),
    insurance_no: col("insurance_no"),
    insurance_type: col("insurance_type"),
    month: col("month"),
    national_id: col("national_id"),
    first_name: col("first_name"),
    last_name: col("last_name"),
    full_name: col("full_name"),
    amount: col("amount"),
    amount_due: due,
    amount_paid: paid,
    note: col("note"),
  };
}

function currentSheet(file) {
  return file.sheets.find((sheet) => sheet.name === file.activeSheet) || file.sheets[0];
}

function render() {
  const hasFiles = state.files.length > 0;
  options.hidden = !hasFiles;
  actionBar.hidden = !hasFiles;
  fileList.innerHTML = state.files.map((file, index) => cardHtml(file, index)).join("");
  document.querySelectorAll(".steps li").forEach((item) => {
    const step = Number(item.dataset.step);
    item.classList.toggle("is-current", (step === 1 && !hasFiles) || (step === 2 && hasFiles && resultBox.hidden) || (step === 3 && !resultBox.hidden));
    item.classList.toggle("is-done", (step === 1 && hasFiles) || (step === 2 && !resultBox.hidden));
  });
}

function cardHtml(file, index) {
  const sheet = currentSheet(file);
  const headers = sheet.headers || [];
  const both = file.role === "both";
  const previewHeaders = headers.slice(0, 12);
  return `
    <article class="file-card" data-id="${esc(file.id)}">
      <div class="file-head">
        <div class="file-index">${faDigits(index + 1)}</div>
        <div>
          <h3>${esc(file.filename)}</h3>
          <p>${faDigits(sheet.row_count)} ردیف داده در برگه «${esc(sheet.name)}»</p>
        </div>
        <button type="button" class="btn danger" data-action="remove">حذف</button>
      </div>
      <div class="seg" role="radiogroup" aria-label="نقش این فایل">
        ${roleButton(file, "due", "قابل‌پرداخت")}
        ${roleButton(file, "paid", "پرداخت‌شده")}
        ${roleButton(file, "both", "هر دو در یک فایل")}
      </div>
      <div class="card-grid">
        <div>
          <label>عنوان این فهرست در گزارش</label>
          <input data-field="label" value="${esc(file.label)}">
        </div>
        <div>
          <label>برگه</label>
          <select data-field="sheet">${file.sheets.map((item) => `<option value="${esc(item.name)}" ${item.name === file.activeSheet ? "selected" : ""}>${esc(item.name)}</option>`).join("")}</select>
        </div>
        <div>
          <label>ردیف عنوان‌ها</label>
          <input class="num" data-field="header" inputmode="numeric" value="${esc(file.headerRow)}">
        </div>
        <div>
          <label>ضریب مبلغ</label>
          <input class="num" data-field="factor" inputmode="decimal" value="${esc(file.factor)}">
        </div>
      </div>
      <p class="help">ستون‌ها را با فایل خودتان مقابله کنید. تطبیق با شماره بیمه انجام می‌شود و اگر نباشد با کد ملی یا نام. ستون ماه را بگذارید تا معلوم شود کدام ماه‌ها مانده است.</p>
      <div class="map-grid">
        ${COLUMNS.map(([key, title]) => selectField(key, title, headers, file.columns[key])).join("")}
        ${both ? "" : selectField("amount", "ستون مبلغ", headers, file.columns.amount)}
        <div ${both ? "" : "hidden"}>
          ${selectField("amount_due", "ستون مبلغ قابل‌پرداخت", headers, file.columns.amount_due)}
        </div>
        <div ${both ? "" : "hidden"}>
          ${selectField("amount_paid", "ستون مبلغ پرداخت‌شده", headers, file.columns.amount_paid)}
        </div>
      </div>
      <div class="preview-wrap">
        <table>
          <thead><tr>${previewHeaders.map((header, column) => `<th data-col-index="${column}" class="${mappedClass(file, column)}">${esc(header)}</th>`).join("")}</tr></thead>
          <tbody>
            ${(sheet.preview || []).map((row) => `<tr>${previewHeaders.map((_, column) => `<td data-col-index="${column}" class="${mappedClass(file, column)} ${isNumeric(row[column]) ? "num" : ""}">${esc(row[column] || "")}</td>`).join("")}</tr>`).join("")}
          </tbody>
        </table>
      </div>
      ${headers.length > 12 ? `<p class="help">پیش‌نمایش فقط ۱۲ ستون اول را نشان می‌دهد. همهٔ ستون‌ها در فهرست انتخاب هستند.</p>` : ""}
    </article>`;
}

function roleButton(file, role, label) {
  return `<button type="button" data-role="${role}" aria-pressed="${file.role === role ? "true" : "false"}">${label}</button>`;
}

function selectField(key, title, headers, selected) {
  const options = [`<option value="">انتخاب نشده</option>`].concat(
    headers.map((header, index) => `<option value="${index}" ${String(selected) === String(index) ? "selected" : ""}>${esc(header)}</option>`)
  );
  return `<div><label>${title}</label><select data-col="${key}">${options.join("")}</select></div>`;
}

function mappedClass(file, column) {
  const used = new Set(Object.values(file.columns).filter((value) => value !== ""));
  return used.has(String(column)) ? "is-mapped" : "";
}

function isNumeric(value) {
  return /^[\d۰-۹0-9,٬.]+$/.test(String(value || ""));
}

function fileFrom(event) {
  const card = event.target.closest("[data-id]");
  if (!card) return null;
  return state.files.find((file) => file.id === card.dataset.id);
}

function onFileInput(event) {
  const file = fileFrom(event);
  if (!file) return;
  const field = event.target.dataset.field;
  if (field === "label") file.label = event.target.value;
  if (field === "factor") file.factor = event.target.value;
  if (event.target.dataset.col) {
    file.columns[event.target.dataset.col] = event.target.value;
    file.touched = true;
    const card = event.target.closest("[data-id]");
    card.querySelectorAll("[data-col-index]").forEach((cell) => {
      cell.classList.toggle("is-mapped", mappedClass(file, cell.dataset.colIndex) === "is-mapped");
    });
  }
}

async function onFileChange(event) {
  const file = fileFrom(event);
  if (!file) return;
  if (event.target.dataset.col) {
    file.columns[event.target.dataset.col] = event.target.value;
    file.touched = true;
  }
  if (event.target.dataset.field === "sheet") {
    file.activeSheet = event.target.value;
    const sheet = currentSheet(file);
    file.headerRow = sheet.header_row;
    file.touched = false;
    file.columns = columnsFrom(sheet.suggestions);
    render();
  }
  if (event.target.dataset.field === "header") {
    file.headerRow = Number(event.target.value || 1);
    await refreshPreview(file);
  }
}

async function onFileClick(event) {
  const file = fileFrom(event);
  if (!file) return;
  if (event.target.dataset.role) {
    file.role = event.target.dataset.role;
    render();
  }
  if (event.target.dataset.action === "remove") {
    await fetch(`/api/files/${file.id}`, { method: "DELETE" });
    state.files = state.files.filter((item) => item.id !== file.id);
    if (!state.files.length) resultBox.hidden = true;
    render();
  }
}

async function refreshPreview(file) {
  try {
    const data = await api(`/api/files/${file.id}/preview`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sheet: file.activeSheet, header_row: Number(file.headerRow) }),
    });
    const index = file.sheets.findIndex((sheet) => sheet.name === file.activeSheet);
    file.sheets[index] = data;
    if (!file.touched) file.columns = columnsFrom(data.suggestions);
    showError("");
    render();
  } catch (error) {
    showError(error.message);
  }
}

function validate(file) {
  const columns = file.columns;
  const hasIdentity = ["policyholder", "insurance_no", "national_id", "first_name", "last_name", "full_name"].some((key) => columns[key] !== "");
  if (!hasIdentity) return `در «${file.label || file.filename}» نام بیمه‌گذار، شماره بیمه یا نام را مشخص کنید.`;
  if (file.role === "both") {
    if (columns.amount_due === "" || columns.amount_paid === "") {
      return `در «${file.label || file.filename}» هر دو ستون مبلغ را مشخص کنید.`;
    }
    if (columns.amount_due === columns.amount_paid) {
      return `در «${file.label || file.filename}» دو ستون مبلغ نباید یکی باشند.`;
    }
  } else if (columns.amount === "") {
    return `در «${file.label || file.filename}» ستون مبلغ را مشخص کنید.`;
  }
  return "";
}

function syncFromDom() {
  state.tolerance = toleranceInput.value;
  document.querySelectorAll(".file-card").forEach((card) => {
    const file = state.files.find((item) => item.id === card.dataset.id);
    if (!file) return;
    const label = card.querySelector('[data-field="label"]');
    const factor = card.querySelector('[data-field="factor"]');
    const header = card.querySelector('[data-field="header"]');
    const sheet = card.querySelector('[data-field="sheet"]');
    if (label) file.label = label.value;
    if (factor) file.factor = factor.value;
    if (header) file.headerRow = Number(header.value || 1);
    if (sheet) file.activeSheet = sheet.value;
    card.querySelectorAll("[data-col]").forEach((select) => {
      file.columns[select.dataset.col] = select.value;
    });
  });
}

async function buildReport() {
  showError("");
  syncFromDom();
  for (const file of state.files) {
    const message = validate(file);
    if (message) {
      showError(message);
      return;
    }
  }
  const roles = new Set(state.files.map((file) => file.role));
  const hasDue = roles.has("due") || roles.has("both");
  const hasPaid = roles.has("paid") || roles.has("both");
  if (!hasDue || !hasPaid) {
    showError("حداقل یک فهرست قابل‌پرداخت و یک فهرست پرداخت‌شده لازم است.");
    return;
  }
  buildBtn.disabled = true;
  buildBtn.innerHTML = `<span class="spinner"></span> در حال ساخت گزارش`;
  try {
    const data = await api("/api/reconcile", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        tolerance: state.tolerance || "0",
        files: state.files.map((file) => ({
          id: file.id,
          sheet: file.activeSheet,
          header_row: Number(file.headerRow),
          label: file.label,
          role: file.role,
          factor: file.factor || "1",
          columns: nulls(file.columns),
        })),
      }),
    });
    renderResult(data);
    const link = document.createElement("a");
    link.href = data.download_url;
    link.download = data.filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
  } catch (error) {
    showError(error.message);
  } finally {
    buildBtn.disabled = false;
    buildBtn.textContent = "ساخت گزارش و دانلود اکسل";
  }
}

function nulls(columns) {
  const output = {};
  Object.entries(columns).forEach(([key, value]) => {
    output[key] = value === "" ? null : Number(value);
  });
  return output;
}

function renderResult(data) {
  const summary = data.summary;
  const balance = Number(summary.balance_total);
  const counts = summary.counts;
  resultBox.hidden = false;
  resultBox.innerHTML = `
    <div class="result-top">
      <div>
        <h2>گزارش آماده است</h2>
        <p class="hint">تاریخ گزارش: ${esc(data.created_at)}. فایل اکسل را برای تحویل باز کنید؛ برگهٔ «خلاصه» برای خواندن سریع است.</p>
      </div>
      <a class="btn primary" href="${esc(data.download_url)}" download="${esc(data.filename)}">دانلود ${esc(data.filename)}</a>
    </div>
    <div class="stats">
      <div class="stat"><b class="num">${faDigits(summary.people)}</b><span>بیمه‌گذار</span></div>
      <div class="stat"><b class="num">${money(summary.due_total)}</b><span>مبلغ کل</span></div>
      <div class="stat"><b class="num">${money(summary.paid_total)}</b><span>تطبیق پرداخت</span></div>
      <div class="stat ${balance > 0 ? "debt" : "ok"}"><b class="num">${money(summary.balance_total)}</b><span>مانده</span></div>
    </div>
    <div class="chips">
      <span class="chip">تسویه‌شده: ${faDigits(counts["تسویه‌شده"] || 0)}</span>
      <span class="chip">کسری پرداخت: ${faDigits(counts["کسری پرداخت"] || 0)}</span>
      <span class="chip">فاقد پرداخت: ${faDigits(counts["فاقد پرداخت"] || 0)}</span>
      <span class="chip">اضافه‌پرداخت: ${faDigits(counts["اضافه‌پرداخت"] || 0)}</span>
      <span class="chip">پرداخت بدون هزینه: ${faDigits(counts["پرداخت بدون هزینه"] || 0)}</span>
    </div>
    ${summary.warnings.length ? `<ul class="warn-list">${summary.warnings.slice(0, 6).map((line) => `<li>${esc(line)}</li>`).join("")}</ul>` : ""}
  `;
  render();
  resultBox.scrollIntoView({ behavior: "smooth", block: "start" });
}
