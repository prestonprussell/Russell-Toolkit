const form = document.getElementById("upload-form");
const vendorTypeSelect = document.getElementById("vendor-type");
const vendorTypeButtons = [...document.querySelectorAll(".vendor-type-btn")];
const invoiceInput = document.getElementById("invoice-file");
const csvInput = document.getElementById("csv-files");
const invoiceUploadSection = document.getElementById("invoice-upload-section");
const csvUploadSection = document.getElementById("csv-upload-section");
const invoiceLabelText = document.getElementById("invoice-label-text");
const csvLabelText = document.getElementById("csv-label-text");
const invoiceDropzone = document.getElementById("invoice-dropzone");
const csvDropzone = document.getElementById("csv-dropzone");
const resultPanel = document.getElementById("result-panel");
const summarySection = document.getElementById("summary-section");
const usersSection = document.getElementById("adobe-users-section");
const nonUserSection = document.getElementById("non-user-section");
const branchAssignmentSection = document.getElementById("branch-assignment-section");
const supportReviewSection = document.getElementById("support-review-section");
const usersSectionTitle = document.getElementById("users-section-title");
const usersSectionHelp = document.getElementById("users-section-help");
const analyzeBtn = document.getElementById("analyze-btn");
const saveBranchesBtn = document.getElementById("save-branches-btn");
const downloadBtn = document.getElementById("download-btn");
const summaryBody = document.getElementById("summary-body");
const usersBody = document.getElementById("adobe-users-body");
const nonUserBody = document.getElementById("non-user-body");
const branchAssignmentBody = document.getElementById("branch-assignment-body");
const integricomBranchOptions = document.getElementById("integricom-branch-options");
const supportReviewBody = document.getElementById("support-review-body");
const supportBranchOptions = document.getElementById("support-branch-options");
const summaryCards = document.getElementById("summary-cards");
const warningsList = document.getElementById("warnings");
const notesList = document.getElementById("file-notes");
const reconciliationList = document.getElementById("reconciliation-notes");
const missingUsersResultsList = document.getElementById("missing-users-results");

let latestCsvText = "";
let latestInvoiceFilename = "";
let currentUserRows = [];
let currentUserVendor = "";
let currentBranchAssignmentPrompts = [];
let currentSupportRows = [];
let currentNonUserRows = [];
let currentSummaryRows = [];

function formatMoney(value) {
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(value || 0);
}

function displayVendorName(value) {
  if (value === "integricom") return "Integricom Licensing";
  if (value === "integricom_support") return "Integricom Support Hours";
  if (value === "adobe") return "Adobe";
  if (value === "hexnode") return "Hexnode";
  if (value === "generic") return "Generic";
  return value || "";
}

function refreshAnalyzeButtonLabel() {
  const vendor = vendorTypeSelect.value;
  if (vendor === "integricom_support" && currentSupportRows.length) {
    analyzeBtn.textContent = "Apply Support Branch Review and Recalculate";
    return;
  }
  if (vendor === "integricom" && currentBranchAssignmentPrompts.length) {
    analyzeBtn.textContent = "Apply Branch Assignments and Recalculate";
    return;
  }
  if ((vendor === "adobe" || vendor === "integricom") && currentUserRows.length) {
    analyzeBtn.textContent = "Save Branch Updates and Recalculate";
    return;
  }
  analyzeBtn.textContent = "Analyze and Build Breakdown";
}

const UPLOAD_CONFIGS = {
  hexnode: {
    invoice: "Invoice PDF",
    csv: "Device Export CSV",
  },
  adobe: {
    invoice: "Invoice PDF",
    csv: "Adobe Users Export CSV (optional)",
  },
  integricom: {
    invoice: "Invoice PDF",
    csv: "Microsoft User Export CSV (optional)",
  },
  integricom_support: {
    invoice: "Invoice PDF",
    csv: null,
  },
  generic: {
    invoice: null,
    csv: "CSV File",
  },
};

function updateUploadSections(vendor) {
  const cfg = UPLOAD_CONFIGS[vendor] || { invoice: "Invoice", csv: "CSV File" };
  invoiceUploadSection.hidden = !cfg.invoice;
  csvUploadSection.hidden = !cfg.csv;
  if (cfg.invoice) invoiceLabelText.textContent = cfg.invoice;
  if (cfg.csv) csvLabelText.textContent = cfg.csv;
}

function setVendorType(nextVendor) {
  vendorTypeSelect.value = nextVendor;
  vendorTypeButtons.forEach((button) => {
    const isActive = button.dataset.vendor === nextVendor;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", isActive ? "true" : "false");
  });

  currentUserRows = [];
  currentUserVendor = "";
  currentBranchAssignmentPrompts = [];
  currentSupportRows = [];
  updateUploadSections(nextVendor);
  refreshAnalyzeButtonLabel();
}

function initializeResizableTables() {
  const tables = document.querySelectorAll("table.resizable-table");
  tables.forEach((table) => {
    if (!table.querySelector("colgroup")) {
      const headerCells = [...table.querySelectorAll("thead th")];
      if (headerCells.length) {
        const colgroup = document.createElement("colgroup");
        headerCells.forEach(() => {
          const col = document.createElement("col");
          colgroup.appendChild(col);
        });
        table.insertBefore(colgroup, table.firstChild);
      }
    }

    table.querySelectorAll("thead th").forEach((headerCell) => {
      if (headerCell.dataset.resizable === "true") {
        return;
      }
      headerCell.dataset.resizable = "true";

      const resizer = document.createElement("span");
      resizer.className = "col-resizer";

      let startX = 0;
      let startWidth = 0;
      const columnIndex = [...headerCell.parentElement.children].indexOf(headerCell);
      const column = table.querySelectorAll("colgroup col")[columnIndex];

      const setColumnWidth = (widthPx) => {
        if (!column) return;
        if (!widthPx) {
          column.style.width = "";
          headerCell.style.width = "";
          return;
        }
        const clampedWidth = Math.max(90, Math.min(700, widthPx));
        column.style.width = `${clampedWidth}px`;
        headerCell.style.width = `${clampedWidth}px`;
      };

      const handlePointerMove = (event) => {
        const delta = event.clientX - startX;
        setColumnWidth(startWidth + delta);
      };

      const handlePointerUp = () => {
        document.body.classList.remove("col-resize-active");
        window.removeEventListener("pointermove", handlePointerMove);
        window.removeEventListener("pointerup", handlePointerUp);
      };

      resizer.addEventListener("pointerdown", (event) => {
        event.preventDefault();
        startX = event.clientX;
        startWidth = headerCell.getBoundingClientRect().width;
        setColumnWidth(startWidth);
        document.body.classList.add("col-resize-active");
        window.addEventListener("pointermove", handlePointerMove);
        window.addEventListener("pointerup", handlePointerUp);
      });

      resizer.addEventListener("dblclick", (event) => {
        event.preventDefault();
        setColumnWidth(null);
      });

      headerCell.appendChild(resizer);
    });
  });
}

function autoFitResizableTables() {
  const tables = document.querySelectorAll("table.resizable-table");
  tables.forEach((table) => {
    const wrap = table.closest(".table-wrap");
    const headerCells = [...table.querySelectorAll("thead th")];
    const columns = [...table.querySelectorAll("colgroup col")];
    if (!headerCells.length || columns.length !== headerCells.length || !wrap) {
      return;
    }

    const measuredWidths = [];
    headerCells.forEach((headerCell, index) => {
      const relatedCells = [
        headerCell,
        ...table.querySelectorAll(`tbody tr td:nth-child(${index + 1})`),
      ];

      let maxWidth = 90;
      relatedCells.forEach((cell) => {
        maxWidth = Math.max(maxWidth, Math.ceil(cell.scrollWidth + 24));
      });
      measuredWidths.push(Math.max(90, Math.min(700, maxWidth)));
    });

    const availableWidth = Math.max(wrap.clientWidth - 4, 320);
    const totalMeasuredWidth = measuredWidths.reduce((sum, value) => sum + value, 0);
    const scale = totalMeasuredWidth > availableWidth ? availableWidth / totalMeasuredWidth : 1;

    measuredWidths.forEach((width, index) => {
      const fittedWidth = Math.max(72, Math.floor(width * scale));
      columns[index].style.width = `${fittedWidth}px`;
      headerCells[index].style.width = `${fittedWidth}px`;
    });
  });
}

function scheduleAutoFitTables() {
  window.requestAnimationFrame(() => {
    window.requestAnimationFrame(() => {
      autoFitResizableTables();
    });
  });
}

function updateDropzoneLabel(dropzone, input, baseText) {
  const label = dropzone?.querySelector(".dropzone-text");
  if (!label || !input) return;
  if (!input.files || !input.files.length) {
    label.textContent = baseText;
    return;
  }
  if (input.multiple) {
    label.textContent = `${input.files.length} file(s) selected`;
  } else {
    label.textContent = input.files[0].name;
  }
}

function setInputFiles(input, files, { append = false } = {}) {
  const dataTransfer = new DataTransfer();
  if (append && input.multiple && input.files?.length) {
    for (const existingFile of input.files) {
      dataTransfer.items.add(existingFile);
    }
  }
  for (const file of files) {
    dataTransfer.items.add(file);
  }
  input.files = dataTransfer.files;
}

function initializeDropzone(dropzone, input, baseText) {
  if (!dropzone || !input) return;

  updateDropzoneLabel(dropzone, input, baseText);
  input.addEventListener("change", () => updateDropzoneLabel(dropzone, input, baseText));

  dropzone.addEventListener("click", (event) => {
    if (event.target === input) return;
    input.click();
  });

  const setDragState = (active) => {
    dropzone.classList.toggle("drag-active", active);
  };

  ["dragenter", "dragover"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      setDragState(true);
    });
  });

  ["dragleave", "dragend"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      setDragState(false);
    });
  });

  dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    setDragState(false);
    const droppedFiles = [...(event.dataTransfer?.files || [])];
    if (!droppedFiles.length) return;

    if (input.multiple) {
      setInputFiles(input, droppedFiles, { append: true });
    } else {
      setInputFiles(input, [droppedFiles[0]]);
    }
    updateDropzoneLabel(dropzone, input, baseText);
  });
}

function clearResults() {
  summaryBody.innerHTML = "";
  usersBody.innerHTML = "";
  nonUserBody.innerHTML = "";
  branchAssignmentBody.innerHTML = "";
  integricomBranchOptions.innerHTML = "";
  supportReviewBody.innerHTML = "";
  supportBranchOptions.innerHTML = "";
  warningsList.innerHTML = "";
  notesList.innerHTML = "";
  reconciliationList.innerHTML = "";
  missingUsersResultsList.innerHTML = "";
  summaryCards.innerHTML = "";
  downloadBtn.hidden = true;
  latestCsvText = "";
  latestInvoiceFilename = "";
  resultPanel.hidden = true;
  summarySection.hidden = true;
  usersSection.hidden = true;
  nonUserSection.hidden = true;
  branchAssignmentSection.hidden = true;
  supportReviewSection.hidden = true;
  saveBranchesBtn.disabled = false;
  saveBranchesBtn.textContent = "Save Branch Changes";
  currentBranchAssignmentPrompts = [];
  currentSupportRows = [];
  currentNonUserRows = [];
  currentSummaryRows = [];
}

function addCard(label, value) {
  const card = document.createElement("div");
  card.className = "card";
  card.innerHTML = `<div class=\"label\">${label}</div><div class=\"value\">${value}</div>`;
  summaryCards.appendChild(card);
}

function renderWarnings(warnings) {
  warningsList.innerHTML = "";
  if (!warnings || !warnings.length) {
    const li = document.createElement("li");
    li.textContent = "No warnings.";
    warningsList.appendChild(li);
    return;
  }

  warnings.forEach((warning) => {
    const li = document.createElement("li");
    li.textContent = warning;
    warningsList.appendChild(li);
  });
}

function renderMissingUsers(missingUsers) {
  missingUsersResultsList.innerHTML = "";
  if (!missingUsers || !missingUsers.length) {
    const li = document.createElement("li");
    li.textContent = "No missing users detected.";
    missingUsersResultsList.appendChild(li);
    return;
  }

  missingUsers.forEach((user) => {
    const li = document.createElement("li");
    const name = `${user.first_name || ""} ${user.last_name || ""}`.trim();
    const label = name ? `${name} (${user.email})` : user.email;
    li.textContent = `${label} - ${user.branch || "Unknown Branch"}`;
    missingUsersResultsList.appendChild(li);
  });
}

function renderSummaryTable(summary) {
  summaryBody.innerHTML = "";
  currentSummaryRows = summary || [];
  summary.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.branch}</td>
      <td>${row.license}</td>
      <td>${formatMoney(row.total_amount)}</td>
    `;
    summaryBody.appendChild(tr);
  });
}

function renderUserTable(rows) {
  usersBody.innerHTML = "";
  currentUserRows = rows || [];

  currentUserRows.forEach((user, idx) => {
    const tr = document.createElement("tr");
    const missingClass = user.branch ? "" : " style=\"background:rgba(245, 158, 11, 0.18);\"";
    tr.innerHTML = `
      <td${missingClass}>${user.first_name || ""}</td>
      <td${missingClass}>${user.last_name || ""}</td>
      <td${missingClass}>${user.email || ""}</td>
      <td${missingClass}><input type="text" id="user-branch-${idx}" value="${user.branch || ""}" placeholder="Branch" /></td>
      <td${missingClass}>${user.license_list || ""}</td>
      <td${missingClass}>${formatMoney(user.user_total || 0)}</td>
    `;
    usersBody.appendChild(tr);
  });
}

function renderNonUserTable(rows) {
  nonUserBody.innerHTML = "";
  currentNonUserRows = rows || [];
  currentNonUserRows.forEach((row, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input type="text" id="non-user-branch-${idx}" value="${row.branch || ""}" placeholder="Branch" /></td>
      <td>${row.license || ""}</td>
      <td>${row.allocation_type || ""}</td>
      <td>${formatMoney(row.total_amount || 0)}</td>
    `;
    nonUserBody.appendChild(tr);
  });
}

function renderBranchAssignmentPrompts(prompts) {
  branchAssignmentBody.innerHTML = "";
  integricomBranchOptions.innerHTML = "";
  currentBranchAssignmentPrompts = prompts || [];

  const optionSet = new Set();
  currentBranchAssignmentPrompts.forEach((prompt) => {
    (prompt.available_branches || []).forEach((branch) => optionSet.add(branch));
  });
  [...optionSet].sort().forEach((branch) => {
    const option = document.createElement("option");
    option.value = branch;
    integricomBranchOptions.appendChild(option);
  });

  currentBranchAssignmentPrompts.forEach((prompt, idx) => {
    const tr = document.createElement("tr");
    const assignedLabel = (prompt.already_assigned_branches || []).join(", ");
    const existingValue = prompt.branch || "";
    const error = prompt.validation_error || "";
    const inputStyle = error ? " style=\"background:rgba(248, 113, 113, 0.2);\"" : "";
    tr.innerHTML = `
      <td>${prompt.license || ""}</td>
      <td>${formatMoney(prompt.unit_price || 0)}</td>
      <td>${prompt.prompt_index || idx + 1}</td>
      <td>${assignedLabel || "None"}</td>
      <td${inputStyle}>
        <input
          type="text"
          id="branch-assignment-${idx}"
          list="integricom-branch-options"
          value="${existingValue}"
          placeholder="Enter branch"
        />
        ${error ? `<div class="cell-error">${error}</div>` : ""}
      </td>
    `;
    branchAssignmentBody.appendChild(tr);
  });
}

function renderSupportReviewTable(rows, branchOptions) {
  supportReviewBody.innerHTML = "";
  supportBranchOptions.innerHTML = "";
  currentSupportRows = rows || [];

  const optionSet = new Set(branchOptions || []);
  currentSupportRows.forEach((row) => optionSet.add(row.branch || ""));
  [...optionSet]
    .filter((value) => value)
    .sort()
    .forEach((branch) => {
      const option = document.createElement("option");
      option.value = branch;
      supportBranchOptions.appendChild(option);
    });

  currentSupportRows.forEach((row, idx) => {
    const needsReview = Boolean(row.needs_review);
    const rowStyle = needsReview ? " style=\"background:rgba(245, 158, 11, 0.18);\"" : "";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td${rowStyle}>${row.charge_summary || ""}</td>
      <td${rowStyle}>${formatMoney(row.amount || 0)}</td>
      <td${rowStyle}>${(row.billable_hours || 0).toFixed(2)}</td>
      <td${rowStyle}>
        <input
          type="text"
          id="support-branch-${idx}"
          list="support-branch-options"
          value="${row.branch || ""}"
          placeholder="Branch"
        />
      </td>
      <td${rowStyle}>${row.confidence || ""}</td>
      <td${rowStyle}>${row.assignment_reason || ""}</td>
    `;
    supportReviewBody.appendChild(tr);
  });
}

function collectUserUpdates() {
  return currentUserRows.map((user, idx) => ({
    email: user.email,
    first_name: user.first_name || "",
    last_name: user.last_name || "",
    branch: (document.getElementById(`user-branch-${idx}`)?.value || user.branch || "").trim(),
  }));
}

function collectBranchAssignmentUpdates() {
  return currentBranchAssignmentPrompts.map((prompt, idx) => ({
    line_key: prompt.line_key,
    prompt_index: prompt.prompt_index,
    branch: (document.getElementById(`branch-assignment-${idx}`)?.value || prompt.branch || "").trim(),
  }));
}

function collectSupportUpdates() {
  return currentSupportRows.map((row, idx) => ({
    row_key: row.row_key,
    branch: (document.getElementById(`support-branch-${idx}`)?.value || row.branch || "").trim(),
  }));
}

function collectNonUserUpdates() {
  return currentNonUserRows.map((row, idx) => ({
    ...row,
    branch: (document.getElementById(`non-user-branch-${idx}`)?.value || row.branch || "").trim(),
  }));
}

function buildSummaryCsv(summaryRows) {
  const branchTotals = new Map();
  (summaryRows || []).forEach((row) => {
    branchTotals.set(row.branch, (branchTotals.get(row.branch) || 0) + Number(row.total_amount || 0));
  });

  const lines = [["Branch", "Total"]];
  [...branchTotals.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([branch, total]) => {
      lines.push([branch, Number(total.toFixed(2))]);
    });

  const grandTotal = [...branchTotals.values()].reduce((sum, value) => sum + value, 0);
  lines.push(["Grand Total", "", Number(grandTotal.toFixed(2))]);
  lines.push([]);
  lines.push(["Branch", "License", "TotalAmount", "BranchTotal"]);

  (summaryRows || []).forEach((row) => {
    lines.push([
      row.branch,
      row.license,
      Number(Number(row.total_amount || 0).toFixed(2)),
      Number((branchTotals.get(row.branch) || 0).toFixed(2)),
    ]);
  });

  return lines
    .map((line) =>
      line
        .map((value) => {
          const text = `${value ?? ""}`;
          return /[",\n]/.test(text) ? `"${text.replace(/"/g, "\"\"")}"` : text;
        })
        .join(","),
    )
    .join("\n");
}

function applyNonUserEditsToSummary(updatedNonUserRows) {
  const grouped = new Map();
  currentSummaryRows.forEach((row) => {
    const key = `${row.branch}|||${row.license}`;
    grouped.set(key, Number(row.total_amount || 0));
  });

  currentNonUserRows.forEach((row, idx) => {
    const updated = updatedNonUserRows[idx];
    const oldBranch = row.branch || "";
    const newBranch = updated.branch || "";
    const license = row.license || "";
    const amount = Number(row.total_amount || 0);
    if (!newBranch || oldBranch === newBranch) {
      return;
    }

    const oldKey = `${oldBranch}|||${license}`;
    const newKey = `${newBranch}|||${license}`;
    grouped.set(oldKey, Number(((grouped.get(oldKey) || 0) - amount).toFixed(2)));
    grouped.set(newKey, Number(((grouped.get(newKey) || 0) + amount).toFixed(2)));
  });

  const nextSummary = [...grouped.entries()]
    .filter(([, total]) => Math.abs(total) >= 0.005)
    .map(([key, total]) => {
      const [branch, license] = key.split("|||");
      return {
        branch,
        license,
        total_amount: Number(total.toFixed(2)),
      };
    })
    .sort((a, b) => {
      const branchCompare = a.branch.localeCompare(b.branch);
      return branchCompare || a.license.localeCompare(b.license);
    });

  currentSummaryRows = nextSummary;
  latestCsvText = buildSummaryCsv(nextSummary);
  renderSummaryTable(nextSummary);
}

function setUsersSectionCopy(vendorType) {
  if (vendorType === "integricom") {
    usersSectionTitle.textContent = "Integricom Users";
  } else {
    usersSectionTitle.textContent = "Adobe Users";
  }
  usersSectionHelp.textContent = "Branch is editable. Save branch edits here, then Analyze to recalculate totals if needed.";
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const payload = new FormData();
  const vendorType = vendorTypeSelect.value;
  const invoiceFile = invoiceInput.files[0];
  const csvFiles = csvInput.files;

  if (!csvFiles.length && !["adobe", "integricom", "integricom_support"].includes(vendorType)) {
    alert("Please add at least one CSV file.");
    return;
  }

  if (
    (vendorType === "adobe" ||
      vendorType === "hexnode" ||
      vendorType === "integricom" ||
      vendorType === "integricom_support") &&
    !invoiceFile
  ) {
    alert("This vendor mode requires an invoice upload.");
    return;
  }

  const pendingUserUpdates =
    (vendorType === "adobe" || vendorType === "integricom") && currentUserRows.length ? collectUserUpdates() : [];
  const pendingBranchAssignmentUpdates =
    vendorType === "integricom" && currentBranchAssignmentPrompts.length ? collectBranchAssignmentUpdates() : [];
  const pendingSupportUpdates =
    vendorType === "integricom_support" && currentSupportRows.length ? collectSupportUpdates() : [];

  clearResults();
  analyzeBtn.disabled = true;
  analyzeBtn.textContent = "Analyzing...";

  try {
    payload.append("vendor_type", vendorType);

    if (invoiceFile) {
      payload.append("invoice_file", invoiceFile);
    }

    if (pendingUserUpdates.length) {
      if (vendorType === "adobe") {
        payload.append("adobe_user_updates", JSON.stringify(pendingUserUpdates));
      }
      if (vendorType === "integricom") {
        payload.append("integricom_user_updates", JSON.stringify(pendingUserUpdates));
      }
    }
    if (vendorType === "integricom" && pendingBranchAssignmentUpdates.length) {
      payload.append("integricom_branch_item_updates", JSON.stringify(pendingBranchAssignmentUpdates));
    }
    if (vendorType === "integricom_support" && pendingSupportUpdates.length) {
      payload.append("integricom_support_updates", JSON.stringify(pendingSupportUpdates));
    }

    for (const file of csvFiles) {
      payload.append("csv_files", file);
    }

    const response = await fetch("/api/analyze", {
      method: "POST",
      body: payload,
    });

    if (!response.ok) {
      const failure = await response.json();
      throw new Error(failure.detail || "Upload failed.");
    }

    const data = await response.json();

    addCard("Grand Total", formatMoney(data.totals.grand_total));
    addCard("Breakdown Lines", data.totals.line_items);
    addCard("Source Files", data.files.length);
    addCard("Mode", displayVendorName(data.vendor_type));

    if (data.invoice) {
      addCard("Invoice Reference", data.invoice.filename);
      const baseName = data.invoice.filename.replace(/\.[^.]+$/, "");
      latestInvoiceFilename = `${baseName}_breakdown.csv`;
    }

    data.files.forEach((entry) => {
      const li = document.createElement("li");
      li.textContent = `${entry.filename}: ingested ${entry.rows_ingested}, skipped ${entry.rows_skipped}`;
      notesList.appendChild(li);
    });

    if (!data.files.length) {
      const li = document.createElement("li");
      li.textContent =
        data.vendor_type === "integricom_support"
          ? "Invoice-only mode: no CSV files were ingested."
          : "No valid CSV rows were ingested.";
      notesList.appendChild(li);
    }

    if (data.reconciliation) {
      const base = document.createElement("li");
      base.textContent = `Source base total: ${formatMoney(data.reconciliation.base_total)}`;
      reconciliationList.appendChild(base);

      const inv = document.createElement("li");
      inv.textContent = `Invoice total: ${formatMoney(data.reconciliation.invoice_total)}`;
      reconciliationList.appendChild(inv);

      const adj = document.createElement("li");
      adj.textContent = `Home Office adjustment applied: ${formatMoney(data.reconciliation.home_office_adjustment)}`;
      reconciliationList.appendChild(adj);
    } else {
      const li = document.createElement("li");
      li.textContent = "No invoice reconciliation applied.";
      reconciliationList.appendChild(li);
    }

    renderMissingUsers(data.missing_users || []);
    renderWarnings(data.warnings || []);

    if (data.vendor_type === "integricom_support") {
      currentUserRows = [];
      currentUserVendor = "";
      renderUserTable([]);
      renderNonUserTable([]);
      renderBranchAssignmentPrompts([]);
      usersSection.hidden = true;
      nonUserSection.hidden = true;
      branchAssignmentSection.hidden = true;

      const supportRows = data.support_rows || data.integricom_support_rows || [];
      renderSupportReviewTable(supportRows, data.support_branch_options || []);
      supportReviewSection.hidden = false;

      renderSummaryTable(data.summary || []);
      summarySection.hidden = false;

      if (data.needs_support_review) {
        const li = document.createElement("li");
        li.textContent =
          data.message ||
          "Some support blocks were defaulted to Home Office with low confidence. Review branch assignments.";
        warningsList.prepend(li);
      }
    } else if (data.vendor_type === "adobe" || data.vendor_type === "integricom") {
      const rows = data.user_rows || data.adobe_user_rows || data.integricom_user_rows || [];
      renderUserTable(rows);
      currentUserVendor = data.vendor_type;
      renderSupportReviewTable([], []);
      supportReviewSection.hidden = true;
      setUsersSectionCopy(data.vendor_type);
      usersSection.hidden = false;
      if (data.vendor_type === "integricom") {
        const nonUserRows = data.non_user_rows || data.integricom_non_user_rows || [];
        const branchPrompts = data.non_user_branch_prompts || data.integricom_non_user_branch_prompts || [];
        renderNonUserTable(nonUserRows);
        renderBranchAssignmentPrompts(branchPrompts);
        nonUserSection.hidden = false;
        branchAssignmentSection.hidden = !branchPrompts.length;
        renderSummaryTable(data.summary || []);
        summarySection.hidden = false;
      } else {
        renderNonUserTable([]);
        renderBranchAssignmentPrompts([]);
        nonUserSection.hidden = true;
        branchAssignmentSection.hidden = true;
        renderSummaryTable(data.summary || []);
        summarySection.hidden = false;
      }

      if (data.needs_user_enrichment) {
        const li = document.createElement("li");
        li.textContent = data.message || "Some users still need branch values.";
        warningsList.prepend(li);
      }
      if (data.needs_non_user_branch_assignment && !data.needs_user_enrichment) {
        const li = document.createElement("li");
        li.textContent =
          data.message ||
          "Some branch-tethered non-user charges need branch assignments before the breakdown can be finalized.";
        warningsList.prepend(li);
      }
    } else {
      currentUserRows = [];
      currentUserVendor = "";
      renderSupportReviewTable([], []);
      renderBranchAssignmentPrompts([]);
      renderSummaryTable(data.summary || []);
      renderNonUserTable([]);
      summarySection.hidden = false;
      usersSection.hidden = true;
      nonUserSection.hidden = true;
      branchAssignmentSection.hidden = true;
      supportReviewSection.hidden = true;
    }

    latestCsvText = data.breakdown_csv || "";
    if (latestCsvText) {
      downloadBtn.hidden = false;
    }

    resultPanel.hidden = false;
    scheduleAutoFitTables();
  } catch (error) {
    alert(error.message);
  } finally {
    analyzeBtn.disabled = false;
    refreshAnalyzeButtonLabel();
  }
});

saveBranchesBtn.addEventListener("click", async () => {
  if (!currentUserRows.length || !currentUserVendor) {
    if (!currentNonUserRows.length) {
      alert("No editable rows loaded yet. Run an Adobe or Integricom analysis first.");
      return;
    }
  }

  const nonUserUpdates = collectNonUserUpdates();
  const hasNonUserBranchChanges = currentNonUserRows.some(
    (row, idx) => (row.branch || "") !== (nonUserUpdates[idx].branch || ""),
  );

  if (!currentUserRows.length || !currentUserVendor) {
    if (hasNonUserBranchChanges) {
      applyNonUserEditsToSummary(nonUserUpdates);
      currentNonUserRows = nonUserUpdates;
      renderNonUserTable(currentNonUserRows);
      const li = document.createElement("li");
      li.textContent = "Applied branch updates to non-user charges and refreshed totals/export.";
      warningsList.prepend(li);
    }
    return;
  }

  const saveEndpoint =
    currentUserVendor === "integricom" ? "/api/integricom/users/save" : "/api/adobe/users/save";

  const updates = collectUserUpdates();
  const originalText = saveBranchesBtn.textContent;
  saveBranchesBtn.disabled = true;
  saveBranchesBtn.textContent = "Saving...";

  try {
    const response = await fetch(saveEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(updates),
    });
    if (!response.ok) {
      const failure = await response.json();
      throw new Error(failure.detail || "Save failed.");
    }

    const result = await response.json();
    currentUserRows = currentUserRows.map((user, idx) => ({
      ...user,
      branch: updates[idx].branch,
    }));
    renderUserTable(currentUserRows);

    if (hasNonUserBranchChanges) {
      applyNonUserEditsToSummary(nonUserUpdates);
      currentNonUserRows = nonUserUpdates;
      renderNonUserTable(currentNonUserRows);
    }

    const li = document.createElement("li");
    li.textContent = hasNonUserBranchChanges
      ? `Saved branch updates for ${result.saved} users and refreshed non-user charge totals.`
      : `Saved branch updates for ${result.saved} users.`;
    warningsList.prepend(li);
    scheduleAutoFitTables();
  } catch (error) {
    alert(error.message);
  } finally {
    saveBranchesBtn.disabled = false;
    saveBranchesBtn.textContent = originalText;
    refreshAnalyzeButtonLabel();
  }
});

downloadBtn.addEventListener("click", () => {
  if (!latestCsvText) {
    return;
  }

  const blob = new Blob([latestCsvText], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = latestInvoiceFilename || "breakdown.csv";
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
});

vendorTypeButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setVendorType(button.dataset.vendor || "integricom");
  });
});

initializeResizableTables();
initializeDropzone(invoiceDropzone, invoiceInput, "Drag and drop invoice file here, or click to browse");
initializeDropzone(csvDropzone, csvInput, "Drag and drop one or more CSV files here, or click to browse");
setVendorType(vendorTypeSelect.value || "integricom");
refreshAnalyzeButtonLabel();
