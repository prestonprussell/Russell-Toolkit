const vendorSelect = document.getElementById("admin-vendor");
const searchInput = document.getElementById("admin-search");
const reloadBtn = document.getElementById("admin-reload-btn");
const syncEntraBtn = document.getElementById("admin-sync-entra-btn");
const importBtn = document.getElementById("admin-import-btn");
const importFileInput = document.getElementById("admin-import-file");
const addBtn = document.getElementById("admin-add-btn");
const saveBtn = document.getElementById("admin-save-btn");
const statusText = document.getElementById("admin-status");
const usersBody = document.getElementById("admin-users-body");

let userRows = [];
let nextRowId = 1;

function setStatus(message, type = "info") {
  statusText.textContent = message;
  statusText.classList.remove("ok", "error");
  if (type === "ok") statusText.classList.add("ok");
  if (type === "error") statusText.classList.add("error");
}

function formatDate(value) {
  if (!value) return "-";
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;
  return parsed.toLocaleString();
}

function escapeHtml(value) {
  return (value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function applyFilter(rows) {
  const query = searchInput.value.trim().toLowerCase();
  if (!query) return rows;

  return rows.filter((row) =>
    [row.first_name, row.last_name, row.email, row.branch].some((value) =>
      (value || "").toLowerCase().includes(query),
    ),
  );
}

function updateRowFromInput(rowId, field, value) {
  const row = userRows.find((item) => item._id === rowId);
  if (!row) return;
  row[field] = value;
}

function renderRows() {
  usersBody.innerHTML = "";
  const visibleRows = applyFilter(userRows);

  if (!visibleRows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = '<td colspan="7">No users match your current filter.</td>';
    usersBody.appendChild(tr);
    return;
  }

  visibleRows.forEach((row) => {
    const tr = document.createElement("tr");
    const emailReadOnly = row.is_new ? "" : "readonly";
    tr.innerHTML = `
      <td><input type="text" id="first-${row._id}" value="${escapeHtml(row.first_name)}" /></td>
      <td><input type="text" id="last-${row._id}" value="${escapeHtml(row.last_name)}" /></td>
      <td><input type="text" id="email-${row._id}" value="${escapeHtml(row.email)}" ${emailReadOnly} /></td>
      <td><input type="text" id="branch-${row._id}" value="${escapeHtml(row.branch)}" /></td>
      <td>${formatDate(row.last_seen_at)}</td>
      <td>${formatDate(row.updated_at)}</td>
      <td>
        <button id="action-${row._id}" type="button" class="${row.is_new ? "btn-secondary" : "btn-danger"}">
          ${row.is_new ? "Remove" : "Deactivate"}
        </button>
      </td>
    `;
    usersBody.appendChild(tr);

    document.getElementById(`first-${row._id}`).addEventListener("input", (event) => {
      updateRowFromInput(row._id, "first_name", event.target.value);
    });
    document.getElementById(`last-${row._id}`).addEventListener("input", (event) => {
      updateRowFromInput(row._id, "last_name", event.target.value);
    });
    document.getElementById(`email-${row._id}`).addEventListener("input", (event) => {
      updateRowFromInput(row._id, "email", event.target.value);
    });
    document.getElementById(`branch-${row._id}`).addEventListener("input", (event) => {
      updateRowFromInput(row._id, "branch", event.target.value);
    });

    document.getElementById(`action-${row._id}`).addEventListener("click", async () => {
      if (row.is_new) {
        userRows = userRows.filter((item) => item._id !== row._id);
        renderRows();
        return;
      }
      await deactivateRow(row);
    });
  });
}

function updateVendorActions() {
  const isIntegricom = vendorSelect.value === "integricom";
  const isAdobe = vendorSelect.value === "adobe";
  syncEntraBtn.hidden = !isIntegricom;
  importBtn.hidden = !isAdobe;
}

async function loadUsers() {
  const vendor = vendorSelect.value;
  updateVendorActions();
  setStatus(`Loading ${vendor} users...`);
  saveBtn.disabled = true;
  addBtn.disabled = true;
  reloadBtn.disabled = true;
  syncEntraBtn.disabled = true;

  try {
    const response = await fetch(`/api/${vendor}/users?active_only=true`);
    if (!response.ok) {
      const err = await response.json();
      throw new Error(err.detail || "Failed to load users.");
    }
    const data = await response.json();
    userRows = (data.users || []).map((row) => ({
      _id: nextRowId++,
      is_new: false,
      email: row.email || "",
      first_name: row.first_name || "",
      last_name: row.last_name || "",
      branch: row.branch || "Home Office",
      last_seen_at: row.last_seen_at || "",
      updated_at: row.updated_at || "",
    }));
    renderRows();
    setStatus(`Loaded ${userRows.length} active ${vendor} users.`, "ok");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    saveBtn.disabled = false;
    addBtn.disabled = false;
    reloadBtn.disabled = false;
    syncEntraBtn.disabled = vendorSelect.value !== "integricom";
  }
}

function addUserRow() {
  userRows.unshift({
    _id: nextRowId++,
    is_new: true,
    email: "",
    first_name: "",
    last_name: "",
    branch: "Home Office",
    last_seen_at: "",
    updated_at: "",
  });
  renderRows();
}

function collectValidatedRows() {
  const cleanRows = [];
  const seenEmails = new Set();

  for (const row of userRows) {
    const email = (row.email || "").trim().toLowerCase();
    const firstName = (row.first_name || "").trim();
    const lastName = (row.last_name || "").trim();
    const branch = (row.branch || "").trim();

    if (!email) {
      throw new Error("Every row must include an email.");
    }
    if (!branch) {
      throw new Error(`Branch is required for ${email}.`);
    }
    if (seenEmails.has(email)) {
      throw new Error(`Duplicate email found: ${email}`);
    }
    seenEmails.add(email);

    cleanRows.push({
      email,
      first_name: firstName,
      last_name: lastName,
      branch,
    });
  }
  return cleanRows;
}

async function saveUsers() {
  const vendor = vendorSelect.value;
  const endpoint = vendor === "adobe" ? "/api/adobe/users/save" : "/api/integricom/users/save";
  let payload = [];
  try {
    payload = collectValidatedRows();
  } catch (error) {
    setStatus(error.message, "error");
    return;
  }

  setStatus(`Saving ${payload.length} ${vendor} users...`);
  saveBtn.disabled = true;

  try {
    const response = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!response.ok) {
      const err = await response.json();
      throw new Error(err.detail || "Failed to save users.");
    }
    const result = await response.json();
    setStatus(`Saved ${result.saved} users.`, "ok");
    await loadUsers();
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    saveBtn.disabled = false;
  }
}

async function deactivateRow(row) {
  const vendor = vendorSelect.value;
  const endpoint = vendor === "adobe" ? "/api/adobe/users/deactivate" : "/api/integricom/users/deactivate";
  const email = (row.email || "").trim().toLowerCase();
  if (!email) {
    setStatus("Cannot deactivate a row with a blank email.", "error");
    return;
  }

  if (!window.confirm(`Deactivate ${email}?`)) {
    return;
  }

  setStatus(`Deactivating ${email}...`);
  try {
    const response = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ emails: [email] }),
    });
    if (!response.ok) {
      const err = await response.json();
      throw new Error(err.detail || "Failed to deactivate user.");
    }
    userRows = userRows.filter((item) => item._id !== row._id);
    renderRows();
    setStatus(`Deactivated ${email}.`, "ok");
  } catch (error) {
    setStatus(error.message, "error");
  }
}

async function syncFromEntra() {
  if (vendorSelect.value !== "integricom") {
    setStatus("Entra sync is only available for Integricom directory.", "error");
    return;
  }

  if (!window.confirm("Sync Integricom users from Microsoft Entra now?")) {
    return;
  }

  setStatus("Syncing from Microsoft Entra...");
  syncEntraBtn.disabled = true;
  try {
    const response = await fetch("/api/integricom/sync/entra", { method: "POST" });
    if (!response.ok) {
      const err = await response.json();
      throw new Error(err.detail || "Entra sync failed.");
    }
    const result = await response.json();
    const warningSuffix = result.warnings?.length ? ` Warnings: ${result.warnings.join(" | ")}` : "";
    setStatus(
      `Entra sync complete. Synced ${result.synced} users from ${result.users_scanned} scanned.${warningSuffix}`,
      "ok",
    );
    await loadUsers();
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    syncEntraBtn.disabled = vendorSelect.value !== "integricom";
  }
}

async function importAdobeStartingList() {
  if (vendorSelect.value !== "adobe") {
    setStatus("Spreadsheet import is only available for Adobe directory.", "error");
    return;
  }

  const file = importFileInput.files && importFileInput.files[0];
  if (!file) {
    setStatus("Choose an Adobe mapping file first.", "error");
    return;
  }

  const payload = new FormData();
  payload.append("mapping_file", file);

  setStatus(`Importing ${file.name}...`);
  importBtn.disabled = true;
  try {
    const response = await fetch("/api/adobe/users/import", {
      method: "POST",
      body: payload,
    });
    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.detail || "Import failed.");
    }

    const warningSuffix = result.warnings?.length ? ` Warnings: ${result.warnings.join(" | ")}` : "";
    setStatus(`Imported ${result.imported} Adobe users from ${result.filename}.${warningSuffix}`, "ok");
    importFileInput.value = "";
    await loadUsers();
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    importBtn.disabled = false;
  }
}

function initializeResizableTables() {
  const tables = document.querySelectorAll("table.resizable-table");
  tables.forEach((table) => {
    table.querySelectorAll("thead th").forEach((headerCell) => {
      if (headerCell.dataset.resizable === "true") {
        return;
      }
      headerCell.dataset.resizable = "true";

      const resizer = document.createElement("span");
      resizer.className = "col-resizer";

      let startX = 0;
      let startWidth = 0;

      const handlePointerMove = (event) => {
        const delta = event.clientX - startX;
        const nextWidth = Math.max(110, startWidth + delta);
        headerCell.style.width = `${nextWidth}px`;
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
        headerCell.style.width = `${startWidth}px`;
        document.body.classList.add("col-resize-active");
        window.addEventListener("pointermove", handlePointerMove);
        window.addEventListener("pointerup", handlePointerUp);
      });

      headerCell.appendChild(resizer);
    });
  });
}

vendorSelect.addEventListener("change", loadUsers);
searchInput.addEventListener("input", renderRows);
reloadBtn.addEventListener("click", loadUsers);
syncEntraBtn.addEventListener("click", syncFromEntra);
importBtn.addEventListener("click", () => importFileInput.click());
importFileInput.addEventListener("change", importAdobeStartingList);
addBtn.addEventListener("click", addUserRow);
saveBtn.addEventListener("click", saveUsers);

initializeResizableTables();
updateVendorActions();
loadUsers();
