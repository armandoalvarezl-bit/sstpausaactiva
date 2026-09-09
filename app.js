const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzDIVViPUDu4KNQph9GfiaVgaBnA6fVW9AyW5-beqik554U4sGkJ3R1J_G1wefLs_LHeQ/exec";
const SHEET_IMAGE_MAX_EDGE = 900;
const SHEET_IMAGE_MAX_PIXELS = 950000;
const SHEET_IMAGE_MAX_BYTES = 1800000;
const SHEET_IMAGE_QUALITY = 0.72;

const state = {
  records: [],
  query: "",
  selectedPhotos: [],
  onlineUrl: "",
  loadingOnline: false,
  editingRecordId: null
};

const els = {
  form: document.querySelector("#recordForm"),
  scriptUrl: document.querySelector("#scriptUrl"),
  workCenter: document.querySelector("#workCenter"),
  monthInput: document.querySelector("#monthInput"),
  dateInput: document.querySelector("#dateInput"),
  photoInput: document.querySelector("#photoInput"),
  photoDrop: document.querySelector(".photo-drop"),
  photoPreview: document.querySelector("#photoPreview"),
  notesInput: document.querySelector("#notesInput"),
  saveBtn: document.querySelector("#saveBtn"),
  searchInput: document.querySelector("#searchInput"),
  calendarBoard: document.querySelector("#calendarBoard"),
  calendarCaption: document.querySelector("#calendarCaption"),
  shiftBars: document.querySelector("#shiftBars"),
  shiftGallery: document.querySelector("#shiftGallery"),
  missingAlert: document.querySelector("#missingAlert"),
  recordsTable: document.querySelector("#recordsTable tbody"),
  metricRecords: document.querySelector("#metricRecords"),
  metricCoverage: document.querySelector("#metricCoverage"),
  metricPhotos: document.querySelector("#metricPhotos"),
  metricLastSync: document.querySelector("#metricLastSync"),
  syncBtn: document.querySelector("#syncBtn"),
  exportBtn: document.querySelector("#exportBtn"),
  reloadOnlineBtn: document.querySelector("#reloadOnlineBtn"),
  toast: document.querySelector("#appToast"),
  statusModal: document.querySelector("#statusModal"),
  statusModalIcon: document.querySelector("#statusModalIcon"),
  statusModalTitle: document.querySelector("#statusModalTitle"),
  statusModalMessage: document.querySelector("#statusModalMessage"),
  statusModalDetails: document.querySelector("#statusModalDetails")
};

init();

function init() {
  const today = new Date();
  const month = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}`;
  const date = today.toISOString().slice(0, 10);

  els.monthInput.value = month;
  els.dateInput.value = date;
  els.scriptUrl.value = APPS_SCRIPT_URL;
  els.workCenter.value = "PEAJE FRAGUA";

  bindEvents();
  render();
  refreshOnlineHistory();
}

function bindEvents() {
  els.form.addEventListener("submit", handleSubmit);
  els.photoInput.addEventListener("change", handlePhotoSelect);
  els.photoPreview.addEventListener("click", handlePhotoPreviewClick);
  els.photoDrop.addEventListener("dragover", handlePhotoDragOver);
  els.photoDrop.addEventListener("dragenter", handlePhotoDragOver);
  els.photoDrop.addEventListener("dragleave", handlePhotoDragLeave);
  els.photoDrop.addEventListener("drop", handlePhotoDrop);
  els.recordsTable.addEventListener("click", handleRecordTableClick);
  els.searchInput.addEventListener("input", (event) => {
    state.query = event.target.value.trim().toLowerCase();
    renderTable();
  });
  els.monthInput.addEventListener("change", () => {
    render();
  });
  els.exportBtn.addEventListener("click", exportExcel);
  els.syncBtn.addEventListener("click", syncRecords);
  els.reloadOnlineBtn.addEventListener("click", () => refreshOnlineHistory({ notify: true }));
}

async function refreshOnlineHistory(options = {}) {
  const { notify = false } = options;
  state.loadingOnline = true;
  els.reloadOnlineBtn.disabled = true;
  try {
    const result = await getOnlineRecords();
    if (!result.ok) {
      throw new Error(result.error || "No se pudo leer el historial en linea.");
    }
    state.onlineUrl = result.spreadsheetUrl || "";
    replaceWithOnlineRecords(result.records || []);
    render();
    if (notify || (result.records && result.records.length)) {
      showToast("Historial cargado desde el Excel en linea.", "success");
    }
  } catch (error) {
    console.warn(error);
    showStatus("error", "No se pudo cargar el Excel en linea", "Revisa que el Apps Script este publicado y que el enlace sea el correcto.", error.message);
  } finally {
    state.loadingOnline = false;
    els.reloadOnlineBtn.disabled = false;
  }
}

async function getOnlineRecords() {
  try {
    const response = await fetch(`${APPS_SCRIPT_URL}?mode=records`, { method: "GET", mode: "cors" });
    return response.json();
  } catch {
    return getOnlineRecordsJsonp();
  }
}

function getOnlineRecordsJsonp() {
  return new Promise((resolve, reject) => {
    const callbackName = `pausasRecords_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    const script = document.createElement("script");
    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error("No se pudo cargar el historial en linea."));
    }, 12000);

    window[callbackName] = (data) => {
      cleanup();
      resolve(data);
    };

    function cleanup() {
      clearTimeout(timeout);
      script.remove();
      delete window[callbackName];
    }

    script.onerror = () => {
      cleanup();
      reject(new Error("No se pudo conectar con el Excel en linea."));
    };
    script.src = `${APPS_SCRIPT_URL}?mode=records&callback=${callbackName}`;
    document.body.appendChild(script);
  });
}

function replaceWithOnlineRecords(records) {
  const pending = state.records.filter((record) => !record.syncedAt && hasLocalPhotos(record));
  const online = Array.isArray(records) ? records.map(normalizeRecord) : [];
  const pendingKeys = new Set(pending.map(recordKey));
  state.records = [
    ...pending,
    ...online.filter((record) => !pendingKeys.has(recordKey(record)))
  ];
}

async function handlePhotoSelect() {
  const files = Array.from(els.photoInput.files || []);
  await addPhotos(files);
}

function handlePhotoDragOver(event) {
  event.preventDefault();
  event.stopPropagation();
  els.photoDrop.classList.add("is-dragging");
}

function handlePhotoDragLeave(event) {
  event.preventDefault();
  event.stopPropagation();
  els.photoDrop.classList.remove("is-dragging");
}

async function handlePhotoDrop(event) {
  event.preventDefault();
  event.stopPropagation();
  els.photoDrop.classList.remove("is-dragging");

  const files = Array.from(event.dataTransfer?.files || []);
  await addPhotos(files);
}

async function addPhotos(files) {
  const validFiles = files.filter((file) => file && file.type && file.type.startsWith("image/"));
  if (!validFiles.length) {
    return;
  }

  const newPhotos = await Promise.all(validFiles.map(fileToPhoto));
  const merged = uniquePhotos([...state.selectedPhotos, ...newPhotos]);
  const duplicates = [...state.selectedPhotos, ...newPhotos].length - merged.length;

  state.selectedPhotos = merged;
  syncPhotoInputFiles(validFiles);

  if (duplicates > 0) {
    showToast("Se omitieron fotos duplicadas del mismo turno y fecha.", "warning");
  }

  renderPhotoPreview();
}

function syncPhotoInputFiles(newFiles) {
  if (typeof DataTransfer === "undefined") {
    els.photoInput.value = "";
    return;
  }

  const dataTransfer = new DataTransfer();
  const currentFiles = Array.from(els.photoInput.files || []);
  currentFiles.forEach((file) => dataTransfer.items.add(file));
  newFiles.forEach((file) => dataTransfer.items.add(file));
  els.photoInput.files = dataTransfer.files;
}

async function handleSubmit(event) {
  event.preventDefault();

  const date = els.dateInput.value;
  const shift = document.querySelector("input[name='shift']:checked").value;
  const existingSignatures = new Set(
    state.records
      .filter((record) => (
        record.date === date &&
        record.shift === shift &&
        record.id !== state.editingRecordId
      ))
      .flatMap((record) => getPhotos(record).map((photo) => photoSignature(photo)))
  );

  const filteredPhotos = uniquePhotos(
    state.selectedPhotos.filter((photo) => !existingSignatures.has(photoSignature(photo)))
  );
  const duplicateCount = state.selectedPhotos.length - filteredPhotos.length;

  if (!filteredPhotos.length) {
    const message = duplicateCount > 0
      ? "Ya existe esa foto para este día y turno. Elige otra o cambia la fecha."
      : "Adjunta al menos una foto para poder guardar el turno en el Excel en linea.";
    showStatus("warning", "Falta evidencia fotografica", message);
    return;
  }

  const payload = {
    date,
    month: els.monthInput.value || date.slice(0, 7),
    shift,
    workCenter: els.workCenter.value.trim() || "PEAJE FRAGUA",
    photos: filteredPhotos.map((photo, index) => ({
      ...photo,
      name: buildPhotoName(photo.name, date, shift, index + 1)
    })),
    notes: els.notesInput.value.trim(),
    syncedAt: ""
  };

  if (duplicateCount > 0) {
    showToast("Se omitieron fotos duplicadas para este día y turno.", "warning");
  }

  let recordsToSync = [];

  if (state.editingRecordId) {
    const targetId = state.editingRecordId;
    state.records = state.records.map((record) =>
      record.id === targetId ? { ...record, ...payload, id: record.id } : record
    );
    recordsToSync = state.records.filter((record) => record.id === targetId);
    showToast("Registro actualizado.", "success");
  } else {
    const newRecord = { id: crypto.randomUUID(), ...payload };
    state.records.unshift(newRecord);
    recordsToSync = [newRecord];
    showToast("Registro guardado.", "success");
  }

  state.editingRecordId = null;
  state.selectedPhotos = [];
  els.photoInput.value = "";
  els.notesInput.value = "";
  render();
  renderPhotoPreview();
  await syncRecords(recordsToSync);
}

function render() {
  renderMetrics();
  renderMissingAlert();
  renderCalendar();
  renderShiftBars();
  renderShiftGallery();
  renderTable();
}

function getMonthRecords() {
  return state.records.filter((record) => getRecordMonth(record) === els.monthInput.value);
}

function renderMetrics() {
  const monthRecords = getMonthRecords();
  const days = daysInMonth(els.monthInput.value);
  const completedSlots = new Set(monthRecords.map((record) => `${record.date}-${record.shift}`));
  const coverage = days ? Math.round((completedSlots.size / (days * 3)) * 100) : 0;
  const photos = monthRecords.reduce((total, record) => total + getPhotos(record).length, 0);
  const lastSync = state.records.find((record) => record.syncedAt)?.syncedAt;

  els.metricRecords.textContent = monthRecords.length;
  els.metricCoverage.textContent = `${coverage}%`;
  els.metricPhotos.textContent = photos;
  els.metricLastSync.textContent = lastSync ? formatDateTime(lastSync) : "Excel en linea";
}

function renderCalendar() {
  const month = els.monthInput.value;
  const totalDays = daysInMonth(month);
  const monthRecords = getMonthRecords();
  const formatter = new Intl.DateTimeFormat("es-CO", { month: "long", year: "numeric" });
  const captionDate = new Date(`${month}-01T00:00:00`);
  els.calendarCaption.textContent = formatter.format(captionDate);
  els.calendarBoard.innerHTML = "";

  for (let day = 1; day <= totalDays; day += 1) {
    const date = `${month}-${String(day).padStart(2, "0")}`;
    const dayRecords = monthRecords.filter((record) => record.date === date);
    const shifts = new Set(dayRecords.map((record) => record.shift));
    const photos = dayRecords.reduce((total, record) => total + getPhotos(record).length, 0);
    const cell = document.createElement("button");
    cell.type = "button";
    cell.className = `day-cell ${shifts.size === 3 ? "complete" : shifts.size > 0 ? "partial" : ""}`;
    cell.innerHTML = `<strong>${day}</strong><span>${shifts.size}/3 turnos</span><em>${photos} fotos</em>`;
    cell.addEventListener("click", () => {
      els.dateInput.value = date;
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
    els.calendarBoard.appendChild(cell);
  }
}

function renderShiftBars() {
  const monthRecords = getMonthRecords();
  const totalDays = daysInMonth(els.monthInput.value);
  els.shiftBars.innerHTML = "";

  ["A", "B", "C"].forEach((shift) => {
    const records = monthRecords.filter((record) => record.shift === shift);
    const dates = new Set(records.map((record) => record.date));
    const photos = records.reduce((total, record) => total + getPhotos(record).length, 0);
    const percent = totalDays ? Math.round((dates.size / totalDays) * 100) : 0;
    const row = document.createElement("div");
    row.className = "shift-row";
    row.innerHTML = `
      <div class="shift-label"><span>Turno ${shift}</span><span>${percent}% - ${photos} fotos</span></div>
      <div class="progress" role="progressbar" aria-valuenow="${percent}" aria-valuemin="0" aria-valuemax="100">
        <div class="progress-bar" style="width: ${percent}%"></div>
      </div>
    `;
    els.shiftBars.appendChild(row);
  });
}

function renderMissingAlert() {
  const date = els.dateInput.value;
  const dateRecords = state.records.filter((record) => record.date === date);
  const missing = ["A", "B", "C"].filter((shift) => {
    return !dateRecords.some((record) => record.shift === shift && getPhotos(record).length > 0);
  });

  if (!missing.length) {
    els.missingAlert.className = "missing-alert ok";
    els.missingAlert.innerHTML = `<i class="bi bi-check-circle"></i><span>${escapeHtml(date)} tiene fotos en los turnos A, B y C.</span>`;
    return;
  }

  els.missingAlert.className = "missing-alert warn";
  els.missingAlert.innerHTML = `<i class="bi bi-exclamation-triangle"></i><span>Faltan fotos para ${escapeHtml(date)} en turno ${missing.join(", turno ")}.</span>`;
}

function renderShiftGallery() {
  const monthRecords = getMonthRecords();
  els.shiftGallery.innerHTML = ["A", "B", "C"].map((shift) => {
    const records = monthRecords.filter((record) => record.shift === shift && getPhotos(record).length);
    const photos = records.flatMap((record) => getPhotos(record).map((photo, index) => ({ ...photo, date: record.date, index })));
    const body = photos.length ? photos.map((photo) => {
      const image = photo.dataUrl
        ? `<img src="${photo.dataUrl}" alt="${escapeHtml(photo.name)}">`
        : getPhotoDisplaySrc(photo)
          ? `<img src="${getPhotoDisplaySrc(photo)}" alt="${escapeHtml(photo.name)}">`
          : `<span class="gallery-file"><i class="bi bi-image"></i></span>`;
      const link = photo.url ? `<a href="${photo.url}" target="_blank" rel="noopener">Abrir</a>` : `<span>Pendiente</span>`;
      return `
        <figure>
          ${image}
          <figcaption>
            <strong>${escapeHtml(photo.date)}</strong>
            ${link}
          </figcaption>
        </figure>
      `;
    }).join("") : `<div class="empty-gallery">Sin fotos registradas</div>`;

    return `
      <article class="shift-gallery-card">
        <header><span>Turno ${shift}</span><strong>${photos.length}</strong></header>
        <div class="gallery-grid">${body}</div>
      </article>
    `;
  }).join("");
}

function renderTable() {
  const rows = state.records
    .filter((record) => !state.query || Object.values({
      date: record.date,
      month: getRecordMonth(record),
      shift: record.shift,
      workCenter: record.workCenter,
      notes: record.notes,
      photos: getPhotos(record).map((photo) => photo.name).join(" ")
    }).join(" ").toLowerCase().includes(state.query))
    .sort((a, b) => b.date.localeCompare(a.date) || a.shift.localeCompare(b.shift));

  els.recordsTable.innerHTML = rows.map((record) => `
    <tr>
      <td>${escapeHtml(record.date)}</td>
      <td>${escapeHtml(getRecordMonth(record))}</td>
      <td><span class="badge text-bg-secondary">Turno ${escapeHtml(record.shift)}</span></td>
      <td>${escapeHtml(record.workCenter)}</td>
      <td>${renderPhotoThumbs(record)}</td>
      <td>${renderPhotoLinks(record)}</td>
      <td>${escapeHtml(record.notes || "-")}</td>
      <td>
        <div class="table-actions">
          <button type="button" class="btn btn-sm btn-outline-primary" data-action="edit-record" data-id="${record.id}">Editar</button>
          <button type="button" class="btn btn-sm btn-outline-danger" data-action="delete-record" data-id="${record.id}">Eliminar</button>
        </div>
      </td>
    </tr>
  `).join("");
}

function renderPhotoPreview() {
  if (!state.selectedPhotos.length) {
    els.photoPreview.innerHTML = "";
    return;
  }

  els.photoPreview.innerHTML = state.selectedPhotos.map((photo, index) => `
    <figure>
      <img src="${photo.dataUrl || getPhotoDisplaySrc(photo)}" alt="${escapeHtml(photo.name)}">
      <figcaption>${escapeHtml(photo.name)}</figcaption>
      <button type="button" class="btn btn-sm btn-outline-danger remove-photo-btn" data-remove-photo="${index}">Quitar</button>
    </figure>
  `).join("");
}

function renderPhotoThumbs(record) {
  const photos = getPhotos(record);
  if (!photos.length) return "-";
  return `<div class="thumb-strip">${photos.map((photo) => {
    if (photo.dataUrl) {
      return `<img src="${photo.dataUrl}" alt="${escapeHtml(photo.name)}">`;
    }
    if (getPhotoDisplaySrc(photo)) {
      return `<img src="${getPhotoDisplaySrc(photo)}" alt="${escapeHtml(photo.name)}">`;
    }
    return `<span class="thumb-empty"><i class="bi bi-image"></i></span>`;
  }).join("")}</div>`;
}

function handleRecordTableClick(event) {
  const button = event.target.closest("button[data-action]");
  if (!button) return;

  const { action, id } = button.dataset;
  if (!id) return;

  if (action === "delete-record") {
    const record = state.records.find((item) => item.id === id);
    if (!record) return;
    if (!window.confirm("¿Seguro que deseas eliminar este registro y sus fotos?")) return;
    state.records = state.records.filter((item) => item.id !== id);
    render();
    return;
  }

  if (action === "edit-record") {
    const record = state.records.find((item) => item.id === id);
    if (!record) return;
    state.editingRecordId = record.id;
    els.dateInput.value = record.date;
    els.workCenter.value = record.workCenter || "PEAJE FRAGUA";
    els.notesInput.value = record.notes || "";
    state.selectedPhotos = getPhotos(record).map((photo) => ({
      ...photo,
      dataUrl: photo.dataUrl || getPhotoDisplaySrc(photo),
      name: photo.name || `foto-${record.date}`
    }));

    const shiftRadio = document.querySelector(`input[name='shift'][value='${record.shift}']`);
    if (shiftRadio) shiftRadio.checked = true;

    renderPhotoPreview();
    showToast("Registro cargado para editar. Guarda para aplicar los cambios.", "info");
  }
}

function handlePhotoPreviewClick(event) {
  const button = event.target.closest("button[data-remove-photo]");
  if (!button) return;

  const index = Number(button.dataset.removePhoto);
  if (Number.isNaN(index)) return;

  state.selectedPhotos = state.selectedPhotos.filter((_, photoIndex) => photoIndex !== index);
  renderPhotoPreview();
}

function renderPhotoLinks(record) {
  const photos = getPhotos(record);
  if (!photos.length) return "-";
  const links = photos.map((photo, index) => {
    if (photo.url) {
      return `<a href="${photo.url}" target="_blank" rel="noopener">Foto ${index + 1}</a>`;
    }
    return `<span>${escapeHtml(photo.name)}</span>`;
  });
  return `<div class="photo-links">${links.join("")}</div>`;
}

async function syncRecords(recordsToSend) {
  const scriptUrl = APPS_SCRIPT_URL;
  const pending = Array.isArray(recordsToSend) ? recordsToSend : state.records.filter((record) => {
    return !record.syncedAt || getPhotos(record).some((photo) => photo.dataUrl);
  });
  if (!pending.length) {
    showStatus("info", "Sin registros pendientes", "No hay fotos nuevas pendientes por enviar al Excel en linea.");
    return;
  }

  els.syncBtn.disabled = true;
  els.saveBtn.disabled = true;
  try {
    const result = await postToAppsScript(scriptUrl, pending);
    if (!result.opaque) {
      markAsSynced(pending, result.photoLinksById || {});
      await refreshOnlineHistory();
      showStatus(
        "success",
        "Evidencia registrada",
        buildSuccessMessage(pending),
        "El historial se actualizo desde la base en linea. Puedes revisar la hoja desde el boton de descarga o continuar registrando otro turno."
      );
    } else {
      showStatus(
        "warning",
        "Solicitud enviada sin confirmacion",
        "El navegador envio los datos al Apps Script, pero Google no devolvio confirmacion por seguridad del iframe. Recarga el Excel para verificar el registro.",
        "Los datos no quedaron guardados localmente en este navegador."
      );
    }
    render();
  } catch (error) {
    showStatus("error", "No se guardo en el Excel", "El registro quedo solo en memoria mientras esta pagina siga abierta. Corrige el error y vuelve a sincronizar.", error.message);
  } finally {
    els.syncBtn.disabled = false;
    els.saveBtn.disabled = false;
  }
}

async function postToAppsScript(scriptUrl, records) {
  const payload = JSON.stringify({ action: "saveRecords", records });
  try {
    const response = await fetch(scriptUrl, {
      method: "POST",
      mode: "cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: payload
    });
    const result = await response.json();
    if (!result.ok) {
      const appError = new Error(result.error || "No se pudo sincronizar.");
      appError.appError = true;
      throw appError;
    }
    return result;
  } catch (error) {
    if (error.appError) {
      throw error;
    }
    await submitViaHiddenForm(scriptUrl, payload);
    return { ok: true, opaque: true, photoLinksById: {} };
  }
}

function submitViaHiddenForm(scriptUrl, payload) {
  return new Promise((resolve) => {
    const iframeName = `syncTarget_${Date.now()}`;
    const iframe = document.createElement("iframe");
    iframe.name = iframeName;
    iframe.style.display = "none";

    const form = document.createElement("form");
    form.method = "POST";
    form.action = scriptUrl;
    form.target = iframeName;
    form.style.display = "none";

    const input = document.createElement("input");
    input.type = "hidden";
    input.name = "payload";
    input.value = payload;

    form.appendChild(input);
    document.body.appendChild(iframe);
    document.body.appendChild(form);
    iframe.addEventListener("load", () => {
      setTimeout(() => {
        form.remove();
        iframe.remove();
        resolve();
      }, 300);
    }, { once: true });
    form.submit();
    setTimeout(resolve, 4000);
  });
}

function markAsSynced(records, photoLinksById) {
  const now = new Date().toISOString();
  records.forEach((record) => {
    record.syncedAt = now;
    const links = photoLinksById[record.id] || [];
    getPhotos(record).forEach((photo, index) => {
      if (links[index]) {
        photo.url = links[index];
      }
    });
  });
}

function buildSuccessMessage(records) {
  if (records.length === 1) {
    const record = records[0];
    const photos = getPhotos(record).length;
    return `La evidencia del ${formatDisplayDate(record.date)} para el turno ${record.shift} quedo guardada correctamente con ${photos} foto(s).`;
  }

  const photos = records.reduce((total, record) => total + getPhotos(record).length, 0);
  return `Se guardaron correctamente ${records.length} registros con ${photos} foto(s) en la base en linea.`;
}

async function exportExcel() {
  if (!window.ExcelJS) {
    showStatus("error", "No se pudo preparar el Excel", "No cargo la libreria necesaria para generar el archivo descargable.", "Revisa la conexion a internet e intenta de nuevo.");
    return;
  }

  try {
    await hydratePhotosForExport(state.records);
    const months = [...new Set(state.records.map(getRecordMonth).filter(Boolean))]
      .sort((a, b) => a.localeCompare(b));
    const exportMonths = months.length ? months : [els.monthInput.value];
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Pausas Activas SST";
    workbook.created = new Date();
    const logoDataUrl = await getLogoDataUrl();

    exportMonths.forEach((month) => {
      const records = state.records.filter((record) => getRecordMonth(record) === month);
      buildMonthWorksheet(workbook, month, records, logoDataUrl);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "REGISTRO-PAUSAS-ACTIVAS-HISTORIAL.xlsx";
    link.click();
    URL.revokeObjectURL(url);
    showToast("Excel descargado con la informacion disponible.", "success");
  } catch (error) {
    showStatus("error", "No se pudo descargar el Excel", "El sistema encontro un dato con formato incorrecto al crear el archivo.", error.message);
  }
}

function buildMonthWorksheet(workbook, month, records, logoDataUrl) {
  const safeMonth = normalizeMonth(month) || els.monthInput.value;
  const sheetName = sanitizeWorksheetName(safeMonth ? safeMonth.replace("-", " ") : "Historial");
  const worksheet = workbook.addWorksheet(sheetName.slice(0, 31), {
    pageSetup: { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
    views: [{ showGridLines: false }]
  });
  const totalDays = daysInMonth(safeMonth);
  const captionDate = new Date(`${safeMonth}-01T00:00:00`);
  const monthLabel = new Intl.DateTimeFormat("es-CO", { month: "long", year: "numeric" })
    .format(captionDate)
    .toUpperCase();
  const center = els.workCenter.value.trim() || "PEAJE FRAGUA";

  worksheet.columns = [
    { width: 18 },
    { width: 38 },
    { width: 38 },
    { width: 38 }
  ];

  worksheet.mergeCells("A2:A4");
  worksheet.mergeCells("B2:D4");
  worksheet.mergeCells("B5:D6");
  worksheet.mergeCells("A7:D9");
  worksheet.mergeCells("A10:D12");

  worksheet.getCell("A2").value = "";
  worksheet.getCell("B2").value = "FORMATO REGISTRO FOTOGRAFICO DE PAUSAS ACTIVAS\nZIMA SEGURIDAD LTDA";
  worksheet.getCell("A5").value = "ESTACION DE RECAUDO";
  worksheet.getCell("B5").value = `CENTRO DE TRABAJO: ${center}`;
  worksheet.getCell("A7").value = `MES: ${monthLabel}`;
  worksheet.getCell("A10").value = "ANEXO FOTOGRAFICO";
  worksheet.getRow(13).values = ["FECHA", "TURNO A", "TURNO B", "TURNO C"];

  [2, 3, 4].forEach((rowNumber) => {
    worksheet.getRow(rowNumber).height = 22;
  });
  [7, 8, 9, 10, 11, 12].forEach((rowNumber) => {
    worksheet.getRow(rowNumber).height = 20;
  });
  worksheet.getRow(13).height = 34;

  styleHeader(worksheet);
  addLogoToWorksheet(workbook, worksheet, logoDataUrl);

  let rowNumber = 14;
  for (let day = 1; day <= totalDays; day += 1) {
    const date = `${safeMonth}-${String(day).padStart(2, "0")}`;
    const row = worksheet.getRow(rowNumber);
    row.height = getDayRowHeight(records, date);
    row.getCell(1).value = formatDateForExcel(date);
    row.getCell(1).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    row.getCell(1).font = { bold: true, name: "Calibri", size: 11 };

    ["A", "B", "C"].forEach((shift, index) => {
      fillShiftCell(workbook, worksheet, records, date, shift, rowNumber, index + 2);
    });
    rowNumber += 1;
  }

  worksheet.eachRow((row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FF1C2630" } },
        left: { style: "thin", color: { argb: "FF1C2630" } },
        bottom: { style: "thin", color: { argb: "FF1C2630" } },
        right: { style: "thin", color: { argb: "FF1C2630" } }
      };
    });
  });
}

function addLogoToWorksheet(workbook, worksheet, logoDataUrl) {
  if (!logoDataUrl) return;
  const imageId = workbook.addImage({
    base64: logoDataUrl,
    extension: "png"
  });
  worksheet.addImage(imageId, {
    tl: { col: 0.12, row: 1.18 },
    ext: { width: 76, height: 66 },
    editAs: "oneCell"
  });
}

function styleHeader(worksheet) {
  const headerGray = "FFD9D9D9";
  const lightGray = "FFEDEDED";

  worksheet.getCell("B2").alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  worksheet.getCell("B2").font = { bold: true, color: { argb: "FF000000" }, size: 14 };
  worksheet.getCell("B2").fill = solidFill(headerGray);
  worksheet.getCell("A5").font = { bold: true, color: { argb: "FF000000" } };
  worksheet.getCell("B5").font = { bold: true, color: { argb: "FF000000" } };
  worksheet.getCell("A5").fill = solidFill(lightGray);
  worksheet.getCell("B5").fill = solidFill(lightGray);
  worksheet.getCell("A7").font = { bold: true, color: { argb: "FF000000" } };
  worksheet.getCell("A7").fill = solidFill("FFFFFFFF");
  worksheet.getCell("A10").alignment = { vertical: "middle", horizontal: "center" };
  worksheet.getCell("A10").font = { bold: true, color: { argb: "FF000000" } };
  worksheet.getCell("A10").fill = solidFill(headerGray);

  for (let column = 1; column <= 4; column += 1) {
    const cell = worksheet.getRow(13).getCell(column);
    cell.font = { bold: true, color: { argb: "FF000000" } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.fill = solidFill(lightGray);
  }
}

function fillShiftCell(workbook, worksheet, records, date, shift, rowNumber, columnNumber) {
  const shiftRecords = records.filter((record) => record.date === date && record.shift === shift);
  const cell = worksheet.getRow(rowNumber).getCell(columnNumber);
  cell.alignment = { vertical: "top", horizontal: "center", wrapText: true };

  if (!shiftRecords.length) {
    cell.value = "";
    return;
  }

  const links = [];
  const notes = [];
  let imageOffset = 0;

  shiftRecords.forEach((record) => {
    if (record.notes) notes.push(record.notes);
    getPhotos(record).forEach((photo, index) => {
      if (photo.dataUrl) {
        addPhotoToWorksheet(workbook, worksheet, photo, rowNumber, columnNumber, imageOffset);
        imageOffset += 1;
      } else if (photo.url) {
        links.push({ text: `Foto ${index + 1}`, hyperlink: photo.url });
      }
    });
  });

  if (links.length) {
    cell.value = {
      richText: links.map((link) => ({ text: `${link.text}: ${link.hyperlink}\n` }))
    };
  }

  if (notes.length) {
    cell.value = `${cell.value?.richText?.map((part) => part.text).join("") || ""}\n${notes.join("\n")}`;
  }
}

function addPhotoToWorksheet(workbook, worksheet, photo, rowNumber, columnNumber, index) {
  const imageId = workbook.addImage({
    base64: photo.dataUrl,
    extension: "jpeg"
  });
  const size = fitImage(photo.width || 1280, photo.height || 960, 230, 118);
  worksheet.addImage(imageId, {
    tl: { col: columnNumber - 1 + 0.08, row: rowNumber - 1 + 0.12 + index * 1.65 },
    ext: { width: size.width, height: size.height },
    editAs: "oneCell"
  });
}

async function hydratePhotosForExport(records) {
  const photos = records.flatMap((record) => getPhotos(record))
    .filter((photo) => !photo.dataUrl && photo.url);

  for (const photo of photos) {
    try {
      const result = await getOnlinePhotoData(photo.url);
      if (result.ok && result.dataUrl) {
        photo.dataUrl = result.dataUrl;
        photo.type = result.type || photo.type || "image/jpeg";
        photo.size = result.size || photo.size || estimateDataUrlBytes(result.dataUrl);
      }
    } catch (error) {
      console.warn("No se pudo convertir foto para Excel", error);
    }
  }
}

function getOnlinePhotoData(url) {
  return new Promise((resolve, reject) => {
    const id = extractDriveFileId(url);
    if (!id) {
      reject(new Error("No se encontro el id de Drive de una foto."));
      return;
    }

    const callbackName = `pausasPhoto_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    const script = document.createElement("script");
    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error("No se pudo leer una foto desde el Apps Script."));
    }, 18000);

    window[callbackName] = (data) => {
      cleanup();
      if (!data || !data.ok) {
        reject(new Error(data?.error || "No se pudo leer una foto desde Drive."));
        return;
      }
      resolve(data);
    };

    function cleanup() {
      clearTimeout(timeout);
      script.remove();
      delete window[callbackName];
    }

    script.onerror = () => {
      cleanup();
      reject(new Error("No se pudo conectar con la foto en linea."));
    };
    script.src = `${APPS_SCRIPT_URL}?mode=photo&id=${encodeURIComponent(id)}&callback=${callbackName}`;
    document.body.appendChild(script);
  });
}

function fitImage(width, height, maxWidth, maxHeight) {
  const ratio = Math.min(maxWidth / width, maxHeight / height, 1);
  return {
    width: Math.round(width * ratio),
    height: Math.round(height * ratio)
  };
}

function getDayRowHeight(records, date) {
  const maxPhotos = Math.max(1, ...["A", "B", "C"].map((shift) => {
    return records
      .filter((record) => record.date === date && record.shift === shift)
      .reduce((total, record) => total + getPhotos(record).length, 0);
  }));
  return Math.max(118, maxPhotos * 94);
}

function solidFill(argb) {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function formatDateForExcel(date) {
  const [year, month, day] = date.split("-");
  return `${Number(day)}/${Number(month)}/${year}`;
}

function recordKey(record) {
  return [
    normalizeDateValue(record.date) || "",
    getRecordMonth(record) || "",
    record.shift || "",
    record.workCenter || "",
    getPhotos(record).map((photo) => photo.url || photo.name).join("|")
  ].join("::");
}

function normalizeRecord(record) {
  const date = normalizeDateValue(record.date);
  return {
    ...record,
    date,
    month: normalizeMonth(record.month) || normalizeMonth(date),
    photos: getPhotos(record)
  };
}

function getRecordMonth(record) {
  return normalizeMonth(record.month) || normalizeMonth(record.date);
}

function normalizeMonth(value) {
  const text = String(value || "").trim();
  if (!text) return "";

  const isoMatch = text.match(/^(\d{4})-(\d{2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2]}`;
  }

  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (slashMatch) {
    return `${slashMatch[3]}-${String(Number(slashMatch[2])).padStart(2, "0")}`;
  }

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}`;
  }

  return "";
}

function normalizeDateValue(value) {
  const text = String(value || "").trim();
  if (!text) return "";

  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  }

  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (slashMatch) {
    return `${slashMatch[3]}-${String(Number(slashMatch[2])).padStart(2, "0")}-${String(Number(slashMatch[1])).padStart(2, "0")}`;
  }

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10);
  }

  return text.slice(0, 10);
}

function sanitizeWorksheetName(name) {
  const cleanName = String(name || "Historial")
    .replace(/[\\*?:/\[\]]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return cleanName || "Historial";
}

function getPhotoDisplaySrc(photo) {
  if (photo.thumbnailUrl) return photo.thumbnailUrl;
  if (!photo.url) return "";
  const id = extractDriveFileId(photo.url);
  return id ? `https://drive.google.com/thumbnail?id=${encodeURIComponent(id)}&sz=w800` : photo.url;
}

function extractDriveFileId(url) {
  const text = String(url || "");
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /\/d\/([a-zA-Z0-9_-]+)/
  ];
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) return match[1];
  }
  return "";
}

function hasLocalPhotos(record) {
  return getPhotos(record).some((photo) => photo.dataUrl);
}

function fileToPhoto(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      resizeImage(reader.result, SHEET_IMAGE_MAX_EDGE, SHEET_IMAGE_QUALITY)
        .then((image) => resolve({
          name: file.name,
          type: "image/jpeg",
          size: estimateDataUrlBytes(image.dataUrl),
          width: image.width,
          height: image.height,
          dataUrl: image.dataUrl,
          url: ""
        }))
        .catch(reject);
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function resizeImage(dataUrl, maxSize, quality) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => {
      let targetMaxSize = maxSize;
      let targetQuality = quality;

      for (let attempt = 0; attempt < 8; attempt += 1) {
        const edgeScale = targetMaxSize / Math.max(image.width, image.height);
        const pixelScale = Math.sqrt(SHEET_IMAGE_MAX_PIXELS / (image.width * image.height));
        const scale = Math.min(1, edgeScale, pixelScale);
        const width = Math.max(1, Math.round(image.width * scale));
        const height = Math.max(1, Math.round(image.height * scale));
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const context = canvas.getContext("2d");
        context.drawImage(image, 0, 0, width, height);

        const resizedDataUrl = canvas.toDataURL("image/jpeg", targetQuality);
        if (estimateDataUrlBytes(resizedDataUrl) <= SHEET_IMAGE_MAX_BYTES && width * height <= SHEET_IMAGE_MAX_PIXELS) {
          resolve({
            dataUrl: resizedDataUrl,
            width,
            height
          });
          return;
        }

        targetMaxSize = Math.max(520, Math.round(targetMaxSize * 0.84));
        targetQuality = Math.max(0.52, targetQuality - 0.06);
      }

      reject(new Error("La foto es demasiado pesada para insertarla en Google Sheets. Intenta con una imagen mas liviana."));
    };
    image.onerror = () => reject(new Error("No se pudo leer una de las fotos."));
    image.src = dataUrl;
  });
}

function estimateDataUrlBytes(dataUrl) {
  const base64 = String(dataUrl || "").split(",")[1] || "";
  return Math.ceil((base64.length * 3) / 4);
}

function getLogoDataUrl() {
  return new Promise((resolve) => {
    const image = new Image();
    image.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = image.naturalWidth;
        canvas.height = image.naturalHeight;
        const context = canvas.getContext("2d");
        context.drawImage(image, 0, 0);
        resolve(canvas.toDataURL("image/png"));
      } catch {
        resolve("");
      }
    };
    image.onerror = () => resolve("");
    image.src = "img/logo-zima.png";
  });
}

function buildPhotoName(originalName, date, shift, index) {
  return `pausa-activa_${date}_turno-${shift}_${String(index).padStart(2, "0")}.jpg`;
}

function uniquePhotos(photos) {
  const seen = new Set();
  return photos.filter((photo) => {
    const signature = photoSignature(photo);
    if (!signature || seen.has(signature)) {
      return false;
    }
    seen.add(signature);
    return true;
  });
}

function photoSignature(photo) {
  const raw = photo?.dataUrl || photo?.url || photo?.name || "";
  return String(raw || "").trim();
}

function getPhotos(record) {
  return Array.isArray(record.photos) ? record.photos : [];
}

function daysInMonth(month) {
  const safeMonth = normalizeMonth(month);
  if (!safeMonth) return 0;
  const [year, monthNumber] = safeMonth.split("-").map(Number);
  return new Date(year, monthNumber, 0).getDate();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#039;");
}

function formatDateTime(value) {
  return new Intl.DateTimeFormat("es-CO", {
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  }).format(new Date(value));
}

function formatDisplayDate(value) {
  const date = normalizeDateValue(value);
  if (!date) return "la fecha seleccionada";
  const [year, month, day] = date.split("-");
  return `${Number(day)}/${Number(month)}/${year}`;
}

function showToast(message, type = "info") {
  els.toast.classList.remove("toast-success", "toast-error", "toast-warning", "toast-info");
  els.toast.classList.add(`toast-${type}`);
  els.toast.querySelector(".toast-body").textContent = message;
  bootstrap.Toast.getOrCreateInstance(els.toast).show();
}

function showStatus(type, title, message, details = "") {
  const icons = {
    success: "bi-check-circle",
    error: "bi-x-circle",
    warning: "bi-exclamation-triangle",
    info: "bi-info-circle"
  };
  els.statusModal.classList.remove("status-success", "status-error", "status-warning", "status-info");
  els.statusModal.classList.add(`status-${type}`);
  els.statusModalIcon.innerHTML = `<i class="bi ${icons[type] || icons.info}"></i>`;
  els.statusModalTitle.textContent = title;
  els.statusModalMessage.textContent = message;
  els.statusModalDetails.textContent = details || "";
  els.statusModalDetails.hidden = !details;
  bootstrap.Modal.getOrCreateInstance(els.statusModal).show();
}
