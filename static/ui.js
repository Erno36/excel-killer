"use strict";

function updateConditionalFields() {
  document.querySelectorAll("[data-show-when]").forEach((container) => {
    const [controlId, valueList] = container.dataset.showWhen.split(":");
    const control = document.getElementById(controlId);
    const value = control?.value ?? "";
    const expectedValues = (valueList ?? "").split(",");
    const visible = expectedValues.includes(value);
    container.hidden = !visible;
    container.querySelectorAll("input, select, textarea").forEach((field) => {
      field.disabled = !visible;
    });
  });
}

function updateColumnChoices() {
  const fileSelector = document.getElementById("file_index");
  const columnSelector = document.getElementById("column");
  if (!fileSelector || !columnSelector) return;

  const selectedFile = fileSelector.value;
  for (const option of columnSelector.options) {
    const [fileIndex] = option.value.split("::");
    const unavailable = option.value !== "" && fileIndex !== selectedFile;
    option.hidden = unavailable;
    option.disabled = unavailable;
  }
  if (columnSelector.selectedOptions[0]?.hidden) {
    const firstVisible = [...columnSelector.options].find((option) => !option.hidden && option.value);
    if (firstVisible) columnSelector.value = firstVisible.value;
  }
}

function updateRecipeActions() {
  const selector = document.querySelector("[data-recipe-control]");
  const automaticButton = document.getElementById("automatic-recipe-button");
  if (selector && automaticButton) {
    automaticButton.hidden = !selector.value;
    automaticButton.disabled = !selector.value;
  }
}

function validateCrossKeyPairs(form) {
  let valid = true;
  const steps = [...form.querySelectorAll("fieldset > legend")].length;
  for (let step = 0; step < steps; step += 1) {
    const left = [...form.querySelectorAll(`[name="left_key_${step}"]`)];
    const right = [...form.querySelectorAll(`[name="right_key_${step}"]`)];
    const hasPair = left.some((field, index) => Boolean(field.value || right[index]?.value));
    let complete = false;
    left.forEach((field, index) => {
      const counterpart = right[index];
      const mismatch = Boolean(field.value) !== Boolean(counterpart?.value);
      field.setAttribute("aria-invalid", mismatch ? "true" : "false");
      counterpart?.setAttribute("aria-invalid", mismatch ? "true" : "false");
      if (mismatch) valid = false;
      if (field.value && counterpart?.value) complete = true;
    });
    if (!hasPair || !complete) valid = false;
  }
  return valid;
}

const crossKeyForm = document.querySelector('form[action$="/run"]');
if (crossKeyForm) {
  crossKeyForm.addEventListener("submit", (event) => {
    if (!validateCrossKeyPairs(crossKeyForm)) {
      event.preventDefault();
      let alert = crossKeyForm.querySelector('[data-cross-validation-alert]');
      if (!alert) {
        alert = document.createElement("div");
        alert.dataset.crossValidationAlert = "true";
        alert.className = "notice warning";
        alert.setAttribute("role", "alert");
        alert.textContent = "Seleccioná una columna en ambos lados de cada par de claves, o dejá ambos campos vacíos.";
        crossKeyForm.prepend(alert);
      }
    }
  });
}

let previewRequest = 0;
function refreshResultPreview() {
  const container = document.getElementById("result-preview");
  const form = document.getElementById("result-filter-form");
  const body = document.getElementById("result-preview-body");
  if (!container || !form || !body) return;

  const params = new URLSearchParams();
  params.set("quick_search", form.elements.quick_search?.value ?? "");
  ["column", "operator", "value", "value2"].forEach((name) => {
    [...form.elements[name]].forEach((field) => params.append(name, field.value));
  });

  const requestId = ++previewRequest;
  fetch(`${container.dataset.previewUrl}?${params.toString()}`, { headers: { "Accept": "application/json" } })
    .then((response) => response.ok ? response.json() : Promise.reject(response))
    .then((data) => {
      if (requestId !== previewRequest) return;
      const columns = data.columns ?? [];
      body.replaceChildren();
      if (!data.rows.length) {
        const row = body.insertRow();
        const cell = row.insertCell();
        cell.colSpan = columns.length + 2;
        cell.textContent = "No hay filas que coincidan con los filtros.";
      } else {
        data.rows.forEach((item) => {
          const row = body.insertRow();
          row.insertCell().textContent = item.source_index + 1;
          columns.forEach((column) => row.insertCell().textContent = item.values[column] ?? "");
          row.insertCell().textContent = "";
        });
      }
      const count = document.getElementById("filtered-result-count");
      if (count) count.textContent = String(data.filtered_count);
      const status = document.getElementById("preview-status");
      if (status) {
        status.textContent = data.preview_truncated
          ? `Vista previa: primeras 50 filas de ${data.filtered_count}.`
          : `Vista previa actualizada: ${data.filtered_count} fila(s).`;
      }
    })
    .catch(() => {
      const status = document.getElementById("preview-status");
      if (status) status.textContent = "No se pudo actualizar la vista previa.";
    });
}

let previewTimer;
document.addEventListener("DOMContentLoaded", () => {
  updateConditionalFields();
  updateColumnChoices();
  updateRecipeActions();
  document.addEventListener("change", (event) => {
    if (event.target.matches("[data-conditional-control]")) updateConditionalFields();
    if (event.target.id === "file_index") updateColumnChoices();
    if (event.target.matches("[data-recipe-control]")) updateRecipeActions();
  });
  const resultFilterForm = document.getElementById("result-filter-form");
  if (resultFilterForm) {
    resultFilterForm.addEventListener("input", () => {
      window.clearTimeout(previewTimer);
      previewTimer = window.setTimeout(refreshResultPreview, 180);
    });
    resultFilterForm.addEventListener("change", refreshResultPreview);
  }
  document.querySelectorAll("form[data-confirm]").forEach((form) => {
    form.addEventListener("submit", (event) => {
      if (!window.confirm(form.dataset.confirm)) event.preventDefault();
    });
  });
});
