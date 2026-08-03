/* global Office, Excel */

// Configuraciones guardadas
let targetCellAddress = localStorage.getItem("targetExcelCell") || "V1436";
let targetSheetName = localStorage.getItem("targetExcelSheet") || "";
let includeSheet = localStorage.getItem("includeSheetName") === "true";
let captureMode = localStorage.getItem("captureMode") || "address";
let autoCopy = localStorage.getItem("autoCopyClipboard") === "true";

// Estado interno
let isTrackingActive = true;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    Office.addin.setStartupBehavior(Office.StartupBehavior.load);

    const input = document.getElementById("targetCellInput");
    const checkbox = document.getElementById("includeSheetName");
    const captureSelect = document.getElementById("captureModeSelect");
    const autoCopyBox = document.getElementById("autoCopyClipboard");
    const button = document.getElementById("saveCellBtn");
    const toggleBtn = document.getElementById("toggleTrackingBtn");
    const goToBtn = document.getElementById("goToTargetBtn");
    const status = document.getElementById("statusMessage");

    if (input) input.value = targetCellAddress;
    if (checkbox) checkbox.checked = includeSheet;
    if (captureSelect) captureSelect.value = captureMode;
    if (autoCopyBox) autoCopyBox.checked = autoCopy;

    loadWorksheetsList();

    // Toggle Play / Pausa
    if (toggleBtn) {
      toggleBtn.onclick = () => {
        isTrackingActive = !isTrackingActive;
        if (isTrackingActive) {
          toggleBtn.innerText = "Pause Tracking";
          toggleBtn.style.backgroundColor = "#0f7b41";
        } else {
          toggleBtn.innerText = "Resume Tracking";
          toggleBtn.style.backgroundColor = "#d9381e";
        }
      };
    }

    // Ir a la Celda Destino
    if (goToBtn) {
      goToBtn.onclick = goToTargetCell;
    }

    // Botón Guardar
    if (button) {
      button.onclick = () => {
        const newValue = input.value.trim();
        const selectedSheet = document.getElementById("sheetSelect").value;

        if (newValue && selectedSheet) {
          targetCellAddress = newValue;
          targetSheetName = selectedSheet;
          includeSheet = checkbox.checked;
          captureMode = captureSelect.value;
          autoCopy = autoCopyBox.checked;

          localStorage.setItem("targetExcelCell", newValue);
          localStorage.setItem("targetExcelSheet", selectedSheet);
          localStorage.setItem("includeSheetName", includeSheet);
          localStorage.setItem("captureMode", captureMode);
          localStorage.setItem("autoCopyClipboard", autoCopy);

          status.innerText = `✓ Settings saved successfully`;
          setTimeout(() => { status.innerText = ""; }, 3000);
        }
      };
    }

    // Registrar escuchador GLOBAL para todo el libro
    registerGlobalSelectionChangeListener();
  }
});

async function loadWorksheetsList() {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const sheetSelect = document.getElementById("sheetSelect");
    if (!sheetSelect) return;

    const currentSelection = sheetSelect.value || targetSheetName;
    sheetSelect.innerHTML = "";

    sheets.items.forEach((sheet) => {
      const option = document.createElement("option");
      option.value = sheet.name;
      option.textContent = sheet.name;
      if (sheet.name === currentSelection) option.selected = true;
      sheetSelect.appendChild(option);
    });

    if (!targetSheetName && sheets.items.length > 0) {
      targetSheetName = sheets.items[0].name;
    }
  });
}

async function goToTargetCell() {
  await Excel.run(async (context) => {
    if (targetSheetName && targetCellAddress) {
      const specificSheet = context.workbook.worksheets.getItem(targetSheetName);
      const targetRange = specificSheet.getRange(targetCellAddress);
      
      specificSheet.activate();
      targetRange.select();
      await context.sync();
    }
  });
}

// 🌟 SOLUCIÓN: Registrar el evento a nivel del libro completo (worksheets collection)
async function registerGlobalSelectionChangeListener() {
  await Excel.run(async (context) => {
    // Escuchar el cambio de selección GLOBAL en CUALQUIER hoja del libro
    context.workbook.worksheets.onSelectionChanged.add(onSelectionChangeHandler);
    
    // También actualizamos la lista de hojas si el usuario añade o activa pestañas
    context.workbook.worksheets.onActivated.add(loadWorksheetsList);
    
    await context.sync();
  });
}

async function onSelectionChangeHandler(args) {
  if (!isTrackingActive) return;

  await Excel.run(async (context) => {
    // getActiveCell() siempre devuelve la celda activa sin importar en qué hoja esté
    const activeCell = context.workbook.getActiveCell();
    activeCell.load(["address", "values", "formulas", "text"]);
    await context.sync();

    // 1. Obtener dirección formateada
    let cellAddr = activeCell.address;
    if (!includeSheet && cellAddr.includes("!")) {
      cellAddr = cellAddr.split("!").pop();
    }

    // 2. Obtener valor visible
    const cellValue = activeCell.text[0][0] || activeCell.values[0][0] || "(empty)";

    // 3. Evaluar qué tipo de dato exportar
    let outputValue = "";

    if (captureMode === "value") {
      outputValue = cellValue;
    } else if (captureMode === "formula") {
      outputValue = activeCell.formulas[0][0];
    } else if (captureMode === "both") {
      outputValue = `${cellValue} / ${cellAddr}`;
    } else {
      outputValue = cellAddr;
    }

    // 📋 Copiar al Portapapeles si está activo
    if (autoCopy && outputValue) {
      try {
        await navigator.clipboard.writeText(String(outputValue));
      } catch (e) {
        console.log("Clipboard error:", e);
      }
    }

    // Escribir en la celda de destino seleccionada
    if (targetSheetName) {
      const specificSheet = context.workbook.worksheets.getItem(targetSheetName);
      const targetRange = specificSheet.getRange(targetCellAddress);
      
      targetRange.values = [[outputValue]];
      await context.sync();

      // 🕒 Actualizar interfaz
      const lastValElem = document.getElementById("lastLoggedValue");
      const lastTimeElem = document.getElementById("lastLoggedTime");
      if (lastValElem) lastValElem.innerText = outputValue;
      if (lastTimeElem) lastTimeElem.innerText = `At ${new Date().toLocaleTimeString()}`;
    }
  });
}