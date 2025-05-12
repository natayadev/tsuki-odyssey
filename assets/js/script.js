const checklistTable = document.querySelector("#checklist tbody");
const downloadJsonBtn = document.getElementById("downloadJson");
const downloadExcelBtn = document.getElementById("downloadExcel");
const tabsContainer = document.querySelector(".tabs");

let data = {};
let currentTab = "";

async function loadExcelFile() {

  try {
    const response = await fetch("data/tsuki_archivements.xlsx");
    if (!response.ok) throw new Error("No se pudo cargar el archivo Excel.");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    const sheetNames = workbook.SheetNames;
    sheetNames.forEach((sheetName, index) => {
      const button = document.createElement("button");
      button.className = `tab-button ${index === 0 ? "active" : ""}`;
      button.setAttribute("data-tab", sheetName);
      button.textContent = sheetName;
      tabsContainer.appendChild(button);

      const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
      data[sheetName] = sheetData;

      button.addEventListener("click", () => {
        if (sheetName !== currentTab) {
          document.querySelector(".tab-button.active").classList.remove("active");
          button.classList.add("active");
          currentTab = sheetName;
          updateChecklist(data[currentTab]);
        }
      });
    });

    currentTab = sheetNames[0];

    if (data[currentTab] && data[currentTab].length > 0) {
      updateChecklist(data[currentTab]);
    } else {
      console.warn("No hay datos disponibles para actualizar la tabla.");
    }
  } catch (error) {
    console.error("Error cargando el archivo Excel:", error);
  }
}

function updateChecklist(items) {
  checklistTable.innerHTML = "";

  const columns = Object.keys(items[0]);

  const headerRow = document.createElement("tr");
  const checkboxHeader = document.createElement("th");
  checkboxHeader.textContent = "Select";
  checkboxHeader.style.cursor = "pointer";
  checkboxHeader.addEventListener("click", () => {
    const checkboxes = checklistTable.querySelectorAll("input[type='checkbox']");
    const allChecked = Array.from(checkboxes).every(checkbox => checkbox.checked);
    checkboxes.forEach(checkbox => {
      checkbox.checked = !allChecked;
      const event = new Event('change');
      checkbox.dispatchEvent(event);
    });
  });
  headerRow.appendChild(checkboxHeader);

  columns.forEach((col) => {
    const headerCell = document.createElement("th");
    headerCell.textContent = col;
    headerRow.appendChild(headerCell);
  });
  checklistTable.appendChild(headerRow);

  items.forEach((item, index) => {
    const row = document.createElement("tr");

    const checkboxCell = document.createElement("td");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.id = `item-${index}`;
    checkbox.checked = item.checked || false;
    checkbox.addEventListener("change", () => {
      item.checked = checkbox.checked;
      saveChecklistState();
      updateRowStyle(row, checkbox.checked);
    });
    checkboxCell.appendChild(checkbox);
    row.appendChild(checkboxCell);

    columns.forEach((col) => {
      const cell = document.createElement("td");
      if (col === 'Image') {
        let imageName = item[col];
      
        if (typeof imageName === 'string') {
          imageName = imageName
            .trim() // Elimina espacios en blanco al inicio y al final
            .replace(/\s+/g, '_') // Reemplaza espacios con guiones bajos
            .replace(/[^\w\-]/g, ''); // Elimina caracteres no válidos para nombres de archivo
        } else {
          imageName = ''; // Si no es una cadena, asignar un valor por defecto vacío
        }
      
        const imgElement = document.createElement("img");
        imgElement.src = `assets/img/${imageName}.webp`;
        imgElement.alt = item[col];
        imgElement.style.width = "50px";
        imgElement.style.height = "auto";
        imgElement.onerror = () => {
          // Manejador para casos donde la imagen no exista
          imgElement.src = "assets/img/logo_tsuki.webp"; // Imagen de respaldo
          imgElement.alt = "Logo Tsuki";
        };
        cell.appendChild(imgElement);
      } else {
        cell.textContent = item[col] || "";
      }
      

      row.appendChild(cell);
    });

    updateRowStyle(row, checkbox.checked);
    checklistTable.appendChild(row);
  });
}

function updateRowStyle(row, isChecked) {
  console.log(`Updating row style: ${row}, Checked: ${isChecked}`);
  const color = isChecked ? 'green' : 'red';
  row.style.color = color;
}

downloadJsonBtn.addEventListener("click", () => {
  const jsonBlob = new Blob([JSON.stringify(data[currentTab], null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(jsonBlob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${currentTab}.json`;
  a.click();
  URL.revokeObjectURL(url);
});

downloadExcelBtn.addEventListener("click", () => {
  const worksheet = XLSX.utils.json_to_sheet(data[currentTab]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, currentTab);
  XLSX.writeFile(workbook, `${currentTab}.xlsx`);
});

function saveChecklistState() {
  localStorage.setItem('checklistState', JSON.stringify(data));
}

function loadChecklistState() {
  const savedState = localStorage.getItem('checklistState');
  if (savedState) {
    data = JSON.parse(savedState);
    if (data[currentTab] && data[currentTab].length > 0) {
      updateChecklist(data[currentTab]);
    }
  }
}

loadExcelFile();
loadChecklistState();
