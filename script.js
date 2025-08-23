// ✅ Attendance Cell Color Styling
function updateStatusClass(cell, value) {
  cell.classList.remove("present", "absent");
  if (value === "Present") {
    cell.classList.add("present");
  } else if (value === "Absent") {
    cell.classList.add("absent");
  }
}

// ✅ Add Row
function addRow() {
  const tableBody = document.getElementById("tableBody");
  const row = document.createElement("tr");

  for (let i = 0; i < document.getElementById("tableHeader").children.length - 1; i++) {
    const td = document.createElement("td");

    // Mobile Number Column
    if (i === 4) {
      td.innerHTML = `
        <span class="phone-display">Not Set</span>
        <button onclick="modifyNumber(this)">Modify</button>
      `;
    } else {
      td.contentEditable = true;
    }
    row.appendChild(td);
  }

  // Attendance Dropdown
  const attendanceTd = document.createElement("td");
  const select = document.createElement("select");
  select.innerHTML = `
    <option value="Present">Present</option>
    <option value="Absent">Absent</option>
  `;
  select.onchange = () => updateStatusClass(attendanceTd, select.value);
  updateStatusClass(attendanceTd, select.value);
  attendanceTd.appendChild(select);
  row.appendChild(attendanceTd);

  tableBody.appendChild(row);
}

// ✅ Modify Number
function modifyNumber(button) {
  const td = button.parentElement;
  const currentSpan = td.querySelector(".phone-display");
  const currentNumber = currentSpan.textContent === "Not Set" ? "" : currentSpan.textContent;

  const newNumber = prompt("Enter Mobile Number:", currentNumber);

  if (newNumber && confirm("Save this number temporarily?")) {
    currentSpan.innerHTML = `<a href="tel:${newNumber}">${newNumber}</a>`;
    saveData(); // update localStorage instantly
  }
}

// ✅ Remove Row
function removeRow() {
  const tableBody = document.getElementById("tableBody");
  if (tableBody.rows.length > 0) {
    tableBody.deleteRow(tableBody.rows.length - 1);
  }
}

// ✅ Add Column
function addColumn() {
  const header = document.getElementById("tableHeader");
  const newTh = document.createElement("th");
  newTh.contentEditable = true;
  newTh.textContent = "New Column";
  header.insertBefore(newTh, header.lastElementChild);

  document.querySelectorAll("#tableBody tr").forEach(row => {
    const newTd = document.createElement("td");
    newTd.contentEditable = true;
    row.insertBefore(newTd, row.lastElementChild);
  });
}

// ✅ Remove Only User-Added Columns
const defaultColumnCount = 5;
function removeColumn() {
  const header = document.getElementById("tableHeader");
  const totalColumns = header.children.length;

  if (totalColumns <= defaultColumnCount + 1) {
    alert("Cannot remove default columns.");
    return;
  }

  const index = totalColumns - 2;
  header.removeChild(header.children[index]);

  document.querySelectorAll("#tableBody tr").forEach(row => {
    row.removeChild(row.children[index]);
  });
}

// ✅ Save Data to File
async function saveData() {
  const headers = [];
  const headerCells = document.querySelectorAll("#tableHeader th");
  for (let i = 0; i < headerCells.length; i++) {
    headers.push(headerCells[i].innerText.trim());
  }

  const rows = [];
  document.querySelectorAll("#tableBody tr").forEach(row => {
    const rowData = [];
    row.querySelectorAll("td").forEach((cell, i) => {
      if (i === 4) {
        const phone = cell.querySelector("a");
        rowData.push(phone ? phone.textContent : "Not Set");
      } else if (cell.querySelector("select")) {
        rowData.push(cell.querySelector("select").value);
      } else {
        rowData.push(cell.textContent.trim());
      }
    });
    rows.push(rowData);
  });

  // Convert to Excel
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  XLSX.utils.book_append_sheet(wb, ws, "Attendance");

  // Show save dialog
  const fileBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([fileBuffer], { type: "application/octet-stream" });

  const handle = await window.showSaveFilePicker({
    suggestedName: "attendance.xlsx",
    types: [{ description: "Excel File", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }],
  });

  const writable = await handle.createWritable();
  await writable.write(blob);
  await writable.close();

  alert("File saved successfully!");
}