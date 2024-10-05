let data = []; // This holds the initial Excel data
let filteredData = []; // This holds the filtered data after user operations

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0]; // Load the first sheet
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        // Initially display the full sheet
        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = ''; // Clear existing content

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');

    // Create table headers
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create table rows
    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null ? 'NULL' : cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Function to apply the selected operations and update the table
function applyOperation() {
    const primaryColumn = document.getElementById('primary-column').value.trim();
    const operationColumnsInput = document.getElementById('operation-columns').value.trim();
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    if (!primaryColumn || !operationColumnsInput) {
        alert("Please fill in both Primary Column and Operation Columns.");
        return;
    }

    const operationColumns = operationColumnsInput.split(',').map(col => col.trim().toUpperCase());
    
    filteredData = data.filter(row => {
        const primaryValue = row[primaryColumn];
        const isPrimaryNull = primaryValue === null;

        return operationColumns.every(col => {
            const colValue = row[col];
            const isColNull = colValue === null;

            if (operation === 'null') {
                return isColNull && (operationType === 'and' ? isPrimaryNull : true);
            } else { // not-null
                return !isColNull && (operationType === 'and' ? !isPrimaryNull : true);
            }
        });
    });

    displaySheet(filteredData);
}

// Function to open the download modal
function openDownloadModal() {
    const modal = document.getElementById('download-modal');
    modal.style.display = "block";
}

// Function to close the download modal
function closeDownloadModal() {
    const modal = document.getElementById('download-modal');
    modal.style.display = "none";
}

// Function to handle file download
function handleDownload() {
    const filename = document.getElementById('filename').value || 'download';
    const format = document.getElementById('file-format').value;

    const worksheet = XLSX.utils.json_to_sheet(filteredData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    if (format === 'xlsx') {
        XLSX.writeFile(workbook, `${filename}.xlsx`);
    } else if (format === 'csv') {
        XLSX.writeFile(workbook, `${filename}.csv`);
    }

    closeDownloadModal();
}

// Event Listeners
document.getElementById('apply-operation').addEventListener('click', applyOperation);
document.getElementById('download-button').addEventListener('click', openDownloadModal);
document.getElementById('close-modal').addEventListener('click', closeDownloadModal);
document.getElementById('confirm-download').addEventListener('click', handleDownload);

// Load the Excel sheet when the page loads
window.onload = () => {
    loadExcelSheet('path/to/your/excel/file.xlsx'); // Update with the actual path to your Excel file
};
