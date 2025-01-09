let workbookData = [];

// Function to trigger the file input dialog
function triggerFileUpload() {
    const fileInput = document.getElementById('fileInput');
    fileInput.click();
}

// Function to load and parse Excel file
function loadExcel() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    const fileName = document.getElementById('fileName');

    if (file) {
        fileName.textContent = `Selected File: ${file.name}`;
    } else {
        fileName.textContent = "No file selected";
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Get the first sheet
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        workbookData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        populateHeaders();
    };
    reader.readAsArrayBuffer(file);
}

// Function to populate dropdown with headers
function populateHeaders() {
    const headerSelector = document.getElementById('headerSelector');
    headerSelector.innerHTML = ''; // Clear previous options

    if (workbookData.length > 0) {
        const headers = workbookData[0]; // First row contains headers
        headers.forEach((header, index) => {
            const option = document.createElement('option');
            option.value = index; // Store column index as value
            option.textContent = header;
            headerSelector.appendChild(option);
        });
    }
}

// Function to calculate Standard Deviation
function calculateSD() {
    const selectedOptions = Array.from(document.getElementById('headerSelector').selectedOptions);
    const resultsBox = document.getElementById('resultsBox');
    resultsBox.innerHTML = ''; // Clear previous results

    selectedOptions.forEach(option => {
        const columnIndex = parseInt(option.value);
        const columnData = workbookData.slice(1) // Skip header row
            .map(row => parseFloat(row[columnIndex]))
            .filter(value => !isNaN(value)); // Filter out non-numeric values

        const sd = standardDeviation(columnData);
        const result = document.createElement('p');
        result.textContent = `SD of ${option.textContent}: ${sd}`;
        resultsBox.appendChild(result);
    });
}

// Helper function to calculate Standard Deviation
function standardDeviation(values) {
    if (values.length === 0) return 0;

    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    const variance = values.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / values.length;
    return Math.sqrt(variance).toFixed(2); // Rounded to 2 decimal places
}

// Event Listeners
document.getElementById('uploadBtn').addEventListener('click', triggerFileUpload);
document.getElementById('fileInput').addEventListener('change', loadExcel);
document.getElementById('calculateBtn').addEventListener('click', calculateSD);
