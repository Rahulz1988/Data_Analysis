/* Global variables and constants */
let jsonData = null;
let excelHeaders = [];  // Store Excel headers

// Constants for headers
const FIXED_HEADERS = ['SN', 'Email ID', 'Full Name', 'Specialization'];
const SECTION_HEADERS = ['MaxScore', 'Score', 'Percentage'];
const SEPARATOR = 'Separator';
const ASSESSMENT_STATUS = 'Assessment Status';

// Constants for normalization
const TARGET_MIN = 28; // Minimum target percentage
const TARGET_MAX = 44; // Maximum target percentage
const UPPER_THRESHOLD = 45; // Upper threshold for test select percentage
const LOWER_THRESHOLD = 20; // Lower threshold for test select percentage

// DOM Elements
const fileInput = document.getElementById('fileInput');
const uploadBtn = document.getElementById('uploadBtn');
const processBtn = document.getElementById('processBtn');
const downloadTemplateBtn = document.getElementById('downloadTemplateBtn');
const downloadResultBtn = document.getElementById('downloadResultBtn');
const resultsBox = document.getElementById('resultsBox');
const addCutoffBtn = document.getElementById('addCutoffBtn');
const cutoffInputs = document.getElementById('cutoffInputs');

// Event Listeners
uploadBtn.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', handleFileUpload);
processBtn.addEventListener('click', processData);
downloadTemplateBtn.addEventListener('click', downloadTemplate);
downloadResultBtn.addEventListener('click', () => {
    if (validateNormalization()) {
        downloadResult();
    }
});
addCutoffBtn.addEventListener('click', addCutoffFilter);

// Helper Functions
function calculateStandardDeviation(values) {
    const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
    const squareDiffs = values.map(value => Math.pow(value - mean, 2));
    const avgSquareDiff = squareDiffs.reduce((sum, val) => sum + val, 0) / values.length;
    return Math.round(Math.sqrt(avgSquareDiff) * 100) / 100;
}

function validateNormalization() {
    return jsonData.some(row => Object.keys(row).some(key => key.endsWith('_changed') && row[key] === true));
}

function getPercentageHeaders() {
    return excelHeaders.filter(header => header.toLowerCase().includes('percentage'));
}

function calculateTestSelectPercentage() {
    const totalCandidates = jsonData.length;
    const selectedCandidates = jsonData.filter(row => row[ASSESSMENT_STATUS] === 'Test Select').length;
    return (selectedCandidates / totalCandidates) * 100;
}

function createSectionHeaders(sectionNumber) {
    return SECTION_HEADERS.map(header => `Section${sectionNumber}_${header}`);
}

// Template Generation
function downloadTemplate() {
    try {
        const headers = [...FIXED_HEADERS];
        for (let i = 1; i <= 4; i++) {
            headers.push(...createSectionHeaders(i));
            headers.push(SEPARATOR);
        }
        headers.push(ASSESSMENT_STATUS);

        const ws = XLSX.utils.aoa_to_sheet([headers]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Template');
        XLSX.writeFile(wb, 'assessment_template.xlsx');
    } catch (error) {
        console.error('Template generation error:', error);
        alert('Error generating template. Please check console for details.');
    }
}

// File Upload Handler
async function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById('fileName').textContent = `File uploaded: ${file.name}`;

    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            jsonData = XLSX.utils.sheet_to_json(worksheet);

            if (jsonData.length === 0) {
                throw new Error('Empty file');
            }

            excelHeaders = Object.keys(jsonData[0]).filter(header => header !== ASSESSMENT_STATUS);
            processBtn.disabled = false;
            addCutoffBtn.disabled = false;
            cutoffInputs.innerHTML = '';
        } catch (error) {
            console.error('File processing error:', error);
            alert('Error processing file. Please ensure it\'s a valid Excel file.');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Cutoff Filter UI
function addCutoffFilter() {
    const filterDiv = document.createElement('div');
    filterDiv.className = 'cutoff-filter';

    const headerOptions = excelHeaders.map(header => `<option value="${header}">${header}</option>`).join('');

    filterDiv.innerHTML = `
        <select class="cutoff-field">${headerOptions}</select>
        <select class="cutoff-operator">
            <option value=">=">>=</option>
            <option value="<="><=</option>
            <option value="=">=</option>
            <option value=">">></option>
            <option value="<"><</option>
        </select>
        <input type="number" class="cutoff-value" step="0.01" />
        <button class="remove-filter">Remove</button>
    `;
    
    filterDiv.querySelector('.remove-filter').addEventListener('click', () => filterDiv.remove());
    cutoffInputs.appendChild(filterDiv);
}

// Main Processing Function
function processData() {
    if (!jsonData) {
        alert('Please upload data first.');
        return;
    }

    applyCutoffFilters();
    const initialPercentage = calculateTestSelectPercentage();

    // Determine the action based on the initial test select percentage
    if (initialPercentage > UPPER_THRESHOLD) {
        adjustScoresWithSD('down');
    } else if (initialPercentage < LOWER_THRESHOLD) {
        adjustScoresWithSD('up');
    }

    downloadResultBtn.style.display = 'block';
    displayResults();
}

// Adjust Scores Based on Standard Deviation
function adjustScoresWithSD(direction) {
    // Use a regular expression to match score headers starting with 'S' followed by any number
    const scoreHeaders = excelHeaders.filter(header => header.match(/^S\d+Score$/));
  
    // Ensure all sections are included
    if (!scoreHeaders.includes('S3Score')) scoreHeaders.push('S3Score');
    if (!scoreHeaders.includes('S4Score')) scoreHeaders.push('S4Score');
  
    const maxScoreHeaders = scoreHeaders.map(scoreHeader => 
      scoreHeader.replace('Score', 'MaxScore')
    );
  
    const standardDeviations = {};
  
    // Calculate SD for each score column
    scoreHeaders.forEach(header => {
      const values = jsonData.map(row => parseFloat(row[header])).filter(val => !isNaN(val));
      standardDeviations[header] = calculateStandardDeviation(values);
    });
  
    let currentPercentage;
  
    do {
      currentPercentage = calculateTestSelectPercentage();
  
      // Randomly select candidates based on direction
      const candidatesToModify = 
        direction === 'down' ? jsonData.filter(row => row[ASSESSMENT_STATUS] === 'Test Select') :
                             jsonData.filter(row => row[ASSESSMENT_STATUS] === 'Test Reject');
  
      if (candidatesToModify.length === 0) break;
  
      const randomCandidateIndex = Math.floor(Math.random() * candidatesToModify.length);
      const randomCandidate = candidatesToModify[randomCandidateIndex];
  
      scoreHeaders.forEach((header, index) => {
        const currentValue = parseFloat(randomCandidate[header]);
        const maxScoreHeader = maxScoreHeaders[index];
        const maxScoreValue = parseFloat(randomCandidate[maxScoreHeader]);
  
        if (!isNaN(currentValue) && !isNaN(maxScoreValue)) {
          const adjustmentValue = standardDeviations[header];
          if (direction === 'down') {
            randomCandidate[header] = Math.max(0, (currentValue - adjustmentValue)).toFixed(2);
          } else { // direction === 'up'
            randomCandidate[header] = Math.min(maxScoreValue, (currentValue + adjustmentValue)).toFixed(2);
          }
  
          // Update the corresponding percentage
          randomCandidate[header.replace('Score', 'Percentage')] = 
            ((parseFloat(randomCandidate[header]) / maxScoreValue) * 100).toFixed(2);
  
          // Mark the cell as changed for highlighting
          randomCandidate[header + '_changed'] = true;
        }
      });
  
      randomCandidate[ASSESSMENT_STATUS] = 
        direction === 'down' ? 'Test Reject' : 'Test Select';
  
    } while ((direction === 'down' && currentPercentage > TARGET_MIN) || 
             (direction === 'up' && currentPercentage < TARGET_MAX));
  }
  
 

// Apply Cutoff Filters
function applyCutoffFilters() {
    const filters = Array.from(document.querySelectorAll('.cutoff-filter')).map(filter => ({
        field: filter.querySelector('.cutoff-field').value,
        operator: filter.querySelector('.cutoff-operator').value,
        value: parseFloat(filter.querySelector('.cutoff-value').value)
    }));

    if (filters.length === 0 || filters.some(f => isNaN(f.value))) {
        alert('Please add valid cutoff filters.');
        return;
    }

    jsonData.forEach(row => {
        const passesAllFilters =
            filters.every(filter => {
                const fieldValue = parseFloat(row[filter.field]);
                if (isNaN(fieldValue)) return false;

                switch (filter.operator) {
                    case '>=': return fieldValue >= filter.value;
                    case '<=': return fieldValue <= filter.value;
                    case '=': return fieldValue === filter.value;
                    case '>': return fieldValue > filter.value;
                    case '<': return fieldValue < filter.value;
                    default: return false;
                }
            });

        row[ASSESSMENT_STATUS] =
            passesAllFilters ? 'Test Select' : 'Test Reject';
    });
}

// Results Display
function displayResults() {
   const totalCandidates= jsonData.length;
   const selected= jsonData.filter(row=> row[ASSESSMENT_STATUS]==='Test Select').length;
   const rejected= totalCandidates - selected;
   const selectionPercentage= ((selected / totalCandidates)*100).toFixed(2);

   resultsBox.innerHTML= `
       <p><strong>Total Candidates:</strong> ${totalCandidates}</p>
       <p><strong>Selected:</strong> ${selected} (${selectionPercentage}%)</p>
       <p><strong>Rejected:</strong> ${rejected}</p>
   `;
}

// Download Results
function downloadResult() {
    try {
      // Ensure only normalized data is included in the download
      const normalizedData = jsonData.map(row => {
        // Keep Assessment Status
        return row;
      });
  
      // Create worksheet from normalized data
      const ws = XLSX.utils.json_to_sheet(normalizedData);
  
      // Highlight changed cells
      normalizedData.forEach((row, rowIndex) => {
        Object.keys(row).forEach(header => {
          if (header.endsWith('_changed') && row[header]) {
            const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: excelHeaders.indexOf(header.replace('_changed', '')) });
            if (ws[cellAddress]) {
              ws[cellAddress].s = { fill: { fgColor: { rgb: "FFFF00" } } }; // Yellow color for changed cells
            }
          }
        });
      });
  
      // Create and save workbook
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Results');
      XLSX.writeFile(wb, `normalized_results_${new Date().toISOString().replace(/[:.]/g,'-')}.xlsx`);
    } catch (error) {
      console.error('Download error:', error);
      alert('Error downloading results. Please try again.');
    }
  }


