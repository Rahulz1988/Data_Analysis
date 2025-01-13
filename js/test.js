/* Global variables and constants */
let jsonData = null;
let originalJsonData = null;  // Store original data
let excelHeaders = [];  // Store Excel headers

// Constants for headers
const FIXED_HEADERS = ['SN', 'Email ID', 'Full Name', 'Specialization', 'Credibility Score'];
const SECTION_HEADERS = ['MaxScore', 'Score', 'Percentage'];
const SEPARATOR = 'Separator';
const ASSESSMENT_STATUS = 'Assessment Status';
const CREDIBILITY_SCORE = 'Credibility Score';

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
        headers.push(CREDIBILITY_SCORE);
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
            originalJsonData = XLSX.utils.sheet_to_json(worksheet);
            jsonData = JSON.parse(JSON.stringify(originalJsonData)); // Deep copy

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
// function processData() {
//     if (!originalJsonData) {
//         alert('Please upload data first.');
//         return;
//     }

//     // Reset to original data before each normalization
//     jsonData = JSON.parse(JSON.stringify(originalJsonData));

//     // Clear any previous change markers
//     jsonData.forEach(row => {
//         Object.keys(row).forEach(key => {
//             if (key.endsWith('_changed')) {
//                 delete row[key];
//             }
//         });
//     });

//     applyCutoffFilters();
//     const initialPercentage = calculateTestSelectPercentage();
//     console.log('Initial percentage:', initialPercentage);

//     // Define target percentage within range
//     let targetPercentage;
//     if (initialPercentage < TARGET_MIN) {
//         targetPercentage = TARGET_MIN + (Math.random() * (TARGET_MAX - TARGET_MIN));
//         adjustScoresWithSD('up', targetPercentage);
//     } else if (initialPercentage > TARGET_MAX) {
//         targetPercentage = TARGET_MIN + (Math.random() * (TARGET_MAX - TARGET_MIN));
//         adjustScoresWithSD('down', targetPercentage);
//     } else {
//         // If already in range, make small random adjustments
//         targetPercentage = TARGET_MIN + (Math.random() * (TARGET_MAX - TARGET_MIN));
//         const direction = Math.random() < 0.5 ? 'up' : 'down';
//         adjustScoresWithSD(direction, targetPercentage);
//     }

//     const finalPercentage = calculateTestSelectPercentage();
//     console.log('Final percentage:', finalPercentage);

//     downloadResultBtn.style.display = 'block';
//     displayResults();
// }

function processData() {
    if (!originalJsonData) {
        alert('Please upload data first.');
        return;
    }

    // Reset to original data before each normalization
    jsonData = JSON.parse(JSON.stringify(originalJsonData));

    // Clear any previous change markers
    jsonData.forEach(row => {
        Object.keys(row).forEach(key => {
            if (key.endsWith('_changed')) {
                delete row[key];
            }
        });
    });

    applyCutoffFilters();
    const initialPercentage = calculateTestSelectPercentage();
    console.log('Initial percentage:', initialPercentage);

    // Define target percentage within range
    let targetPercentage = TARGET_MIN + (Math.random() * (TARGET_MAX - TARGET_MIN));
    console.log('Target percentage:', targetPercentage);

    if (initialPercentage < targetPercentage) {
        adjustScoresWithSD('up', targetPercentage);
    } else {
        adjustScoresWithSD('down', targetPercentage);
    }

    const finalPercentage = calculateTestSelectPercentage();
    console.log('Final percentage:', finalPercentage);

    // Verify that jsonData has been modified
    console.log('Data has been normalized:', 
        jsonData.some(row => Object.keys(row).some(key => key.endsWith('_changed')))
    );

    downloadResultBtn.style.display = 'block';
    displayResults();
}

// // Score Adjustment Function
// function adjustScoresWithSD(direction, targetPercentage) {
//     const scoreHeaders = excelHeaders.filter(header => header.match(/^S\d+Score$/));
    
//     if (!scoreHeaders.includes('S3Score')) scoreHeaders.push('S3Score');
//     if (!scoreHeaders.includes('S4Score')) scoreHeaders.push('S4Score');
    
//     const maxScoreHeaders = scoreHeaders.map(scoreHeader => 
//         scoreHeader.replace('Score', 'MaxScore')
//     );
    
//     const standardDeviations = {};
    
//     scoreHeaders.forEach(header => {
//         const values = jsonData.map(row => parseFloat(row[header])).filter(val => !isNaN(val));
//         standardDeviations[header] = calculateStandardDeviation(values);
//     });
    
//     let currentPercentage = calculateTestSelectPercentage();
//     let iterationCount = 0;
//     const maxIterations = 1000; // Prevent infinite loops
    
//     while (iterationCount < maxIterations) {
//         currentPercentage = calculateTestSelectPercentage();
//         console.log('Current percentage:', currentPercentage);
        
//         // Break if we're close enough to target
//         if (Math.abs(currentPercentage - targetPercentage) < 0.5) {
//             break;
//         }

//         // Determine if we need to go up or down based on current vs target
//         const effectiveDirection = currentPercentage < targetPercentage ? 'up' : 'down';
        
//         const candidatesToModify = 
//             effectiveDirection === 'down' ? 
//                 jsonData.filter(row => row[ASSESSMENT_STATUS] === 'Test Select') :
//                 jsonData.filter(row => row[ASSESSMENT_STATUS] === 'Test Reject');
        
//         if (candidatesToModify.length === 0) break;
        
//         // Adjust the number of candidates to modify based on how far we are from target
//         const percentageDiff = Math.abs(currentPercentage - targetPercentage);
//         const numCandidatesToModify = Math.max(1, Math.min(
//             Math.ceil(candidatesToModify.length * (percentageDiff / 100)),
//             Math.ceil(candidatesToModify.length * 0.1) // Max 10% at once
//         ));
        
//         for (let i = 0; i < numCandidatesToModify; i++) {
//             const randomCandidateIndex = Math.floor(Math.random() * candidatesToModify.length);
//             const randomCandidate = candidatesToModify[randomCandidateIndex];
            
//             scoreHeaders.forEach((header, index) => {
//                 const currentValue = parseFloat(randomCandidate[header]);
//                 const maxScoreHeader = maxScoreHeaders[index];
//                 const maxScoreValue = parseFloat(randomCandidate[maxScoreHeader]);
                
//                 if (!isNaN(currentValue) && !isNaN(maxScoreValue)) {
//                     // Adjust the random factor based on how far we are from target
//                     const randomFactor = 0.5 + (Math.random() * (percentageDiff / 50));
//                     const adjustmentValue = standardDeviations[header] * randomFactor;
                    
//                     if (effectiveDirection === 'down') {
//                         randomCandidate[header] = Math.max(0, (currentValue - adjustmentValue)).toFixed(2);
//                     } else {
//                         randomCandidate[header] = Math.min(maxScoreValue, (currentValue + adjustmentValue)).toFixed(2);
//                     }
                    
//                     randomCandidate[header.replace('Score', 'Percentage')] = 
//                         ((parseFloat(randomCandidate[header]) / maxScoreValue) * 100).toFixed(2);
                    
//                     randomCandidate[header + '_changed'] = true;
//                 }
//             });
            
//             randomCandidate[ASSESSMENT_STATUS] = 
//                 effectiveDirection === 'down' ? 'Test Reject' : 'Test Select';
//         }
        
//         iterationCount++;
//     }
// }


function adjustScoresWithSD(direction, targetPercentage) {
    const scoreHeaders = excelHeaders.filter(header => header.match(/^S\d+Score$/));
    
    if (!scoreHeaders.includes('S3Score')) scoreHeaders.push('S3Score');
    if (!scoreHeaders.includes('S4Score')) scoreHeaders.push('S4Score');
    
    const maxScoreHeaders = scoreHeaders.map(scoreHeader => 
        scoreHeader.replace('Score', 'MaxScore')
    );
    
    const standardDeviations = {};
    
    scoreHeaders.forEach(header => {
        const values = jsonData.map(row => parseFloat(row[header])).filter(val => !isNaN(val));
        standardDeviations[header] = calculateStandardDeviation(values);
    });
    
    let currentPercentage = calculateTestSelectPercentage();
    let iterationCount = 0;
    const maxIterations = 1000; // Prevent infinite loops
    
    while (iterationCount < maxIterations) {
        currentPercentage = calculateTestSelectPercentage();
        console.log('Current percentage:', currentPercentage);
        
        // Break if we're close enough to target
        if (Math.abs(currentPercentage - targetPercentage) < 0.5) {
            break;
        }

        // Determine if we need to go up or down based on current vs target
        const effectiveDirection = currentPercentage < targetPercentage ? 'up' : 'down';
        
        const candidatesToModify = 
            effectiveDirection === 'down' ? 
                jsonData.filter(row => row[ASSESSMENT_STATUS] === 'Test Select') :
                jsonData.filter(row => row[ASSESSMENT_STATUS] === 'Test Reject');
        
        if (candidatesToModify.length === 0) break;
        
        // Adjust the number of candidates to modify based on how far we are from target
        const percentageDiff = Math.abs(currentPercentage - targetPercentage);
        const numCandidatesToModify = Math.max(1, Math.min(
            Math.ceil(candidatesToModify.length * (percentageDiff / 100)),
            Math.ceil(candidatesToModify.length * 0.1) // Max 10% at once
        ));
        
        for (let i = 0; i < numCandidatesToModify; i++) {
            const randomCandidateIndex = Math.floor(Math.random() * candidatesToModify.length);
            const randomCandidate = candidatesToModify[randomCandidateIndex];
            
            scoreHeaders.forEach((header, index) => {
                const currentValue = parseFloat(randomCandidate[header]);
                const maxScoreHeader = maxScoreHeaders[index];
                const maxScoreValue = parseFloat(randomCandidate[maxScoreHeader]);
                
                if (!isNaN(currentValue) && !isNaN(maxScoreValue)) {
                    // Adjust the random factor based on how far we are from target
                    const randomFactor = 0.5 + (Math.random() * (percentageDiff / 50));
                    const adjustmentValue = standardDeviations[header] * randomFactor;
                    
                    let newValue;
                    if (effectiveDirection === 'down') {
                        newValue = Math.max(0, (currentValue - adjustmentValue));
                    } else {
                        newValue = Math.min(maxScoreValue, (currentValue + adjustmentValue));
                    }
                    
                    // Round the value to the nearest integer
                    randomCandidate[header] = Math.round(newValue);
                    
                    randomCandidate[header.replace('Score', 'Percentage')] = 
                        Math.round((parseFloat(randomCandidate[header]) / maxScoreValue) * 100);
                    
                    randomCandidate[header + '_changed'] = true;
                }
            });
            
            randomCandidate[ASSESSMENT_STATUS] = 
                effectiveDirection === 'down' ? 'Test Reject' : 'Test Select';
        }
        
        iterationCount++;
    }
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

        row[ASSESSMENT_STATUS] = passesAllFilters ? 'Test Select' : 'Test Reject';
    });
}

// Results Display
function displayResults() {
    const totalCandidates = jsonData.length;
    const selected = jsonData.filter(row => row[ASSESSMENT_STATUS] === 'Test Select').length;
    const rejected = totalCandidates - selected;
    const selectionPercentage = ((selected / totalCandidates) * 100).toFixed(2);

    resultsBox.innerHTML = `
        <p><strong>Total Candidates:</strong> ${totalCandidates}</p>
        <p><strong>Selected:</strong> ${selected} (${selectionPercentage}%)</p>
        <p><strong>Rejected:</strong> ${rejected}</p>
    `;
}

// Download Results
// function downloadResult() {
//     try {
//         const ws = XLSX.utils.json_to_sheet(jsonData);

//         // Highlight changed cells
//         jsonData.forEach((row, rowIndex) => {
//             Object.keys(row).forEach(header => {
//                 if (header.endsWith('_changed') && row[header]) {
//                     const cellAddress = XLSX.utils.encode_cell({
//                         r: rowIndex + 1,
//                         c: excelHeaders.indexOf(header.replace('_changed', ''))
//                     });
//                     if (ws[cellAddress]) {
//                         ws[cellAddress].s = { fill: { fgColor: { rgb: "FFFF00" } } };
//                     }
//                 }
//             });
//         });

//         const wb = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(wb, ws, 'Results');
//         XLSX.writeFile(wb, `normalized_results_${new Date().toISOString().replace(/[:.]/g, '-')}.xlsx`);
//     } catch (error) {
//         console.error('Download error:', error);
//         alert('Error downloading results. Please try again.');
//     }
// }

// Download Results
// function downloadResult() {
//     try {
//         // Create a deep copy of jsonData without the _changed flags
//         const downloadData = jsonData.map(row => {
//             const cleanRow = {};
//             Object.keys(row).forEach(key => {
//                 // Skip the _changed marker properties
//                 if (!key.endsWith('_changed')) {
//                     cleanRow[key] = row[key];
//                 }
//             });
//             return cleanRow;
//         });

//         // Create worksheet from the normalized data
//         const ws = XLSX.utils.json_to_sheet(downloadData);

//         // Add cell styling for changed values
//         downloadData.forEach((row, rowIndex) => {
//             Object.keys(row).forEach(key => {
//                 // Check if this cell was changed (by looking at the _changed flag in original jsonData)
//                 if (jsonData[rowIndex][key + '_changed']) {
//                     const cellAddress = XLSX.utils.encode_cell({
//                         r: rowIndex + 1, // Add 1 to account for header row
//                         c: Object.keys(downloadData[0]).indexOf(key)
//                     });
                    
//                     if (!ws[cellAddress]) {
//                         ws[cellAddress] = { v: row[key] };
//                     }
                    
//                     // Set yellow background for changed cells
//                     ws[cellAddress].s = {
//                         fill: {
//                             patternType: 'solid',
//                             fgColor: { rgb: "FFFF00" }
//                         }
//                     };
//                 }
//             });
//         });

//         // Set column widths
//         const maxWidth = 15;
//         const colWidths = {};
//         Object.keys(downloadData[0]).forEach(key => {
//             colWidths[key] = Math.min(maxWidth, key.length + 2);
//         });
//         ws['!cols'] = Object.values(colWidths).map(width => ({ width }));

//         // Create workbook and append worksheet
//         const wb = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(wb, ws, 'Normalized Results');

//         // Generate filename with timestamp
//         const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
//         const filename = `normalized_results_${timestamp}.xlsx`;

//         // Write file
//         XLSX.writeFile(wb, filename);

//         console.log('Download completed successfully');
//     } catch (error) {
//         console.error('Download error:', error);
//         alert('Error downloading results. Please try again.');
//     }
// }

function downloadResult() {
    try {
        // Create enhanced data with change tracking columns
        const downloadData = jsonData.map(row => {
            const enhancedRow = {};
            
            // First add all original columns
            Object.keys(row).forEach(key => {
                if (!key.endsWith('_changed')) {
                    enhancedRow[key] = row[key];
                }
            });
            
            // Add change tracking columns for score and percentage fields
            Object.keys(row).forEach(key => {
                if (key.endsWith('_changed') && row[key] === true) {
                    // Get the original column name without '_changed'
                    const originalKey = key.replace('_changed', '');
                    // Create a new column indicating the change
                    enhancedRow[`${originalKey}_IsChanged`] = 'Yes';
                }
            });
            
            return enhancedRow;
        });

        // // Add 'No' for unchanged values
        // const allColumns = new Set();
        // downloadData.forEach(row => {
        //     Object.keys(row).forEach(key => allColumns.add(key));
        // });
        
        // downloadData.forEach(row => {
        //     allColumns.forEach(col => {
        //         if (col.endsWith('_IsChanged') && !row[col]) {
        //             row[col] = 'No';
        //         }
        //     });
        // });

        // Create worksheet
        const ws = XLSX.utils.json_to_sheet(downloadData);

        // Set column widths for better readability
        const maxWidth = 15;
        ws['!cols'] = Object.keys(downloadData[0]).map(key => ({
            width: Math.min(maxWidth, key.length + 2)
        }));

        // Create workbook and append worksheet
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Normalized Results');

        // Generate filename with timestamp
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const filename = `normalized_results_${timestamp}.xlsx`;

        // Write file
        const wopts = {
            bookSST: false,
            bookType: 'xlsx',
            compression: true
        };

        XLSX.writeFile(wb, filename, wopts);
        console.log('Download completed with change tracking columns');
    } catch (error) {
        console.error('Download error:', error);
        alert('Error downloading results. Please try again.');
    }
}