// Using SheetJS (xlsx) library for Excel handling

// Function to read an Excel file
const readExcelFile = (file) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: 'array' });
                console.log("Workbook read successfully:", workbook);
                resolve(workbook);
            } catch (error) {
                console.error("Error reading workbook:", error);
                reject(error);
            }
        };

        reader.onerror = (error) => {
            console.error("File reading error:", error);
            reject(error);
        };
        
        reader.readAsArrayBuffer(file);
    });
};

// Main function to process the Excel file
export const processExcelFile = async (file) => {
    try {
        console.log("Starting to process Excel file...");
        
        const workbook = await readExcelFile(file);
        
        if (!workbook || !workbook.SheetNames.length) {
            throw new Error("No sheets found in workbook.");
        }

        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        
        if (!worksheet) {
            throw new Error("No worksheet found.");
        }

        // Added defval: 0 to handle empty cells
        const data = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1,
            defval: 0
        });
        
        // Process headers to identify sections
        const headers = data[0];
        const emailIndex = headers.indexOf('Email');
        const testDateIndex = headers.indexOf('TestDate');
        
        if (emailIndex === -1 || testDateIndex === -1) {
            throw new Error("Required headers 'Email' and 'TestDate' not found.");
        }

        // Find sections based on separator columns
        const sections = [];
        let currentSection = [];
        
        for (let i = testDateIndex + 1; i<headers.length; i++) {
            const header = headers[i];
            if (header &&header.toLowerCase().includes('separator')) {
                if (currentSection.length> 0) {
                    sections.push(currentSection);
                }
                currentSection = [];
            } else if (header &&header.startsWith('C')) {
                currentSection.push({
                    index: i,
                    name: header
                });
            }
        }
        
        // Add the last section if it exists
        if (currentSection.length> 0) {
            sections.push(currentSection);
        }

        // Process each row of data
        const processedData = data.slice(1).map(row => {
            // Ensure row has enough elements, pad with zeros if needed
            while (row.length<headers.length) {
                row.push(0);
            }

            const results = {
                email: row[emailIndex] || '',
                testDate: row[testDateIndex] || '',
                sections: []
            };

            // Process each section
            sections.forEach((section, sectionIndex) => {
                // Calculate sum of raw values for the section, converting any non-numeric or undefined to 0
                const rawValues = section.map(col => {
                    const value = row[col.index];
                    return (value !== undefined && !isNaN(Number(value))) ? Number(value) : 0;
                });
                const rawSum = rawValues.reduce((sum, val) => sum + val, 0);

                // Get all values from this section type across all rows for z-score calculation
                const allSectionValues = data.slice(1).map(r =>
                    section.reduce((sum, col) => {
                        const value = r[col.index];
                        return sum + ((value !== undefined && !isNaN(Number(value))) ? Number(value) : 0);
                    }, 0)
                );

                const zScore = calculateZScore(rawSum, allSectionValues);
                const stenScore = calculateStenScore(zScore);

                results.sections.push({
                    name: `Section ${sectionIndex + 1}`,
                    rawSum,
                    zScore,
                    stenScore,
                    rawValues // Added for debugging if needed
                });
            });
            
            return results;
        });

        const consolidatedWorkbook = createConsolidatedWorkbook(processedData);
        downloadWorkbook(consolidatedWorkbook, 'processed_results.xlsx');
        
        return consolidatedWorkbook;
       
    } catch (error) {
        console.error("Error during processing:", error);
        throw error;
    }
};

// Function to calculate Z-Score using sum of raw values
const calculateZScore = (rawSum, allSectionSums) => {
    const mean = allSectionSums.reduce((sum, val) => sum + val, 0) / allSectionSums.length;
    const variance = allSectionSums.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / allSectionSums.length;
    const stdDev = Math.sqrt(variance);
    
    // Handle case where all values are same (stdDev = 0)
    if (stdDev === 0) return 0;
    
    return (rawSum - mean) / stdDev;
};

// Function to calculate Sten Score from Z-Score
const calculateStenScore = (zScore) => {
    const stenScore = 5.5 + (zScore * 2);
    return Math.min(Math.max(Math.round(stenScore), 1), 10);
};

// Function to create a consolidated workbook from processed data
const createConsolidatedWorkbook = (processedData) => {
    const wb = XLSX.utils.book_new();
    
    // Create consolidated results sheet
    const consolidatedData = processedData.map(row => ({
        Email: row.email || '',
        TestDate: row.testDate || '',
        ...row.sections.reduce((acc, section) => ({
            ...acc,
            [`${section.name}_RawSum`]: section.rawSum,
            [`${section.name}_ZScore`]: section.zScore.toFixed(2),
            [`${section.name}_StenScore`]: section.stenScore
        }), {})
    }));
    
    const ws = XLSX.utils.json_to_sheet(consolidatedData);
    XLSX.utils.book_append_sheet(wb, ws, 'Consolidated Results');
    
    return wb;
};

// Function to download the workbook
const downloadWorkbook = (workbook, filename) => {
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
    
    const downloadLink = document.createElement('a');
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = filename;
    
    document.body.appendChild(downloadLink);
    downloadLink.click();
    
    document.body.removeChild(downloadLink);
    URL.revokeObjectURL(downloadLink.href);
};

// Utility function to convert string to array buffer
const s2ab = (s) => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i<s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xFF;
    }
    return buf;
};
