export const processDynamicSections = async (file, benchmarks) => {
    try {
        console.log("Starting dynamic section processing...");

        // Read the Excel file
        const workbook = await readExcelFile(file);
        if (!workbook || !workbook.SheetNames.length) {
            throw new Error("No sheets found in workbook.");
        }

        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: 0 });

        // Extract headers and identify columns
        const headers = data[0];
        const emailIndex = headers.indexOf("Email");
        const testDateIndex = headers.indexOf("TestDate");

        if (emailIndex === -1 || testDateIndex === -1) {
            throw new Error("Required headers 'Email' and 'TestDate' not found.");
        }

        // Identify sections by finding Separator columns
        const sections = [];
        let currentSection = [];
        let sectionIndex = 0;

        headers.forEach((header, index) => {
            if (header &&header.toLowerCase().includes("separator")) {
                if (currentSection.length> 0) {
                    sections.push({
                        columns: currentSection,
                        benchmark: benchmarks[sectionIndex++]
                    });
                }
                currentSection = [];
            } else if (header &&header.startsWith("C")) {
                currentSection.push(index);
            }
        });

        // Add the last section if it exists
        if (currentSection.length> 0) {
            sections.push({
                columns: currentSection,
                benchmark: benchmarks[sectionIndex]
            });
        }

        // Validate sections against benchmarks
        if (sections.length !== benchmarks.length) {
            throw new Error(
                `Mismatch between detected sections (${sections.length}) and provided benchmarks (${benchmarks.length}).`
            );
        }

        // Process data rows
        const processedData = data.slice(1).map((row) => {
            const results = {
                email: row[emailIndex] || "",
                testDate: row[testDateIndex] || "",
                sections: []
            };

            sections.forEach((section, idx) => {
                // Calculate raw sum for the section
                const rawSum = section.columns.reduce((sum, colIndex) => {
                    const value = row[colIndex];
                    return sum + (value !== undefined && !isNaN(Number(value)) ? Number(value) : 0);
                }, 0);

                // Calculate Z-score and Sten score using the section's benchmark
                const zScore = (rawSum - section.benchmark.mean) / section.benchmark.sd;
                const stenScore = calculateStenScore(zScore);

                results.sections.push({
                    name: `Section ${idx + 1}`,
                    rawSum,
                    zScore,
                    stenScore
                });
            });

            return results;
        });

        // Create output workbook with detailed results
        const wb = XLSX.utils.book_new();

        const consolidatedData = processedData.map((row) => ({
            Email: row.email,
            TestDate: row.testDate,
            ...row.sections.reduce((acc, section, idx) => ({
                ...acc,
                [`${section.name}_RawSum`]: section.rawSum,
                [`${section.name}_ZScore`]: Number(section.zScore.toFixed(2)),
                [`${section.name}_StenScore`]: section.stenScore,
                [`${section.name}_Benchmark_Mean`]: benchmarks[idx].mean,
                [`${section.name}_Benchmark_SD`]: benchmarks[idx].sd
            }), {})
        }));

        const ws = XLSX.utils.json_to_sheet(consolidatedData);
        XLSX.utils.book_append_sheet(wb, ws, "Results with Benchmarks");

        // Generate and download the file
        const wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
        const blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });

        const downloadLink = document.createElement("a");
        downloadLink.href = URL.createObjectURL(blob);
        downloadLink.download = "benchmark_results.xlsx";
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
        URL.revokeObjectURL(downloadLink.href);

        return wb;

    } catch (error) {
        console.error("Error during section processing:", error);
        throw error;
    }
};

// Helper function for Sten score calculation
const calculateStenScore = (zScore) => {
    const stenScore = Math.round(5.5 + zScore * 2);
    return Math.min(Math.max(stenScore, 1), 10); // Clamp to range [1, 10]
};

// Helper function to convert string to array buffer
const s2ab = (s) => {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i<s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xff;
    }
    return buf;
};

// Function to read the Excel file
const readExcelFile = (file) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            resolve(workbook);
        };
        reader.onerror = (error) => {
            reject(error);
        };
        reader.readAsBinaryString(file);
    });
};
