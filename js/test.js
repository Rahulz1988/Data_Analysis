class ExcelProcessor {
    // ... (keep existing constructor and other methods the same)

    // Update the downloadSelectedColumns method
    downloadSelectedColumns() {
        if (this.selectedHeaders.size === 0) {
            alert('Please select at least one column to download');
            return;
        }

        // Get headers to include
        const headersToInclude = [...this.selectedHeaders];
        if (!headersToInclude.includes('Assessment Status')) {
            headersToInclude.push('Assessment Status');
        }

        // Get all records based on value filters
        let filteredData = [...this.rawData];
        const filterGroups = this.filtersContainer.querySelectorAll('.filter-group');

        // Apply value filters
        filterGroups.forEach(group => {
            if (!group.querySelector('input')) { // Value filter
                const headerSelect = group.querySelector('select');
                const valueSelect = group.querySelector('select:nth-child(2)');
                const header = headerSelect.value;
                const selectedValue = valueSelect.value;
                
                if (selectedValue) {
                    filteredData = filteredData.filter(row => 
                        row[header]?.toString() === selectedValue
                    );
                }
            }
        });

        // Apply cutoff filters and mark status
        filteredData.forEach(row => {
            let passesAllCutoffs = true;
            
            filterGroups.forEach(group => {
                const inputElement = group.querySelector('input');
                if (inputElement) { // Cutoff filter
                    const headerSelect = group.querySelector('select');
                    const operatorSelect = group.querySelector('select:nth-child(2)');
                    const header = headerSelect.value;
                    const operator = operatorSelect.value;
                    const cutoffValue = parseFloat(inputElement.value);

                    if (!isNaN(cutoffValue)) {
                        const value = parseFloat(row[header]);
                        if (!isNaN(value)) {
                            let passes = false;
                            switch (operator) {
                                case '>=': passes = value >= cutoffValue; break;
                                case '<=': passes = value <= cutoffValue; break;
                                case '=': passes = value === cutoffValue; break;
                                case '>': passes = value > cutoffValue; break;
                                case '<': passes = value < cutoffValue; break;
                            }
                            if (!passes) passesAllCutoffs = false;
                        } else {
                            passesAllCutoffs = false;
                        }
                    }
                }
            });

            row['Assessment Status'] = passesAllCutoffs ? 'Test Select' : 'Test Reject';
        });

        // Process data with selected columns
        const processedData = filteredData.map(row => {
            const newRow = {};
            headersToInclude.forEach(header => {
                newRow[header] = row[header];
            });
            return newRow;
        });

        // Calculate statistics
        const totalRecords = filteredData.length;
        const selectedRecords = filteredData.filter(row => row['Assessment Status'] === 'Test Select').length;
        const rejectedRecords = totalRecords - selectedRecords;
        const selectPercentage = ((selectedRecords / totalRecords) * 100).toFixed(2);

        // Add statistics rows
        const statsRows = [
            {
                [headersToInclude[0]]: 'Statistics',
                [headersToInclude[1]]: `Total Records: ${totalRecords}`,
                [headersToInclude[2]]: `Selected: ${selectedRecords}`,
                [headersToInclude[3]]: `Rejected: ${rejectedRecords}`
            },
            {
                [headersToInclude[0]]: '',
                [headersToInclude[1]]: `Selection %: ${selectPercentage}%`,
                [headersToInclude[2]]: `Reject %: ${(100 - parseFloat(selectPercentage)).toFixed(2)}%`,
                [headersToInclude[3]]: ''
            }
        ];

        // Download the data
        this.downloadExcel([...statsRows, ...processedData], 'selected_columns');
    }

    // Update the downloadAllFilteredData method
    downloadAllFilteredData() {
        // Get all records based on value filters
        let filteredData = [...this.rawData];
        const filterGroups = this.filtersContainer.querySelectorAll('.filter-group');

        // Apply value filters
        filterGroups.forEach(group => {
            if (!group.querySelector('input')) { // Value filter
                const headerSelect = group.querySelector('select');
                const valueSelect = group.querySelector('select:nth-child(2)');
                const header = headerSelect.value;
                const selectedValue = valueSelect.value;
                
                if (selectedValue) {
                    filteredData = filteredData.filter(row => 
                        row[header]?.toString() === selectedValue
                    );
                }
            }
        });

        // Apply cutoff filters and mark status
        filteredData.forEach(row => {
            let passesAllCutoffs = true;
            
            filterGroups.forEach(group => {
                const inputElement = group.querySelector('input');
                if (inputElement) { // Cutoff filter
                    const headerSelect = group.querySelector('select');
                    const operatorSelect = group.querySelector('select:nth-child(2)');
                    const header = headerSelect.value;
                    const operator = operatorSelect.value;
                    const cutoffValue = parseFloat(inputElement.value);

                    if (!isNaN(cutoffValue)) {
                        const value = parseFloat(row[header]);
                        if (!isNaN(value)) {
                            let passes = false;
                            switch (operator) {
                                case '>=': passes = value >= cutoffValue; break;
                                case '<=': passes = value <= cutoffValue; break;
                                case '=': passes = value === cutoffValue; break;
                                case '>': passes = value > cutoffValue; break;
                                case '<': passes = value < cutoffValue; break;
                            }
                            if (!passes) passesAllCutoffs = false;
                        } else {
                            passesAllCutoffs = false;
                        }
                    }
                }
            });

            row['Assessment Status'] = passesAllCutoffs ? 'Test Select' : 'Test Reject';
        });

        // Calculate statistics
        const totalRecords = filteredData.length;
        const selectedRecords = filteredData.filter(row => row['Assessment Status'] === 'Test Select').length;
        const rejectedRecords = totalRecords - selectedRecords;
        const selectPercentage = ((selectedRecords / totalRecords) * 100).toFixed(2);

        // Add statistics rows
        const statsRows = [
            {
                [this.headers[0]]: 'Statistics',
                [this.headers[1]]: `Total Records: ${totalRecords}`,
                [this.headers[2]]: `Selected: ${selectedRecords}`,
                [this.headers[3]]: `Rejected: ${rejectedRecords}`
            },
            {
                [this.headers[0]]: '',
                [this.headers[1]]: `Selection %: ${selectPercentage}%`,
                [this.headers[2]]: `Reject %: ${(100 - parseFloat(selectPercentage)).toFixed(2)}%`,
                [this.headers[3]]: ''
            }
        ];

        // Download the data
        this.downloadExcel([...statsRows, ...filteredData], 'all_filtered_data');
    }

    downloadExcel(data, prefix) {
        if (data.length === 0) {
            alert('No data to download');
            return;
        }

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, 'Data');

        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const filename = `${prefix}_${timestamp}.xlsx`;

        XLSX.writeFile(wb, filename);
    }
}