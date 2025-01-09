class ExcelProcessor {
    constructor() {
        this.rawData = [];
        this.filteredData = [];
        this.headers = [];
        this.filters = new Map();
        this.sortDirection = new Map();
        this.selectedHeaders = new Set();
        this.itemsPerPage = 10;
        this.currentPage = 1;
        this.initializeElements();
        this.setupEventListeners();
    }

    initializeElements() {
        this.dropZone = document.getElementById('dropZone');
        this.fileInput = document.getElementById('fileInput');
        this.fileName = document.getElementById('fileName');
        this.progressBar = document.getElementById('progressBar');
        this.filtersContainer = document.getElementById('filtersContainer');
        this.addCutoffBtn = document.getElementById('addCutoffBtn');
        this.addValueFilterBtn = document.getElementById('addValueFilterBtn');
        this.applyFiltersBtn = document.getElementById('applyFiltersBtn');
        this.tableContainer = document.getElementById('tableContainer');
        this.dataTable = document.getElementById('dataTable');
        this.tableNav = document.getElementById('tableNav');
        this.downloadSelectedBtn = document.getElementById('downloadSelected');
        this.downloadAllFilteredBtn = document.getElementById('downloadAllFilteredData');
        this.headerSelector = document.getElementById('headerSelector');
        this.statsContainer = document.getElementById('statsContainer');
        this.selectedHeadersBox = document.getElementById('selectedHeadersBox');
        
        this.itemsPerPage = 10;
        this.currentPage = 1;
    }

    setupEventListeners() {
        // Previous event listeners remain the same
        this.dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.dropZone.classList.add('dragover');
        });

        this.dropZone.addEventListener('dragleave', () => {
            this.dropZone.classList.remove('dragover');
        });

        this.dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            this.dropZone.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            this.processExcelFile(file);
        });

        this.fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            this.processExcelFile(file);
        });

        // Add event listener for downloading all filtered data
        if (this.downloadAllFilteredBtn) {
            this.downloadAllFilteredBtn.addEventListener('click', () => {
                this.downloadAllFilteredData();
            });
        }

        // Download selected columns button
        if (this.downloadSelectedBtn) {
            this.downloadSelectedBtn.addEventListener('click', () => {
                this.downloadSelectedColumns();
            });
        }

        // Header selection handler
        if (this.headerSelector) {
            this.headerSelector.addEventListener('change', () => {
                const selectedOptions = [...this.headerSelector.selectedOptions];
                this.selectedHeaders = new Set(selectedOptions.map(option => option.value));
                this.updateSelectedHeadersDisplay();
            });
        }

        // Filter buttons handlers
        if (this.addCutoffBtn) {
            this.addCutoffBtn.addEventListener('click', () => this.createCutoffFilter());
        }
        if (this.addValueFilterBtn) {
            this.addValueFilterBtn.addEventListener('click', () => this.createValueFilter());
        }
    }

    updateSelectedHeadersDisplay() {
        if (!this.selectedHeadersBox) return;

        this.selectedHeadersBox.innerHTML = '<strong>Selected Columns:</strong><br>';
        [...this.selectedHeaders].forEach(header => {
            const headerSpan = document.createElement('span');
            headerSpan.className = 'selected-header-tag';
            headerSpan.textContent = header;
            this.selectedHeadersBox.appendChild(headerSpan);
        });
    }

    async processExcelFile(file) {
        if (!file || !file.name.match(/\.(xlsx|xls)$/)) {
            alert('Please select a valid Excel file');
            return;
        }

        this.fileName.textContent = file.name;
        this.showProgress();

        try {
            const data = await this.readExcelFile(file);
            this.rawData = data;
            this.filteredData = [...data];
            this.headers = Object.keys(data[0]);
            this.initializeHeaderSelector();
            this.displayData();
            this.enableButtons();
            this.updateStats();
        } catch (error) {
            console.error('Error processing file:', error);
            alert('Error processing the Excel file');
        } finally {
            this.hideProgress();
        }
    }

    initializeHeaderSelector() {
        if (!this.headerSelector) return;
    
        this.headerSelector.innerHTML = '';
        this.headers.forEach(header => {
            const option = document.createElement('option');
            option.value = header;
            option.textContent = header;
            this.headerSelector.appendChild(option);
        });
    
        this.selectedHeaders = new Set();
    
        // Event Listener to Update Selected Headers
        this.headerSelector.addEventListener('change', () => {
            const selectedOptions = [...this.headerSelector.selectedOptions];
            this.selectedHeaders = new Set(selectedOptions.map(option => option.value));
            this.updateSelectedHeadersDisplay();
        });
    
        this.updateSelectedHeadersDisplay();
    }
    

    updateStats() {
        if (!this.statsContainer) return;

        const totalRecords = this.rawData.length;
        const selectedRecords = this.filteredData.length;
        const selectPercentage = ((selectedRecords / totalRecords) * 100).toFixed(2);

        this.statsContainer.innerHTML = `
            <div class="stats">
                <p>Total Records: ${totalRecords}</p>
                <p>Selected Records: ${selectedRecords}</p>
                <p>Selection Percentage: ${selectPercentage}%</p>
            </div>
        `;
    }

    // downloadSelectedColumns() {
    //     if (this.selectedHeaders.size === 0) {
    //         alert('Please select at least one column to download');
    //         return;
    //     }

    //     const headersToInclude = [...this.selectedHeaders];
    //     if (!headersToInclude.includes('Assessment Status')) {
    //         headersToInclude.push('Assessment Status');
    //     }

    //     const processedData = this.filteredData.map(row => {
    //         const newRow = {};
    //         headersToInclude.forEach(header => {
    //             newRow[header] = row[header];
    //         });
    //         return newRow;
    //     });

    //     this.downloadExcel(processedData, 'selected_columns');
    // }

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
    // downloadAllFilteredData() {
    //     if (this.filteredData.length === 0) {
    //         alert('No filtered data to download');
    //         return;
    //     }

    //     // Create workbook
    //     const wb = XLSX.utils.book_new();
        
    //     // Add statistics row at the top
    //     const totalRecords = this.rawData.length;
    //     const selectedRecords = this.filteredData.filter(row => row['Assessment Status'] === 'Test Select').length;
    //     const selectPercentage = ((selectedRecords / totalRecords) * 100).toFixed(2);

    //     const statsRow = {
    //         [this.headers[0]]: 'Statistics',
    //         [this.headers[1]]: `Total Records: ${totalRecords}`,
    //         [this.headers[2]]: `Selected Records: ${selectedRecords}`,
    //         [this.headers[3]]: `Selection Percentage: ${selectPercentage}%`
    //     };

    //     const dataToDownload = [statsRow, ...this.filteredData];
        
    //     // Convert to worksheet
    //     const ws = XLSX.utils.json_to_sheet(dataToDownload);
    //     XLSX.utils.book_append_sheet(wb, ws, 'Filtered Data');

    //     // Generate filename with timestamp
    //     const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    //     const filename = `all_filtered_data_${timestamp}.xlsx`;

    //     // Write and download file
    //     XLSX.writeFile(wb, filename);
    // }

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
    // downloadExcel(data, prefix) {
    //     // Add statistics row
    //     const totalRecords = data.length;
    //     const selectedRecords = data.filter(row => row['Assessment Status'] === 'Test Select').length;
    //     const selectPercentage = ((selectedRecords / totalRecords) * 100).toFixed(2);

    //     const statsRow = {
    //         [Object.keys(data[0])[0]]: 'Statistics',
    //         [Object.keys(data[0])[1]]: `Total Records: ${totalRecords}`,
    //         [Object.keys(data[0])[2]]: `Test Select: ${selectedRecords}`,
    //         [Object.keys(data[0])[3]]: `Selection %: ${selectPercentage}%`
    //     };

    //     const dataToDownload = [statsRow, ...data];

    //     // Create and download Excel file
    //     const wb = XLSX.utils.book_new();
    //     const ws = XLSX.utils.json_to_sheet(dataToDownload);
    //     XLSX.utils.book_append_sheet(wb, ws, 'Data');

    //     const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    //     const filename = `${prefix}_${timestamp}.xlsx`;

    //     XLSX.writeFile(wb, filename);
    // }

    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    resolve(jsonData);
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    createCutoffFilter() {
        const filterGroup = document.createElement('div');
        filterGroup.className = 'filter-group';

        // Header select
        const headerSelect = document.createElement('select');
        headerSelect.className = 'filter-input';
        this.headers.forEach(header => {
            const option = document.createElement('option');
            option.value = header;
            option.textContent = header;
            headerSelect.appendChild(option);
        });

        // Operator select
        const operatorSelect = document.createElement('select');
        operatorSelect.className = 'filter-input';
        const operators = ['>=', '<=', '=', '>', '<'];
        operators.forEach(op => {
            const option = document.createElement('option');
            option.value = op;
            option.textContent = op;
            operatorSelect.appendChild(option);
        });

        // Cutoff input
        const cutoffInput = document.createElement('input');
        cutoffInput.type = 'number';
        cutoffInput.className = 'filter-input';
        cutoffInput.placeholder = 'Enter cutoff value';
        cutoffInput.step = 'any';

        // Remove button
        const removeBtn = document.createElement('button');
        removeBtn.className = 'button secondary';
        removeBtn.innerHTML = '<i class="fas fa-times"></i>';
        removeBtn.onclick = () =>filterGroup.remove();

        filterGroup.appendChild(headerSelect);
        filterGroup.appendChild(operatorSelect);
        filterGroup.appendChild(cutoffInput);
        filterGroup.appendChild(removeBtn);
        this.filtersContainer.appendChild(filterGroup);
    }

    createValueFilter() {
        const filterGroup = document.createElement('div');
        filterGroup.className = 'filter-group';

        // Header select
        const headerSelect = document.createElement('select');
        headerSelect.className = 'filter-input';
        this.headers.forEach(header => {
            const option = document.createElement('option');
            option.value = header;
            option.textContent = header;
            headerSelect.appendChild(option);
        });

        // Filter value dropdown
        const valueSelect = document.createElement('select');
        valueSelect.className = 'filter-input';

        // Update values when header is selected
        headerSelect.addEventListener('change', () => {
            const selectedHeader = headerSelect.value;
            valueSelect.innerHTML = '';
            
            // Add a default "Select value" option
            const defaultOption = document.createElement('option');
            defaultOption.value = '';
            defaultOption.textContent = 'Select value';
            valueSelect.appendChild(defaultOption);
            
            const uniqueValues = [...new Set(this.rawData.map(row => row[selectedHeader]))];
            uniqueValues.sort((a, b) =>a?.toString().localeCompare(b?.toString()));
            
            uniqueValues.forEach(value => {
                if (value !== undefined && value !== null) {
                    const option = document.createElement('option');
                    option.value = value;
                    option.textContent = value;
                    valueSelect.appendChild(option);
                }
            });
        });

        // Trigger initial population of values
        headerSelect.dispatchEvent(new Event('change'));

        // Remove button
        const removeBtn = document.createElement('button');
        removeBtn.className = 'button secondary';
        removeBtn.innerHTML = '<i class="fas fa-times"></i>';
        removeBtn.onclick = () =>filterGroup.remove();

        filterGroup.appendChild(headerSelect);
        filterGroup.appendChild(valueSelect);
        filterGroup.appendChild(removeBtn);
        this.filtersContainer.appendChild(filterGroup);
    }

    // applyFilters() {
    //     // Reset filtered data
    //     this.filteredData = [...this.rawData];
        
    //     const filterGroups = this.filtersContainer.querySelectorAll('.filter-group');
        
    //     // Apply value filters first
    //     filterGroups.forEach(group => {
    //         if (!group.querySelector('input')) { // Value filter
    //             const headerSelect = group.querySelector('select');
    //             const valueSelect = group.querySelector('select:nth-child(2)');
    //             const header = headerSelect.value;
    //             const value = valueSelect.value;
                
    //             if (value) {
    //                 this.filteredData = this.filteredData.filter(row => 
    //                     row[header]?.toString() === value
    //                 );
    //             }
    //         }
    //     });

    //     // Then apply cutoff filters and mark Test Select/Reject
    //     const cutoffFilteredData = [...this.filteredData];
    //     filterGroups.forEach(group => {
    //         const inputElement = group.querySelector('input');
    //         if (inputElement) { // Cutoff filter
    //             const headerSelect = group.querySelector('select');
    //             const operatorSelect = group.querySelector('select:nth-child(2)');
    //             const header = headerSelect.value;
    //             const operator = operatorSelect.value;
    //             const cutoffValue = parseFloat(inputElement.value);

    //             if (!isNaN(cutoffValue)) {
    //                 cutoffFilteredData.forEach(row => {
    //                     const value = parseFloat(row[header]);
    //                     let passes = false;
                        
    //                     if (!isNaN(value)) {
    //                         switch (operator) {
    //                             case '>=': passes = value >= cutoffValue; break;
    //                             case '<=': passes = value <= cutoffValue; break;
    //                             case '=': passes = value === cutoffValue; break;
    //                             case '>': passes = value > cutoffValue; break;
    //                             case '<': passes = value < cutoffValue; break;
    //                         }
    //                     }
                        
    //                     row['Assessment Status'] = passes ? 'Test Select' : 'Test Reject';
    //                 });
    //             }
    //         }
    //     });

    //     this.filteredData = cutoffFilteredData;
    //     this.currentPage = 1;
    //     this.displayData();
    //     this.updateStats();
    // }

     applyFilters() {
        // Reset filtered data
        this.filteredData = [...this.rawData];
        
        const filterGroups = this.filtersContainer.querySelectorAll('.filter-group');
        
        // Apply all filters
        filterGroups.forEach(group => {
            const headerSelect = group.querySelector('select');
            const header = headerSelect.value;
            const inputElement = group.querySelector('input');
            
            if (!inputElement) { // Value filter
                const valueSelect = group.querySelector('select:nth-child(2)');
                const selectedValue = valueSelect.value;
                if (selectedValue) {
                    this.filteredData = this.filteredData.filter(row => 
                        row[header]?.toString() === selectedValue
                    );
                }
            } else { // Cutoff filter
                const operatorSelect = group.querySelector('select:nth-child(2)');
                const operator = operatorSelect.value;
                const cutoffValue = parseFloat(inputElement.value);

                if (!isNaN(cutoffValue)) {
                    this.filteredData = this.filteredData.filter(row => {
                        const value = parseFloat(row[header]);
                        if (isNaN(value)) return false;
                        
                        switch (operator) {
                            case '>=': return value >= cutoffValue;
                            case '<=': return value <= cutoffValue;
                            case '=': return value === cutoffValue;
                            case '>': return value > cutoffValue;
                            case '<': return value < cutoffValue;
                            default: return false;
                        }
                    });
                }
            }
        });

        // Mark filtered data as Test Select and unfiltered as Test Reject
        const filteredIds = new Set(this.filteredData.map(row => JSON.stringify(row)));
        this.rawData.forEach(row => {
            row['Assessment Status'] = filteredIds.has(JSON.stringify(row)) ? 'Test Select' : 'Test Reject';
        });

        this.currentPage = 1;
        this.displayData();
        this.updateStats();
    }
    updateStats() {
        if (!this.statsContainer) return;

        // Check if any value filters are applied
        const hasValueFilters = Array.from(this.filtersContainer.querySelectorAll('.filter-group'))
            .some(group => !group.querySelector('input')); // Value filters don't have input elements

        let baseData;
        if (hasValueFilters) {
            // If value filters exist, apply only value filters first to get the base data
            baseData = [...this.rawData];
            const filterGroups = this.filtersContainer.querySelectorAll('.filter-group');
            
            filterGroups.forEach(group => {
                const headerSelect = group.querySelector('select');
                const header = headerSelect.value;
                const inputElement = group.querySelector('input');
                
                if (!inputElement) { // Value filter
                    const valueSelect = group.querySelector('select:nth-child(2)');
                    const selectedValue = valueSelect.value;
                    if (selectedValue) {
                        baseData = baseData.filter(row => 
                            row[header]?.toString() === selectedValue
                        );
                    }
                }
            });
        } else {
            // If no value filters, use all raw data as base
            baseData = this.rawData;
        }

        const totalRecords = baseData.length;
        const selectedRecords = this.filteredData.filter(row => row['Assessment Status'] === 'Test Select').length;
        const selectPercentage = totalRecords > 0 ? ((selectedRecords / totalRecords) * 100).toFixed(2) : '0.00';

        this.statsContainer.innerHTML = `
            <div class="stats">
                <p>Total Records: ${totalRecords}</p>
                <p>Selected Records: ${selectedRecords}</p>
                <p>Selection Percentage: ${selectPercentage}%</p>
            </div>
        `;
    }
  
    downloadData() {
        if (this.rawData.length === 0) {
            alert('No data to download');
            return;
        }

        // Get selected headers or use all if none selected
        const headersToInclude = [...this.selectedHeaders];
        if (headersToInclude.length === 0) {
            alert('Please select at least one column to download');
            return;
        }

        // Add Assessment Status to headers if not already included
        if (!headersToInclude.includes('Assessment Status')) {
            headersToInclude.push('Assessment Status');
        }

        // Create a copy of raw data to process
        let allData = this.rawData.map(row => ({ ...row }));

        // First, identify records that pass value filters
        const valueFilterGroups = Array.from(this.filtersContainer.querySelectorAll('.filter-group'))
            .filter(group => !group.querySelector('input'));

        if (valueFilterGroups.length > 0) {
            // Apply value filters and mark records
            allData = allData.map(row => {
                let passesValueFilters = true;
                
                valueFilterGroups.forEach(group => {
                    const headerSelect = group.querySelector('select');
                    const valueSelect = group.querySelector('select:nth-child(2)');
                    const header = headerSelect.value;
                    const selectedValue = valueSelect.value;
                    
                    if (selectedValue && row[header]?.toString() !== selectedValue) {
                        passesValueFilters = false;
                    }
                });

                return {
                    ...row,
                    passesValueFilters
                };
            });
        } else {
            // If no value filters, all records pass value filter stage
            allData = allData.map(row => ({
                ...row,
                passesValueFilters: true
            }));
        }

        // Apply cutoff filters to records that passed value filters
        const cutoffFilterGroups = Array.from(this.filtersContainer.querySelectorAll('.filter-group'))
            .filter(group => group.querySelector('input'));

        allData = allData.map(row => {
            // Only apply cutoff filters to rows that passed value filters
            if (row.passesValueFilters) {
                let passesCutoffFilters = true;

                cutoffFilterGroups.forEach(group => {
                    const headerSelect = group.querySelector('select');
                    const operatorSelect = group.querySelector('select:nth-child(2)');
                    const inputElement = group.querySelector('input');
                    
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
                            if (!passes) {
                                passesCutoffFilters = false;
                            }
                        } else {
                            passesCutoffFilters = false;
                        }
                    }
                });

                row['Assessment Status'] = passesCutoffFilters ? 'Test Select' : 'Test Reject';
            } else {
                // Records that didn't pass value filters are marked as Test Reject
                row['Assessment Status'] = 'Test Reject';
            }

            return row;
        });

        // Remove the temporary passesValueFilters property
        allData = allData.map(({ passesValueFilters, ...rest }) => rest);

        // Prepare data for download with selected columns
        const processedData = allData.map(row => {
            const newRow = {};
            headersToInclude.forEach(header => {
                newRow[header] = row[header];
            });
            return newRow;
        });

        // Calculate statistics
        const totalRecords = processedData.length;
        const selectedRecords = processedData.filter(row => row['Assessment Status'] === 'Test Select').length;
        const rejectedRecords = processedData.filter(row => row['Assessment Status'] === 'Test Reject').length;
        const selectPercentage = totalRecords > 0 ? ((selectedRecords / totalRecords) * 100).toFixed(2) : '0.00';

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

        const dataToDownload = [...statsRows, ...processedData];

        // Create and download Excel file
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(dataToDownload);
        XLSX.utils.book_append_sheet(wb, ws, 'Data');

        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const filename = `processed_data_${timestamp}.xlsx`;

        XLSX.writeFile(wb, filename);
    }


    // downloadData() {
    //     if (this.filteredData.length === 0) {
    //         alert('No data to download');
    //         return;
    //     }

    //     // Get selected headers or use all if none selected
    //     const headersToInclude = [...this.selectedHeaders];
    //     if (headersToInclude.length === 0) {
    //         alert('Please select at least one column to download');
    //         return;
    //     }

    //     // Add Assessment Status to headers if not already included
    //     if (!headersToInclude.includes('Assessment Status')) {
    //         headersToInclude.push('Assessment Status');
    //     }

    //     // Check if value filters are applied
    //     const hasValueFilters = Array.from(this.filtersContainer.querySelectorAll('.filter-group'))
    //         .some(group => !group.querySelector('input'));

    //     let baseData;
    //     if (hasValueFilters) {
    //         // Apply only value filters to get base data
    //         baseData = [...this.rawData];
    //         const filterGroups = this.filtersContainer.querySelectorAll('.filter-group');
            
    //         filterGroups.forEach(group => {
    //             const headerSelect = group.querySelector('select');
    //             const header = headerSelect.value;
    //             const inputElement = group.querySelector('input');
                
    //             if (!inputElement) { // Value filter
    //                 const valueSelect = group.querySelector('select:nth-child(2)');
    //                 const selectedValue = valueSelect.value;
    //                 if (selectedValue) {
    //                     baseData = baseData.filter(row => 
    //                         row[header]?.toString() === selectedValue
    //                     );
    //                 }
    //             }
    //         });
    //     } else {
    //         baseData = this.rawData;
    //     }

    //     // Filter data to include only selected headers
    //     const processedData = this.filteredData.map(row => {
    //         const newRow = {};
    //         headersToInclude.forEach(header => {
    //             newRow[header] = row[header];
    //         });
    //         return newRow;
    //     });

    //     // Add statistics row with correct total based on filter state
    //     const totalRecords = baseData.length;
    //     const selectedRecords = this.filteredData.filter(row => row['Assessment Status'] === 'Test Select').length;
    //     const selectPercentage = totalRecords > 0 ? ((selectedRecords / totalRecords) * 100).toFixed(2) : '0.00';

    //     const statsRow = {
    //         [headersToInclude[0]]: 'Statistics',
    //         [headersToInclude[1]]: `Total Records: ${totalRecords}`,
    //         [headersToInclude[2]]: `Selected: ${selectedRecords}`,
    //         [headersToInclude[3]]: `Selection %: ${selectPercentage}%`
    //     };

    //     const dataToDownload = [statsRow, ...processedData];

    //     // Create workbook and worksheet
    //     const wb = XLSX.utils.book_new();
    //     const ws = XLSX.utils.json_to_sheet(dataToDownload);
    //     XLSX.utils.book_append_sheet(wb, ws, 'Data');

    //     // Generate filename
    //     const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    //     const filename = `processed_data_${timestamp}.xlsx`;

    //     // Download file
    //     XLSX.writeFile(wb, filename);
    // }

    // downloadData(type) {
    //     let dataToDownload;
        
    //     if (type === 'selected') {
    //         // For selected data, use the filtered data with 'Test Select' status
    //         dataToDownload = this.filteredData;
    //     } else {
    //         // For rejected data, get the complement of filtered data and set status to 'Test Reject'
    //         const filteredIds = new Set(this.filteredData.map(row =>JSON.stringify(row)));
    //         dataToDownload = this.rawData
    //             .filter(row => !filteredIds.has(JSON.stringify({...row, 'Assessment Status': 'Test Select'})))
    //             .map(row => ({
    //                 ...row,
    //                 'Assessment Status': 'Test Reject'
    //             }));
    //     }

    //     if (dataToDownload.length === 0) {
    //         alert('No data to download');
    //         return;
    //     }

    //     // Create workbook and worksheet
    //     const wb = XLSX.utils.book_new();
    //     const ws = XLSX.utils.json_to_sheet(dataToDownload);
    //     XLSX.utils.book_append_sheet(wb, ws, 'Data');

    //     // Generate filename
    //     const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    //     const filename = `${type}_data_${timestamp}.xlsx`;

    //     // Download file
    //     XLSX.writeFile(wb, filename);
    // }

    displayData() {
        if (!this.dataTable) return;
        
        this.dataTable.innerHTML = '';
        
        // Check if there's data to display
        const dataToDisplay = this.filteredData;
        if (!dataToDisplay || dataToDisplay.length === 0) {
            const noDataRow = document.createElement('tr');
            const noDataCell = document.createElement('td');
            noDataCell.colSpan = this.headers.length || 1;
            noDataCell.textContent = 'No data available';
            noDataCell.style.textAlign = 'center';
            noDataRow.appendChild(noDataCell);
            this.dataTable.appendChild(noDataRow);
            return;
        }

        // Create header row
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        
        this.headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            th.style.cursor = 'pointer';
            th.onclick = () => this.sortData(header);
            
            // Add sort direction indicator
            const direction = this.sortDirection.get(header);
            if (direction) {
                th.textContent += direction === 'asc' ? ' ↑' : ' ↓';
            }
            
            headerRow.appendChild(th);
        });
        
        thead.appendChild(headerRow);
        this.dataTable.appendChild(thead);

        // Create table body
        const tbody = document.createElement('tbody');
        
        // Calculate pagination
        const start = (this.currentPage - 1) * this.itemsPerPage;
        const end = start + this.itemsPerPage;
        const paginatedData = dataToDisplay.slice(start, end);

        // Create data rows
        paginatedData.forEach(row => {
            const tr = document.createElement('tr');
            this.headers.forEach(header => {
                const td = document.createElement('td');
                td.textContent = row[header] !== undefined ? row[header] : '';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        this.dataTable.appendChild(tbody);
        this.initializeTableNav();
    }

    initializeTableNav() {
        if (!this.tableNav) return;
        
        const totalPages = Math.ceil(this.filteredData.length / this.itemsPerPage);
        this.tableNav.innerHTML = '';

        if (totalPages <= 1) {
            return; // Don't show navigation if there's only one page
        }

        // Previous button
        const prevBtn = document.createElement('button');
        prevBtn.textContent = '←';
        prevBtn.className = 'button secondary';
        prevBtn.onclick = () => this.changePage(this.currentPage - 1);
        prevBtn.disabled = this.currentPage === 1;

        // Page number
        const pageText = document.createElement('span');
        pageText.textContent = `Page ${this.currentPage} of ${totalPages}`;
        pageText.className = 'page-text';

        // Next button
        const nextBtn = document.createElement('button');
        nextBtn.textContent = '→';
        nextBtn.className = 'button secondary';
        nextBtn.onclick = () => this.changePage(this.currentPage + 1);
        nextBtn.disabled = this.currentPage === totalPages;

        this.tableNav.appendChild(prevBtn);
        this.tableNav.appendChild(pageText);
        this.tableNav.appendChild(nextBtn);
    }

    changePage(pageNumber) {
        const totalPages = Math.ceil(this.filteredData.length / this.itemsPerPage);
        if (pageNumber < 1 || pageNumber > totalPages) return;

        this.currentPage = pageNumber;
        this.displayData();
    }

    sortData(header) {
        const direction = this.sortDirection.get(header) === 'asc' ? 'desc' : 'asc';
        this.sortDirection.clear(); // Clear previous sorting
        this.sortDirection.set(header, direction);
        
        const isNumeric = this.rawData.every(row => !isNaN(row[header]));
        
        this.filteredData.sort((a, b) => {
            let comparison;
            if (isNumeric) {
                comparison = parseFloat(a[header]) - parseFloat(b[header]);
            } else {
                comparison = String(a[header]).localeCompare(String(b[header]));
            }
            return direction === 'asc' ? comparison : -comparison;
        });

        this.displayData(this.filteredData);
    }

    showProgress() {
        if (this.progressBar) {
            this.progressBar.style.display = 'block';
            this.progressBar.querySelector('.progress-bar').style.width = '50%';
        }
    }

    hideProgress() {
        if (this.progressBar) {
            this.progressBar.style.display = 'none';
            this.progressBar.querySelector('.progress-bar').style.width = '0%';
        }
    }

    enableButtons() {
        if (this.downloadSelectedBtn) {
            this.downloadSelectedBtn.disabled = false;
        }
        if (this.downloadAllFilteredBtn) {
            this.downloadAllFilteredBtn.disabled = false;
        }
    }


    downloadData(type) {
        const dataToDownload = type === 'selected' ?this.filteredData : 
            this.rawData.filter(row => !this.filteredData.includes(row));
        
        if (dataToDownload.length === 0) {
            alert('No data to download');
            return;
        }

        // Create workbook and worksheet
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(dataToDownload);
        XLSX.utils.book_append_sheet(wb, ws, 'Data');

        // Generate filename
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const filename = `${type}_data_${timestamp}.xlsx`;

        // Download file
        XLSX.writeFile(wb, filename);
    }
}

// Initialize the processor when the DOM is fully loaded
document.addEventListener('DOMContentLoaded', () => {
    window.excelProcessor = new ExcelProcessor();
});
