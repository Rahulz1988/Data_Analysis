class ExcelProcessor {
    constructor() {
        this.rawData = [];
        this.filteredData = [];
        this.headers = [];
        this.filters = new Map();
        this.sortDirection = new Map();
        this.initializeElements();
        this.setupEventListeners();
    }

    initializeElements() {
        // Upload elements
        this.dropZone = document.getElementById('dropZone');
        this.fileInput = document.getElementById('fileInput');
        this.fileName = document.getElementById('fileName');
        this.progressBar = document.getElementById('progressBar');

        // Filter elements
        this.filtersContainer = document.getElementById('filtersContainer');
        this.addCutoffBtn = document.getElementById('addCutoffBtn');
        this.addValueFilterBtn = document.getElementById('addValueFilterBtn');
        this.applyFiltersBtn = document.getElementById('applyFiltersBtn');
        
        // Table elements
        this.tableContainer = document.getElementById('tableContainer');
        this.dataTable = document.getElementById('dataTable');
        this.tableNav = document.getElementById('tableNav');
        
        // Action buttons
        this.downloadSelectedBtn = document.getElementById('downloadSelected');
        this.downloadRejectedBtn = document.getElementById('downloadRejected');
        
        // Pagination
        this.itemsPerPage = 10;
        this.currentPage = 1;
    }

    setupEventListeners() {
        // File upload listeners
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

        // Filter buttons listeners
        if (this.addCutoffBtn) {
            this.addCutoffBtn.addEventListener('click', () =>this.createCutoffFilter());
        }
        if (this.addValueFilterBtn) {
            this.addValueFilterBtn.addEventListener('click', () =>this.createValueFilter());
        }
        if (this.applyFiltersBtn) {
            this.applyFiltersBtn.addEventListener('click', () =>this.applyFilters());
        }

        // Download buttons listeners
        if (this.downloadSelectedBtn) {
            this.downloadSelectedBtn.addEventListener('click', () =>this.downloadData('selected'));
        }
        if (this.downloadRejectedBtn) {
            this.downloadRejectedBtn.addEventListener('click', () =>this.downloadData('rejected'));
        }
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
            this.displayData(this.filteredData);
            this.enableButtons();
            this.initializeTableNav();
        } catch (error) {
            console.error('Error processing file:', error);
            alert('Error processing the Excel file');
        } finally {
            this.hideProgress();
        }
    }

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

    applyFilters() {
        // Start with all data
        this.filteredData = [...this.rawData];
        
        // Get all filter groups
        const filterGroups = this.filtersContainer.querySelectorAll('.filter-group');
        
        filterGroups.forEach(group => {
            const headerSelect = group.querySelector('select');
            const header = headerSelect.value;
            
            // Check if it's a cutoff filter (has input element) or value filter
            const inputElement = group.querySelector('input');
            
            if (inputElement) {
                // Cutoff filter
                const operator = group.querySelector('select:nth-child(2)').value;
                const cutoffValue = parseFloat(inputElement.value);
                
                if (!isNaN(cutoffValue)) {
                    this.filteredData = this.filteredData.filter(row => {
                        const value = parseFloat(row[header]);
                        switch (operator) {
                            case '>=': return value >= cutoffValue;
                            case '<=': return value <= cutoffValue;
                            case '=': return value === cutoffValue;
                            case '>': return value >cutoffValue;
                            case '<': return value <cutoffValue;
                            default: return true;
                        }
                    });
                }
            } else {
                // Value filter
                const valueSelect = group.querySelector('select:nth-child(2)');
                const selectedValue = valueSelect.value;
                
                if (selectedValue) {
                    this.filteredData = this.filteredData.filter(row =>
                        row[header]?.toString() === selectedValue
                    );
                }
            }
        });

        // Update Assessment Status to 'Test Select' for filtered data
        this.filteredData = this.filteredData.map(row => ({
            ...row,
            'Assessment Status': 'Test Select'
        }));

        this.currentPage = 1;
        this.displayData(this.filteredData);
    }

    downloadData(type) {
        let dataToDownload;
        
        if (type === 'selected') {
            // For selected data, use the filtered data with 'Test Select' status
            dataToDownload = this.filteredData;
        } else {
            // For rejected data, get the complement of filtered data and set status to 'Test Reject'
            const filteredIds = new Set(this.filteredData.map(row =>JSON.stringify(row)));
            dataToDownload = this.rawData
                .filter(row => !filteredIds.has(JSON.stringify({...row, 'Assessment Status': 'Test Select'})))
                .map(row => ({
                    ...row,
                    'Assessment Status': 'Test Reject'
                }));
        }

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

    displayData(data) {
        if (!this.dataTable) return;
        
        this.dataTable.innerHTML = '';
        
        if (!data || data.length === 0) {
            this.dataTable.innerHTML = '<tr><td colspan="100%">No data available</td></tr>';
            return;
        }

        // Create header row
        const headerRow = document.createElement('tr');
        this.headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            th.onclick = () =>this.sortData(header);
            th.style.cursor = 'pointer';
            
            // Add sort direction indicator
            const direction = this.sortDirection.get(header);
            if (direction) {
                th.textContent += direction === 'asc' ? ' ↑' : ' ↓';
            }
            
            headerRow.appendChild(th);
        });
        this.dataTable.appendChild(headerRow);

        // Calculate pagination
        const start = (this.currentPage - 1) * this.itemsPerPage;
        const end = start + this.itemsPerPage;
        const paginatedData = data.slice(start, end);

        // Create data rows
        paginatedData.forEach(row => {
            const tr = document.createElement('tr');
            this.headers.forEach(header => {
                const td = document.createElement('td');
                td.textContent = row[header];
                tr.appendChild(td);
            });
            this.dataTable.appendChild(tr);
        });

        this.initializeTableNav();
    }

    initializeTableNav() {
        if (!this.tableNav) return;
        
        const totalPages = Math.ceil(this.filteredData.length / this.itemsPerPage);
        this.tableNav.innerHTML = '';

        // Previous button
        const prevBtn = document.createElement('button');
        prevBtn.textContent = '←';
        prevBtn.className = 'button secondary';
        prevBtn.onclick = () =>this.changePage(this.currentPage - 1);
        prevBtn.disabled = this.currentPage === 1;

        // Page number
        const pageText = document.createElement('span');
        pageText.textContent = `Page ${this.currentPage} of ${totalPages}`;
        pageText.className = 'page-text';

        // Next button
        const nextBtn = document.createElement('button');
        nextBtn.textContent = '→';
        nextBtn.className = 'button secondary';
        nextBtn.onclick = () =>this.changePage(this.currentPage + 1);
        nextBtn.disabled = this.currentPage === totalPages;

        this.tableNav.appendChild(prevBtn);
        this.tableNav.appendChild(pageText);
        this.tableNav.appendChild(nextBtn);
    }

    changePage(pageNumber) {
        const totalPages = Math.ceil(this.filteredData.length / this.itemsPerPage);
        if (pageNumber< 1 || pageNumber>totalPages) return;

        this.currentPage = pageNumber;
        this.displayData(this.filteredData);
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
        if (this.downloadSelectedBtn) this.downloadSelectedBtn.disabled = false;
        if (this.downloadRejectedBtn) this.downloadRejectedBtn.disabled = false;
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
