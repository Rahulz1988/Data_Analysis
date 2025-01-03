import { processExcelFile } from './excel-processor.js';


document.getElementById('downloadTemplate').addEventListener('click', () => {
    // Path to your Excel template file
    const fileUrl = 'js/templates/test_template.xlsx'; // Ensure this path matches your file structure
    const link = document.createElement('a'); // Create a temporary link element
    link.href = fileUrl; // Set the URL of the file
    link.download = 'test_template.xlsx'; // Set the desired file name for download
    document.body.appendChild(link); // Append the link to the DOM
    link.click(); // Trigger the click event to start the download
    document.body.removeChild(link); // Remove the link after download starts
});


document.addEventListener('DOMContentLoaded', () => {
    // Get DOM elements
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileName = document.getElementById('fileName');
    const calculateBtn = document.getElementById('calculateBtn');
    const browseBtn = document.querySelector('.browse-btn');
    const progressBar = document.getElementById('progressBar');

    // Initialize SheetJS
    if (typeof XLSX === 'undefined') {
        console.error('SheetJS library not loaded');
        showMessage('Required libraries not loaded. Please refresh the page.', 'error');
        return;
    }

    // File reading function
    const readExcelFile = (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = e.target.result;
                    const workbook = XLSX.read(data, { type: 'array' });
                    console.log("Workbook read successfully:", workbook); // Debug log
                    resolve(workbook);
                } catch (error) {
                    console.error("Error reading workbook:", error); // Debug log
                    reject(error);
                }
            };

            reader.onerror = (error) => {
                console.error("File reading error:", error); // Debug log
                reject(error);
            };
            reader.readAsArrayBuffer(file);
        });
    };

    // File selection handler
    const handleFileSelection = (files) => {
        if (files &&files.length> 0) {
            const file = files[0];

            // Check if file is an Excel file
            const validTypes = [
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'application/vnd.ms-excel',
                '.xlsx',
                '.xls'
            ];

            const isValidType = validTypes.some(type =>
                file.type === type || file.name.toLowerCase().endsWith(type)
            );

            if (isValidType) {
                fileName.textContent = file.name;
                calculateBtn.disabled = false;
                showMessage('File selected successfully!', 'success');

                // Store the actual File object for later processing
                dropZone.fileObject = file;
            } else {
                showMessage('Please select a valid Excel file (.xlsx or .xls)', 'error');
                fileName.textContent = '';
                calculateBtn.disabled = true;
                delete dropZone.fileObject; // Clear invalid entry
            }
        }
    };

    // Message display function
    const showMessage = (message, type) => {
        const existingMessage = document.querySelector('.alert');
        if (existingMessage) {
            existingMessage.remove();
        }

        const messageDiv = document.createElement('div');
        messageDiv.className = `alert alert-${type}`;
        messageDiv.textContent = message;

        const uploadSection = document.querySelector('.section');
        uploadSection.insertBefore(messageDiv, uploadSection.firstChild);

        setTimeout(() =>messageDiv.remove(), 3000);
    };

    // Event Listeners
    browseBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        handleFileSelection(e.target.files);
    });

    // Drag and drop handlers
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        handleFileSelection(e.dataTransfer.files);
    });

    // Calculate button handler
    calculateBtn.addEventListener('click', async () => {
        console.log("Calculate button clicked");

        // Retrieve the actual File object from dropZone
        const file = dropZone.fileObject;

        // Check if the file exists and is an instance of File
        if (!file || !(file instanceof File)) {
            console.error('No valid file selected for processing.');
            showMessage('No valid file selected for processing.', 'error');
            return;
        }

        try {
            console.log("Reading Excel file...");
            
            calculateBtn.disabled = true;
            calculateBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';
            progressBar.style.display = 'block';

            // Use processExcelFile function from excel-processor.js to process the uploaded Excel file
            await processExcelFile(file);

            showMessage('File processed successfully!', 'success');
        } catch (error) {
            console.error('Processing error:', error);
            showMessage('Error processing file: ' + error.message, 'error');
        } finally {
            calculateBtn.disabled = false;
            calculateBtn.innerHTML = '<i class="fas fa-calculator"></i> Calculate';
            progressBar.style.display = 'none';
        }
    });
});

