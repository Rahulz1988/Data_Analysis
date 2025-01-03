import { processDynamicSections } from './benchmark-processor.js';

document.addEventListener('DOMContentLoaded', () => {
    const addSectionBtn = document.getElementById('addSectionBtn');
    const sectionsContainer = document.getElementById('sectionsContainer');
    const calculateBtn = document.getElementById('calculateBtnRight');
    const fileInput = document.getElementById('fileInputRight');
    const dropZone = document.getElementById('dropZoneRight');
    const fileName = document.getElementById('fileNameRight');
    const browseBtn = dropZone.querySelector('.browse-btn');
    let sectionCount = 0;

    const showMessage = (message, type) => {
        const existingMessage = document.querySelector('.alert');
        if (existingMessage) {
            existingMessage.remove();
        }
        const messageDiv = document.createElement('div');
        messageDiv.className = `alert alert-${type}`;
        messageDiv.textContent = message;

        document.body.insertBefore(messageDiv, document.body.firstChild);

        setTimeout(() =>messageDiv.remove(), 3000);
    };

    const createSection = () => {
        sectionCount++;
        const sectionDiv = document.createElement('div');
        sectionDiv.className = 'benchmark-section';
        sectionDiv.innerHTML = `
            <h3>Section ${sectionCount}</h3>
            <div class="input-group">
                <div class="input-field">
                    <label for="mean-${sectionCount}">Mean:</label>
                    <input type="number" id="mean-${sectionCount}" class="mean-input" placeholder="Enter Mean">
                </div>
                <div class="input-field">
                    <label for="sd-${sectionCount}">Standard Deviation:</label>
                    <input type="number" id="sd-${sectionCount}" class="sd-input" placeholder="Enter SD">
                </div>
                <button class="remove-section-btn" data-section="${sectionCount}">
                    <i class="fas fa-trash"></i>
                </button>
            </div>
        `;
        sectionsContainer.appendChild(sectionDiv);

        sectionDiv.querySelector('.remove-section-btn').addEventListener('click', () => {
            sectionDiv.remove();
            validateInputs();
        });

        const inputs = sectionDiv.querySelectorAll('input');
        inputs.forEach(input =>input.addEventListener('input', validateInputs));
    };

    const validateInputs = () => {
        const sections = document.querySelectorAll('.benchmark-section');
        let isValid = true;

        sections.forEach(section => {
            const meanInput = section.querySelector('.mean-input');
            const sdInput = section.querySelector('.sd-input');

            if (!meanInput.value || !sdInput.value || isNaN(meanInput.value) || Number(sdInput.value) <= 0) {
                isValid = false;
            }
        });

        calculateBtn.disabled= !isValid || sections.length === 0 || !dropZone.fileObject;
    };

    const handleFileSelection = (files) => {
        if (files &&files.length> 0) {
            const file = files[0];
            const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];

            if (validTypes.includes(file.type) || file.name.toLowerCase().endsWith('.xlsx')) {
                fileName.textContent = file.name;
                dropZone.fileObject = file;
                validateInputs();
                showMessage('File selected successfully!', 'success');
            } else {
                showMessage('Invalid file type. Please select an Excel file.', 'error');
                fileName.textContent = '';
                dropZone.fileObject = null;
                validateInputs();
            }
        }
    };

    addSectionBtn.addEventListener('click', createSection);

    browseBtn.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) =>handleFileSelection(e.target.files));
    dropZone.addEventListener('dragover', (e) =>e.preventDefault());
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        handleFileSelection(e.dataTransfer.files);
    });

    calculateBtn.addEventListener('click', async () => {
        const file = dropZone.fileObject;
        if (!file) {
            showMessage('No file selected.', 'error');
            return;
        }

        const sections = Array.from(document.querySelectorAll('.benchmark-section'));
        const benchmarks = sections.map(section => ({
            mean: parseFloat(section.querySelector('.mean-input').value),
            sd: parseFloat(section.querySelector('.sd-input').value)
        }));

        try {
            await processDynamicSections(file, benchmarks);
            showMessage('File processed and results generated successfully!', 'success');
        } catch (error) {
            showMessage(`Error: ${error.message}`, 'error');
        }
    });
});

