// document.addEventListener('DOMContentLoaded', () => {
//     const fileInput = document.getElementById('file-input');
//     const uploadBox = document.getElementById('upload-box');
//     const browseBtn = document.getElementById('browse-btn');
//     const processBtn = document.getElementById('process-btn');
//     const processOptions = document.getElementById('process-options');
//     const fileNameDisplay = document.getElementById('file-name');
//     const fileSearchInput = document.getElementById('file-search');
//     let uploadedFile = null;

//     // Handle file selection or drag-and-drop
//     uploadBox.addEventListener('click', () => fileInput.click());
//     browseBtn.addEventListener('click', () => fileInput.click());
    
//     // Add drag and drop functionality
//     uploadBox.addEventListener('dragover', (e) => {
//         e.preventDefault();
//         uploadBox.classList.add('dragover');
//     });

//     uploadBox.addEventListener('dragleave', () => {
//         uploadBox.classList.remove('dragover');
//     });

//     uploadBox.addEventListener('drop', (e) => {
//         e.preventDefault();
//         uploadBox.classList.remove('dragover');
//         const file = e.dataTransfer.files[0];
//         handleFileSelect({ target: { files: [file] } });
//     });

//     fileInput.addEventListener('change', handleFileSelect);

//     // Display file name when selected
//     function handleFileSelect(e) {
//         const file = e.target.files[0];
//         if (file && file.name.endsWith('.zip')) {
//             uploadedFile = file;
//             fileNameDisplay.textContent = `Selected File: ${file.name}`;
//             enableProcessButton();
//         } else {
//             alert('Please upload a valid ZIP file.');
//         }
//     }

//     // Enable Process Button
//     function enableProcessButton() {
//         processBtn.disabled = !processOptions.checked || !uploadedFile;
//     }

//     processOptions.addEventListener('change', enableProcessButton);

//     // Handle file processing
//     processBtn.addEventListener('click', async () => {
//         if (uploadedFile) {
//             const searchQuery = fileSearchInput.value.trim();
//             if (searchQuery) {
//                 try {
//                     await processZipFile(uploadedFile, searchQuery);
//                 } catch (error) {
//                     alert(`Error processing file: ${error.message}`);
//                 }
//             } else {
//                 alert('Please enter a filename or substring.');
//             }
//         } else {
//             alert('No file selected');
//         }
//     });

//     // Process ZIP file and search for matching files
//     async function processZipFile(zipFile, searchQuery) {
//         const zip = new JSZip();
        
//         try {
//             // Load the ZIP file
//             const loadedZip = await zip.loadAsync(zipFile);
            
//             // Create a new ZIP for matching files
//             const matchingZip = new JSZip();
//             let matchCount = 0;

//             // Show loading state
//             processBtn.textContent = 'Processing...';
//             processBtn.disabled = true;

//             // Search through all files in the ZIP
//             const promises = [];
//             loadedZip.forEach((relativePath, zipEntry) => {
//                 // Skip if it's a directory
//                 if (!zipEntry.dir) {
//                     const fileName = zipEntry.name.split('/').pop();
//                     // Check if filename matches the search query
//                     if (fileName.toLowerCase().includes(searchQuery.toLowerCase())) {
//                         // Add file to the new ZIP
//                         promises.push(
//                             zipEntry.async('uint8array').then(content => {
//                                 matchingZip.file(zipEntry.name, content);
//                                 matchCount++;
//                             })
//                         );
//                     }
//                 }
//             });

//             // Wait for all matching files to be processed
//             await Promise.all(promises);

//             if (matchCount > 0) {
//                 // Generate the new ZIP file
//                 const content = await matchingZip.generateAsync({
//                     type: 'blob',
//                     compression: 'DEFLATE'
//                 });

//                 // Create download link
//                 const url = URL.createObjectURL(content);
//                 const a = document.createElement('a');
//                 a.href = url;
//                 a.download = `matched_files_${searchQuery}.zip`;
//                 document.body.appendChild(a);
//                 a.click();
//                 document.body.removeChild(a);
//                 URL.revokeObjectURL(url);

//                 alert(`Found and zipped ${matchCount} matching files!`);
//             } else {
//                 alert('No matching files found.');
//             }
//         } catch (error) {
//             throw new Error('Error processing ZIP file: ' + error.message);
//         } finally {
//             // Reset button state
//             processBtn.textContent = 'Process Files';
//             processBtn.disabled = false;
//         }
//     }
// });

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('file-input');
    const uploadBox = document.getElementById('upload-box');
    const browseBtn = document.getElementById('browse-btn');
    const processBtn = document.getElementById('process-btn');
    const processOptions = document.getElementById('process-options');
    const fileNameDisplay = document.getElementById('file-name');
    const fileSearchInput = document.getElementById('file-search');
    let uploadedFile = null;

    // Handle file selection or drag-and-drop
    uploadBox.addEventListener('click', () => fileInput.click());
    browseBtn.addEventListener('click', () => fileInput.click());
    
    // Add drag and drop functionality
    uploadBox.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadBox.classList.add('dragover');
    });

    uploadBox.addEventListener('dragleave', () => {
        uploadBox.classList.remove('dragover');
    });

    uploadBox.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadBox.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        handleFileSelect({ target: { files: [file] } });
    });

    fileInput.addEventListener('change', handleFileSelect);

    // Display file name when selected
    function handleFileSelect(e) {
        const file = e.target.files[0];
        if (file && file.name.endsWith('.zip')) {
            uploadedFile = file;
            fileNameDisplay.textContent = `Selected File: ${file.name}`;
            enableProcessButton();
        } else {
            alert('Please upload a valid ZIP file.');
        }
    }

    // Enable Process Button
    function enableProcessButton() {
        processBtn.disabled = !processOptions.checked || !uploadedFile;
    }

    processOptions.addEventListener('change', enableProcessButton);

    // Handle file processing
    processBtn.addEventListener('click', async () => {
        if (uploadedFile) {
            const searchQuery = fileSearchInput.value.trim();
            if (searchQuery) {
                try {
                    await processZipFile(uploadedFile, searchQuery);
                } catch (error) {
                    alert(`Error processing file: ${error.message}`);
                }
            } else {
                alert('Please enter a filename or substring.');
            }
        } else {
            alert('No file selected');
        }
    });

    // Process ZIP file and search for matching files
    async function processZipFile(zipFile, searchQuery) {
        const zip = new JSZip();
        
        try {
            // Load the ZIP file
            const loadedZip = await zip.loadAsync(zipFile);
            
            // Create a new ZIP for matching files
            const matchingZip = new JSZip();
            let matchCount = 0;

            // Show loading state
            processBtn.textContent = 'Processing...';
            processBtn.disabled = true;

            // Search through all files in the ZIP
            const promises = [];
            loadedZip.forEach((relativePath, zipEntry) => {
                // Skip if it's a directory
                if (!zipEntry.dir) {
                    const fileName = zipEntry.name.split('/').pop();
                    // Check if filename matches the search query
                    if (fileName.toLowerCase().includes(searchQuery.toLowerCase())) {
                        // Add file to the new ZIP under 'processed_data' folder
                        promises.push(
                            zipEntry.async('uint8array').then(content => {
                                // Place the file inside 'processed_data' folder in the new ZIP
                                matchingZip.file(`processed_data/${fileName}`, content);
                                matchCount++;
                            })
                        );
                    }
                }
            });

            // Wait for all matching files to be processed
            await Promise.all(promises);

            if (matchCount > 0) {
                // Generate the new ZIP file
                const content = await matchingZip.generateAsync({
                    type: 'blob',
                    compression: 'DEFLATE'
                });

                // Create download link
                const url = URL.createObjectURL(content);
                const a = document.createElement('a');
                a.href = url;
                a.download = `matched_files_${searchQuery}.zip`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);

                alert(`Found and zipped ${matchCount} matching files!`);
            } else {
                alert('No matching files found.');
            }
        } catch (error) {
            throw new Error('Error processing ZIP file: ' + error.message);
        } finally {
            // Reset button state
            processBtn.textContent = 'Process Files';
            processBtn.disabled = false;
        }
    }
});
