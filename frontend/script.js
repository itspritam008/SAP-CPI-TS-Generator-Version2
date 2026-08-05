document.addEventListener('DOMContentLoaded', () => {
    const API_URL = '/generate';
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileChip = document.getElementById('fileChip');
    const fileNameEl = document.getElementById('fileName');
    const submitBtn = document.getElementById('submitBtn');
    const form = document.getElementById('uploadForm');
    const statusMsg = document.getElementById('statusMessage');
    const progressContainer = document.getElementById('progressContainer');
    const progressFill = document.getElementById('progressFill');

    // Drag and Drop Logic
    dropZone.onclick = () => fileInput.click();

    ['dragover', 'dragenter'].forEach(evt => {
        dropZone.addEventListener(evt, (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });
    });

    ['dragleave', 'drop'].forEach(evt => {
        dropZone.addEventListener(evt, (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
        });
    });

    dropZone.addEventListener('drop', (e) => {
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            handleFileSelection();
        }
    });

    fileInput.addEventListener('change', handleFileSelection);

    function handleFileSelection() {
        if (!fileInput.files.length) return;
        
        const file = fileInput.files[0];
        if (!file.name.toLowerCase().endsWith('.zip')) {
            statusMsg.className = 'status-message error';
            statusMsg.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> Invalid file type. Please upload a .zip file.';
            submitBtn.disabled = true;
            fileChip.style.display = 'none';
            return;
        }

        // Show animated file chip and enable button
        fileChip.style.display = 'inline-flex';
        fileNameEl.innerHTML = `<i class="fa-solid fa-file-zipper"></i> ${file.name}`;
        submitBtn.disabled = false;
        
        // Reset status
        statusMsg.className = 'status-message';
        statusMsg.textContent = '';
        progressContainer.style.display = 'none';
    }

    // Form Submission Logic
    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        if (!fileInput.files.length) return;

        const formData = new FormData(form);
        
        // UI State: Loading
        submitBtn.disabled = true;
        progressContainer.style.display = 'block';
        progressFill.style.width = '15%'; // Initial loading state
        statusMsg.className = 'status-message loading';
        statusMsg.innerHTML = 'Analyzing iFlow and generating specification...';

        try {
            // Fake progression to make UI feel responsive during LLM wait time
            const progressInterval = setInterval(() => {
                let currentWidth = parseInt(progressFill.style.width);
                if (currentWidth < 85) {
                    progressFill.style.width = (currentWidth + 5) + '%';
                }
            }, 800);

            const response = await fetch(API_URL, {
                method: 'POST',
                body: formData
            });

            clearInterval(progressInterval);
            progressFill.style.width = '100%';

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.error || 'Server connection failed.');
            }

            // Handle File Download
            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            
            // Extract filename from headers if available
            let filename = "Technical_Specification.docx";
            const disposition = response.headers.get('content-disposition');
            if (disposition && disposition.includes('filename=')) {
                filename = disposition.split('filename=')[1].replace(/"/g, '');
            }

            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            // UI State: Success
            statusMsg.className = 'status-message success';
            statusMsg.innerHTML = '<i class="fa-solid fa-circle-check"></i> Document generated successfully!';
            
        } catch (err) {
            // UI State: Error
            progressFill.style.width = '100%';
            progressFill.style.background = 'var(--error)';
            statusMsg.className = 'status-message error';
            statusMsg.innerHTML = `<i class="fa-solid fa-circle-xmark"></i> Error: ${err.message}`;
        } finally {
            submitBtn.disabled = false;
            // Reset progress bar color in case it turned red from an error
            setTimeout(() => {
                progressFill.style.background = ''; 
            }, 3000);
        }
    });
});