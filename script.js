/**
 * Excel Data Matcher & Merger
 * Pure client-side application using SheetJS
 */

// ===== State Management =====
const state = {
    file1: null,
    file2: null,
    file1Data: null,
    file2Data: null,
    processedData: null,
    matchedCount: 0,
    unmatchedCount: 0,
    isProcessing: false
};

// ===== DOM Elements =====
const elements = {
    file1Input: document.getElementById('file1'),
    file2Input: document.getElementById('file2'),
    dropZone1: document.getElementById('dropZone1'),
    dropZone2: document.getElementById('dropZone2'),
    fileInfo1: document.getElementById('fileInfo1'),
    fileInfo2: document.getElementById('fileInfo2'),
    status1: document.getElementById('status1'),
    status2: document.getElementById('status2'),
    processBtn: document.getElementById('processBtn'),
    processSection: document.getElementById('processSection'),
    progressSection: document.getElementById('progressSection'),
    progressBar: document.getElementById('progressBar'),
    progressPercent: document.getElementById('progressPercent'),
    progressStatus: document.getElementById('progressStatus'),
    resultsSection: document.getElementById('resultsSection'),
    matchedCount: document.getElementById('matchedCount'),
    unmatchedCount: document.getElementById('unmatchedCount'),
    totalCount: document.getElementById('totalCount'),
    previewTableBody: document.getElementById('previewTableBody'),
    downloadBtn: document.getElementById('downloadBtn'),
    toastContainer: document.getElementById('toastContainer')
};

// ===== Utility Functions =====

/**
 * Format file size to human readable
 */
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Show toast notification
 */
function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    const icons = {
        success: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#10b981" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',
        error: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#ef4444" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>',
        info: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#3b82f6" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>'
    };
    
    toast.innerHTML = `
        <div class="toast-icon">${icons[type]}</div>
        <div class="toast-message">${message}</div>
    `;
    
    elements.toastContainer.appendChild(toast);
    
    setTimeout(() => {
        toast.remove();
    }, 5000);
}

/**
 * Update progress bar
 */
function updateProgress(percent, status) {
    elements.progressBar.style.width = percent + '%';
    elements.progressPercent.textContent = percent + '%';
    if (status) elements.progressStatus.textContent = status;
}

/**
 * Enable/disable process button
 */
function updateProcessButton() {
    const enabled = state.file1 && state.file2 && !state.isProcessing;
    elements.processBtn.disabled = !enabled;
}

/**
 * Get column letter from index (0 = A, 1 = B, etc.)
 */
function getColumnLetter(index) {
    return String.fromCharCode(65 + index);
}

/**
 * Read Excel file and convert to array of arrays
 */
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
                resolve(jsonData);
            } catch (error) {
                reject(new Error('Invalid Excel file: ' + error.message));
            }
        };
        
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsArrayBuffer(file);
    });
}

// ===== File Upload Handlers =====

/**
 * Handle file selection
 */
function handleFileSelect(file, fileNum) {
    if (!file) return;
    
    // Validate file type
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'application/octet-stream'
    ];
    
    const validExtensions = ['.xlsx', '.xls'];
    const hasValidExtension = validExtensions.some(ext => 
        file.name.toLowerCase().endsWith(ext)
    );
    
    if (!validTypes.includes(file.type) && !hasValidExtension) {
        showToast('Please upload a valid Excel file (.xlsx or .xls)', 'error');
        return;
    }
    
    // Update state
    if (fileNum === 1) {
        state.file1 = file;
        state.file1Data = null;
    } else {
        state.file2 = file;
        state.file2Data = null;
    }
    
    // Update UI
    const fileInfo = fileNum === 1 ? elements.fileInfo1 : elements.fileInfo2;
    const status = fileNum === 1 ? elements.status1 : elements.status2;
    const dropZone = fileNum === 1 ? elements.dropZone1 : elements.dropZone2;
    
    fileInfo.querySelector('.file-name').textContent = file.name;
    fileInfo.querySelector('.file-size').textContent = formatFileSize(file.size);
    fileInfo.classList.add('active');
    dropZone.classList.add('has-file');
    status.textContent = 'Ready';
    status.classList.add('ready');
    
    updateProcessButton();
    showToast(`File ${fileNum} uploaded successfully`, 'success');
}

/**
 * Remove uploaded file
 */
function removeFile(fileNum) {
    if (fileNum === 1) {
        state.file1 = null;
        state.file1Data = null;
        elements.file1Input.value = '';
        elements.fileInfo1.classList.remove('active');
        elements.dropZone1.classList.remove('has-file');
        elements.status1.textContent = 'Waiting...';
        elements.status1.classList.remove('ready');
    } else {
        state.file2 = null;
        state.file2Data = null;
        elements.file2Input.value = '';
        elements.fileInfo2.classList.remove('active');
        elements.dropZone2.classList.remove('has-file');
        elements.status2.textContent = 'Waiting...';
        elements.status2.classList.remove('ready');
    }
    
    // Hide results if visible
    elements.resultsSection.style.display = 'none';
    elements.progressSection.style.display = 'none';
    updateProcessButton();
}

// ===== Drag & Drop =====

function setupDragAndDrop(dropZone, fileNum) {
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
        });
    });
    
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.add('drag-over');
        });
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.remove('drag-over');
        });
    });
    
    dropZone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFileSelect(files[0], fileNum);
        }
    });
}

// ===== Data Processing =====

/**
 * Main processing function
 */
async function processData() {
    if (!state.file1 || !state.file2) {
        showToast('Please upload both files first', 'error');
        return;
    }
    
    state.isProcessing = true;
    updateProcessButton();
    
    // Show progress section
    elements.progressSection.style.display = 'block';
    elements.resultsSection.style.display = 'none';
    elements.processBtn.classList.add('processing');
    
    try {
        updateProgress(10, 'Reading File 1 (Master Data)...');
        
        // Read File 1
        const file1Raw = await readExcelFile(state.file1);
        state.file1Data = file1Raw;
        
        updateProgress(30, 'Reading File 2 (Reference Data)...');
        
        // Read File 2
        const file2Raw = await readExcelFile(state.file2);
        state.file2Data = file2Raw;
        
        updateProgress(50, 'Building match index from File 1...');
        
        // Small delay for UI update
        await new Promise(resolve => setTimeout(resolve, 100));
        
        // Process data
        const result = await performMatching(file1Raw, file2Raw);
        
        updateProgress(100, 'Complete!');
        
        // Store result
        state.processedData = result.data;
        state.matchedCount = result.matched;
        state.unmatchedCount = result.unmatched;
        
        // Show results
        setTimeout(() => {
            displayResults();
            elements.progressSection.style.display = 'none';
            elements.resultsSection.style.display = 'block';
            elements.processBtn.classList.remove('processing');
            state.isProcessing = false;
            updateProcessButton();
            showToast(`Processing complete! ${result.matched} rows matched, ${result.unmatched} rows unmatched.`, 'success');
        }, 500);
        
    } catch (error) {
        console.error(error);
        showToast(error.message, 'error');
        elements.progressSection.style.display = 'none';
        elements.processBtn.classList.remove('processing');
        state.isProcessing = false;
        updateProcessButton();
    }
}

/**
 * Perform the matching logic
 * File 1: Master Data
 * File 2: Reference Data
 * Match on E+F+G (indices 4,5,6)
 * Copy A,B,C,D (indices 0,1,2,3) from File 1 to File 2
 */
async function performMatching(file1Data, file2Data) {
    return new Promise((resolve) => {
        // Skip header rows if present (assume first row might be headers)
        // We'll treat all rows as data since user said columns A-I
        
        const file1Rows = file1Data;
        const file2Rows = file2Data;
        
        // Build lookup map from File 1
        // Key: E+F+G (columns 4,5,6), Value: row with A,B,C,D (columns 0,1,2,3)
        const lookupMap = new Map();
        
        for (let i = 0; i < file1Rows.length; i++) {
            const row = file1Rows[i];
            if (row.length < 7) continue; // Need at least up to G
            
            const key = `${String(row[4]).trim()}|${String(row[5]).trim()}|${String(row[6]).trim()}`;
            
            // Store the A-D values (columns 0-3)
            const values = {
                A: row[0] !== undefined ? row[0] : '',
                B: row[1] !== undefined ? row[1] : '',
                C: row[2] !== undefined ? row[2] : '',
                D: row[3] !== undefined ? row[3] : ''
            };
            
            // If duplicate keys exist, first one wins (or we could store array)
            if (!lookupMap.has(key)) {
                lookupMap.set(key, values);
            }
        }
        
        updateProgress(70, 'Matching rows and merging data...');
        
        // Process File 2
        const processedRows = [];
        let matched = 0;
        let unmatched = 0;
        
        for (let i = 0; i < file2Rows.length; i++) {
            const row = [...file2Rows[i]]; // Clone row
            const originalLength = row.length;
            
            // Ensure row has at least 9 columns (A-I)
            while (row.length < 9) {
                row.push('');
            }
            
            if (row.length >= 7) {
                const key = `${String(row[4]).trim()}|${String(row[5]).trim()}|${String(row[6]).trim()}`;
                
                if (lookupMap.has(key)) {
                    const matchData = lookupMap.get(key);
                    // Update A-D with matched data from File 1
                    row[0] = matchData.A;
                    row[1] = matchData.B;
                    row[2] = matchData.C;
                    row[3] = matchData.D;
                    row._matched = true;
                    matched++;
                } else {
                    row._matched = false;
                    unmatched++;
                }
            } else {
                row._matched = false;
                unmatched++;
            }
            
            processedRows.push(row);
            
            // Update progress periodically
            if (i % 100 === 0) {
                const progress = 70 + Math.floor((i / file2Rows.length) * 25);
                updateProgress(progress, `Processing row ${i + 1} of ${file2Rows.length}...`);
            }
        }
        
        resolve({
            data: processedRows,
            matched: matched,
            unmatched: unmatched
        });
    });
}

// ===== Results Display =====

/**
 * Display processing results
 */
function displayResults() {
    // Update stats
    elements.matchedCount.textContent = state.matchedCount.toLocaleString();
    elements.unmatchedCount.textContent = state.unmatchedCount.toLocaleString();
    elements.totalCount.textContent = (state.matchedCount + state.unmatchedCount).toLocaleString();
    
    // Build preview table (first 50 rows)
    const tbody = elements.previewTableBody;
    tbody.innerHTML = '';
    
    const rowsToShow = state.processedData.slice(0, 50);
    
    rowsToShow.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.className = row._matched ? 'matched-row' : 'unmatched-row';
        
        // Columns A-I (0-8)
        for (let i = 0; i < 9; i++) {
            const td = document.createElement('td');
            td.textContent = row[i] !== undefined && row[i] !== '' ? row[i] : '-';
            td.title = String(row[i] || ''); // Tooltip for truncated text
            tr.appendChild(td);
        }
        
        // Status column
        const statusTd = document.createElement('td');
        statusTd.innerHTML = row._matched 
            ? '<span class="status-badge matched">Matched</span>'
            : '<span class="status-badge unmatched">Unmatched</span>';
        tr.appendChild(statusTd);
        
        tbody.appendChild(tr);
    });
    
    // Scroll to results
    elements.resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ===== Download =====

/**
 * Download processed file as Excel
 */
function downloadFile() {
    if (!state.processedData || state.processedData.length === 0) {
        showToast('No data to download', 'error');
        return;
    }
    
    try {
        // Prepare data for export (remove _matched flag)
        const exportData = state.processedData.map(row => {
            const cleanRow = [...row];
            delete cleanRow._matched;
            return cleanRow.slice(0, 9); // Ensure only A-I
        });
        
        // Create worksheet
        const ws = XLSX.utils.aoa_to_sheet(exportData);
        
        // Set column widths for better formatting
        const colWidths = [
            { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 },
            { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }
        ];
        ws['!cols'] = colWidths;
        
        // Create workbook
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Updated Data');
        
        // Generate filename with timestamp
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
        const filename = `Updated_Data_${timestamp}.xlsx`;
        
        // Download
        XLSX.writeFile(wb, filename);
        
        showToast('File downloaded successfully!', 'success');
    } catch (error) {
        console.error(error);
        showToast('Failed to download file: ' + error.message, 'error');
    }
}

// ===== Event Listeners =====

function initEventListeners() {
    // File 1
    elements.dropZone1.addEventListener('click', () => elements.file1Input.click());
    elements.file1Input.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0], 1);
        }
    });
    setupDragAndDrop(elements.dropZone1, 1);
    
    // File 2
    elements.dropZone2.addEventListener('click', () => elements.file2Input.click());
    elements.file2Input.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0], 2);
        }
    });
    setupDragAndDrop(elements.dropZone2, 2);
}

// ===== Initialize =====
document.addEventListener('DOMContentLoaded', () => {
    initEventListeners();
    updateProcessButton();
    console.log('Excel Data Matcher initialized');
});