// DOM Elements
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const processBtn = document.getElementById('processBtn');
const clearAllBtn = document.getElementById('clearAllBtn');
const filesSection = document.getElementById('filesSection');
const filesList = document.getElementById('filesList');
const resultsSection = document.getElementById('resultsSection');
const resultsTableBody = document.getElementById('resultsTableBody');
const totalSummary = document.getElementById('totalSummary');
const emptyState = document.getElementById('emptyState');
const loadingOverlay = document.getElementById('loadingOverlay');
const viewGraphBtn = document.getElementById('viewGraphBtn');
const graphModal = document.getElementById('graphModal');
const closeModal = document.getElementById('closeModal');
const viewAgentCollectionBtn = document.getElementById('viewAgentCollectionBtn');
const agentModal = document.getElementById('agentModal');
const closeAgentModal = document.getElementById('closeAgentModal');

// State
let selectedFiles = [];
let processedResults = [];

// Event Listeners
uploadArea.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', handleFileSelect);
processBtn.addEventListener('click', processFiles);
clearAllBtn.addEventListener('click', clearAll);
viewGraphBtn.addEventListener('click', showGraph);
closeModal.addEventListener('click', hideGraph);
viewAgentCollectionBtn.addEventListener('click', showAgentCollection);
closeAgentModal.addEventListener('click', hideAgentCollection);

// Close modal when clicking outside
graphModal.addEventListener('click', (e) => {
    if (e.target === graphModal) {
        hideGraph();
    }
});

agentModal.addEventListener('click', (e) => {
    if (e.target === agentModal) {
        hideAgentCollection();
    }
});

// Drag and drop
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = Array.from(e.dataTransfer.files).filter(file => 
        file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    );
    addFiles(files);
});

// File handling
function handleFileSelect(e) {
    const files = Array.from(e.target.files);
    addFiles(files);
}

function addFiles(files) {
    files.forEach(file => {
        if (!selectedFiles.find(f => f.name === file.name)) {
            selectedFiles.push(file);
        }
    });
    updateFilesList();
    updateProcessButton();
}

function removeFile(fileName) {
    selectedFiles = selectedFiles.filter(f => f.name !== fileName);
    updateFilesList();
    updateProcessButton();
}

function updateFilesList() {
    if (selectedFiles.length === 0) {
        filesSection.style.display = 'none';
        return;
    }

    filesSection.style.display = 'block';
    filesList.innerHTML = selectedFiles.map(file => `
        <div class="file-item">
            <div class="file-info">
                <svg class="file-icon" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                    <polyline points="14 2 14 8 20 8"></polyline>
                    <line x1="12" y1="18" x2="12" y2="12"></line>
                    <line x1="9" y1="15" x2="15" y2="15"></line>
                </svg>
                <span class="file-name">${file.name}</span>
            </div>
            <button class="file-remove" onclick="removeFile('${file.name}')">Remove</button>
        </div>
    `).join('');
}

function updateProcessButton() {
    processBtn.disabled = selectedFiles.length === 0;
    clearAllBtn.style.display = selectedFiles.length > 0 ? 'block' : 'none';
}

function clearAll() {
    // Clear selected files
    selectedFiles = [];
    processedResults = [];
    
    // Reset file input
    fileInput.value = '';
    
    // Hide sections
    filesSection.style.display = 'none';
    resultsSection.style.display = 'none';
    totalSummary.style.display = 'none';
    clearAllBtn.style.display = 'none';
    
    // Show empty state
    emptyState.style.display = 'flex';
    
    // Disable process button
    processBtn.disabled = true;
    
    // Clear file list
    filesList.innerHTML = '';
    
    // Clear results table
    resultsTableBody.innerHTML = '';
}

// File processing
async function processFiles() {
    loadingOverlay.style.display = 'flex';
    resultsSection.style.display = 'none';
    emptyState.style.display = 'none';
    resultsTableBody.innerHTML = '';

    const results = [];
    const totalFiles = selectedFiles.length;
    const startTime = Date.now();
    
    // Progress elements
    const progressFill = document.getElementById('progressFill');
    const progressPercent = document.getElementById('progressPercent');
    const progressStatus = document.getElementById('progressStatus');
    const progressTime = document.getElementById('progressTime');
    
    // Initialize
    progressFill.style.width = '0%';
    progressPercent.textContent = '0%';
    progressStatus.textContent = `Processing 0 of ${totalFiles} files...`;
    progressTime.textContent = 'Calculating...';

    for (let i = 0; i < selectedFiles.length; i++) {
        const file = selectedFiles[i];
        
        // Update progress BEFORE processing (starting file)
        const startProgress = Math.round((i / totalFiles) * 100);
        progressFill.style.width = startProgress + '%';
        progressPercent.textContent = startProgress + '%';
        progressStatus.textContent = `Processing file ${i + 1} of ${totalFiles}: ${file.name}`;
        
        // Calculate estimated time remaining BEFORE processing
        if (i > 0) {
            const elapsed = Date.now() - startTime;
            const avgTimePerFile = elapsed / i;
            const remainingFiles = totalFiles - i;
            const estimatedRemainingMs = avgTimePerFile * remainingFiles;
            const estimatedRemainingSec = Math.ceil(estimatedRemainingMs / 1000);
            
            if (estimatedRemainingSec > 60) {
                const minutes = Math.floor(estimatedRemainingSec / 60);
                const seconds = estimatedRemainingSec % 60;
                progressTime.textContent = `Estimated time: ~${minutes}m ${seconds}s`;
            } else {
                progressTime.textContent = `Estimated time: ~${estimatedRemainingSec}s`;
            }
        } else {
            progressTime.textContent = 'Calculating time...';
        }
        
        try {
            console.log('Processing file:', file.name);
            const result = await processExcelFile(file);
            console.log('Result:', result);
            results.push(result);
            
            // Update progress AFTER processing (completed file)
            const completedProgress = Math.round(((i + 1) / totalFiles) * 100);
            progressFill.style.width = completedProgress + '%';
            progressPercent.textContent = completedProgress + '%';
            
            // Update time estimate after completion
            if (i < totalFiles - 1) {
                const elapsed = Date.now() - startTime;
                const avgTimePerFile = elapsed / (i + 1);
                const remainingFiles = totalFiles - (i + 1);
                const estimatedRemainingMs = avgTimePerFile * remainingFiles;
                const estimatedRemainingSec = Math.ceil(estimatedRemainingMs / 1000);
                
                if (estimatedRemainingSec > 60) {
                    const minutes = Math.floor(estimatedRemainingSec / 60);
                    const seconds = estimatedRemainingSec % 60;
                    progressTime.textContent = `Estimated time: ~${minutes}m ${seconds}s`;
                } else {
                    progressTime.textContent = `Estimated time: ~${estimatedRemainingSec}s`;
                }
            }
            
        } catch (error) {
            console.error('Error processing file:', file.name, error);
            results.push({
                fileName: file.name,
                error: error.message
            });
            
            // Still update progress even on error
            const completedProgress = Math.round(((i + 1) / totalFiles) * 100);
            progressFill.style.width = completedProgress + '%';
            progressPercent.textContent = completedProgress + '%';
        }
    }
    
    // Set to 100% when complete
    progressFill.style.width = '100%';
    progressPercent.textContent = '100%';
    progressStatus.textContent = `Completed processing ${totalFiles} file${totalFiles > 1 ? 's' : ''}`;
    
    const totalTime = Math.ceil((Date.now() - startTime) / 1000);
    if (totalTime > 60) {
        const minutes = Math.floor(totalTime / 60);
        const seconds = totalTime % 60;
        progressTime.textContent = `Total time: ${minutes}m ${seconds}s`;
    } else {
        progressTime.textContent = `Total time: ${totalTime}s`;
    }

    console.log('All results:', results);
    processedResults = results; // Store for graph
    displayResults(results);
    
    // Delay hiding overlay to show completion
    setTimeout(() => {
        loadingOverlay.style.display = 'none';
    }, 800);
}

async function processExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Process first sheet
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
                
                // Find columns
                const accountNoColumn = findColumn(jsonData, 'Account No.');
                const remarkTypeColumn = findColumn(jsonData, 'Remark Type');
                const relationColumn = findColumn(jsonData, 'Relation');
                const ptpAmountColumn = findColumn(jsonData, 'PTP Amount');
                const claimPaidAmountColumn = findColumn(jsonData, 'Claim Paid Amount');
                const remarkByColumn = findColumn(jsonData, 'Remark By');
                const balanceColumn = findColumn(jsonData, 'Balance');
                
                if (!accountNoColumn) {
                    throw new Error('Column "Account No." not found');
                }
                
                if (!remarkTypeColumn) {
                    throw new Error('Column "Remark Type" not found');
                }
                
                if (!relationColumn) {
                    throw new Error('Column "Relation" not found');
                }
                
                if (!ptpAmountColumn) {
                    throw new Error('Column "PTP Amount" not found');
                }
                
                if (!claimPaidAmountColumn) {
                    throw new Error('Column "Claim Paid Amount" not found');
                }
                
                // Remark By is optional - just log if not found
                if (!remarkByColumn) {
                    console.warn('Column "Remark By" not found - agent data will not be available');
                }
                
                // Extract date from filename BEFORE processing
                const fileDate = extractDateFromFilename(file.name);
                
                // Step 1: Filter rows and collect Account No. values
                const accountNumbers = [];
                const predictiveAccountNumbers = [];
                const debtorAccountNumbers = [];
                const seenAccountsForBalance = new Set(); // track first occurrence only
                let totalRows = jsonData.length;
                let excludedSMS = 0;
                let excludedBlanks = 0;
                let ptpCount = 0;
                let ptpTotalAmount = 0;
                let predictivePtpCount = 0;
                let predictivePtpAmount = 0;
                let claimPaidCount = 0;
                let claimPaidTotalAmount = 0;
                let totalBalance = 0;
                
                // Collect agent data
                const agentData = {};
                const agentDataByDate = {};
                
                jsonData.forEach(row => {
                    const accountNo = row[accountNoColumn];
                    const remarkType = row[remarkTypeColumn];
                    const relation = row[relationColumn];
                    const ptpAmount = row[ptpAmountColumn];
                    const claimPaidAmount = row[claimPaidAmountColumn];
                    const remarkBy = remarkByColumn ? row[remarkByColumn] : null;
                    
                    // Get agent name
                    const agentName = remarkBy ? remarkBy.toString().trim() : '';
                    
                    // Initialize agent data if not exists
                    if (agentName) {
                        if (!agentData[agentName]) {
                            agentData[agentName] = {
                                ptpCount: 0,
                                ptpAmount: 0,
                                claimPaidCount: 0,
                                claimPaidAmount: 0
                            };
                        }
                        
                        // Initialize date-based tracking
                        if (fileDate) {
                            if (!agentDataByDate[agentName]) {
                                agentDataByDate[agentName] = {};
                            }
                            if (!agentDataByDate[agentName][fileDate]) {
                                agentDataByDate[agentName][fileDate] = {
                                    ptpCount: 0,
                                    ptpAmount: 0,
                                    claimPaidCount: 0,
                                    claimPaidAmount: 0
                                };
                            }
                        }
                    }
                    
                    // Process PTP Amount
                    if (ptpAmount !== null && ptpAmount !== undefined && ptpAmount !== '') {
                        const ptpStr = ptpAmount.toString().trim();
                        if (ptpStr !== '' && ptpStr !== '0' && ptpStr !== '0.00') {
                            const cleanStr = ptpStr.replace(/,/g, '').replace(/[^\d.-]/g, '');
                            const ptpValue = parseFloat(cleanStr);
                            if (!isNaN(ptpValue) && ptpValue > 0) {
                                ptpCount++;
                                ptpTotalAmount += ptpValue;
                                
                                // Track Predictive PTP separately
                                const remarkTypeStr = row[remarkTypeColumn] ? row[remarkTypeColumn].toString().trim().toUpperCase() : '';
                                if (remarkTypeStr === 'PREDICTIVE') {
                                    predictivePtpCount++;
                                    predictivePtpAmount += ptpValue;
                                }
                                
                                // Add to agent data
                                if (agentName && agentData[agentName]) {
                                    agentData[agentName].ptpCount++;
                                    agentData[agentName].ptpAmount += ptpValue;
                                    
                                    // Add to date-based tracking
                                    if (fileDate && agentDataByDate[agentName] && agentDataByDate[agentName][fileDate]) {
                                        agentDataByDate[agentName][fileDate].ptpCount++;
                                        agentDataByDate[agentName][fileDate].ptpAmount += ptpValue;
                                    }
                                }
                            }
                        }
                    }
                    
                    // Process Claim Paid Amount
                    if (claimPaidAmount !== null && claimPaidAmount !== undefined && claimPaidAmount !== '') {
                        const claimPaidStr = claimPaidAmount.toString().trim();
                        if (claimPaidStr !== '' && claimPaidStr !== '0' && claimPaidStr !== '0.00') {
                            const cleanStr = claimPaidStr.replace(/,/g, '').replace(/[^\d.-]/g, '');
                            const claimPaidValue = parseFloat(cleanStr);
                            if (!isNaN(claimPaidValue) && claimPaidValue > 0) {
                                claimPaidCount++;
                                claimPaidTotalAmount += claimPaidValue;
                                
                                // Add to agent data
                                if (agentName && agentData[agentName]) {
                                    agentData[agentName].claimPaidCount++;
                                    agentData[agentName].claimPaidAmount += claimPaidValue;
                                    
                                    // Add to date-based tracking
                                    if (fileDate && agentDataByDate[agentName] && agentDataByDate[agentName][fileDate]) {
                                        agentDataByDate[agentName][fileDate].claimPaidCount++;
                                        agentDataByDate[agentName][fileDate].claimPaidAmount += claimPaidValue;
                                    }
                                }
                            }
                        }
                    }
                    
                    // Only process rows that have an Account No.
                    if (accountNo && accountNo.toString().trim() !== '') {
                        const remarkTypeStr = remarkType ? remarkType.toString().trim().toUpperCase() : '';
                        const relationStr = relation ? relation.toString().trim().toUpperCase() : '';
                        
                        // Count Debtor separately
                        if (relationStr === 'DEBTOR') {
                            debtorAccountNumbers.push(accountNo.toString().trim());
                        }
                        
                        // Count Predictive separately (before exclusions)
                        if (remarkTypeStr === 'PREDICTIVE') {
                            predictiveAccountNumbers.push(accountNo.toString().trim());
                        }
                        
                        // Exclude SMS
                        if (remarkTypeStr === 'SMS') {
                            excludedSMS++;
                            return;
                        }
                        
                        // Exclude blanks
                        if (remarkTypeStr === '' || remarkTypeStr === '(BLANKS)') {
                            excludedBlanks++;
                            return;
                        }
                        
                        // This row passes the filter - add account number
                        accountNumbers.push(accountNo.toString().trim());
                        
                        // Sum balance only for the first occurrence of each account (matches Worked on Ticket dedup logic)
                        if (balanceColumn && !seenAccountsForBalance.has(accountNo.toString().trim())) {
                            seenAccountsForBalance.add(accountNo.toString().trim());
                            const balVal = row[balanceColumn];
                            if (balVal !== null && balVal !== undefined && balVal !== '') {
                                const cleanBal = balVal.toString().replace(/,/g, '').replace(/[^\d.-]/g, '');
                                const parsed = parseFloat(cleanBal);
                                if (!isNaN(parsed)) totalBalance += parsed;
                            }
                        }
                    }
                });
                
                // Step 2: Remove duplicates
                const uniqueAccountNumbers = [...new Set(accountNumbers)];
                const uniquePredictiveNumbers = [...new Set(predictiveAccountNumbers)];
                const uniqueDebtorNumbers = [...new Set(debtorAccountNumbers)];
                
                // Step 3: Count unique values
                const uniqueCount = uniqueAccountNumbers.length;
                const duplicatesRemoved = accountNumbers.length - uniqueCount;
                const predictiveCount = uniquePredictiveNumbers.length;
                const predictiveDuplicates = predictiveAccountNumbers.length - predictiveCount;
                const debtorCount = uniqueDebtorNumbers.length;
                const debtorDuplicates = debtorAccountNumbers.length - debtorCount;
                
                resolve({
                    fileName: file.name,
                    fileDate: fileDate,
                    uniqueCount: uniqueCount,
                    countBeforeDedup: accountNumbers.length,
                    totalRows: totalRows,
                    excludedSMS: excludedSMS,
                    excludedBlanks: excludedBlanks,
                    totalExcluded: excludedSMS + excludedBlanks,
                    duplicatesRemoved: duplicatesRemoved,
                    predictiveCount: predictiveCount,
                    predictiveBeforeDedup: predictiveAccountNumbers.length,
                    predictiveDuplicates: predictiveDuplicates,
                    debtorCount: debtorCount,
                    debtorBeforeDedup: debtorAccountNumbers.length,
                    debtorDuplicates: debtorDuplicates,
                    ptpCount: ptpCount,
                    ptpTotalAmount: ptpTotalAmount,
                    predictivePtpCount: predictivePtpCount,
                    predictivePtpAmount: predictivePtpAmount,
                    claimPaidCount: claimPaidCount,
                    claimPaidTotalAmount: claimPaidTotalAmount,
                    totalBalance: totalBalance,
                    hasBalanceColumn: !!balanceColumn,
                    agentData: agentData,
                    agentDataByDate: agentDataByDate,
                    // Store arrays for cross-file deduplication
                    accountNumbersArray: uniqueAccountNumbers,
                    predictiveNumbersArray: uniquePredictiveNumbers,
                    debtorNumbersArray: uniqueDebtorNumbers
                });
                
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsArrayBuffer(file);
    });
}

function findColumn(data, headerName) {
    if (data.length === 0) return null;
    
    const firstRow = data[0];
    for (const key in firstRow) {
        if (key.toLowerCase().trim() === headerName.toLowerCase().trim()) {
            return key;
        }
    }
    return null;
}

function extractDateFromFilename(filename) {
    // Remove file extension first
    const nameWithoutExt = filename.replace(/\.(xlsx|xls)$/i, '');
    
    // Pattern to match DD-MMM-YY at the end of filename
    // Example: 01-Feb-26, 15-Mar-25, etc.
    const datePattern = /(\d{2})-([A-Za-z]{3})-(\d{2})$/;
    const match = nameWithoutExt.match(datePattern);
    
    if (match) {
        const day = match[1];
        const month = match[2];
        const year = match[3];
        
        // Convert month abbreviation to month number
        const monthMap = {
            'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
            'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
            'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
        };
        
        const monthNum = monthMap[month.toLowerCase()];
        if (monthNum) {
            // Assume 20XX for year (e.g., 26 becomes 2026)
            const fullYear = '20' + year;
            // Return ISO format: YYYY-MM-DD
            return `${fullYear}-${monthNum}-${day}`;
        }
    }
    
    return null;
}

function displayResults(results) {
    console.log('displayResults called with:', results);
    const validResults = results.filter(r => !r.error);
    console.log('Valid results:', validResults);
    
    if (validResults.length === 0) {
        console.log('No valid results, showing empty state');
        emptyState.style.display = 'flex';
        resultsSection.style.display = 'none';
        totalSummary.style.display = 'none';
        
        // Show errors if any
        const errors = results.filter(r => r.error);
        if (errors.length > 0) {
            alert('Errors processing files:\n' + errors.map(e => `${e.fileName}: ${e.error}`).join('\n'));
        }
        return;
    }
    
    console.log('Showing results');
    emptyState.style.display = 'none';
    resultsSection.style.display = 'block';
    
    // Calculate totals - MERGE ALL FILES FIRST then count unique
    // Collect all account numbers from all files
    const allAccountNumbers = [];
    const allPredictiveNumbers = [];
    const allDebtorNumbers = [];
    
    // Merge all account numbers from all files
    validResults.forEach(result => {
        allAccountNumbers.push(...result.accountNumbersArray);
        allPredictiveNumbers.push(...result.predictiveNumbersArray);
        allDebtorNumbers.push(...result.debtorNumbersArray);
    });
    
    // Remove duplicates ACROSS all files
    const uniqueAccountsAcrossFiles = [...new Set(allAccountNumbers)];
    const uniquePredictiveAcrossFiles = [...new Set(allPredictiveNumbers)];
    const uniqueDebtorAcrossFiles = [...new Set(allDebtorNumbers)];
    
    // For other stats, we sum them
    const totalBeforeDedup = validResults.reduce((sum, r) => sum + r.countBeforeDedup, 0);
    const totalPtpCount = validResults.reduce((sum, r) => sum + r.ptpCount, 0);
    const totalPtpAmount = validResults.reduce((sum, r) => sum + r.ptpTotalAmount, 0);
    const totalPredictivePtpCount = validResults.reduce((sum, r) => sum + r.predictivePtpCount, 0);
    const totalPredictivePtpAmount = validResults.reduce((sum, r) => sum + r.predictivePtpAmount, 0);
    const totalClaimPaidCount = validResults.reduce((sum, r) => sum + r.claimPaidCount, 0);
    const totalClaimPaidAmount = validResults.reduce((sum, r) => sum + r.claimPaidTotalAmount, 0);
    const totalBalance = validResults.reduce((sum, r) => sum + (r.totalBalance || 0), 0);
    const hasBalance = validResults.some(r => r.hasBalanceColumn);
    
    // Calculate Collection Rate: (Claim Paid Amount / Total Balance) * 100
    const collectionRate = totalBalance > 0 
        ? Math.round((totalClaimPaidAmount / totalBalance) * 100) 
        : 0;
    
    // Calculate Penetration: DIALS / WORKED ON TICKET (whole number, no decimals)
    const penetration = uniqueAccountsAcrossFiles.length > 0 
        ? Math.round(totalBeforeDedup / uniqueAccountsAcrossFiles.length * 100) 
        : 0;
    
    // Calculate Connected Rate: CONNECTED / WORKED ON TICKET (whole number, no decimals)
    const connectedRate = uniqueAccountsAcrossFiles.length > 0 
        ? Math.round(uniquePredictiveAcrossFiles.length / uniqueAccountsAcrossFiles.length * 100) 
        : 0;
    
    // Display total summary
    totalSummary.style.display = 'block';
    totalSummary.innerHTML = `
        <div class="summary-header-controls">
            <h3>Overall Summary (${validResults.length} file${validResults.length > 1 ? 's' : ''}) - Merged & Deduplicated</h3>
            <div class="summary-header-actions">
                <button id="togglePredictivePtp" class="btn-toggle btn-toggle-ptp">
                    <span class="toggle-icon">🔍</span>
                    <span class="toggle-text">Show Predictive PTP</span>
                </button>
                <button id="toggleMonthlyView" class="btn-toggle">
                    <span class="toggle-icon">📅</span>
                    <span class="toggle-text">View by Month</span>
                </button>
            </div>
        </div>
        <div id="overallView" class="summary-view">
            <div class="summary-container">
                <div class="summary-main">
                    <div class="summary-stat">
                        <div class="summary-stat-value">${totalBeforeDedup}</div>
                        <div class="summary-stat-label">Total Dials</div>
                    </div>
                    <div class="summary-stat">
                        <div class="summary-stat-value">${uniqueAccountsAcrossFiles.length}</div>
                        <div class="summary-stat-label">Worked on Ticket</div>
                    </div>
                    ${hasBalance ? `
                    <div class="summary-stat balance-highlight">
                        <div class="summary-stat-value">${totalBalance.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
                        <div class="summary-stat-label">Total Balance</div>
                    </div>` : ''}
                    <div class="summary-stat highlight">
                        <div class="summary-stat-value">${uniquePredictiveAcrossFiles.length}</div>
                        <div class="summary-stat-label">Connected</div>
                    </div>
                    <div class="summary-stat debtor-highlight">
                        <div class="summary-stat-value">${uniqueDebtorAcrossFiles.length}</div>
                        <div class="summary-stat-label">RPC</div>
                    </div>
                    <div class="summary-stat ptp-highlight">
                        <div class="summary-stat-value">${totalPtpCount}</div>
                        <div class="summary-stat-label">PTP Count</div>
                    </div>
                    <div class="summary-stat ptp-highlight">
                        <div class="summary-stat-value">${totalPtpAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
                        <div class="summary-stat-label">PTP Amount</div>
                    </div>
                    <div class="summary-stat claim-highlight">
                        <div class="summary-stat-value">${totalClaimPaidCount}</div>
                        <div class="summary-stat-label">Claim Paid Count</div>
                    </div>
                    <div class="summary-stat claim-highlight">
                        <div class="summary-stat-value">${totalClaimPaidAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
                        <div class="summary-stat-label">Claim Paid Amount</div>
                    </div>
                </div>
                <div class="summary-sidebar">
                    <div class="summary-stat-large penetration-highlight">
                        <div class="summary-stat-value-large">${penetration}%</div>
                        <div class="summary-stat-label-large">Penetration</div>
                    </div>
                    <div class="summary-stat-large connected-rate-highlight">
                        <div class="summary-stat-value-large">${connectedRate}%</div>
                        <div class="summary-stat-label-large">Connected Rate</div>
                    </div>
                    ${hasBalance ? `
                    <div class="summary-stat-large collection-rate-highlight">
                        <div class="summary-stat-value-large">${collectionRate}%</div>
                        <div class="summary-stat-label-large">Collection Rate</div>
                    </div>` : ''}
                </div>
            </div>

            <!-- Predictive PTP Breakdown - hidden by default -->
            <div id="predictivePtpBreakdown" class="predictive-ptp-breakdown" style="display: none;">
                <div class="breakdown-header">
                    <span class="breakdown-label">🔍 Predictive PTP Breakdown</span>
                    <span class="breakdown-sub">PTP records where Remark Type = PREDICTIVE</span>
                </div>
                <div class="breakdown-grid">
                    <div class="breakdown-card">
                        <div class="breakdown-card-section all">
                            <div class="breakdown-card-value">${totalPtpCount}</div>
                            <div class="breakdown-card-label">All PTP Count</div>
                        </div>
                        <div class="breakdown-divider">→</div>
                        <div class="breakdown-card-section predictive">
                            <div class="breakdown-card-value">${totalPredictivePtpCount}</div>
                            <div class="breakdown-card-label">Predictive PTP Count</div>
                        </div>
                        <div class="breakdown-divider">+</div>
                        <div class="breakdown-card-section other">
                            <div class="breakdown-card-value">${totalPtpCount - totalPredictivePtpCount}</div>
                            <div class="breakdown-card-label">Other PTP Count</div>
                        </div>
                    </div>
                    <div class="breakdown-card">
                        <div class="breakdown-card-section all">
                            <div class="breakdown-card-value">${totalPtpAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
                            <div class="breakdown-card-label">All PTP Amount</div>
                        </div>
                        <div class="breakdown-divider">→</div>
                        <div class="breakdown-card-section predictive">
                            <div class="breakdown-card-value">${totalPredictivePtpAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
                            <div class="breakdown-card-label">Predictive PTP Amount</div>
                        </div>
                        <div class="breakdown-divider">+</div>
                        <div class="breakdown-card-section other">
                            <div class="breakdown-card-value">${(totalPtpAmount - totalPredictivePtpAmount).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
                            <div class="breakdown-card-label">Other PTP Amount</div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <div id="monthlyView" class="summary-view" style="display: none;">
            <!-- Monthly breakdown will be generated here -->
        </div>
    `;
    
    // Set up toggle button event listeners
    setTimeout(() => {
        const toggleBtn = document.getElementById('toggleMonthlyView');
        if (toggleBtn) toggleBtn.addEventListener('click', () => toggleMonthlyView(validResults));

        const predictiveToggle = document.getElementById('togglePredictivePtp');
        if (predictiveToggle) predictiveToggle.addEventListener('click', () => togglePredictivePtp());
    }, 0);
    
    // Populate Excel table — store results for table re-render
    window._lastResults = results;
    window._predictivePtpOn = false;
    renderResultsTable(results, false);
    
    // Scroll to results (smooth scroll to top of results)
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// Graph Modal Functions
let chartInstance = null;
let currentGraphData = null;

function showGraph() {
    const validResults = processedResults.filter(r => !r.error && r.fileDate);
    
    if (validResults.length === 0) {
        alert('No valid data with dates to display in graph');
        return;
    }
    
    // Store data for filtering
    currentGraphData = validResults;
    
    // Show modal
    graphModal.style.display = 'flex';
    
    // Set up month filter
    const monthFilter = document.getElementById('monthFilter');
    const monthCompareSelectors = document.getElementById('monthCompareSelectors');
    const month1Select = document.getElementById('month1Select');
    const month2Select = document.getElementById('month2Select');
    const applyCompare = document.getElementById('applyCompare');
    
    // Populate month options
    const monthsData = getMonthsFromData(validResults);
    month1Select.innerHTML = monthsData.map(m => 
        `<option value="${m.key}">${m.name}</option>`
    ).join('');
    month2Select.innerHTML = monthsData.map(m => 
        `<option value="${m.key}">${m.name}</option>`
    ).join('');
    
    // Set default selections (most recent two months)
    if (monthsData.length >= 2) {
        month1Select.value = monthsData[0].key;
        month2Select.value = monthsData[1].key;
    }
    
    // Month filter change handler
    monthFilter.onchange = () => {
        if (monthFilter.value === 'compare') {
            monthCompareSelectors.style.display = 'flex';
        } else {
            monthCompareSelectors.style.display = 'none';
            renderAllDatesGraph(validResults);
        }
    };
    
    // Apply comparison button
    applyCompare.onclick = () => {
        const month1 = month1Select.value;
        const month2 = month2Select.value;
        renderComparisonGraph(validResults, month1, month2);
    };
    
    // Initial render - all dates
    renderAllDatesGraph(validResults);
}

function getMonthsFromData(results) {
    const monthsMap = {};
    
    results.forEach(result => {
        const date = new Date(result.fileDate);
        const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
        const monthName = date.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
        
        if (!monthsMap[monthKey]) {
            monthsMap[monthKey] = monthName;
        }
    });
    
    return Object.keys(monthsMap)
        .sort()
        .reverse()
        .map(key => ({ key, name: monthsMap[key] }));
}

function renderAllDatesGraph(results) {
    results.sort((a, b) => new Date(a.fileDate) - new Date(b.fileDate));
    const labels = results.map(r => new Date(r.fileDate).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' }));
    const ptpData = results.map(r => r.ptpTotalAmount);
    const claimData = results.map(r => r.claimPaidTotalAmount);
    destroyMainCharts();
    window.ptpChartInstance = buildSingleChart('ptpChart', labels, [{ label: 'PTP Amount', data: ptpData, borderColor: '#722ed1', backgroundColor: 'rgba(114,46,209,0.1)' }]);
    window.claimChartInstance = buildSingleChart('claimChart', labels, [{ label: 'Claim Paid Amount', data: claimData, borderColor: '#eb2f96', backgroundColor: 'rgba(235,47,150,0.1)' }]);
}

function renderComparisonGraph(results, month1Key, month2Key) {
    const month1Data = {}, month2Data = {};
    results.forEach(result => {
        const date = new Date(result.fileDate);
        const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
        const day = date.getDate();
        if (monthKey === month1Key) month1Data[day] = result;
        else if (monthKey === month2Key) month2Data[day] = result;
    });
    const allDays = [...new Set([...Object.keys(month1Data), ...Object.keys(month2Data)])].map(Number).sort((a, b) => a - b);
    const labels = allDays.map(day => `Day ${day}`);
    const m1 = new Date(month1Key + '-01').toLocaleDateString('en-US', { month: 'short', year: '2-digit' });
    const m2 = new Date(month2Key + '-01').toLocaleDateString('en-US', { month: 'short', year: '2-digit' });

    destroyMainCharts();
    window.ptpChartInstance = buildSingleChart('ptpChart', labels, [
        { label: m1, data: allDays.map(d => month1Data[d]?.ptpTotalAmount || null), borderColor: '#722ed1', backgroundColor: 'rgba(114,46,209,0.1)' },
        { label: m2, data: allDays.map(d => month2Data[d]?.ptpTotalAmount || null), borderColor: '#1890ff', backgroundColor: 'rgba(24,144,255,0.1)', borderDash: [5, 5] }
    ]);
    window.claimChartInstance = buildSingleChart('claimChart', labels, [
        { label: m1, data: allDays.map(d => month1Data[d]?.claimPaidTotalAmount || null), borderColor: '#eb2f96', backgroundColor: 'rgba(235,47,150,0.1)' },
        { label: m2, data: allDays.map(d => month2Data[d]?.claimPaidTotalAmount || null), borderColor: '#52c41a', backgroundColor: 'rgba(82,196,26,0.1)', borderDash: [5, 5] }
    ]);
}

function buildSingleChart(canvasId, labels, datasets) {
    const ctx = document.getElementById(canvasId).getContext('2d');
    return new Chart(ctx, {
        type: 'line',
        data: {
            labels,
            datasets: datasets.map(ds => ({
                label: ds.label,
                data: ds.data,
                borderColor: ds.borderColor,
                backgroundColor: ds.backgroundColor,
                borderWidth: 3,
                tension: 0.4,
                fill: true,
                pointRadius: 6,
                pointHoverRadius: 8,
                pointBackgroundColor: ds.borderColor,
                pointBorderColor: '#fff',
                pointBorderWidth: 2,
                borderDash: ds.borderDash || [],
                spanGaps: true
            }))
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            interaction: { mode: 'index', intersect: false },
            plugins: {
                legend: { display: true, position: 'top', labels: { font: { size: 13, weight: '600' }, padding: 16, usePointStyle: true } },
                tooltip: {
                    backgroundColor: 'rgba(0,0,0,0.8)', padding: 12,
                    callbacks: {
                        label: ctx => {
                            const val = ctx.parsed.y !== null
                                ? ctx.parsed.y.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
                                : 'No data';
                            return `${ctx.dataset.label}: ${val}`;
                        }
                    }
                }
            },
            scales: {
                y: { beginAtZero: true, ticks: { callback: v => v.toLocaleString('en-US'), font: { size: 11 }, color: 'rgba(0,0,0,0.65)' }, grid: { color: 'rgba(0,0,0,0.05)' } },
                x: { ticks: { font: { size: 11 }, color: 'rgba(0,0,0,0.65)' }, grid: { display: false } }
            }
        }
    });
}

function destroyMainCharts() {
    if (window.ptpChartInstance) { window.ptpChartInstance.destroy(); window.ptpChartInstance = null; }
    if (window.claimChartInstance) { window.claimChartInstance.destroy(); window.claimChartInstance = null; }
}

function hideGraph() {
    graphModal.style.display = 'none';
    destroyMainCharts();
}

// Close modal with Escape key
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
        if (graphModal.style.display === 'flex') {
            hideGraph();
        }
        if (agentModal.style.display === 'flex') {
            hideAgentCollection();
        }
    }
});

// Agent Collection Modal Functions
function showAgentCollection() {
    const validResults = processedResults.filter(r => !r.error && r.agentData);
    
    if (validResults.length === 0) {
        alert('No agent data available');
        return;
    }
    
    // Aggregate agent data from all files
    const allAgentsData = {};
    const allAgentsDataByDate = {};
    const allAgentsDataByMonth = {}; // NEW: Data grouped by month
    
    validResults.forEach(result => {
        Object.keys(result.agentData).forEach(agentName => {
            if (!allAgentsData[agentName]) {
                allAgentsData[agentName] = {
                    ptpCount: 0,
                    ptpAmount: 0,
                    claimPaidCount: 0,
                    claimPaidAmount: 0
                };
            }
            allAgentsData[agentName].ptpCount += result.agentData[agentName].ptpCount;
            allAgentsData[agentName].ptpAmount += result.agentData[agentName].ptpAmount;
            allAgentsData[agentName].claimPaidCount += result.agentData[agentName].claimPaidCount;
            allAgentsData[agentName].claimPaidAmount += result.agentData[agentName].claimPaidAmount;
        });
        
        // Collect date-based data
        if (result.agentDataByDate) {
            Object.keys(result.agentDataByDate).forEach(agentName => {
                if (!allAgentsDataByDate[agentName]) {
                    allAgentsDataByDate[agentName] = {};
                }
                Object.keys(result.agentDataByDate[agentName]).forEach(date => {
                    if (!allAgentsDataByDate[agentName][date]) {
                        allAgentsDataByDate[agentName][date] = {
                            ptpCount: 0,
                            ptpAmount: 0,
                            claimPaidCount: 0,
                            claimPaidAmount: 0
                        };
                    }
                    const dateData = result.agentDataByDate[agentName][date];
                    allAgentsDataByDate[agentName][date].ptpCount += dateData.ptpCount;
                    allAgentsDataByDate[agentName][date].ptpAmount += dateData.ptpAmount;
                    allAgentsDataByDate[agentName][date].claimPaidCount += dateData.claimPaidCount;
                    allAgentsDataByDate[agentName][date].claimPaidAmount += dateData.claimPaidAmount;
                    
                    // Group by month
                    const d = new Date(date);
                    const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
                    
                    if (!allAgentsDataByMonth[monthKey]) {
                        allAgentsDataByMonth[monthKey] = {};
                    }
                    if (!allAgentsDataByMonth[monthKey][agentName]) {
                        allAgentsDataByMonth[monthKey][agentName] = {
                            ptpCount: 0,
                            ptpAmount: 0,
                            claimPaidCount: 0,
                            claimPaidAmount: 0
                        };
                    }
                    allAgentsDataByMonth[monthKey][agentName].ptpCount += dateData.ptpCount;
                    allAgentsDataByMonth[monthKey][agentName].ptpAmount += dateData.ptpAmount;
                    allAgentsDataByMonth[monthKey][agentName].claimPaidCount += dateData.claimPaidCount;
                    allAgentsDataByMonth[monthKey][agentName].claimPaidAmount += dateData.claimPaidAmount;
                });
            });
        }
    });
    
    // Filter out agents with 0 PTP AND 0 Claim Paid
    // Keep agents with at least 1 PTP OR at least 1 Claim Paid
    Object.keys(allAgentsData).forEach(agentName => {
        const agent = allAgentsData[agentName];
        if (agent.ptpCount === 0 && agent.claimPaidCount === 0) {
            // Remove this agent
            delete allAgentsData[agentName];
            delete allAgentsDataByDate[agentName];
        }
    });
    
    // Show modal
    agentModal.style.display = 'flex';
    
    // Populate month filter
    const agentMonthFilter = document.getElementById('agentMonthFilter');
    const availableMonths = Object.keys(allAgentsDataByMonth).sort().reverse();
    agentMonthFilter.innerHTML = '<option value="all">All Months</option>' + 
        availableMonths.map(monthKey => {
            const d = new Date(monthKey + '-01');
            const monthName = d.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            return `<option value="${monthKey}">${monthName}</option>`;
        }).join('');
    
    // Populate agent filter
    const agentFilter = document.getElementById('agentFilter');
    const agentNames = Object.keys(allAgentsData).sort();
    agentFilter.innerHTML = '<option value="all">All Agents</option>' + 
        agentNames.map(name => `<option value="${name}">${name}</option>`).join('');
    
    // Month filter change handler
    agentMonthFilter.onchange = () => {
        const selectedMonth = agentMonthFilter.value;
        const selectedAgent = agentFilter.value;
        
        if (selectedMonth === 'all') {
            renderAgentData(allAgentsData, allAgentsDataByDate, selectedAgent);
        } else {
            // Filter data by selected month
            const monthAgentsData = allAgentsDataByMonth[selectedMonth] || {};
            
            // Filter out agents with 0 PTP AND 0 Claim Paid in this month
            const filteredMonthAgentsData = {};
            Object.keys(monthAgentsData).forEach(agentName => {
                const agent = monthAgentsData[agentName];
                if (agent.ptpCount > 0 || agent.claimPaidCount > 0) {
                    filteredMonthAgentsData[agentName] = agent;
                }
            });
            
            const monthAgentsDataByDate = {};
            
            // Filter date data for selected month
            Object.keys(allAgentsDataByDate).forEach(agentName => {
                // Only include agents that have collections in this month
                if (filteredMonthAgentsData[agentName]) {
                    monthAgentsDataByDate[agentName] = {};
                    Object.keys(allAgentsDataByDate[agentName]).forEach(date => {
                        const d = new Date(date);
                        const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
                        if (monthKey === selectedMonth) {
                            monthAgentsDataByDate[agentName][date] = allAgentsDataByDate[agentName][date];
                        }
                    });
                }
            });
            
            renderAgentData(filteredMonthAgentsData, monthAgentsDataByDate, selectedAgent);
        }
    };
    
    // Agent filter change handler
    agentFilter.onchange = () => {
        const selectedMonth = agentMonthFilter.value;
        const selectedAgent = agentFilter.value;
        
        if (selectedMonth === 'all') {
            renderAgentData(allAgentsData, allAgentsDataByDate, selectedAgent);
        } else {
            const monthAgentsData = allAgentsDataByMonth[selectedMonth] || {};
            
            // Filter out agents with 0 PTP AND 0 Claim Paid in this month
            const filteredMonthAgentsData = {};
            Object.keys(monthAgentsData).forEach(agentName => {
                const agent = monthAgentsData[agentName];
                if (agent.ptpCount > 0 || agent.claimPaidCount > 0) {
                    filteredMonthAgentsData[agentName] = agent;
                }
            });
            
            const monthAgentsDataByDate = {};
            
            Object.keys(allAgentsDataByDate).forEach(agentName => {
                // Only include agents that have collections in this month
                if (filteredMonthAgentsData[agentName]) {
                    monthAgentsDataByDate[agentName] = {};
                    Object.keys(allAgentsDataByDate[agentName]).forEach(date => {
                        const d = new Date(date);
                        const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
                        if (monthKey === selectedMonth) {
                            monthAgentsDataByDate[agentName][date] = allAgentsDataByDate[agentName][date];
                        }
                    });
                }
            });
            
            renderAgentData(filteredMonthAgentsData, monthAgentsDataByDate, selectedAgent);
        }
    };
    
    // Initial render - all agents, all months
    renderAgentData(allAgentsData, allAgentsDataByDate, 'all');
}

function renderAgentData(allAgentsData, allAgentsDataByDate, selectedAgent) {
    const agentStatsContainer = document.getElementById('agentStatsContainer');
    const agentTableContainer = document.getElementById('agentTableContainer');
    const agentGraphContainer = document.getElementById('agentGraphContainer');
    
    if (selectedAgent === 'all') {
        // Hide graph for all agents view
        agentGraphContainer.style.display = 'none';
        
        // Show summary stats for all agents
        const totalPtpCount = Object.values(allAgentsData).reduce((sum, agent) => sum + agent.ptpCount, 0);
        const totalPtpAmount = Object.values(allAgentsData).reduce((sum, agent) => sum + agent.ptpAmount, 0);
        const totalClaimPaidCount = Object.values(allAgentsData).reduce((sum, agent) => sum + agent.claimPaidCount, 0);
        const totalClaimPaidAmount = Object.values(allAgentsData).reduce((sum, agent) => sum + agent.claimPaidAmount, 0);
        
        agentStatsContainer.innerHTML = `
            <div class="agent-summary-stats">
                <div class="agent-stat-card ptp">
                    <div class="agent-stat-label">Total PTP Count</div>
                    <div class="agent-stat-value">${totalPtpCount.toLocaleString('en-US')}</div>
                </div>
                <div class="agent-stat-card ptp">
                    <div class="agent-stat-label">Total PTP Amount</div>
                    <div class="agent-stat-value">${totalPtpAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
                </div>
                <div class="agent-stat-card claim">
                    <div class="agent-stat-label">Total Claim Paid Count</div>
                    <div class="agent-stat-value">${totalClaimPaidCount.toLocaleString('en-US')}</div>
                </div>
                <div class="agent-stat-card claim">
                    <div class="agent-stat-label">Total Claim Paid Amount</div>
                    <div class="agent-stat-value">${totalClaimPaidAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
                </div>
            </div>
        `;
        
        // Show table of all agents
        const agentNames = Object.keys(allAgentsData).sort();
        agentTableContainer.innerHTML = `
            <h3 style="margin: 24px 0 16px 0; font-size: 16px; font-weight: 600; color: rgba(0, 0, 0, 0.85);">Agent Collection Details</h3>
            <div class="table-container">
                <table class="excel-table agent-table">
                    <thead>
                        <tr>
                            <th>Agent Name</th>
                            <th>PTP Count</th>
                            <th>PTP Amount</th>
                            <th>Claim Paid Count</th>
                            <th>Claim Paid Amount</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${agentNames.map(name => {
                            const agent = allAgentsData[name];
                            return `
                                <tr>
                                    <td>${name}</td>
                                    <td>${agent.ptpCount.toLocaleString('en-US')}</td>
                                    <td>${agent.ptpAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                                    <td>${agent.claimPaidCount.toLocaleString('en-US')}</td>
                                    <td>${agent.claimPaidAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                                </tr>
                            `;
                        }).join('')}
                    </tbody>
                </table>
            </div>
        `;
    } else {
        // Show details for selected agent
        const agent = allAgentsData[selectedAgent];
        const totalCollection = agent.ptpAmount + agent.claimPaidAmount;
        
        agentStatsContainer.innerHTML = `
            <div class="agent-detail-header">
                <h3>Agent: ${selectedAgent}</h3>
            </div>
            <div class="agent-summary-stats">
                <div class="agent-stat-card ptp">
                    <div class="agent-stat-label">PTP Count</div>
                    <div class="agent-stat-value">${agent.ptpCount.toLocaleString('en-US')}</div>
                </div>
                <div class="agent-stat-card ptp">
                    <div class="agent-stat-label">PTP Amount</div>
                    <div class="agent-stat-value">${agent.ptpAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
                </div>
                <div class="agent-stat-card claim">
                    <div class="agent-stat-label">Claim Paid Count</div>
                    <div class="agent-stat-value">${agent.claimPaidCount.toLocaleString('en-US')}</div>
                </div>
                <div class="agent-stat-card claim">
                    <div class="agent-stat-label">Claim Paid Amount</div>
                    <div class="agent-stat-value">${agent.claimPaidAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
                </div>
            </div>
        `;
        
        agentTableContainer.innerHTML = '';
        
        // Show graph container
        agentGraphContainer.style.display = 'block';
        
        // Set up graph controls
        setupAgentGraphControls(selectedAgent, allAgentsDataByDate[selectedAgent]);
        
        // Render initial daily graph
        renderAgentDailyGraph(selectedAgent, allAgentsDataByDate[selectedAgent]);
    }
}

function hideAgentCollection() {
    agentModal.style.display = 'none';
    destroyAgentCharts();
}

// Agent Graph Functions
function setupAgentGraphControls(agentName, agentDateData) {
    const graphType = document.getElementById('agentGraphType');
    const monthlyControls = document.getElementById('monthlyComparisonControls');
    const month1Select = document.getElementById('agentMonth1');
    const month2Select = document.getElementById('agentMonth2');
    const applyBtn = document.getElementById('applyAgentMonthCompare');
    
    // Get available months
    const months = getMonthsFromAgentData(agentDateData);
    
    // Populate month selectors
    month1Select.innerHTML = months.map(m => 
        `<option value="${m.key}">${m.name}</option>`
    ).join('');
    month2Select.innerHTML = months.map(m => 
        `<option value="${m.key}">${m.name}</option>`
    ).join('');
    
    // Set defaults
    if (months.length >= 2) {
        month1Select.value = months[0].key;
        month2Select.value = months[1].key;
    }
    
    // Graph type change handler
    graphType.onchange = () => {
        if (graphType.value === 'monthly') {
            monthlyControls.style.display = 'block';
        } else {
            monthlyControls.style.display = 'none';
            renderAgentDailyGraph(agentName, agentDateData);
        }
    };
    
    // Apply comparison handler
    applyBtn.onclick = () => {
        const month1 = month1Select.value;
        const month2 = month2Select.value;
        renderAgentMonthlyComparisonGraph(agentName, agentDateData, month1, month2);
    };
}

function getMonthsFromAgentData(agentDateData) {
    const monthsMap = {};
    
    Object.keys(agentDateData).forEach(date => {
        const d = new Date(date);
        const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
        const monthName = d.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
        monthsMap[monthKey] = monthName;
    });
    
    return Object.keys(monthsMap)
        .sort()
        .reverse()
        .map(key => ({ key, name: monthsMap[key] }));
}

function renderAgentDailyGraph(agentName, agentDateData) {
    const dates = Object.keys(agentDateData).sort();
    const labels = dates.map(date => new Date(date).toLocaleDateString('en-US', { month: 'short', day: 'numeric' }));
    const ptpData = dates.map(date => agentDateData[date].ptpAmount);
    const claimData = dates.map(date => agentDateData[date].claimPaidAmount);

    destroyAgentCharts();
    window.agentPtpChartInstance = buildSingleChart('agentPtpChart', labels, [
        { label: 'PTP Amount', data: ptpData, borderColor: '#722ed1', backgroundColor: 'rgba(114,46,209,0.1)' }
    ]);
    window.agentClaimChartInstance = buildSingleChart('agentClaimChart', labels, [
        { label: 'Claim Paid Amount', data: claimData, borderColor: '#eb2f96', backgroundColor: 'rgba(235,47,150,0.1)' }
    ]);
}

function renderAgentMonthlyComparisonGraph(agentName, agentDateData, month1Key, month2Key) {
    const month1Data = {}, month2Data = {};
    Object.keys(agentDateData).forEach(date => {
        const d = new Date(date);
        const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
        const day = d.getDate();
        if (monthKey === month1Key) month1Data[day] = agentDateData[date];
        else if (monthKey === month2Key) month2Data[day] = agentDateData[date];
    });

    const allDays = [...new Set([...Object.keys(month1Data), ...Object.keys(month2Data)])].map(Number).sort((a, b) => a - b);
    const labels = allDays.map(day => `Day ${day}`);
    const m1 = new Date(month1Key + '-01').toLocaleDateString('en-US', { month: 'short', year: '2-digit' });
    const m2 = new Date(month2Key + '-01').toLocaleDateString('en-US', { month: 'short', year: '2-digit' });

    destroyAgentCharts();
    window.agentPtpChartInstance = buildSingleChart('agentPtpChart', labels, [
        { label: m1, data: allDays.map(d => month1Data[d]?.ptpAmount || null), borderColor: '#722ed1', backgroundColor: 'rgba(114,46,209,0.1)' },
        { label: m2, data: allDays.map(d => month2Data[d]?.ptpAmount || null), borderColor: '#1890ff', backgroundColor: 'rgba(24,144,255,0.1)', borderDash: [5, 5] }
    ]);
    window.agentClaimChartInstance = buildSingleChart('agentClaimChart', labels, [
        { label: m1, data: allDays.map(d => month1Data[d]?.claimPaidAmount || null), borderColor: '#eb2f96', backgroundColor: 'rgba(235,47,150,0.1)' },
        { label: m2, data: allDays.map(d => month2Data[d]?.claimPaidAmount || null), borderColor: '#52c41a', backgroundColor: 'rgba(82,196,26,0.1)', borderDash: [5, 5] }
    ]);
}

function destroyAgentCharts() {
    if (window.agentPtpChartInstance) { window.agentPtpChartInstance.destroy(); window.agentPtpChartInstance = null; }
    if (window.agentClaimChartInstance) { window.agentClaimChartInstance.destroy(); window.agentClaimChartInstance = null; }
}

// Monthly View Toggle
let isMonthlyView = false;

function renderResultsTable(results, showPredictive) {
    const thead = document.querySelector('#resultsTable thead tr');
    const hasBalance = results.some(r => r.hasBalanceColumn);
    const balCol = hasBalance ? `<th class="col-balance">Total Balance</th>` : '';
    const balColPred = hasBalance ? `<th class="col-balance">Total Balance</th>` : '';

    if (showPredictive) {
        thead.innerHTML = `
            <th>Files</th>
            <th>Worked on Ticket</th>
            ${balColPred}
            <th>Total Dials</th>
            <th>Connected</th>
            <th>RPC</th>
            <th>PTP Count</th>
            <th class="col-predictive">↳ Predictive</th>
            <th class="col-predictive">↳ Other</th>
            <th>PTP Amount</th>
            <th class="col-predictive">↳ Predictive</th>
            <th class="col-predictive">↳ Other</th>
            <th>Claim Paid Count</th>
            <th>Claim Paid Amount</th>
        `;
    } else {
        thead.innerHTML = `
            <th>Files</th>
            <th>Worked on Ticket</th>
            ${balCol}
            <th>Total Dials</th>
            <th>Connected</th>
            <th>RPC</th>
            <th>PTP Count</th>
            <th>PTP Amount</th>
            <th>Claim Paid Count</th>
            <th>Claim Paid Amount</th>
        `;
    }

    resultsTableBody.innerHTML = results.map(result => {
        if (result.error) {
            const span = showPredictive ? (hasBalance ? 14 : 13) : (hasBalance ? 10 : 9);
            return `<tr>
                <td>${result.fileName}</td>
                <td colspan="${span - 1}" style="color:#e53e3e;text-align:center;">Error: ${result.error}</td>
            </tr>`;
        }
        const fmt = (n) => n.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
        const otherPtpCount = result.ptpCount - (result.predictivePtpCount || 0);
        const otherPtpAmount = result.ptpTotalAmount - (result.predictivePtpAmount || 0);
        const balCell = hasBalance
            ? `<td class="col-balance">${fmt(result.totalBalance || 0)}</td>`
            : '';

        if (showPredictive) {
            return `<tr>
                <td>${result.fileName}</td>
                <td>${result.uniqueCount}</td>
                ${balCell}
                <td>${result.countBeforeDedup}</td>
                <td>${result.predictiveCount}</td>
                <td>${result.debtorCount}</td>
                <td>${result.ptpCount}</td>
                <td class="col-predictive">${result.predictivePtpCount || 0}</td>
                <td class="col-predictive">${otherPtpCount}</td>
                <td>${fmt(result.ptpTotalAmount)}</td>
                <td class="col-predictive">${fmt(result.predictivePtpAmount || 0)}</td>
                <td class="col-predictive">${fmt(otherPtpAmount)}</td>
                <td>${result.claimPaidCount}</td>
                <td>${fmt(result.claimPaidTotalAmount)}</td>
            </tr>`;
        } else {
            return `<tr>
                <td>${result.fileName}</td>
                <td>${result.uniqueCount}</td>
                ${balCell}
                <td>${result.countBeforeDedup}</td>
                <td>${result.predictiveCount}</td>
                <td>${result.debtorCount}</td>
                <td>${result.ptpCount}</td>
                <td>${fmt(result.ptpTotalAmount)}</td>
                <td>${result.claimPaidCount}</td>
                <td>${fmt(result.claimPaidTotalAmount)}</td>
            </tr>`;
        }
    }).join('');
}

function togglePredictivePtp() {
    window._predictivePtpOn = !window._predictivePtpOn;
    const btn = document.getElementById('togglePredictivePtp');
    const breakdown = document.getElementById('predictivePtpBreakdown');

    if (window._predictivePtpOn) {
        btn.classList.add('active');
        btn.querySelector('.toggle-text').textContent = 'Hide Predictive PTP';
        if (breakdown) breakdown.style.display = 'block';
    } else {
        btn.classList.remove('active');
        btn.querySelector('.toggle-text').textContent = 'Show Predictive PTP';
        if (breakdown) breakdown.style.display = 'none';
    }
    renderResultsTable(window._lastResults, window._predictivePtpOn);
}

function toggleMonthlyView(results) {
    const overallView = document.getElementById('overallView');
    const monthlyView = document.getElementById('monthlyView');
    const toggleBtn = document.getElementById('toggleMonthlyView');
    const toggleText = toggleBtn.querySelector('.toggle-text');
    const toggleIcon = toggleBtn.querySelector('.toggle-icon');
    
    isMonthlyView = !isMonthlyView;
    
    if (isMonthlyView) {
        // Show monthly view
        overallView.style.display = 'none';
        monthlyView.style.display = 'block';
        toggleText.textContent = 'View Overall';
        toggleIcon.textContent = '📊';
        
        // Generate monthly breakdown
        generateMonthlyBreakdown(results);
    } else {
        // Show overall view
        overallView.style.display = 'block';
        monthlyView.style.display = 'none';
        toggleText.textContent = 'View by Month';
        toggleIcon.textContent = '📅';
    }
}

function generateMonthlyBreakdown(results) {
    const monthlyView = document.getElementById('monthlyView');
    
    // Group results by month
    const monthlyData = {};
    
    results.forEach(result => {
        if (result.fileDate) {
            const date = new Date(result.fileDate);
            const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
            const monthName = date.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            
            if (!monthlyData[monthKey]) {
                monthlyData[monthKey] = {
                    monthName: monthName,
                    monthKey: monthKey,
                    files: [],
                    accountNumbers: [],
                    predictiveNumbers: [],
                    debtorNumbers: [],
                    totalDials: 0,
                    ptpCount: 0,
                    ptpAmount: 0,
                    predictivePtpCount: 0,
                    predictivePtpAmount: 0,
                    claimPaidCount: 0,
                    claimPaidAmount: 0,
                    totalBalance: 0
                };
            }
            
            monthlyData[monthKey].files.push(result);
            monthlyData[monthKey].accountNumbers.push(...result.accountNumbersArray);
            monthlyData[monthKey].predictiveNumbers.push(...result.predictiveNumbersArray);
            monthlyData[monthKey].debtorNumbers.push(...result.debtorNumbersArray);
            monthlyData[monthKey].totalDials += result.countBeforeDedup;
            monthlyData[monthKey].ptpCount += result.ptpCount;
            monthlyData[monthKey].ptpAmount += result.ptpTotalAmount;
            monthlyData[monthKey].predictivePtpCount += (result.predictivePtpCount || 0);
            monthlyData[monthKey].predictivePtpAmount += (result.predictivePtpAmount || 0);
            monthlyData[monthKey].claimPaidCount += result.claimPaidCount;
            monthlyData[monthKey].claimPaidAmount += result.claimPaidTotalAmount;
            monthlyData[monthKey].totalBalance += (result.totalBalance || 0);
        }
    });
    
    // Sort by month (most recent first)
    const sortedMonths = Object.keys(monthlyData).sort().reverse();
    
    // Calculate unique counts for each month
    const monthlyMetrics = sortedMonths.map(monthKey => {
        const data = monthlyData[monthKey];
        const uniqueAccounts = [...new Set(data.accountNumbers)];
        const uniquePredictive = [...new Set(data.predictiveNumbers)];
        const uniqueDebtor = [...new Set(data.debtorNumbers)];
        
        const penetration = uniqueAccounts.length > 0 
            ? Math.round(data.totalDials / uniqueAccounts.length * 100) 
            : 0;
        const connectedRate = uniqueAccounts.length > 0 
            ? Math.round(uniquePredictive.length / uniqueAccounts.length * 100) 
            : 0;
        
        // Calculate Collection Rate: (Claim Paid Amount / Total Balance) * 100
        const collectionRate = data.totalBalance > 0
            ? Math.round((data.claimPaidAmount / data.totalBalance) * 100)
            : 0;
        
        return {
            monthName: data.monthName,
            totalDials: data.totalDials,
            workedOnTicket: uniqueAccounts.length,
            connected: uniquePredictive.length,
            debtor: uniqueDebtor.length,
            totalBalance: data.totalBalance,
            ptpCount: data.ptpCount,
            ptpAmount: data.ptpAmount,
            predictivePtpCount: data.predictivePtpCount,
            predictivePtpAmount: data.predictivePtpAmount,
            otherPtpCount: data.ptpCount - data.predictivePtpCount,
            otherPtpAmount: data.ptpAmount - data.predictivePtpAmount,
            claimPaidCount: data.claimPaidCount,
            claimPaidAmount: data.claimPaidAmount,
            collectionRate: collectionRate,
            penetration: penetration,
            connectedRate: connectedRate
        };
    });
    
    // Helper function to generate trend indicator
    function getTrendIndicator(current, previous) {
        if (previous === undefined) return '';
        const diff = current - previous;
        if (diff === 0) return '';
        const color = diff > 0 ? '#10b981' : '#ef4444';
        const sign = diff > 0 ? '+' : '';
        return `<span class="trend-indicator" style="color:${color}">(${sign}${diff.toLocaleString('en-US')})</span>`;
    }
    
    function getTrendIndicatorPercent(current, previous) {
        if (previous === undefined) return '';
        const diff = current - previous;
        if (diff === 0) return '';
        const color = diff > 0 ? '#10b981' : '#ef4444';
        const sign = diff > 0 ? '+' : '';
        return `<span class="trend-indicator" style="color:${color}">(${sign}${diff}%)</span>`;
    }
    
    function getTrendIndicatorAmount(current, previous) {
        if (previous === undefined) return '';
        const diff = current - previous;
        if (diff === 0) return '';
        const color = diff > 0 ? '#10b981' : '#ef4444';
        const sign = diff > 0 ? '+' : '';
        return `<span class="trend-indicator" style="color:${color}">(${sign}${diff.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})})</span>`;
    }

    const fmt2 = (n) => n.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});

    // Generate month cards — one card per month, metrics grouped by category
    let monthlyHTML = `<div class="monthly-cards-container">`;

    monthlyMetrics.forEach((month, index) => {
        const prev = monthlyMetrics[index + 1];
        monthlyHTML += `
        <div class="monthly-card">
            <div class="monthly-card-header">
                <span class="monthly-card-title">${month.monthName}</span>
                <span class="monthly-card-files">${results.filter(r => {
                    if (!r.fileDate) return false;
                    const d = new Date(r.fileDate);
                    const mk = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0');
                    const sm = sortedMonths[index];
                    return mk === sm;
                }).length} file(s)</span>
            </div>
            <div class="monthly-card-body">

                <!-- General metrics -->
                <div class="monthly-section general-section">
                    <div class="monthly-metric-row">
                        <span class="monthly-metric-name">Total Dials</span>
                        <span class="monthly-metric-value">${month.totalDials.toLocaleString('en-US')} ${getTrendIndicator(month.totalDials, prev?.totalDials)}</span>
                    </div>
                    <div class="monthly-metric-row">
                        <span class="monthly-metric-name">Worked on Ticket</span>
                        <span class="monthly-metric-value">${month.workedOnTicket.toLocaleString('en-US')} ${getTrendIndicator(month.workedOnTicket, prev?.workedOnTicket)}</span>
                    </div>
                    ${month.totalBalance > 0 ? `
                    <div class="monthly-metric-row">
                        <span class="monthly-metric-name">Total Balance</span>
                        <span class="monthly-metric-value balance-value">${fmt2(month.totalBalance)} ${getTrendIndicatorAmount(month.totalBalance, prev?.totalBalance)}</span>
                    </div>` : ''}
                    <div class="monthly-metric-row">
                        <span class="monthly-metric-name">Connected</span>
                        <span class="monthly-metric-value">${month.connected.toLocaleString('en-US')} ${getTrendIndicator(month.connected, prev?.connected)}</span>
                    </div>
                    <div class="monthly-metric-row">
                        <span class="monthly-metric-name">RPC</span>
                        <span class="monthly-metric-value">${month.debtor.toLocaleString('en-US')} ${getTrendIndicator(month.debtor, prev?.debtor)}</span>
                    </div>
                    <div class="monthly-metric-row">
                        <span class="monthly-metric-name">Penetration</span>
                        <span class="monthly-metric-value">${month.penetration}% ${getTrendIndicatorPercent(month.penetration, prev?.penetration)}</span>
                    </div>
                    <div class="monthly-metric-row">
                        <span class="monthly-metric-name">Connected Rate</span>
                        <span class="monthly-metric-value">${month.connectedRate}% ${getTrendIndicatorPercent(month.connectedRate, prev?.connectedRate)}</span>
                    </div>
                    ${month.totalBalance > 0 ? `
                    <div class="monthly-metric-row">
                        <span class="monthly-metric-name">Collection Rate</span>
                        <span class="monthly-metric-value collection-rate-value">${month.collectionRate}% ${getTrendIndicatorPercent(month.collectionRate, prev?.collectionRate)}</span>
                    </div>` : ''}
                </div>

                <!-- PTP section with clear separator -->
                <div class="monthly-section ptp-section">
                    <div class="monthly-section-label ptp-section-label">📌 PTP</div>
                    <div class="monthly-metric-row">
                        <span class="monthly-metric-name">Total PTP Count</span>
                        <span class="monthly-metric-value ptp-value">${month.ptpCount.toLocaleString('en-US')} ${getTrendIndicator(month.ptpCount, prev?.ptpCount)}</span>
                    </div>
                    <div class="monthly-metric-row">
                        <span class="monthly-metric-name">Total PTP Amount</span>
                        <span class="monthly-metric-value ptp-value">${fmt2(month.ptpAmount)} ${getTrendIndicatorAmount(month.ptpAmount, prev?.ptpAmount)}</span>
                    </div>
                    <div class="monthly-ptp-breakdown">
                        <div class="monthly-ptp-sub predictive-sub">
                            <div class="monthly-ptp-sub-label">Predictive</div>
                            <div class="monthly-ptp-sub-row">
                                <span>Count</span>
                                <strong>${month.predictivePtpCount.toLocaleString('en-US')} ${getTrendIndicator(month.predictivePtpCount, prev?.predictivePtpCount)}</strong>
                            </div>
                            <div class="monthly-ptp-sub-row">
                                <span>Amount</span>
                                <strong>${fmt2(month.predictivePtpAmount)} ${getTrendIndicatorAmount(month.predictivePtpAmount, prev?.predictivePtpAmount)}</strong>
                            </div>
                        </div>
                        <div class="monthly-ptp-divider">+</div>
                        <div class="monthly-ptp-sub other-sub">
                            <div class="monthly-ptp-sub-label">Other</div>
                            <div class="monthly-ptp-sub-row">
                                <span>Count</span>
                                <strong>${month.otherPtpCount.toLocaleString('en-US')} ${getTrendIndicator(month.otherPtpCount, prev?.otherPtpCount)}</strong>
                            </div>
                            <div class="monthly-ptp-sub-row">
                                <span>Amount</span>
                                <strong>${fmt2(month.otherPtpAmount)} ${getTrendIndicatorAmount(month.otherPtpAmount, prev?.otherPtpAmount)}</strong>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- Claim Paid section -->
                <div class="monthly-section claim-section">
                    <div class="monthly-section-label claim-section-label">💳 Claim Paid</div>
                    <div class="monthly-metric-row">
                        <span class="monthly-metric-name">Claim Paid Count</span>
                        <span class="monthly-metric-value claim-value">${month.claimPaidCount.toLocaleString('en-US')} ${getTrendIndicator(month.claimPaidCount, prev?.claimPaidCount)}</span>
                    </div>
                    <div class="monthly-metric-row">
                        <span class="monthly-metric-name">Claim Paid Amount</span>
                        <span class="monthly-metric-value claim-value">${fmt2(month.claimPaidAmount)} ${getTrendIndicatorAmount(month.claimPaidAmount, prev?.claimPaidAmount)}</span>
                    </div>
                </div>

            </div>
        </div>`;
    });

    monthlyHTML += `</div>`;
    monthlyView.innerHTML = monthlyHTML;
}