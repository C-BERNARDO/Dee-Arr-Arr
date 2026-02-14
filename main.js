// DOM Elements
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const processBtn = document.getElementById('processBtn');
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
}

// File processing
async function processFiles() {
    loadingOverlay.style.display = 'flex';
    resultsSection.style.display = 'none';
    emptyState.style.display = 'none';
    resultsTableBody.innerHTML = '';

    const results = [];

    for (const file of selectedFiles) {
        try {
            console.log('Processing file:', file.name);
            const result = await processExcelFile(file);
            console.log('Result:', result);
            results.push(result);
        } catch (error) {
            console.error('Error processing file:', file.name, error);
            results.push({
                fileName: file.name,
                error: error.message
            });
        }
    }

    console.log('All results:', results);
    processedResults = results; // Store for graph
    displayResults(results);
    loadingOverlay.style.display = 'none';
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
                let totalRows = jsonData.length;
                let excludedSMS = 0;
                let excludedBlanks = 0;
                let ptpCount = 0;
                let ptpTotalAmount = 0;
                let claimPaidCount = 0;
                let claimPaidTotalAmount = 0;
                
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
                        // Convert to string first to handle different formats
                        const ptpStr = ptpAmount.toString().trim();
                        if (ptpStr !== '' && ptpStr !== '0' && ptpStr !== '0.00') {
                            // Remove commas and any currency symbols before parsing
                            const cleanStr = ptpStr.replace(/,/g, '').replace(/[^\d.-]/g, '');
                            const ptpValue = parseFloat(cleanStr);
                            // Check if it's a valid number and not zero
                            if (!isNaN(ptpValue) && ptpValue > 0) {
                                ptpCount++;
                                ptpTotalAmount += ptpValue;
                                
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
                        // Convert to string first to handle different formats
                        const claimPaidStr = claimPaidAmount.toString().trim();
                        if (claimPaidStr !== '' && claimPaidStr !== '0' && claimPaidStr !== '0.00') {
                            // Remove commas and any currency symbols before parsing
                            const cleanStr = claimPaidStr.replace(/,/g, '').replace(/[^\d.-]/g, '');
                            const claimPaidValue = parseFloat(cleanStr);
                            // Check if it's a valid number and not zero
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
                    claimPaidCount: claimPaidCount,
                    claimPaidTotalAmount: claimPaidTotalAmount,
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
    const totalClaimPaidCount = validResults.reduce((sum, r) => sum + r.claimPaidCount, 0);
    const totalClaimPaidAmount = validResults.reduce((sum, r) => sum + r.claimPaidTotalAmount, 0);
    
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
            <button id="toggleMonthlyView" class="btn-toggle">
                <span class="toggle-icon">📅</span>
                <span class="toggle-text">View by Month</span>
            </button>
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
                    <div class="summary-stat highlight">
                        <div class="summary-stat-value">${uniquePredictiveAcrossFiles.length}</div>
                        <div class="summary-stat-label">Connected</div>
                    </div>
                    <div class="summary-stat debtor-highlight">
                        <div class="summary-stat-value">${uniqueDebtorAcrossFiles.length}</div>
                        <div class="summary-stat-label">Debtor</div>
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
                </div>
            </div>
        </div>
        <div id="monthlyView" class="summary-view" style="display: none;">
            <!-- Monthly breakdown will be generated here -->
        </div>
    `;
    
    // Set up toggle button event listener
    setTimeout(() => {
        const toggleBtn = document.getElementById('toggleMonthlyView');
        if (toggleBtn) {
            toggleBtn.addEventListener('click', () => toggleMonthlyView(validResults));
        }
    }, 0);
    
    // Populate Excel table
    resultsTableBody.innerHTML = results.map(result => {
        if (result.error) {
            return `
                <tr>
                    <td>${result.fileName}</td>
                    <td colspan="8" style="color: #e53e3e; text-align: center;">Error: ${result.error}</td>
                </tr>
            `;
        }
        
        return `
            <tr>
                <td>${result.fileName}</td>
                <td>${result.uniqueCount}</td>
                <td>${result.countBeforeDedup}</td>
                <td>${result.predictiveCount}</td>
                <td>${result.debtorCount}</td>
                <td>${result.ptpCount}</td>
                <td>${result.ptpTotalAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                <td>${result.claimPaidCount}</td>
                <td>${result.claimPaidTotalAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
            </tr>
        `;
    }).join('');
    
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
    // Sort by date
    results.sort((a, b) => new Date(a.fileDate) - new Date(b.fileDate));
    
    // Prepare data
    const labels = results.map(r => {
        const date = new Date(r.fileDate);
        return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' });
    });
    
    const ptpData = results.map(r => r.ptpTotalAmount);
    const claimPaidData = results.map(r => r.claimPaidTotalAmount);
    
    renderChart(labels, [
        {
            label: 'PTP Amount',
            data: ptpData,
            borderColor: '#8b5cf6',
            backgroundColor: 'rgba(139, 92, 246, 0.1)',
        },
        {
            label: 'Claim Paid Amount',
            data: claimPaidData,
            borderColor: '#ec4899',
            backgroundColor: 'rgba(236, 72, 153, 0.1)',
        }
    ]);
}

function renderComparisonGraph(results, month1Key, month2Key) {
    // Group by month and day
    const month1Data = {};
    const month2Data = {};
    
    results.forEach(result => {
        const date = new Date(result.fileDate);
        const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
        const day = date.getDate();
        
        if (monthKey === month1Key) {
            month1Data[day] = result;
        } else if (monthKey === month2Key) {
            month2Data[day] = result;
        }
    });
    
    // Get all unique days
    const allDays = [...new Set([...Object.keys(month1Data), ...Object.keys(month2Data)])].map(Number).sort((a, b) => a - b);
    
    // Prepare labels and data
    const labels = allDays.map(day => `Day ${day}`);
    
    const month1Name = new Date(month1Key + '-01').toLocaleDateString('en-US', { month: 'short', year: '2-digit' });
    const month2Name = new Date(month2Key + '-01').toLocaleDateString('en-US', { month: 'short', year: '2-digit' });
    
    const month1PtpData = allDays.map(day => month1Data[day]?.ptpTotalAmount || null);
    const month1ClaimData = allDays.map(day => month1Data[day]?.claimPaidTotalAmount || null);
    const month2PtpData = allDays.map(day => month2Data[day]?.ptpTotalAmount || null);
    const month2ClaimData = allDays.map(day => month2Data[day]?.claimPaidTotalAmount || null);
    
    renderChart(labels, [
        {
            label: `${month1Name} - PTP Amount`,
            data: month1PtpData,
            borderColor: '#8b5cf6',
            backgroundColor: 'rgba(139, 92, 246, 0.1)',
        },
        {
            label: `${month1Name} - Claim Paid Amount`,
            data: month1ClaimData,
            borderColor: '#ec4899',
            backgroundColor: 'rgba(236, 72, 153, 0.1)',
        },
        {
            label: `${month2Name} - PTP Amount`,
            data: month2PtpData,
            borderColor: '#6366f1',
            backgroundColor: 'rgba(99, 102, 241, 0.1)',
            borderDash: [5, 5],
        },
        {
            label: `${month2Name} - Claim Paid Amount`,
            data: month2ClaimData,
            borderColor: '#f97316',
            backgroundColor: 'rgba(249, 115, 22, 0.1)',
            borderDash: [5, 5],
        }
    ]);
}

function renderChart(labels, datasets) {
    // Destroy previous chart if exists
    if (chartInstance) {
        chartInstance.destroy();
    }
    
    // Create chart
    const ctx = document.getElementById('amountsChart').getContext('2d');
    chartInstance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
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
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        font: {
                            size: 14,
                            weight: '600'
                        },
                        padding: 20,
                        usePointStyle: true
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    padding: 12,
                    titleFont: {
                        size: 14,
                        weight: '600'
                    },
                    bodyFont: {
                        size: 13
                    },
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) {
                                label += ': ';
                            }
                            if (context.parsed.y !== null) {
                                label += context.parsed.y.toLocaleString('en-US', {
                                    minimumFractionDigits: 2,
                                    maximumFractionDigits: 2
                                });
                            } else {
                                label += 'No data';
                            }
                            return label;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        font: {
                            size: 12
                        },
                        callback: function(value) {
                            return value.toLocaleString('en-US', {
                                minimumFractionDigits: 0,
                                maximumFractionDigits: 0
                            });
                        }
                    },
                    grid: {
                        color: 'rgba(0, 0, 0, 0.05)'
                    }
                },
                x: {
                    ticks: {
                        font: {
                            size: 12
                        }
                    },
                    grid: {
                        display: false
                    }
                }
            }
        }
    });
}

function hideGraph() {
    graphModal.style.display = 'none';
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
                            <th>Total Collection</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${agentNames.map(name => {
                            const agent = allAgentsData[name];
                            const totalCollection = agent.ptpAmount + agent.claimPaidAmount;
                            return `
                                <tr>
                                    <td>${name}</td>
                                    <td>${agent.ptpCount.toLocaleString('en-US')}</td>
                                    <td>${agent.ptpAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                                    <td>${agent.claimPaidCount.toLocaleString('en-US')}</td>
                                    <td>${agent.claimPaidAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
                                    <td style="font-weight: 600;">${totalCollection.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>
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
                <div class="agent-stat-card total">
                    <div class="agent-stat-label">Total Collection</div>
                    <div class="agent-stat-value">${totalCollection.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</div>
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
    // Destroy agent chart if exists
    if (window.agentChartInstance) {
        window.agentChartInstance.destroy();
        window.agentChartInstance = null;
    }
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
    
    const labels = dates.map(date => {
        const d = new Date(date);
        return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
    });
    
    const ptpData = dates.map(date => agentDateData[date].ptpAmount);
    const claimPaidData = dates.map(date => agentDateData[date].claimPaidAmount);
    const totalData = dates.map(date => 
        agentDateData[date].ptpAmount + agentDateData[date].claimPaidAmount
    );
    
    renderAgentChart(labels, [
        {
            label: 'PTP Amount',
            data: ptpData,
            borderColor: '#722ed1',
            backgroundColor: 'rgba(114, 46, 209, 0.1)',
        },
        {
            label: 'Claim Paid Amount',
            data: claimPaidData,
            borderColor: '#eb2f96',
            backgroundColor: 'rgba(235, 47, 150, 0.1)',
        },
        {
            label: 'Total Collection',
            data: totalData,
            borderColor: '#1890ff',
            backgroundColor: 'rgba(24, 144, 255, 0.1)',
        }
    ], `${agentName} - Daily Performance`);
}

function renderAgentMonthlyComparisonGraph(agentName, agentDateData, month1Key, month2Key) {
    // Group data by month and day
    const month1Data = {};
    const month2Data = {};
    
    Object.keys(agentDateData).forEach(date => {
        const d = new Date(date);
        const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
        const day = d.getDate();
        
        if (monthKey === month1Key) {
            month1Data[day] = agentDateData[date];
        } else if (monthKey === month2Key) {
            month2Data[day] = agentDateData[date];
        }
    });
    
    // Get all unique days
    const allDays = [...new Set([...Object.keys(month1Data), ...Object.keys(month2Data)])].map(Number).sort((a, b) => a - b);
    const labels = allDays.map(day => `Day ${day}`);
    
    const month1Name = new Date(month1Key + '-01').toLocaleDateString('en-US', { month: 'short', year: '2-digit' });
    const month2Name = new Date(month2Key + '-01').toLocaleDateString('en-US', { month: 'short', year: '2-digit' });
    
    const month1PtpData = allDays.map(day => month1Data[day]?.ptpAmount || null);
    const month1ClaimData = allDays.map(day => month1Data[day]?.claimPaidAmount || null);
    const month1TotalData = allDays.map(day => {
        if (!month1Data[day]) return null;
        return month1Data[day].ptpAmount + month1Data[day].claimPaidAmount;
    });
    
    const month2PtpData = allDays.map(day => month2Data[day]?.ptpAmount || null);
    const month2ClaimData = allDays.map(day => month2Data[day]?.claimPaidAmount || null);
    const month2TotalData = allDays.map(day => {
        if (!month2Data[day]) return null;
        return month2Data[day].ptpAmount + month2Data[day].claimPaidAmount;
    });
    
    renderAgentChart(labels, [
        {
            label: `${month1Name} - PTP`,
            data: month1PtpData,
            borderColor: '#722ed1',
            backgroundColor: 'rgba(114, 46, 209, 0.1)',
        },
        {
            label: `${month1Name} - Claim Paid`,
            data: month1ClaimData,
            borderColor: '#eb2f96',
            backgroundColor: 'rgba(235, 47, 150, 0.1)',
        },
        {
            label: `${month1Name} - Total`,
            data: month1TotalData,
            borderColor: '#1890ff',
            backgroundColor: 'rgba(24, 144, 255, 0.1)',
        },
        {
            label: `${month2Name} - PTP`,
            data: month2PtpData,
            borderColor: '#52c41a',
            backgroundColor: 'rgba(82, 196, 26, 0.1)',
            borderDash: [5, 5],
        },
        {
            label: `${month2Name} - Claim Paid`,
            data: month2ClaimData,
            borderColor: '#fa8c16',
            backgroundColor: 'rgba(250, 140, 22, 0.1)',
            borderDash: [5, 5],
        },
        {
            label: `${month2Name} - Total`,
            data: month2TotalData,
            borderColor: '#13c2c2',
            backgroundColor: 'rgba(19, 194, 194, 0.1)',
            borderDash: [5, 5],
        }
    ], `${agentName} - Monthly Comparison`);
}

function renderAgentChart(labels, datasets, title) {
    // Destroy previous chart if exists
    if (window.agentChartInstance) {
        window.agentChartInstance.destroy();
    }
    
    const ctx = document.getElementById('agentChart').getContext('2d');
    window.agentChartInstance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: datasets.map(ds => ({
                label: ds.label,
                data: ds.data,
                borderColor: ds.borderColor,
                backgroundColor: ds.backgroundColor,
                borderWidth: 3,
                tension: 0.4,
                fill: true,
                pointRadius: 5,
                pointHoverRadius: 7,
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
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                title: {
                    display: true,
                    text: title,
                    font: {
                        size: 16,
                        weight: '600'
                    },
                    color: 'rgba(0, 0, 0, 0.85)',
                    padding: 20
                },
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        font: {
                            size: 12,
                            weight: '500'
                        },
                        padding: 15,
                        usePointStyle: true,
                        color: 'rgba(0, 0, 0, 0.85)'
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    padding: 12,
                    titleFont: {
                        size: 14,
                        weight: '600'
                    },
                    bodyFont: {
                        size: 13
                    },
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) {
                                label += ': ';
                            }
                            if (context.parsed.y !== null) {
                                label += context.parsed.y.toLocaleString('en-US', {
                                    minimumFractionDigits: 2,
                                    maximumFractionDigits: 2
                                });
                            } else {
                                label += 'No data';
                            }
                            return label;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        font: {
                            size: 11
                        },
                        color: 'rgba(0, 0, 0, 0.65)',
                        callback: function(value) {
                            return value.toLocaleString('en-US', {
                                minimumFractionDigits: 0,
                                maximumFractionDigits: 0
                            });
                        }
                    },
                    grid: {
                        color: 'rgba(0, 0, 0, 0.06)'
                    }
                },
                x: {
                    ticks: {
                        font: {
                            size: 11
                        },
                        color: 'rgba(0, 0, 0, 0.65)'
                    },
                    grid: {
                        display: false
                    }
                }
            }
        }
    });
}

// Monthly View Toggle
let isMonthlyView = false;

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
                    claimPaidCount: 0,
                    claimPaidAmount: 0
                };
            }
            
            monthlyData[monthKey].files.push(result);
            monthlyData[monthKey].accountNumbers.push(...result.accountNumbersArray);
            monthlyData[monthKey].predictiveNumbers.push(...result.predictiveNumbersArray);
            monthlyData[monthKey].debtorNumbers.push(...result.debtorNumbersArray);
            monthlyData[monthKey].totalDials += result.countBeforeDedup;
            monthlyData[monthKey].ptpCount += result.ptpCount;
            monthlyData[monthKey].ptpAmount += result.ptpTotalAmount;
            monthlyData[monthKey].claimPaidCount += result.claimPaidCount;
            monthlyData[monthKey].claimPaidAmount += result.claimPaidTotalAmount;
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
        
        return {
            monthName: data.monthName,
            totalDials: data.totalDials,
            workedOnTicket: uniqueAccounts.length,
            connected: uniquePredictive.length,
            debtor: uniqueDebtor.length,
            ptpCount: data.ptpCount,
            ptpAmount: data.ptpAmount,
            claimPaidCount: data.claimPaidCount,
            claimPaidAmount: data.claimPaidAmount,
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
        
        return `<span class="trend-indicator" style="color: ${color}">(${sign}${diff.toLocaleString('en-US')})</span>`;
    }
    
    function getTrendIndicatorPercent(current, previous) {
        if (previous === undefined) return '';
        const diff = current - previous;
        if (diff === 0) return '';
        
        const color = diff > 0 ? '#10b981' : '#ef4444';
        const sign = diff > 0 ? '+' : '';
        
        return `<span class="trend-indicator" style="color: ${color}">(${sign}${diff}%)</span>`;
    }
    
    function getTrendIndicatorAmount(current, previous) {
        if (previous === undefined) return '';
        const diff = current - previous;
        if (diff === 0) return '';
        
        const color = diff > 0 ? '#10b981' : '#ef4444';
        const sign = diff > 0 ? '+' : '';
        
        return `<span class="trend-indicator" style="color: ${color}">(${sign}${diff.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})})</span>`;
    }
    
    // Generate combined list HTML
    let monthlyHTML = `
        <div class="monthly-list-container">
            <div class="monthly-list-main">
                <div class="metric-group">
                    <h4 class="metric-title">Total Dials</h4>
                    <ul class="metric-list">
                        ${monthlyMetrics.map((month, index) => `
                            <li>
                                <span class="month-name">${month.monthName}</span>
                                <span class="month-value">${month.totalDials.toLocaleString('en-US')} ${getTrendIndicator(month.totalDials, monthlyMetrics[index + 1]?.totalDials)}</span>
                            </li>
                        `).join('')}
                    </ul>
                </div>
                
                <div class="metric-group">
                    <h4 class="metric-title">Worked on Ticket</h4>
                    <ul class="metric-list">
                        ${monthlyMetrics.map((month, index) => `
                            <li>
                                <span class="month-name">${month.monthName}</span>
                                <span class="month-value">${month.workedOnTicket.toLocaleString('en-US')} ${getTrendIndicator(month.workedOnTicket, monthlyMetrics[index + 1]?.workedOnTicket)}</span>
                            </li>
                        `).join('')}
                    </ul>
                </div>
                
                <div class="metric-group">
                    <h4 class="metric-title">Connected</h4>
                    <ul class="metric-list">
                        ${monthlyMetrics.map((month, index) => `
                            <li>
                                <span class="month-name">${month.monthName}</span>
                                <span class="month-value">${month.connected.toLocaleString('en-US')} ${getTrendIndicator(month.connected, monthlyMetrics[index + 1]?.connected)}</span>
                            </li>
                        `).join('')}
                    </ul>
                </div>
                
                <div class="metric-group">
                    <h4 class="metric-title">Debtor</h4>
                    <ul class="metric-list">
                        ${monthlyMetrics.map((month, index) => `
                            <li>
                                <span class="month-name">${month.monthName}</span>
                                <span class="month-value">${month.debtor.toLocaleString('en-US')} ${getTrendIndicator(month.debtor, monthlyMetrics[index + 1]?.debtor)}</span>
                            </li>
                        `).join('')}
                    </ul>
                </div>
                
                <div class="metric-group">
                    <h4 class="metric-title">PTP Count</h4>
                    <ul class="metric-list">
                        ${monthlyMetrics.map((month, index) => `
                            <li>
                                <span class="month-name">${month.monthName}</span>
                                <span class="month-value">${month.ptpCount.toLocaleString('en-US')} ${getTrendIndicator(month.ptpCount, monthlyMetrics[index + 1]?.ptpCount)}</span>
                            </li>
                        `).join('')}
                    </ul>
                </div>
                
                <div class="metric-group">
                    <h4 class="metric-title">PTP Amount</h4>
                    <ul class="metric-list">
                        ${monthlyMetrics.map((month, index) => `
                            <li>
                                <span class="month-name">${month.monthName}</span>
                                <span class="month-value">${month.ptpAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})} ${getTrendIndicatorAmount(month.ptpAmount, monthlyMetrics[index + 1]?.ptpAmount)}</span>
                            </li>
                        `).join('')}
                    </ul>
                </div>
                
                <div class="metric-group">
                    <h4 class="metric-title">Claim Paid Count</h4>
                    <ul class="metric-list">
                        ${monthlyMetrics.map((month, index) => `
                            <li>
                                <span class="month-name">${month.monthName}</span>
                                <span class="month-value">${month.claimPaidCount.toLocaleString('en-US')} ${getTrendIndicator(month.claimPaidCount, monthlyMetrics[index + 1]?.claimPaidCount)}</span>
                            </li>
                        `).join('')}
                    </ul>
                </div>
                
                <div class="metric-group">
                    <h4 class="metric-title">Claim Paid Amount</h4>
                    <ul class="metric-list">
                        ${monthlyMetrics.map((month, index) => `
                            <li>
                                <span class="month-name">${month.monthName}</span>
                                <span class="month-value">${month.claimPaidAmount.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})} ${getTrendIndicatorAmount(month.claimPaidAmount, monthlyMetrics[index + 1]?.claimPaidAmount)}</span>
                            </li>
                        `).join('')}
                    </ul>
                </div>
            </div>
            
            <div class="monthly-list-sidebar">
                <div class="metric-group">
                    <h4 class="metric-title">Penetration</h4>
                    <ul class="metric-list">
                        ${monthlyMetrics.map((month, index) => `
                            <li>
                                <span class="month-name">${month.monthName}</span>
                                <span class="month-value">${month.penetration}% ${getTrendIndicatorPercent(month.penetration, monthlyMetrics[index + 1]?.penetration)}</span>
                            </li>
                        `).join('')}
                    </ul>
                </div>
                
                <div class="metric-group">
                    <h4 class="metric-title">Connected Rate</h4>
                    <ul class="metric-list">
                        ${monthlyMetrics.map((month, index) => `
                            <li>
                                <span class="month-name">${month.monthName}</span>
                                <span class="month-value">${month.connectedRate}% ${getTrendIndicatorPercent(month.connectedRate, monthlyMetrics[index + 1]?.connectedRate)}</span>
                            </li>
                        `).join('')}
                    </ul>
                </div>
            </div>
        </div>
    `;
    
    monthlyView.innerHTML = monthlyHTML;
}