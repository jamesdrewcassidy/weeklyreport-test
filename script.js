/* ======================
   Global Variables & Local Storage Keys
   ====================== */
const LS_COMPLETED_KEY  = 'weekly_report_completed_v9'; // Version bump for new data structure
const LS_ONGOING_KEY    = 'weekly_report_ongoing_v9';
const LS_NOTES_KEY      = 'weekly_report_notes_v9';
const LS_DATE_RANGE_KEY = 'weekly_report_dateRange_v9';
const HOURLY_RATE = 25; // The rate for calculating actual price
let confirmCallback = null;

/* ======================
   Date & Calculation Helper Functions
   ====================== */
function parseLocalDate(dateString) {
  if (!dateString) return null;
  const [year, month, day] = dateString.split('-').map(Number);
  return new Date(Date.UTC(year, month - 1, day));
}
function formatDate(dateObj) {
  if (!dateObj || isNaN(dateObj)) return '';
  const year = dateObj.getUTCFullYear();
  const month = ('0' + (dateObj.getUTCMonth() + 1)).slice(-2);
  const day = ('0' + dateObj.getUTCDate()).slice(-2);
  return `${month}/${day}/${year}`;
}
function formatDateStringForInput(dateString) {
    if (!dateString) return '';
    const d = parseLocalDate(dateString);
    if (!d) return '';
    const year = d.getUTCFullYear();
    const month = ('0' + (d.getUTCMonth() + 1)).slice(-2);
    const day = ('0' + d.getUTCDate()).slice(-2);
    return `${year}-${month}-${day}`;
}

/* ======================
   Data Retrieval & Storage
   ====================== */
function getStoredData(key) {
  return JSON.parse(localStorage.getItem(key) || '[]');
}
function setStoredData(key, data) {
  localStorage.setItem(key, JSON.stringify(data));
}

/* ======================
   Modal & Confirmation Dialog
   ====================== */
function showModal(id) {
  document.getElementById(id).style.display = 'flex';
}
function closeModal(id) {
  document.getElementById(id).style.display = 'none';
}
function showConfirmation(title, text, onConfirm) {
  document.getElementById('confirmModalTitle').textContent = title;
  document.getElementById('confirmModalText').textContent = text;
  confirmCallback = onConfirm;
  showModal('confirmModal');
}

/* ======================
   Rendering Logic
   ====================== */
function renderAllTables() {
  renderCompletedTable();
  renderOngoingTable();
}

function renderCompletedTable() {
  const data = getStoredData(LS_COMPLETED_KEY);
  const tbody = document.getElementById('completedTableBody');
  tbody.innerHTML = '';
  if (data.length === 0) {
      tbody.innerHTML = '<tr><td colspan="5" class="empty-cell">No completed inspections recorded yet.</td></tr>';
  } else {
    data.sort((a,b) => (parseLocalDate(b.dateCompleted) || 0) - (parseLocalDate(a.dateCompleted) || 0)); // Sort by most recent
    data.forEach(item => {
      const tr = document.createElement('tr');
      
      const gainLoss = (parseFloat(item.bidPrice) || 0) - ((parseFloat(item.actualHours) || 0) * HOURLY_RATE);
      const diff = (parseFloat(item.actualHours) || 0) - (parseFloat(item.bidHours) || 0);
      const diffPercent = (item.bidHours && item.bidHours != 0) ? (diff / item.bidHours) * 100 : 0;

      let gainLossClass = diff === 0 ? '' : (gainLoss > 0 ? 'text-success' : 'text-danger');

      let historyHTML = (item.hoursHistory && item.hoursHistory.length > 0) 
        ? `<h5>Historical Data</h5>
           <table class="history-table">
             <thead><tr><th>Yr</th><th>Bid</th><th>Actual</th><th>% Diff</th><th>Bid $</th><th>Actual $</th><th>Gain/Loss</th></tr></thead>
             <tbody>` +
           item.hoursHistory.map(h => {
                const hBid = parseFloat(h.bid) || 0;
                const hActual = parseFloat(h.actual) || 0;
                const hBidPrice = parseFloat(h.bidPrice) || 0;
                const hActualPrice = hActual * HOURLY_RATE;
                const hGainLoss = hBidPrice - hActualPrice;
                const hDiffPercent = hBid !== 0 ? ((hActual - hBid) / hBid) * 100 : 0;
                let hGainLossClass = hGainLoss >= 0 ? 'text-success' : 'text-danger';
                return `<tr>
                          <td>${h.year}</td>
                          <td>${hBid.toFixed(1)}</td>
                          <td>${hActual.toFixed(1)}</td>
                          <td class="${hDiffPercent > 0 ? 'text-danger' : 'text-success'}">${hDiffPercent.toFixed(0)}%</td>
                          <td>$${hBidPrice.toFixed(2)}</td>
                          <td>$${hActualPrice.toFixed(2)}</td>
                          <td class="${hGainLossClass}">$${hGainLoss.toFixed(2)}</td>
                        </tr>`;
           }).join('') + '</tbody></table>'
        : '<p>No historical data entered.</p>';

      tr.innerHTML = `
        <td>${item.siteName}</td>
        <td>${item.projectNumber}</td>
        <td>${formatDate(parseLocalDate(item.dateCompleted))}</td>
        <td class="completed-info-cell">
            <p><strong>This Year:</strong> Bid ${parseFloat(item.bidHours || 0).toFixed(1)}h, Actual ${parseFloat(item.actualHours || 0).toFixed(1)}h
              <span class="(${diffPercent > 0 ? 'text-danger' : 'text-success'})"> (${diffPercent.toFixed(0)}%)</span>
            </p>
             <p><strong>Financials:</strong> Bid $${(parseFloat(item.bidPrice) || 0).toFixed(2)}, Actual $${((parseFloat(item.actualHours) || 0) * HOURLY_RATE).toFixed(2)}
              <strong class="${gainLossClass}"> (Gain/Loss: $${gainLoss.toFixed(2)})</strong>
            </p>
            <p><strong>Discrepancies:</strong> ${item.discrepancies || 'None'}</p>
            <p><strong>Deficiencies:</strong> ${item.deficiencies || 'None'}</p>
            <p><strong>Notes for Marty:</strong> ${item.notes || 'None'}</p>
            <p><strong>Report Sent:</strong> ${item.reportSent ? 'Yes' : 'No'}</p>
            <hr>
            ${historyHTML}
        </td>
        <td class="no-pdf action-cell">
          <button class="btn-secondary" onclick="editCompletedItem('${item.id}')">Edit</button>
          <button class="btn-danger" onclick="deleteCompletedItem('${item.id}')">Delete</button>
        </td>
      `;
      tbody.appendChild(tr);
    });
  }
}

function renderOngoingTable() {
  const data = getStoredData(LS_ONGOING_KEY);
  const tbody = document.getElementById('ongoingTableBody');
  tbody.innerHTML = '';
  if (data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" class="empty-cell">No ongoing inspections.</td></tr>';
  } else {
    data.sort((a,b) => (parseLocalDate(a.estCompletion) || 0) - (parseLocalDate(b.estCompletion) || 0)); // Sort by soonest completion
    data.forEach(item => {
      const tr = document.createElement('tr');
      const hoursDiff = (parseFloat(item.hoursWorked) || 0) - (parseFloat(item.bidHours) || 0);
      let diffClass = '';
      if (hoursDiff > 0) diffClass = 'text-danger';

      tr.innerHTML = `
        <td>
          ${item.siteName}
          ${item.notes ? `<div class="notes-display"><strong>Notes:</strong> ${item.notes}</div>` : ''}
        </td>
        <td>${item.projectNumber}</td>
        <td>${parseFloat(item.bidHours || 0).toFixed(1)}</td>
        <td class="${diffClass}">${parseFloat(item.hoursWorked || 0).toFixed(1)}</td>
        <td>${formatDate(parseLocalDate(item.estCompletion))}</td>
        <td class="no-pdf action-cell">
          <button class="btn-primary" onclick="markAsComplete('${item.id}')">Complete</button>
          <button class="btn-secondary" onclick="editOngoingItem('${item.id}')">Edit</button>
          <button class="btn-danger" onclick="deleteOngoingItem('${item.id}')">Delete</button>
        </td>
      `;
      tbody.appendChild(tr);
    });
  }
}


/* ======================
   COMPLETED Form & Data Handling
   ====================== */
document.getElementById('completedForm').addEventListener('submit', function(e) {
  e.preventDefault();
  const data = getStoredData(LS_COMPLETED_KEY);
  const id = document.getElementById('completedEntryId').value;

  const hoursHistory = [];
  document.querySelectorAll('.history-year-entry').forEach(div => {
      const year = div.querySelector('input[name="historyYear"]').value;
      const bid = div.querySelector('input[name="historyBid"]').value;
      const actual = div.querySelector('input[name="historyActual"]').value;
      const bidPrice = div.querySelector('input[name="historyBidPrice"]').value;
      if (year && bid && actual && bidPrice) {
          hoursHistory.push({ year, bid, actual, bidPrice });
      }
  });

  const entry = {
    siteName: document.getElementById('completedSiteName').value,
    projectNumber: document.getElementById('completedProjectNumber').value,
    dateCompleted: document.getElementById('completedDate').value,
    bidHours: document.getElementById('completedBidHours').value,
    actualHours: document.getElementById('completedActualHours').value,
    bidPrice: document.getElementById('completedBidPrice').value,
    discrepancies: document.getElementById('completedDiscrepancies').value,
    deficiencies: document.getElementById('completedDeficiencies').value,
    notes: document.getElementById('completedNotes').value,
    reportSent: document.getElementById('completedReportSent').checked,
    hoursHistory,
  };

  if (id) {
    const index = data.findIndex(item => item.id === id);
    data[index] = { ...data[index], ...entry };
  } else {
    entry.id = 'comp_' + Date.now();
    data.push(entry);
  }

  setStoredData(LS_COMPLETED_KEY, data);
  this.reset();
  document.getElementById('completedEntryId').value = '';
  document.getElementById('completedFormTitle').textContent = 'Add New Completed Inspection';
  document.getElementById('hoursHistoryContainer').innerHTML = '';
  renderCompletedTable();
});

function editCompletedItem(id) {
    const data = getStoredData(LS_COMPLETED_KEY);
    const item = data.find(i => i.id === id);
    if (!item) return;

    document.getElementById('completedFormTitle').textContent = 'Edit Completed Inspection';
    document.getElementById('completedEntryId').value = item.id;
    document.getElementById('completedSiteName').value = item.siteName;
    document.getElementById('completedProjectNumber').value = item.projectNumber;
    document.getElementById('completedDate').value = formatDateStringForInput(item.dateCompleted);
    document.getElementById('completedBidHours').value = item.bidHours;
    document.getElementById('completedActualHours').value = item.actualHours;
    document.getElementById('completedBidPrice').value = item.bidPrice;
    document.getElementById('completedDiscrepancies').value = item.discrepancies || '';
    document.getElementById('completedDeficiencies').value = item.deficiencies || '';
    document.getElementById('completedNotes').value = item.notes || '';
    document.getElementById('completedReportSent').checked = item.reportSent || false;

    const historyContainer = document.getElementById('hoursHistoryContainer');
    historyContainer.innerHTML = '';
    (item.hoursHistory || []).forEach(h => addHistoryYear(h));

    document.getElementById('completedSiteName').focus();
}

function deleteCompletedItem(id) {
    showConfirmation('Delete Completed Entry?', 'Are you sure you want to permanently delete this completed inspection record?', () => {
        let data = getStoredData(LS_COMPLETED_KEY);
        data = data.filter(item => item.id !== id);
        setStoredData(LS_COMPLETED_KEY, data);
        renderCompletedTable();
    });
}

function addHistoryYear(data = {}) {
    const container = document.getElementById('hoursHistoryContainer');
    const div = document.createElement('div');
    div.className = 'history-year-entry';
    div.innerHTML = `
        <input type="number" name="historyYear" placeholder="Year" value="${data.year || ''}" required>
        <input type="number" step="0.1" name="historyBid" placeholder="Bid Hrs" value="${data.bid || ''}" required>
        <input type="number" step="0.1" name="historyActual" placeholder="Actual Hrs" value="${data.actual || ''}" required>
        <input type="number" step="0.01" name="historyBidPrice" placeholder="Bid Price ($)" value="${data.bidPrice || ''}" required>
        <button type="button" class="btn-danger" onclick="this.parentElement.remove()">X</button>
    `;
    container.appendChild(div);
}


/* ======================
   ONGOING Form & Data Handling
   ====================== */
document.getElementById('ongoingForm').addEventListener('submit', function(e) {
  e.preventDefault();
  const data = getStoredData(LS_ONGOING_KEY);
  const id = document.getElementById('ongoingEntryId').value;
  const entry = {
    siteName: document.getElementById('ongoingSiteName').value,
    projectNumber: document.getElementById('ongoingProjectNumber').value,
    bidHours: document.getElementById('ongoingBidHours').value,
    hoursWorked: document.getElementById('ongoingHoursWorked').value,
    estCompletion: document.getElementById('ongoingEstCompletion').value,
    notes: document.getElementById('ongoingNotes').value,
  };

  if (id) {
    const index = data.findIndex(item => item.id === id);
    data[index] = { ...data[index], ...entry };
  } else {
    entry.id = 'ongo_' + Date.now();
    data.push(entry);
  }

  setStoredData(LS_ONGOING_KEY, data);
  this.reset();
  document.getElementById('ongoingEntryId').value = '';
  document.getElementById('ongoingFormTitle').textContent = 'Add New Ongoing Inspection';
  renderOngoingTable();
});

function editOngoingItem(id) {
    const data = getStoredData(LS_ONGOING_KEY);
    const item = data.find(i => i.id === id);
    if (!item) return;

    document.getElementById('ongoingFormTitle').textContent = 'Edit Ongoing Inspection';
    document.getElementById('ongoingEntryId').value = item.id;
    document.getElementById('ongoingSiteName').value = item.siteName;
    document.getElementById('ongoingProjectNumber').value = item.projectNumber;
    document.getElementById('ongoingBidHours').value = item.bidHours;
    document.getElementById('ongoingHoursWorked').value = item.hoursWorked;
    document.getElementById('ongoingEstCompletion').value = formatDateStringForInput(item.estCompletion);
    document.getElementById('ongoingNotes').value = item.notes || '';
    document.getElementById('ongoingSiteName').focus();
}

function deleteOngoingItem(id) {
    showConfirmation('Delete Ongoing Entry?', 'Are you sure you want to permanently delete this ongoing inspection record?', () => {
        let data = getStoredData(LS_ONGOING_KEY);
        data = data.filter(item => item.id !== id);
        setStoredData(LS_ONGOING_KEY, data);
        renderOngoingTable();
    });
}

function markAsComplete(id) {
    const ongoingData = getStoredData(LS_ONGOING_KEY);
    const item = ongoingData.find(i => i.id === id);
    if (!item) return;
    
    // Pre-fill the "Completed" form
    document.getElementById('completedFormTitle').textContent = 'Finalize Completed Inspection';
    document.getElementById('completedSiteName').value = item.siteName;
    document.getElementById('completedProjectNumber').value = item.projectNumber;
    document.getElementById('completedBidHours').value = item.bidHours;
    document.getElementById('completedActualHours').value = item.hoursWorked; // Pre-fill actual with hours worked
    document.getElementById('completedNotes').value = item.notes || '';
    
    // Set today's date
    document.getElementById('completedDate').value = new Date().toISOString().split('T')[0];

    // Clear history and ID from previous edits
    document.getElementById('hoursHistoryContainer').innerHTML = '';
    document.getElementById('completedEntryId').value = '';


    // Remove from ongoing list after confirmation
    showConfirmation('Mark as Complete?', `This will move "${item.siteName}" to the completed list. Please fill in the final details and save.`, () => {
        let data = ongoingData.filter(i => i.id !== id);
        setStoredData(LS_ONGOING_KEY, data);
        renderOngoingTable();
        document.getElementById('completedBidPrice').focus();
    });
}


/* ======================
   General Report Actions (Date Range, Notes, PDF)
   ====================== */
function updateDateRange() {
  const start = document.getElementById('startDate').value;
  const end = document.getElementById('endDate').value;
  localStorage.setItem(LS_DATE_RANGE_KEY, JSON.stringify({ start, end }));
  updateDateRangeDisplay();
}

function updateDateRangeDisplay() {
  const savedRange = JSON.parse(localStorage.getItem(LS_DATE_RANGE_KEY) || '{}');
  const display = document.getElementById('reportDateRange');
  if (savedRange.start && savedRange.end) {
    const startDate = formatDate(parseLocalDate(savedRange.start));
    const endDate = formatDate(parseLocalDate(savedRange.end));
    display.textContent = `For the week of ${startDate} to ${endDate}`;
  } else {
    display.textContent = 'Date range not set.';
  }
}

function saveNotes() {
    const notes = {
        trendsNoticed: document.getElementById('trendsNoticed').value,
        resourceChallenges: document.getElementById('resourceChallenges').value,
        suggestedImprovements: document.getElementById('suggestedImprovements').value,
    };
    localStorage.setItem(LS_NOTES_KEY, JSON.stringify(notes));
}

function downloadPDF() {
    const element = document.getElementById('pdfContainer');
    const savedRange = JSON.parse(localStorage.getItem(LS_DATE_RANGE_KEY) || '{}');
    const startDate = savedRange.start || 'YYYY-MM-DD';
    const endDate = savedRange.end || 'YYYY-MM-DD';
    const fileName = `Omni_IT_Report_${startDate}_to_${endDate}.pdf`;

    const opt = {
        margin:       [0.5, 0.5, 0.5, 0.5], // top, left, bottom, right in inches
        filename:     fileName,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true, letterRendering: true },
        jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' }
    };
    
    // Temporarily add a class to the body to apply PDF-specific styles
    document.body.classList.add('pdf-render-mode');
    
    html2pdf().from(element).set(opt).save().then(() => {
        // Remove the class after the PDF has been generated
        document.body.classList.remove('pdf-render-mode');
    });
}

function clearAllData() {
    showConfirmation('Clear All Report Data?', 'WARNING: This will permanently delete all completed inspections, ongoing inspections, and notes. This cannot be undone.', () => {
        localStorage.removeItem(LS_COMPLETED_KEY);
        localStorage.removeItem(LS_ONGOING_KEY);
        localStorage.removeItem(LS_NOTES_KEY);
        renderAllTables();
    });
}


/* ======================
   Initialization
   ====================== */
document.addEventListener('DOMContentLoaded', () => {
    // Set default date range to the current Mon-Sat week
    const today = new Date();
    const firstDayOfWeek = new Date(today);
    firstDayOfWeek.setDate(today.getDate() - today.getDay() + (today.getDay() === 0 ? -6 : 1)); // Monday
    const lastDayOfWeek = new Date(today);
    lastDayOfWeek.setDate(today.getDate() + (6 - today.getDay())); // Saturday

    const formatForInput = d => d.toISOString().split('T')[0];
    document.getElementById('startDate').value = formatForInput(firstDayOfWeek);
    document.getElementById('endDate').value = formatForInput(lastDayOfWeek);

    // Load saved date range if it exists
    const savedRange = JSON.parse(localStorage.getItem(LS_DATE_RANGE_KEY) || '{}');
    if (savedRange.start) document.getElementById('startDate').value = savedRange.start;
    if (savedRange.end) document.getElementById('endDate').value = savedRange.end;
    updateDateRangeDisplay();

    // Load saved notes if they exist
    const savedNotes = JSON.parse(localStorage.getItem(LS_NOTES_KEY) || '{}');
    if (savedNotes.trendsNoticed) document.getElementById('trendsNoticed').value = savedNotes.trendsNoticed;
    if (savedNotes.resourceChallenges) document.getElementById('resourceChallenges').value = savedNotes.resourceChallenges;
    if (savedNotes.suggestedImprovements) document.getElementById('suggestedImprovements').value = savedNotes.suggestedImprovements;

    // Set up confirmation modal buttons
    document.getElementById('confirmCancelBtn').addEventListener('click', () => closeModal('confirmModal'));
    document.getElementById('confirmOkBtn').addEventListener('click', () => {
        if (confirmCallback) confirmCallback();
        closeModal('confirmModal');
    });

    // Render initial data
    renderAllTables();
});

/* ======================
   Excel Generation (SheetJS)
   ====================== */
/**
 * Main function to trigger the download of all completed inspections as separate Excel files.
 */
function downloadAllExcelFiles() {
    const completedData = JSON.parse(localStorage.getItem(LS_COMPLETED_KEY) || '[]');

    if (completedData.length === 0) {
        // Using the custom modal instead of alert
        showConfirmation('No Data', 'There are no completed inspections to download.', () => {});
        // Need to adjust the button text for this case
        document.getElementById('confirmOkBtn').textContent = 'OK';
        document.getElementById('confirmCancelBtn').style.display = 'none';

        // Add a one-time event listener to reset button styles when the modal closes
        const modal = document.getElementById('confirmModal');
        const observer = new MutationObserver((mutationsList, observer) => {
            for(const mutation of mutationsList) {
                if (mutation.type === 'attributes' && mutation.attributeName === 'style') {
                    if (modal.style.display === 'none') {
                        document.getElementById('confirmOkBtn').textContent = 'OK';
                        document.getElementById('confirmCancelBtn').style.display = 'inline-block';
                        observer.disconnect(); // Clean up the observer
                    }
                }
            }
        });
        observer.observe(modal, { attributes: true });
        
        return;
    }

    // Loop through each completed item and generate a separate Excel file for it.
    completedData.forEach(item => {
        generateAndDownloadExcel(item);
    });
}

/**
 * Generates and downloads a single .xlsx file for a given inspection item.
 * @param {object} item - The completed inspection data object.
 */
function generateAndDownloadExcel(item) {
    // 1. Prepare data for the "History" sheet 
    const historyHeader = ['Year', 'Bid Hours', 'Actual Hours', '% Diff', 'Bid Price ($)', 'Actual Price ($)', 'Gain/Loss ($)'];
    const historyRows = (item.hoursHistory || []).map(h => {
        const bid = parseFloat(h.bid) || 0;
        const actual = parseFloat(h.actual) || 0;
        const bidPrice = parseFloat(h.bidPrice) || 0;
        const actualPrice = actual * HOURLY_RATE;
        const gainLoss = bidPrice - actualPrice;
        const percentDiff = bid !== 0 ? (actual - bid) / bid : 0;
        
        return [
            h.year,
            bid,
            actual,
            percentDiff, // Format as percentage in Excel
            bidPrice,
            actualPrice,
            gainLoss
        ];
    });

    const historySheet = XLSX.utils.aoa_to_sheet([historyHeader, ...historyRows]);
    // Apply formatting to specific columns
    historySheet['!cols'] = [ { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 15 }, { wch: 15 }, { wch: 15 } ];
    historyRows.forEach((_, index) => {
        const i = index + 2; // +2 because sheetJS is 1-based and we have a header
        if(historySheet[`D${i}`]) historySheet[`D${i}`].z = '0.0%'; // Percentage
        if(historySheet[`E${i}`]) historySheet[`E${i}`].z = '"$"#,##0.00'; // Currency
        if(historySheet[`F${i}`]) historySheet[`F${i}`].z = '"$"#,##0.00'; // Currency
        if(historySheet[`G${i}`]) historySheet[`G${i}`].z = '"$"#,##0.00'; // Currency
    });


    // 2. Prepare data for the "Summary" sheet 
    let totalGainLoss = 0, avgBid = 0, avgActual = 0, avgDiff = 0, mostRecentYear = 0;
    if (historyRows.length > 0) {
        totalGainLoss = historyRows.reduce((sum, row) => sum + row[6], 0);
        avgBid = historyRows.reduce((sum, row) => sum + row[1], 0) / historyRows.length;
        avgActual = historyRows.reduce((sum, row) => sum + row[2], 0) / historyRows.length;
        avgDiff = historyRows.reduce((sum, row) => sum + row[3], 0) / historyRows.length;
        mostRecentYear = Math.max(...historyRows.map(row => row[0]));
    }
    
    const summaryData = [
        ['Company/Site:', item.siteName || 'N/A'],
        [],
        ['KPIs (auto-calculated from History)'],
        [],
        ['Average Bid Hours', avgBid.toFixed(2)],
        ['Average Actual Hours', avgActual.toFixed(2)],
        ['Average % Difference', avgDiff],
        ['Total Gain/Loss ($)', totalGainLoss],
        [],
        ['Most Recent Year', mostRecentYear || 'N/A']
    ];
    const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
    // Apply formatting
    if(summarySheet['B7']) summarySheet['B7'].z = '0.0%';
    if(summarySheet['B8']) summarySheet['B8'].z = '"$"#,##0.00';
    summarySheet['!cols'] = [ { wch: 35 }, { wch: 20 } ];


    // 3. Prepare data for the "Notes" sheet 
    const notesData = [
        ['Inspector Notes (free text)'],
        ['Use this sheet to capture context from each inspection (mirrors the PDF fields):'],
        [],
        ['- Date Completed:', formatDate(parseLocalDate(item.dateCompleted))],
        ['- Deficiencies Found:', item.deficiencies],
        ['- Report Sent (Yes/No):', item.reportSent ? 'Yes' : 'No'],
        ['- Field Notes (access issues, special equipment, returns needed, etc.):'],
        [item.discrepancies || ''],
        [],
        ['- Contract/Quote Notes for Marty:'],
        [item.notes || '']
    ];
    const notesSheet = XLSX.utils.aoa_to_sheet(notesData);
    notesSheet['!cols'] = [{ wch: 80 }];


    // 4. Create the Workbook and trigger the download
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, historySheet, 'History');
    XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');
    XLSX.utils.book_append_sheet(workbook, notesSheet, 'Notes');

    // Sanitize the site name for use in a filename
    const safeFileName = (item.siteName || 'Untitled').replace(/[^a-z0-9]/gi, '_').toLowerCase();
    const fileName = `${safeFileName}_${item.projectNumber}.xlsx`;

    XLSX.writeFile(workbook, fileName);
}
