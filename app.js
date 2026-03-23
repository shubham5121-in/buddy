import { db, auth, googleProvider } from './firebase-config.js';
import {
    collection,
    addDoc,
    setDoc,
    getDocs,
    updateDoc,
    deleteDoc,
    doc,
    query,
    orderBy,
    where,
    onSnapshot,
    setLogLevel
} from "https://www.gstatic.com/firebasejs/11.0.1/firebase-firestore.js";
import {
    signInWithEmailAndPassword,
    signInWithPopup,
    signOut,
    onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/11.0.1/firebase-auth.js";

// App State & Data Management
const COLLECTION_NAME = 'Loans';

// Global state
let loans = [];
let editingId = null;
let currentUserRole = null;
let assignedName = null;
let authorizedUsers = [];
let customColumns = []; // Custom column names (strings)
let columnOrder = [];   // Master list of all visible column keys in order
let executiveAdjustments = []; // Debit/Credit adjustments for executives
let unsubscribeAdjustments = null; // Firestore listener for adjustments
// Stores last-set payout % per "ExecutiveName|||BankName" — persists across modal re-renders
const bankPayoutConfig = {};

const SYSTEM_COLUMNS = [
    { id: 'date', label: 'Entry Date' },
    { id: 'disbursementDate', label: 'Disb. Date' },
    { id: 'losNo', label: 'LOS No.' },
    { id: 'customerName', label: 'Customer' },
    { id: 'caseType', label: 'Type' },
    { id: 'bankName', label: 'Bank' },
    { id: 'amount', label: 'Amount' },
    { id: 'tenure', label: 'Tenure' },
    { id: 'location', label: 'Loc' },
    { id: 'status', label: 'Status' },
    { id: 'executiveName', label: 'Executive' },
    { id: 'remarks', label: 'Remarks' }
];


// DOM Elements
const contentArea = document.getElementById('content-area');
const navItems = document.querySelectorAll('.nav-item');
const pageTitle = document.getElementById('page-title');
const loginForm = document.getElementById('login-form');
const googleLoginBtn = document.getElementById('google-login-btn');
const logoutBtn = document.getElementById('logout-btn');
const authError = document.getElementById('auth-error');
const authOverlay = document.getElementById('auth-overlay');

window.addEventListener('online', () => showToast('Back Online'));
window.addEventListener('offline', () => showToast('Working Offline'));

// Global Error Catcher
window.onerror = function (msg, url, line, col, error) {
    console.error("GLOBAL ERROR:", msg, "at", url, ":", line);
    // Only alert for non-extension errors to avoid noise
    if (!url || url.includes('app.js')) {
        alert("CRITICAL APP ERROR:\n" + msg + "\nLine: " + line);
    }
    return false;
};

// Utility Functions
const sanitizeHTML = (str) => {
    if (str === null || str === undefined) return '';
    const temp = document.createElement('div');
    temp.textContent = str.toString();
    return temp.innerHTML;
};

const formatCurrency = (amount) => {
    return new Intl.NumberFormat('en-IN', {
        style: 'currency',
        currency: 'INR',
        maximumFractionDigits: 0
    }).format(amount);
};

const formatDate = (dateString) => {
    if (!dateString) return '';
    return new Date(dateString).toLocaleDateString('en-IN');
};

const getStatusClass = (status) => {
    switch (status?.toLowerCase()) {
        case 'disbursed': return 'status-disbursed';
        case 'approved': return 'status-approved';
        case 'rejected': return 'status-rejected';
        case 'underwriting': return 'status-underwriting';
        case 'underwriting forward': return 'status-forward';
        default: return 'status-default';
    }
};

const saveToLocalStorage = () => {
    // Keeping this for legacy/backup purposes, but main data is in Firestore
    localStorage.setItem('SBE_Loans_Backup', JSON.stringify(loans));
};

// --- DATA SEEDING FOR TESTING ---
window.seedTestData = async () => {
    if (!confirm("Add 100 test entries to the system? This will sync to Firebase.")) return;

    const banks = ['HDFC Bank', 'ICICI Bank', 'Axis Bank', 'Kotak Mahindra', 'Bajaj Finserv'];
    const statuses = ['Underwriting', 'Underwriting Forward', 'Approved', 'Disbursed', 'Rejected'];
    const execs = ['Test Sam', 'Demo Bob', 'Mock Charlie'];
    const types = ['Normal PL', 'Golden Edge', 'BT', 'Business Loan'];

    showToast("Seeding 100 entries... Please wait.");

    try {
        for (let i = 1; i <= 100; i++) {
            const entry = {
                date: new Date(Date.now() - Math.random() * 60 * 24 * 60 * 60 * 1000).toISOString().split('T')[0],
                customerName: `Test User #${i}`,
                losNo: `TEST-${1000 + i}`,
                bankName: banks[Math.floor(Math.random() * banks.length)],
                amount: Math.floor(Math.random() * 800000) + 100000,
                interestRate: (Math.random() * 5 + 10).toFixed(2),
                tenure: (Math.floor(Math.random() * 5) + 1) * 12,
                caseType: types[Math.floor(Math.random() * types.length)],
                location: 'Mumbai',
                status: statuses[Math.floor(Math.random() * statuses.length)],
                executiveName: execs[Math.floor(Math.random() * execs.length)],
                remarks: 'Auto-seeded for calculation testing.',
                customData: {}
            };
            await addDoc(collection(db, COLLECTION_NAME), entry);
        }
        showToast("✅ Successfully added 100 test entries!");
    } catch (err) {
        console.error("Seed error:", err);
        alert("Error seeding data: " + err.message);
    }
};

// Navigation Logic
const updateNavActive = (view) => {
    document.querySelectorAll('.nav-item').forEach(nav => {
        if (nav.getAttribute('data-view') === view) {
            nav.classList.add('active');
        } else {
            nav.classList.remove('active');
        }

        // Dynamically update sidebar text
        if (nav.getAttribute('data-view') === 'entry') {
            const span = nav.querySelector('span');
            if (span) {
                span.textContent = (currentUserRole === 'ADMIN' || currentUserRole === 'BACK_OFFICE') ? 'Daily Entry' : 'YOUR FILES';
            }
        }
    });
};

navItems.forEach(item => {
    item.addEventListener('click', () => {
        const view = item.getAttribute('data-view');
        loadView(view);
    });
});

const loadView = (view) => {
    // Security check for views
    const isAdmin = currentUserRole === 'ADMIN';
    if ((view === 'dashboard' || view === 'exec-files' || view === 'users' || view === 'settings') && !isAdmin) {
        loadView('entry');
        return;
    }
    if (view === 'users' && currentUserRole !== 'ADMIN') {
        loadView('entry');
        return;
    }
    if (view === 'settings' && currentUserRole !== 'ADMIN') {
        loadView('entry');
        return;
    }

    updateNavActive(view);


    contentArea.innerHTML = '';
    editingId = null; // Reset edit mode on view change
    if (view === 'entry') {
        renderEntryPage();
    } else if (view === 'exec-files') {
        renderDashboardPage('ALL_PIPELINE'); // Mode 1: All Files by Entry Date
    } else if (view === 'dashboard') {
        renderDashboardPage('FINANCIAL_PAYOUT'); // Mode 2: Disbursed only by Disb. Date
    } else if (view === 'users' && (currentUserRole === 'ADMIN' || auth.currentUser?.email === 'sharmashubham22657@gmail.com')) {
        renderUserManagement();
    } else if (view === 'settings' && currentUserRole === 'ADMIN') {
        renderSettingsPage();
    }
};






// --- VIEW: DAILY ENTRY ---
const renderEntryPage = () => {
    pageTitle.textContent = (currentUserRole === 'ADMIN' || currentUserRole === 'BACK_OFFICE') ? 'Daily Entries' : 'YOUR FILES';

    const container = document.createElement('div');

    // Form Section
    const formHtml = `
    <div class="card">
            <h3 id="form-title" style="margin-bottom:1.5rem;">Add New Case</h3>
            <form id="entry-form" onsubmit="handleFormSubmit(event)">
                <div class="form-row">
                    <div>
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">Entry Date</label>
                        <input type="date" id="date" class="form-control" required>
                    </div>
                    <div>
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569; border:1px solid #c7d2fe; border-radius:4px; padding:0 4px; display:inline-block; margin-left:0;">Disbursement Date (Optional)</label>
                        <input type="date" id="disbursementDate" class="form-control" style="background:#f0f9ff !important;">
                    </div>
                    <div>
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">Customer Name</label>
                        <input type="text" id="customerName" class="form-control" placeholder="Enter Name" required>
                    </div>
                    <div>
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">LOS Number</label>
                        <input type="text" id="losNo" class="form-control" placeholder="Enter LOS No">
                    </div>
                    <div>
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">Bank Name</label>
                        <input type="text" id="bankName" class="form-control" placeholder="Select Bank" list="bank-list">
                        <datalist id="bank-list">
                            <option value="HDFC Bank">
                            <option value="ICICI Bank">
                            <option value="Axis Bank">
                            <option value="Axis Finance">
                            <option value="Chola MS">
                            <option value="Kotak Mahindra">
                            <option value="Bajaj Finserv">
                        </datalist>
                    </div>
                </div>

                <div class="form-row">
                    <div>
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">Loan Amount</label>
                        <input type="number" id="amount" class="form-control" placeholder="₹ Amount" required min="0">
                    </div>
                    <div>
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">Tenure (Months)</label>
                        <input type="number" id="tenure" class="form-control" placeholder="e.g. 60" required list="tenure-list">
                        <datalist id="tenure-list">
                            <option value="12">
                            <option value="24">
                            <option value="36">
                            <option value="48">
                            <option value="60">
                            <option value="120">
                            <option value="180">
                            <option value="240">
                        </datalist>
                    </div>
                    <div>
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">Interest Rate (%)</label>
                        <input type="number" id="interestRate" class="form-control" placeholder="Rate" step="0.01">
                    </div>
                </div>

                <div class="form-row">
                    <div>
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">Case Type</label>
                        <input type="text" id="caseType" class="form-control" placeholder="Select Type" list="case-type-list" required>
                        <datalist id="case-type-list">
                            <option value="Normal PL">
                            <option value="Golden Edge">
                            <option value="BT">
                            <option value="Ex BT">
                            <option value="Business Loan">
                            <option value="Home Loan">
                        </datalist>
                    </div>
                    <div>
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">Location</label>
                        <input type="text" id="location" class="form-control" placeholder="City / Area">
                    </div>
                    <!-- Dynamic Custom Columns in Row 3 -->
                    ${customColumns.slice(0, 1).map(col => `
                        <div>
                            <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">${col}</label>
                            <input type="text" id="custom-${col}" class="form-control" placeholder="Enter ${col}">
                        </div>
                    `).join('')}
                </div>

                <div class="form-row">
                    <!-- Additional Custom Columns in a separate row if many -->
                    ${customColumns.slice(1).map(col => `
                        <div>
                            <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">${col}</label>
                            <input type="text" id="custom-${col}" class="form-control" placeholder="Enter ${col}">
                        </div>
                    `).join('')}
                    <div>
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">Status</label>
                        <input type="text" id="status" class="form-control" placeholder="Select Status" required list="status-list" 
                            oninput="if(this.value==='Disbursed') document.getElementById('disbursementDate').value='${new Date().toISOString().split('T')[0]}'">
                        <datalist id="status-list">
                            <option value="Underwriting">
                            <option value="Underwriting Forward">
                            <option value="Approved">
                            <option value="Disbursed">
                            <option value="Rejected">
                        </datalist>
                    </div>
                    <div>
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">Executive Name</label>
                        <input type="text" id="executiveName" class="form-control" placeholder="Select Executive" required list="exec-list">
                        <datalist id="exec-list">
                            <!-- Auto-populated from existing data and user list -->
                            ${getExecutiveListForDropdown().map(name => `<option value="${name}">`).join('')}
                        </datalist>
                    </div>
                    <div style="flex:2;">
                        <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">Remarks / Notes</label>
                        <input type="text" id="remarks" class="form-control" placeholder="Any comments...">
                    </div>
                </div>
                
                <div style="display:flex; gap:1rem; margin-top:2rem; align-items:center;">
                    <button type="submit" id="submit-btn" class="btn btn-primary"><i class="fas fa-plus"></i> Add Entry</button>
                    ${currentUserRole === 'ADMIN' ? `
                        <button type="button" onclick="seedTestData()" class="btn" style="background:#f8fafc; color:#64748b; border:1px solid #e2e8f0; font-size:0.8rem;">
                            <i class="fas fa-database"></i> Seed 100 Test Cases
                        </button>
                    ` : ''}
                    <button type="button" id="cancel-btn" class="btn btn-danger" style="display:none;" onclick="cancelEdit()">Cancel</button>
                </div>
            </form>
        </div>
    `;

    // Table Section
    const currentYear = new Date().getFullYear();
    const tableHtml = `
        <div class="controls-bar" style="display:flex; justify-content:space-between; margin-bottom:1rem; gap:1rem; flex-wrap:wrap; align-items:center;">
            <div style="display:flex; gap:0.5rem; align-items:center; flex-wrap:wrap;">
                <input type="text" id="search-input" class="form-control" placeholder="Search..." oninput="window.renderTableRows()" style="max-width:250px;">
                
                <!-- Date Filters -->
                <select id="filter-year" class="filter-select" onchange="window.renderTableRows()">
                    <option value="">All Years</option>
                    ${Array.from({ length: currentYear + 5 - 2020 + 1 }, (_, i) => 2020 + i).map(year =>
        `<option value="${year}">${year}</option>`
    ).join('')}
                </select>
                
                <select id="filter-month" class="filter-select" onchange="window.renderTableRows()">
                    <option value="">All Months</option>
                    <option value="0">January</option>
                    <option value="1">February</option>
                    <option value="2">March</option>
                    <option value="3">April</option>
                    <option value="4">May</option>
                    <option value="5">June</option>
                    <option value="6">July</option>
                    <option value="7">August</option>
                    <option value="8">September</option>
                    <option value="9">October</option>
                    <option value="10">November</option>
                    <option value="11">December</option>
                </select>

                <!-- Bulk Action Button -->
                <button id="bulk-delete-btn" class="btn btn-danger" onclick="deleteSelectedEntries()" style="display:none; padding: 0.5rem 1rem; font-size: 0.85rem;">
                    <i class="fas fa-trash"></i> Delete Selected (<span id="selected-count">0</span>)
                </button>
            </div>
            ${currentUserRole === 'ADMIN' ? `
            <div style="display:flex; gap:0.5rem; align-items:center;">
                <button class="btn btn-excel" onclick="exportToCSV()" title="Export All Data">
                    <i class="fas fa-file-excel"></i> Export Excel
                </button>
                <div style="position:relative;">
                    <button class="btn btn-primary" onclick="triggerImport()" style="background:#0f172a; border:1px solid #1e293b;">
                        <i class="fas fa-file-import"></i> Import Excel
                    </button>
                    <input type="file" id="excel-input" accept=".xlsx, .xls" style="display:none;" onchange="handleExcelImport(this)">
                </div>
                <button class="btn" onclick="downloadImportTemplate()" style="background:none; color:#64748b; font-size:0.85rem; padding:0.5rem; text-decoration:underline;">
                    <i class="fas fa-download"></i> Template
                </button>
            </div>
            ` : ''}
        </div>

        <!-- Floating Ghost Scrollbar (Fixed at bottom of viewport) -->
        <div id="ghost-scrollbar-container" style="position:fixed; bottom:0; height:20px; 
            overflow-x:auto; overflow-y:hidden; z-index:1000; display:none; background:transparent;">
            <div id="ghost-scrollbar-content" style="height:1px;"></div>
        </div>

        <div class="table-container" id="main-table-container">
            <table style="font-size: 0.85rem;">
                <thead>
                    <tr>
                        <th style="width: 40px; text-align: center;">
                            <input type="checkbox" id="select-all" onclick="toggleSelectAll(this)">
                        </th>
                        ${columnOrder.map(key => {
        const sys = SYSTEM_COLUMNS.find(c => c.id === key);
        return `<th>${sys ? sys.label : key}</th>`;
    }).join('')}
                        <th>Action</th>
                    </tr>
                </thead>

                <tbody id="entries-body">
                    <!-- Rows injected here -->
                </tbody>
            </table>
        </div>
        <div class="total-summary">
            <span>Total Disbursed Volume</span>
            <strong id="grand-total">₹0</strong>
        </div>
    `;

    container.innerHTML = ((currentUserRole === 'ADMIN' || currentUserRole === 'BACK_OFFICE') ? formHtml : '') + tableHtml;
    contentArea.innerHTML = '';
    contentArea.appendChild(container);

    const dateInput = document.getElementById('date');
    if (dateInput) {
        dateInput.valueAsDate = new Date();
    }

    // Auto-fill and Lock Executive Name ONLY for stringently restricted EXECUTIVE role
    const execInput = document.getElementById('executiveName');
    if (execInput && currentUserRole === 'EXECUTIVE' && assignedName) {
        execInput.value = assignedName;
        execInput.readOnly = true;
        execInput.style.backgroundColor = '#f1f5f9';
        execInput.style.cursor = 'not-allowed';
    }

    renderTableRows();

    // Initialize sticky scrollbar
    setTimeout(() => initStickyScrollbar(), 200);
};

const getUniqueExecutives = () => {
    const executives = new Set(loans.map(l => l.executiveName));
    return Array.from(executives).sort();
};

const getExecutiveListForDropdown = () => {
    // Combine names from existing loans and the authorized users list
    const fromLoans = loans.map(l => l.executiveName);
    const fromUsers = authorizedUsers.filter(u => u.role === 'EXECUTIVE').map(u => u.assignedName);
    const combined = new Set([...fromLoans, ...fromUsers]);
    return Array.from(combined).filter(Boolean).sort();
};

// Explicitly attach to window for HTML access
window.handleFormSubmit = (e) => {
    e.preventDefault();

    const amount = parseFloat(document.getElementById('amount').value);

    const entryData = {
        id: editingId || Date.now().toString(),
        date: document.getElementById('date').value,
        disbursementDate: document.getElementById('disbursementDate')?.value || '',
        customerName: document.getElementById('customerName').value,
        losNo: document.getElementById('losNo').value,
        bankName: document.getElementById('bankName').value,
        amount: amount,
        interestRate: parseFloat(document.getElementById('interestRate').value) || 0,
        tenure: document.getElementById('tenure').value,
        caseType: document.getElementById('caseType').value,
        location: document.getElementById('location').value,
        status: document.getElementById('status').value,
        executiveName: document.getElementById('executiveName').value.trim(),
        remarks: document.getElementById('remarks').value,
        customData: {}
    };

    // Capture dynamic custom columns
    customColumns.forEach(col => {
        const val = document.getElementById(`custom-${col}`)?.value;
        if (val) entryData.customData[col] = val;
    });

    const performSave = async () => {

        // Auto-apply payout % from bankPayoutConfig if the entry is Disbursed
        // and no payoutPercent has been manually set
        if (entryData.status === 'Disbursed' && entryData.executiveName && entryData.bankName) {
            const configKey = `${entryData.executiveName}|||${entryData.bankName}`;
            const savedPct = bankPayoutConfig[configKey];
            if (savedPct !== undefined && !entryData.payoutPercent) {
                entryData.payoutPercent = savedPct;
            }
            // Also check existing loans of same exec+bank if config has no entry yet
            if (entryData.payoutPercent === undefined) {
                const matchingLoan = loans.find(
                    l => l.executiveName === entryData.executiveName &&
                        l.bankName === entryData.bankName &&
                        l.status === 'Disbursed' &&
                        l.payoutPercent > 0
                );
                if (matchingLoan) entryData.payoutPercent = matchingLoan.payoutPercent;
            }
        }

        try {
            if (editingId) {
                const loanRef = doc(db, COLLECTION_NAME, editingId);
                const { id, ...dataToSave } = entryData;
                await updateDoc(loanRef, dataToSave);
                editingId = null;
                cancelEdit();
                showToast("Entry updated successfully.");
            } else {
                const { id, ...dataToSave } = entryData;
                await addDoc(collection(db, COLLECTION_NAME), dataToSave);

                const date = document.getElementById('date').value;
                e.target.reset();
                if (document.getElementById('date')) document.getElementById('date').value = date;
                if (document.getElementById('status')) document.getElementById('status').value = 'Underwriting';
                showToast("Saved to database!");
            }
        } catch (error) {
            console.error("Firebase Error:", error.message);
            alert("Error saving data: " + error.message);
        }
    };

    performSave();

    // saveToLocalStorage(); // Optional
    // renderTableRows(); // Handled by onSnapshot
};

window.renderTableRows = () => {
    const tbody = document.getElementById('entries-body');
    if (!tbody) return;

    const searchInput = document.getElementById('search-input');
    const searchTerm = (searchInput?.value || '').toString().toLowerCase().trim();

    const monthFilterStr = document.getElementById('filter-month')?.value;
    const yearFilterStr = document.getElementById('filter-year')?.value;

    console.log("Searching for:", searchTerm, "Month:", monthFilterStr, "Year:", yearFilterStr); // Debugging

    const filteredLoans = loans.filter(loan => {
        // Safe String Casting helper
        const safeStr = (val) => String(val || '').toLowerCase();

        const textMatch =
            safeStr(loan.customerName).includes(searchTerm) ||
            safeStr(loan.losNo).includes(searchTerm) ||
            safeStr(loan.bankName).includes(searchTerm) ||
            safeStr(loan.caseType).includes(searchTerm) ||
            safeStr(loan.location).includes(searchTerm) ||
            safeStr(loan.executiveName).includes(searchTerm) ||
            safeStr(loan.amount).includes(searchTerm) ||
            safeStr(loan.status).includes(searchTerm);

        // Date Logic
        let dateMatch = true;
        if (loan.date && (monthFilterStr !== '' || yearFilterStr !== '')) {
            const loanDate = new Date(loan.date);
            if (!isNaN(loanDate)) {
                if (monthFilterStr !== '') {
                    if (loanDate.getMonth() !== parseInt(monthFilterStr)) {
                        dateMatch = false;
                    }
                }
                if (yearFilterStr !== '') {
                    if (loanDate.getFullYear() !== parseInt(yearFilterStr)) {
                        dateMatch = false;
                    }
                }
            }
        }

        return textMatch && dateMatch;
    });

    tbody.innerHTML = filteredLoans.map(loan => {
        const rowClass = {
            'Disbursed': 'row-disbursed',
            'Approved': 'row-approved',
            'Rejected': 'row-rejected',
            'Underwriting Forward': 'row-forward',
            'Underwriting': 'row-underwriting'
        }[loan.status] || '';
        return `
        <tr class="${rowClass}">
            <td style="text-align: center;">
                <input type="checkbox" class="entry-checkbox" value="${loan.id}" onchange="updateBulkState()">
            </td>
            ${columnOrder.map(key => {
            const val = loan[key] !== undefined ? loan[key] : (loan.customData ? loan.customData[key] : undefined);
            let displayVal = val !== undefined && val !== '' ? sanitizeHTML(val) : '-';

            // Special rendering for specific columns
            if (key === 'date') displayVal = formatDate(val);
            if (key === 'disbursementDate') displayVal = val ? `<span style="color:#059669; font-weight:600;">${formatDate(val)}</span>` : '-';
            if (key === 'amount') return `<td class="amount">${formatCurrency(loan.amount)}</td>`;
            if (key === 'tenure') return `<td>${loan.tenure || '-'} M</td>`;
            if (key === 'status') {
                return `
                    <td>
                        ${(currentUserRole === 'ADMIN' || currentUserRole === 'BACK_OFFICE') ? `
                        <select class="inline-status-select ${getStatusClass(loan.status)}" onchange="updateStatus('${loan.id}', this)">
                            <option value="Underwriting" ${loan.status === 'Underwriting' ? 'selected' : ''}>Underwriting</option>
                            <option value="Underwriting Forward" ${loan.status === 'Underwriting Forward' ? 'selected' : ''}>Underwriting Forward</option>
                            <option value="Approved" ${loan.status === 'Approved' ? 'selected' : ''}>Approved</option>
                            <option value="Disbursed" ${loan.status === 'Disbursed' ? 'selected' : ''}>Disbursed</option>
                            <option value="Rejected" ${loan.status === 'Rejected' ? 'selected' : ''}>Rejected</option>
                        </select>
                        ` : `<span class="status-badge ${getStatusClass(loan.status)}">${loan.status}</span>`}
                    </td>`;
            }
            if (key === 'remarks') {
                return `<td style="font-size:0.8rem; color:var(--text-secondary); max-width:150px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${displayVal}">${displayVal}</td>`;
            }

            return `<td>${displayVal}</td>`;
        }).join('')}
            <td>
                <div class="action-buttons">
                ${(currentUserRole === 'ADMIN' || currentUserRole === 'BACK_OFFICE') ? `
                <button onclick="editEntry('${loan.id}')" style="color:var(--primary-color); background:none; border:none; cursor:pointer; margin-right:0.5rem;" title="Edit">
                    <i class="fas fa-edit"></i>
                </button>
                <button onclick="deleteEntry('${loan.id}')" style="color:red; background:none; border:none; cursor:pointer;" title="Delete">
                    <i class="fas fa-trash"></i>
                </button>
                ` : '<span style="color:#94a3b8; font-style:italic; font-size:0.75rem;">View Only</span>'}
            </td>
        </tr>
        `;
    }).join('');

    // Reset Select All checkbox
    const selectAllBox = document.getElementById('select-all');
    if (selectAllBox) selectAllBox.checked = false;
    updateBulkState();

    // Calculate totals based on filtered visible rows - ONLY Disbursed
    const disbursedLoans = filteredLoans.filter(l => l.status === 'Disbursed');

    // Total Volume (Disbursed Only)
    const totalVolume = disbursedLoans.reduce((sum, loan) => sum + loan.amount, 0);
    const grandTotalEl = document.getElementById('grand-total');
    if (grandTotalEl) {
        grandTotalEl.textContent = formatCurrency(totalVolume);
    }
};

// Inline Status Update - saves directly to Firestore on dropdown change
window.updateStatus = async (id, selectEl) => {
    const newStatus = selectEl.value;
    // Update class immediately for visual feedback
    selectEl.className = `inline-status-select ${getStatusClass(newStatus)}`;
    
    try {
        const updateData = { status: newStatus };
        
        // AUTO-SET DISBURSEMENT DATE:
        // If status changed to Disbursed, and it's not already set, use today.
        if (newStatus === 'Disbursed') {
            const currentLoan = loans.find(l => l.id === id);
            if (currentLoan && !currentLoan.disbursementDate) {
                updateData.disbursementDate = new Date().toISOString().split('T')[0];
            }
        }

        await updateDoc(doc(db, COLLECTION_NAME, id), updateData);
        showToast(`Status updated to "${newStatus}"`);
    } catch (error) {
        console.error('Status update error:', error);
        alert('Failed to update status: ' + error.message);
    }
};

// Sticky Scrollbar Functionality (Ghost Scrollbar)
// We keep global listeners but find elements dynamically to handle page navigation
let scrollListenersAttached = false;

const initStickyScrollbar = () => {
    const tableContainer = document.getElementById('main-table-container');
    const ghostContainer = document.getElementById('ghost-scrollbar-container');
    const ghostContent = document.getElementById('ghost-scrollbar-content');

    if (!tableContainer || !ghostContainer || !ghostContent) return;

    // 1. Setup Table-Specific Listeners (Must re-attach on every page render)
    const syncScroll = (source, target) => {
        if (Math.abs(target.scrollLeft - source.scrollLeft) > 1) {
            target.scrollLeft = source.scrollLeft;
        }
    };

    // Remove old listeners implicitly by the element being replaced, but we add fresh ones
    tableContainer.addEventListener('scroll', () => syncScroll(tableContainer, ghostContainer));
    ghostContainer.addEventListener('scroll', () => syncScroll(ghostContainer, tableContainer));

    // 2. Setup Global Visibility Logic (Attached only once)
    const checkVisibility = () => {
        // Find elements fresh in case of navigation
        const currentTable = document.getElementById('main-table-container');
        const currentGhost = document.getElementById('ghost-scrollbar-container');
        const currentContent = document.getElementById('ghost-scrollbar-content');

        if (!currentTable || !currentGhost || !currentContent) return;

        const rect = currentTable.getBoundingClientRect();
        const viewportHeight = window.innerHeight;

        const needsScroll = currentTable.scrollWidth > currentTable.clientWidth;
        const topOfTableVisible = rect.top < viewportHeight;
        const bottomOfTableBelowView = rect.bottom > viewportHeight;

        if (needsScroll && topOfTableVisible && bottomOfTableBelowView) {
            currentContent.style.width = currentTable.scrollWidth + 'px';
            currentGhost.style.left = rect.left + 'px';
            currentGhost.style.width = rect.width + 'px';
            currentGhost.style.display = 'block';
            currentGhost.scrollLeft = currentTable.scrollLeft;
        } else {
            currentGhost.style.display = 'none';
        }
    };

    if (!scrollListenersAttached) {
        window.addEventListener('scroll', checkVisibility);
        window.addEventListener('resize', checkVisibility);
        setInterval(checkVisibility, 1000); // Periodic check for content changes
        scrollListenersAttached = true;
    }

    // Initial check
    setTimeout(checkVisibility, 200);
};

// Undo History
let actionHistory = [];

// Toast Notification
const showToast = (message) => {
    let toast = document.getElementById('toast-notification');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'toast-notification';
        toast.style.cssText = `
            position: fixed;
            bottom: 2rem;
            right: 2rem;
            background: #1e293b;
            color: white;
            padding: 1rem 1.5rem;
            border-radius: 8px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            z-index: 1000;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            transform: translateY(100px);
            transition: transform 0.3s ease-out;
            font-size: 0.9rem;
        `;
        document.body.appendChild(toast);
    }
    toast.innerHTML = `<i class="fas fa-info-circle"></i> ${message}`;
    toast.style.transform = 'translateY(0)';

    setTimeout(() => {
        toast.style.transform = 'translateY(100px)';
    }, 4000);
};

// Custom Confirmation Modal
const showConfirm = (message, onConfirm) => {
    let overlay = document.getElementById('confirm-overlay');
    if (!overlay) {
        overlay = document.createElement('div');
        overlay.id = 'confirm-overlay';
        overlay.style.cssText = `
            position: fixed;
            top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(0, 0, 0, 0.5);
            display: flex; align-items: center; justify-content: center;
            z-index: 9999;
            backdrop-filter: blur(4px);
        `;
        document.body.appendChild(overlay);
    }

    overlay.innerHTML = `
        <div class="card" style="max-width: 400px; width: 90%; text-align: center; padding: 2rem; border-radius: 12px; box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1);">
            <div style="width: 60px; height: 60px; background: #fee2e2; color: #ef4444; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 1.5rem; font-size: 1.5rem;">
                <i class="fas fa-exclamation-triangle"></i>
            </div>
            <h3 style="margin-bottom: 0.5rem; color: #1e293b;">Are you sure?</h3>
            <p style="color: #64748b; margin-bottom: 2rem; font-size: 0.95rem; line-height: 1.5;">${message}</p>
            <div style="display: flex; gap: 1rem; justify-content: center;">
                <button id="confirm-cancel" class="btn" style="background: #f1f5f9; color: #475569; flex: 1;">Cancel</button>
                <button id="confirm-ok" class="btn btn-danger" style="flex: 1;">Confirm</button>
            </div>
        </div>
    `;

    overlay.style.display = 'flex';

    const close = () => overlay.style.display = 'none';

    document.getElementById('confirm-cancel').onclick = close;
    document.getElementById('confirm-ok').onclick = () => {
        close();
        onConfirm();
    };
};

// Global Keyboard Shortcuts
document.addEventListener('keydown', (e) => {
    // Ctrl+Z: Undo
    if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
        e.preventDefault();
        undoLastAction();
    }

    // Ctrl+S: Save/Submit form (if on Daily Entry page)
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        const form = document.getElementById('entry-form');
        if (form) {
            form.requestSubmit(); // Trigger form submission
        }
    }

    // Esc: Close modal
    if (e.key === 'Escape') {
        const modal = document.getElementById('exec-modal');
        if (modal && modal.classList.contains('open')) {
            closeModal();
        }
    }
});

const undoLastAction = () => {
    if (actionHistory.length === 0) {
        showToast("Nothing to undo.");
        return;
    }

    const lastAction = actionHistory.pop();
    if (lastAction.type === 'delete') {
        loans = [...loans, ...lastAction.data];
        saveToLocalStorage();
        renderTableRows();
        showToast(`Restored ${lastAction.data.length} entries.`);
    }
};

// Bulk Actions Logic
window.toggleSelectAll = (source) => {
    const checkboxes = document.querySelectorAll('.entry-checkbox');
    checkboxes.forEach(cb => cb.checked = source.checked);
    updateBulkState();
};

window.updateBulkState = () => {
    const checkboxes = document.querySelectorAll('.entry-checkbox:checked');
    const btn = document.getElementById('bulk-delete-btn');
    const countSpan = document.getElementById('selected-count');

    if (btn && countSpan) { // Ensure elements exist before manipulating
        if (checkboxes.length > 0 && (currentUserRole === 'ADMIN' || currentUserRole === 'BACK_OFFICE')) {
            btn.style.display = 'inline-flex';
            countSpan.textContent = checkboxes.length;
        } else {
            btn.style.display = 'none';
        }
    }
};

window.deleteSelectedEntries = () => {
    const checkboxes = document.querySelectorAll('.entry-checkbox:checked');
    const ids = Array.from(checkboxes).map(cb => cb.value);

    if (ids.length === 0) return;

    showConfirm(`Delete ${ids.length} selected entries permanently?`, async () => {
        const deletedItems = loans.filter(l => ids.includes(l.id));
        actionHistory.push({ type: 'delete', data: deletedItems });

        // Optimistic UI Update
        loans = loans.filter(l => !ids.includes(l.id));
        renderTableRows();

        try {
            for (const id of ids) {
                await deleteDoc(doc(db, COLLECTION_NAME, id));
            }
            showToast(`${ids.length} entries deleted.`);
        } catch (error) {
            console.error("Bulk Delete Error:", error);
            loans = [...loans, ...deletedItems]; // Revert
            renderTableRows();
            alert("SERVER ERROR: Failed to delete some entries.\nCheck your internet or database permissions.");
        }
    });
};

// Edit Logic
window.editEntry = (id) => {
    const loan = loans.find(l => l.id === id);
    if (!loan) {
        alert("Error: Entry not found in memory. Please refresh.");
        return;
    }

    editingId = id;

    // Scroll to top
    document.querySelector('.main-content').scrollTop = 0;

    // Populate Form
    document.getElementById('date').value = loan.date;
    document.getElementById('customerName').value = loan.customerName;
    document.getElementById('losNo').value = loan.losNo || '';
    document.getElementById('bankName').value = loan.bankName || '';
    document.getElementById('amount').value = loan.amount;
    document.getElementById('interestRate').value = loan.interestRate || '';
    document.getElementById('tenure').value = loan.tenure || '';
    document.getElementById('caseType').value = loan.caseType || '';
    document.getElementById('location').value = loan.location || '';
    document.getElementById('status').value = loan.status;
    document.getElementById('executiveName').value = loan.executiveName;
    document.getElementById('remarks').value = loan.remarks || '';

    // Populate dynamic custom columns
    customColumns.forEach(col => {
        const input = document.getElementById(`custom-${col}`);
        if (input) {
            input.value = (loan.customData && loan.customData[col]) || '';
        }
    });

    // Re-lock if executive role ONLY (not Back Office)

    const execInput = document.getElementById('executiveName');
    if (execInput && currentUserRole === 'EXECUTIVE') {
        execInput.readOnly = true;
    } else if (execInput) {
        execInput.readOnly = false;
        execInput.style.backgroundColor = '';
        execInput.style.cursor = '';
    }

    // Update UI
    document.getElementById('form-title').textContent = 'Edit Case';
    const submitBtn = document.getElementById('submit-btn');
    submitBtn.innerHTML = '<i class="fas fa-save"></i> Update Entry';
    submitBtn.classList.remove('btn-primary');
    submitBtn.classList.add('btn-success');
    // Removed undefined background-color override to allow class color to show

    document.getElementById('cancel-btn').style.display = 'inline-block';
};

window.cancelEdit = () => {
    editingId = null;
    document.getElementById('entry-form').reset();
    document.getElementById('date').valueAsDate = new Date(); // Reset to today
    document.getElementById('status').value = 'Underwriting';

    // Reset UI
    document.getElementById('form-title').textContent = 'Add New Case';
    const submitBtn = document.getElementById('submit-btn');
    submitBtn.innerHTML = '<i class="fas fa-plus"></i> Add Entry';
    submitBtn.classList.add('btn-primary');
    submitBtn.style.backgroundColor = '';

    document.getElementById('cancel-btn').style.display = 'none';
};

window.deleteEntry = (id) => {
    if (!id || id === 'undefined') {
        alert("CRITICAL ERROR: No ID provided to delete function!");
        return;
    }

    showConfirm('Delete this entry permanently?', () => {
        const deletedItem = loans.find(l => l.id === id);

        if (!deletedItem) {
            alert(`DATA ERROR: Entry with ID [${id}] not found in the list.`);
            return;
        }

        actionHistory.push({ type: 'delete', data: [deletedItem] });

        // Optimistic UI Update
        loans = loans.filter(l => l.id !== id);
        renderTableRows();

        const loanRef = doc(db, COLLECTION_NAME, id);
        deleteDoc(loanRef)
            .then(() => {
                showToast("Entry deleted.");
            })
            .catch(error => {
                console.error("Firebase Delete Error:", error);
                // Revert optimistic update
                loans.push(deletedItem);
                renderTableRows();
                alert(`SERVER REJECTION: ${error.message} (Code: ${error.code}). Check database permissions.`);
            });
    });
};

window.triggerImport = () => {
    document.getElementById('excel-input').click();
};

window.downloadImportTemplate = () => {
    const headers = [
        "Date (YYYY-MM-DD)", "Customer Name", "LOS No", "Bank Name", "Amount",
        "Tenure", "Interest Rate", "Case Type", "Location", "Status", "Executive Name", "Remarks"
    ];
    const ws = XLSX.utils.aoa_to_sheet([headers]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, "SBE_Import_Template.xlsx");
};

window.handleExcelImport = (input) => {
    const file = input.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

            if (jsonData.length === 0) {
                alert("File appears to be empty.");
                return;
            }

            let importedCount = 0;

            // Helper to parse Excel dates (Serial or String)
            const parseExcelDate = (raw) => {
                if (!raw) return new Date().toISOString().split('T')[0];

                // Handle Excel Serial Date (Numbers like 44562)
                if (typeof raw === 'number') {
                    // Excel base date is Dec 30, 1899 (crazy, but true due to leap year bug)
                    const date = new Date(Math.round((raw - 25569) * 86400 * 1000));
                    return date.toISOString().split('T')[0];
                }

                // Handle Strings
                const date = new Date(raw);
                if (!isNaN(date.getTime())) {
                    return date.toISOString().split('T')[0];
                }

                // Fallback for custom formats if simple parse fails (e.g., DD/MM/YYYY)
                // This is a basic implementation; relying on ISO is safest
                return new Date().toISOString().split('T')[0];
            };

            jsonData.forEach(row => {
                // Fuzzy mapping for column names
                const getVal = (keys) => {
                    for (let k of keys) {
                        const found = Object.keys(row).find(rk => rk.toLowerCase().includes(k.toLowerCase()));
                        if (found) return row[found];
                    }
                    return "";
                };

                // Extract data
                const dateRaw = getVal(["Date"]);
                const customer = getVal(["Customer", "Name"]);
                const amountRaw = getVal(["Amount"]);

                if (customer && amountRaw) {
                    const finalDate = parseExcelDate(dateRaw);

                    const newEntry = {
                        id: Date.now().toString() + Math.random().toString(36).substr(2, 5),
                        date: finalDate,
                        customerName: customer,
                        losNo: getVal(["LOS", "Application"]),
                        bankName: getVal(["Bank"]),
                        amount: parseFloat(amountRaw) || 0,
                        tenure: getVal(["Tenure"]),
                        interestRate: parseFloat(getVal(["Rate", "Interest"])) || 0,
                        caseType: getVal(["Type", "Case"]),
                        location: getVal(["Location", "City"]),
                        status: getVal(["Status"]) || "Underwriting",
                        executiveName: getVal(["Executive"]) || "Unassigned",
                        remarks: getVal(["Remark", "Note"])
                    };
                    // Add to Firestore
                    addDoc(collection(db, COLLECTION_NAME), newEntry);
                    importedCount++;
                }
            });

            // saveToLocalStorage(); // Optional
            // renderTableRows(); // Handled by onSnapshot
            alert(`Successfully imported ${importedCount} records to Firebase!`);
        } catch (error) {
            console.error(error);
            alert("Error parsing Excel file. Please ensure it is a valid format.");
        }
    };
    reader.readAsArrayBuffer(file);
    input.value = ""; // Reset
};

window.exportToCSV = () => {
    if (loans.length === 0) {
        alert("No data to export!");
        return;
    }

    // CSV Headers
    const getLabel = (key) => {
        const sys = SYSTEM_COLUMNS.find(c => c.id === key);
        return sys ? sys.label : key;
    };

    let csvContent = "data:text/csv;charset=utf-8,";
    csvContent += columnOrder.map(key => `"${getLabel(key)}"`).join(",") + "\n";

    // CSV Rows
    loans.forEach(loan => {
        const row = columnOrder.map(key => {
            const val = loan[key] !== undefined ? loan[key] : (loan.customData ? loan.customData[key] : '');
            return `"${val}"`;
        }).join(",");
        csvContent += row + "\r\n";
    });



    // Create Download Link
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `SBE_Data_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
};

// Dashboard Filter State
let dashboardDateFilter = { start: '', end: '' };
let activeDashboardMode = 'FINANCIAL_PAYOUT';

// --- VIEW: OWNER DASHBOARD ---
const renderDashboardPage = (mode = 'FINANCIAL_PAYOUT') => {
    activeDashboardMode = mode;
    pageTitle.textContent = mode === 'FINANCIAL_PAYOUT' ? 'Owner Dashboard' : 'Executive All Files';

    const executiveStats = {};
    let totalCompanyAmount = 0;

    loans.forEach(loan => {
        // --- GLOBAL DATE FILTER ---
        if (dashboardDateFilter.start && dashboardDateFilter.end) {
            if (loan.date < dashboardDateFilter.start || loan.date > dashboardDateFilter.end) {
                return; // Skip this loan if outside range
            }
        }

        const isDisbursed = loan.status === 'Disbursed';
        const isFileProcessed = loan.status !== 'Rejected';

        // --- DUAL MODE DATE LOGIC ---
        // Mode: Payout -> Only show Disbursed, group by Disb. Date
        // Mode: Pipeline -> Show ALL processed, group by Entry Date
        let reportingDate = '';
        if (mode === 'FINANCIAL_PAYOUT') {
            if (!isDisbursed) return; // SKIP non-disbursed files in Payout mode
            reportingDate = loan.disbursementDate || loan.date;
        } else {
            if (!isFileProcessed) return; // SKIP only rejected files in Pipeline mode
            reportingDate = loan.date;
        }

        const dObj = new Date(reportingDate);
        const monthKey = !isNaN(dObj) ? dObj.toLocaleString('default', { month: 'short', year: 'numeric' }) : 'Unknown';

        if (!executiveStats[loan.executiveName]) {
            executiveStats[loan.executiveName] = {
                name: loan.executiveName,
                count: 0,
                disbursedCount: 0,
                totalAmount: 0,
                monthlyVolume: {},
                loans: []
            };
        }

        executiveStats[loan.executiveName].count++;
        if (isDisbursed) executiveStats[loan.executiveName].disbursedCount++;
        
        executiveStats[loan.executiveName].totalAmount += loan.amount;
        totalCompanyAmount += loan.amount;

        // Increment Monthly Volume
        if (!executiveStats[loan.executiveName].monthlyVolume[monthKey]) {
            executiveStats[loan.executiveName].monthlyVolume[monthKey] = 0;
        }
        executiveStats[loan.executiveName].monthlyVolume[monthKey] += loan.amount;

        executiveStats[loan.executiveName].loans.push(loan);
    });

    const totalCompanyVolume = Object.values(executiveStats).reduce((sum, e) => sum + e.totalAmount, 0);
    const totalCompanyCount = Object.values(executiveStats).reduce((sum, e) => sum + e.count, 0);

    const execArray = Object.values(executiveStats).sort((a, b) => b.totalAmount - a.totalAmount);

    const container = document.createElement('div');

    // Date Filter UI
    const filterHtml = `
        <div class="card" style="margin-bottom:1.5rem; display:flex; flex-wrap:wrap; gap:1rem; align-items:center;">
            <div style="font-weight:600; color:var(--text-secondary);"><i class="far fa-calendar-alt"></i> Filter by Date:</div>
            <input type="date" id="startDate" class="form-control" style="width:auto;" value="${dashboardDateFilter.start}">
            <span style="color:var(--text-secondary);">to</span>
            <input type="date" id="endDate" class="form-control" style="width:auto;" value="${dashboardDateFilter.end}">
            <button class="btn btn-primary" onclick="applyDashboardFilter()" style="padding:0.4rem 1rem; font-size:0.9rem;">Apply</button>
            <button class="btn btn-danger" onclick="clearDashboardFilter()" style="padding:0.4rem 1rem; font-size:0.9rem; background: #e5e7eb; color: #374151; border:none;">Clear</button>
        </div>
    `;

    // Top Summary
    const summaryHtml = `
        <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap:1.5rem; margin-bottom:2rem;">
            <!-- Formal Volume Card -->
            <div class="card" style="background: white; padding:1.5rem; margin-bottom:0; border:1px solid #e2e8f0; border-top: 4px solid #0f172a; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                <div style="display:flex; justify-content:space-between; align-items:start;">
                    <div>
                        <h3 style="color:#64748b; font-size:0.85rem; font-weight:600; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:0.5rem;">Total Business Volume</h3>
                        <h1 style="color:#0f172a; font-size:2.2rem; font-weight:700; margin-bottom:0; letter-spacing:-0.03em;">${formatCurrency(totalCompanyAmount)}</h1>
                        ${dashboardDateFilter.start ? `
                        <div style="margin-top:0.75rem; display:inline-flex; align-items:center; background:#f8fafc; border:1px solid #e2e8f0; padding:4px 8px; border-radius:4px;">
                            <i class="far fa-calendar-alt" style="color:#64748b; font-size:0.75rem; margin-right:6px;"></i>
                            <span style="color:#334155; font-size:0.75rem; font-weight:600;">${formatDate(dashboardDateFilter.start)} - ${formatDate(dashboardDateFilter.end)}</span>
                        </div>` : ''}
                    </div>
                    <div style="width:48px; height:48px; background:#f1f5f9; border-radius:8px; display:flex; align-items:center; justify-content:center; color:#0f172a;">
                        <i class="fas fa-chart-pie" style="font-size:1.2rem;"></i>
                    </div>
                </div>
            </div>

            <!-- Total Count Card -->
            <div class="card" style="background: white; padding:1.5rem; margin-bottom:0; border:1px solid #e2e8f0; border-top: 4px solid #4f46e5; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                <div style="display:flex; justify-content:space-between; align-items:start;">
                    <div>
                        <h3 style="color:#64748b; font-size:0.85rem; font-weight:600; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:0.5rem;">${mode === 'FINANCIAL_PAYOUT' ? 'Total Disbursed' : 'Total Files'}</h3>
                        <h1 style="color:#4f46e5; font-size:2.2rem; font-weight:700; margin-bottom:0; letter-spacing:-0.03em;">${totalCompanyCount}</h1>
                    </div>
                    <div style="width:48px; height:48px; background:#eef2ff; border-radius:8px; display:flex; align-items:center; justify-content:center; color:#4f46e5;">
                        <i class="fas ${mode === 'FINANCIAL_PAYOUT' ? 'fa-check-double' : 'fa-copy'}" style="font-size:1.2rem;"></i>
                    </div>
                </div>
            </div>

            <!-- Average Amount View -->
            <div class="card" style="background: white; padding:1.5rem; margin-bottom:0; border:1px solid #e2e8f0; border-top: 4px solid #64748b; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                <div style="display:flex; justify-content:space-between; align-items:start;">
                    <div>
                        <h3 style="color:#64748b; font-size:0.85rem; font-weight:600; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:0.5rem;">Average File Size</h3>
                        <h1 style="color:#0f172a; font-size:2.2rem; font-weight:700; margin-bottom:0; letter-spacing:-0.03em;">${formatCurrency(totalCompanyCount > 0 ? totalCompanyVolume / totalCompanyCount : 0)}</h1>
                    </div>
                    <div style="width:48px; height:48px; background:#f1f5f9; border-radius:8px; display:flex; align-items:center; justify-content:center; color:#64748b;">
                        <i class="fas fa-hand-holding-usd" style="font-size:1.2rem;"></i>
                    </div>
                </div>
            </div>
        </div>
    `;

    // Grid of Executives
    const gridHtml = `
        <h3 style="margin-bottom:1rem; color:var(--text-secondary);">Executive Performance</h3>
        <div class="stats-grid">
            ${execArray.map(exec => {
        // Calculate adjustment totals for this executive
        const execAdjustments = executiveAdjustments.filter(a => a.executiveName === exec.name);
        const totalDebits = execAdjustments.filter(a => a.type === 'Debit').reduce((s, a) => s + (a.amount || 0), 0);
        const totalCredits = execAdjustments.filter(a => a.type === 'Credit').reduce((s, a) => s + (a.amount || 0), 0);
        
        // Calculate payment metrics for FINANCIAL_PAYOUT mode
        let netRevenue = 0;
        let hasAdjustments = execAdjustments.length > 0;
        
        if (mode === 'FINANCIAL_PAYOUT') {
            const loanRevenue = exec.loans.reduce((sum, l) => sum + (l.amount * (l.payoutPercent || 0) / 100), 0);
            netRevenue = loanRevenue - totalDebits + totalCredits;
        }

        // Current & Previous Month Keys
        const now = new Date();
        const curMonthKey = now.toLocaleString('default', { month: 'short', year: 'numeric' });
        const lastMonthDate = new Date();
        lastMonthDate.setMonth(lastMonthDate.getMonth() - 1);
        const lastMonthKey = lastMonthDate.toLocaleString('default', { month: 'short', year: 'numeric' });

        // Modal Action
        const escapedName = exec.name.replace(/'/g, "\\'");
        const openModalAction = mode === 'FINANCIAL_PAYOUT' 
            ? `openExecutiveDetails('${escapedName}', null, 'DISBURSED_ONLY')` 
            : `openExecutiveDetails('${escapedName}', null, 'ALL_FILES')`;

        const currentMonthVol = exec.monthlyVolume[curMonthKey] || 0;
        const lastMonthVol = exec.monthlyVolume[lastMonthKey] || 0;

        return `
                <div class="stat-card" onclick="${openModalAction}" style="display:flex; flex-direction:column; align-items:stretch; gap:1rem;">
                    <div style="display:flex; justify-content:space-between; align-items:flex-start;">
                        <div class="stat-info">
                            <h3 style="margin-bottom:0.25rem;">${sanitizeHTML(exec.name)}</h3>
                            <p style="font-size:1.5rem;">${formatCurrency(exec.totalAmount)} ${mode === 'FINANCIAL_PAYOUT' ? '<span style="font-size:0.75rem; font-weight:normal; opacity:0.6;">Disbursed</span>' : ''}</p>
                        </div>
                        <div class="stat-icon" style="width:40px; height:40px; font-size:1.1rem; background:${mode === 'FINANCIAL_PAYOUT' ? '#ecfdf5' : '#f0f9ff'}; color:${mode === 'FINANCIAL_PAYOUT' ? '#059669' : '#0369a1'};">
                            <i class="fas ${mode === 'FINANCIAL_PAYOUT' ? 'fa-wallet' : 'fa-folder-open'}"></i>
                        </div>
                    </div>

                    <div style="display:grid; grid-template-columns: 1fr 1fr; gap:0.75rem; border-top:1px solid #f1f5f9; padding-top:0.75rem;">
                        <div style="padding:0.5rem; background:#f8fafc; border-radius:6px;">
                            <div style="font-size:0.65rem; text-transform:uppercase; color:#64748b; font-weight:700;">${curMonthKey}</div>
                            <div style="font-size:0.95rem; font-weight:700; color:#0f172a;">${formatCurrency(currentMonthVol)}</div>
                        </div>
                        <div style="padding:0.5rem; background:#f1f5f9; border-radius:6px; opacity:0.8;">
                            <div style="font-size:0.65rem; text-transform:uppercase; color:#64748b; font-weight:700;">${lastMonthKey}</div>
                            <div style="font-size:0.95rem; font-weight:700; color:#475569;">${formatCurrency(lastMonthVol)}</div>
                        </div>
                    </div>

                    <div style="font-size:0.8rem; color:var(--text-secondary); display:flex; gap:10px; align-items:center;">
                        <span><i class="fas fa-file-invoice" style="color:#6366f1;"></i> ${exec.count} ${mode === 'FINANCIAL_PAYOUT' ? 'Disbursed' : 'Files'}</span>
                        <div style="flex:1; text-align:right; display:flex; justify-content:flex-end; gap:8px; align-items:center;">
                            ${(mode === 'FINANCIAL_PAYOUT') ? `
                            <button onclick="event.stopPropagation(); openMonthlySlipModal('${escapedName}')" 
                                class="btn" title="Generate Monthly Slip"
                                style="padding:6px 12px; border-radius:6px; background:#4f46e5; color:white; border:none; display:flex; align-items:center; gap:6px; cursor:pointer; font-size:0.75rem; font-weight:700; box-shadow:0 4px 6px -1px rgba(79, 70, 229, 0.4);">
                                <i class="fas fa-file-invoice-dollar"></i> MONTHLY SLIP
                            </button>
                            <span style="padding:4px 8px; background:${netRevenue >= 0 ? '#ecfdf5' : '#fef2f2'}; border-radius:4px; font-size:0.8rem; font-weight:700; color:${netRevenue >= 0 ? '#047857' : '#b91c1c'}; border:1px solid ${netRevenue >= 0 ? '#bbf7d0' : '#fecaca'};">
                                Net Payout: ${formatCurrency(netRevenue)}
                            </span>` : ''}
                        </div>
                    </div>
                </div>
            `}).join('')}
        </div>
    `;

    const modalHtml = `
        <div id="exec-modal" class="modal-overlay">
            <div class="modal modal-xl">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:2rem;">
                    <h2 id="modal-title" style="margin:0; color:var(--primary-color);">Executive Details</h2>
                    <button onclick="closeModal()" class="close-modal" style="background:none; border:none; font-size:1.5rem; cursor:pointer;">
                        <i class="fas fa-times"></i>
                    </button>
                </div>
                <div id="modal-content"></div>
            </div>
        </div>
    `;

    container.innerHTML = filterHtml + summaryHtml + gridHtml + modalHtml;
    contentArea.innerHTML = '';
    contentArea.appendChild(container);
};

window.applyDashboardFilter = () => {
    const start = document.getElementById('startDate').value;
    const end = document.getElementById('endDate').value;

    if (start && end) {
        dashboardDateFilter = { start, end };
        renderDashboardPage(activeDashboardMode);
    } else {
        alert("Please select both Start and End date.");
    }
};

window.clearDashboardFilter = () => {
    dashboardDateFilter = { start: '', end: '' };
    renderDashboardPage(activeDashboardMode);
};

// --- MONTHLY SLIP (WHATSAPP/IMAGE) SYSTEM ---
window.openMonthlySlipModal = (name) => {
    const now = new Date();
    const curMonthKey = now.toLocaleString('default', { month: 'short', year: 'numeric' });
    
    // We'll reuse the executive data from dashboard logic but in a specialized layout
    // First, ask user for the month if they want to change from current
    
    const userObj = authorizedUsers.find(u => u.assignedName === name);
    const mobileNumber = userObj ? (userObj.mobile || '') : '';

    document.getElementById('modal-title').innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; width:100%; gap:20px;">
            <span><i class="fas fa-file-invoice-dollar"></i> Payout Slip</span>
            <div style="display:flex; gap:10px;">
                <button onclick="downloadSlipImage('${name.replace(/'/g, "\\'")}')" class="btn" style="background:#0f172a; color:white; padding:8px 16px; font-size:0.85rem; height:40px; border-radius:8px;">
                    <i class="fas fa-image"></i> Download
                </button>
                <button onclick="shareSlipToWhatsApp('${name.replace(/'/g, "\\'")}')" class="btn" style="background:#25d366; color:white; border:none; padding:8px 16px; font-size:0.85rem; height:40px; border-radius:8px;">
                    <i class="fab fa-whatsapp"></i> WhatsApp
                </button>
            </div>
        </div>
    `;

    const content = `
        <div style="padding:10px;">
            <div id="payout-slip-capture" style="background:white; border:1px solid #e2e8f0; border-radius:12px; padding:2rem; width:100%; max-width:750px; margin:0 auto; box-shadow:0 10px 15px -3px rgba(0,0,0,0.1);">
                <div style="text-align:center; border-bottom:2px solid #f1f5f9; padding-bottom:1.5rem; margin-bottom:1.5rem;">
                    <img src="360_F_204812227_fVnI2OTNSY7FYF5ZaFU5kuZjNst0kpBF-removebg-preview.png" style="height:80px; margin-bottom:0.5rem;">
                    <h2 style="margin:0; text-transform:uppercase; color:#0f172a; letter-spacing:0.1em; font-size:1.2rem;">Shri Balaji Enterprises</h2>
                    <p style="margin:4px 0 0; color:#64748b; font-size:0.85rem;">Monthly Executive Payout Report</p>
                </div>

                <div style="display:flex; justify-content:space-between; margin-bottom:1.5rem; font-size:0.9rem;">
                    <div>
                        <div style="color:#64748b; font-size:0.75rem; font-weight:700; text-transform:uppercase;">Executive</div>
                        <div style="font-weight:700; color:#0f172a; font-size:1.1rem;">${sanitizeHTML(name)}</div>
                        ${mobileNumber ? `<div style="color:#6366f1; font-weight:600;">+91 ${mobileNumber}</div>` : ''}
                    </div>
                    <div style="text-align:right;">
                        <div style="color:#64748b; font-size:0.75rem; font-weight:700; text-transform:uppercase;">Period</div>
                        <div style="font-weight:700; color:#0369a1; font-size:1.1rem;" id="slip-month-label">${curMonthKey}</div>
                    </div>
                </div>

                <div style="margin-bottom:1.5rem;">
                    <table style="width:100%; border-collapse:collapse; font-size:0.85rem;">
                        <thead>
                            <tr style="border-bottom:2px solid #0f172a; background:#f8fafc;">
                                <th style="padding:10px; text-align:left;">Customer</th>
                                <th style="padding:10px; text-align:center;">Bank</th>
                                <th style="padding:10px; text-align:center;">Date</th>
                                <th style="padding:10px; text-align:right;">Amount</th>
                                <th style="padding:10px; text-align:right;">Revenue</th>
                            </tr>
                        </thead>
                        <tbody id="slip-rows-container">
                            <!-- Rows injected here -->
                        </tbody>
                    </table>
                </div>

                <div style="background:#f1f5f9; border-radius:8px; padding:1.2rem;">
                    <div style="display:flex; justify-content:space-between; margin-bottom:0.5rem; color:#475569;">
                        <span>Total Monthly Volume:</span>
                        <span id="slip-total-volume" style="font-weight:700;">₹0</span>
                    </div>
                    <div style="display:flex; justify-content:space-between; border-top:1px solid #e2e8f0; padding-top:0.5rem; color:#0f172a; font-size:1.2rem;">
                        <span style="font-weight:800;">NET REVENUE:</span>
                        <span id="slip-net-revenue" style="font-weight:800; color:#059669;">₹0</span>
                    </div>
                </div>

                <div style="margin-top:2rem; display:flex; justify-content:space-between; align-items:flex-end;">
                    <div style="color:#94a3b8; font-size:0.7rem; font-style:italic;">
                        This is a computer-generated performance slip — Shri Balaji Enterprises
                    </div>
                    <div style="text-align:right;">
                        <div style="margin-bottom:8px; border-bottom:1px solid #e2e8f0; width:150px;"></div>
                        <div style="font-size:0.7rem; font-weight:700; color:#64748b; text-transform:uppercase;">Authorized Signatory</div>
                    </div>
                </div>
            </div>
        </div>
    `;

    document.getElementById('modal-content').innerHTML = content;
    document.getElementById('exec-modal').classList.add('open');
    
    // Auto-calculate data for the slip
    updateSlipContent(name, curMonthKey);
};

window.updateSlipContent = (name, monthKey) => {
    // Collect specific month disbursed files
    const monthLoans = loans.filter(l => {
        if (l.executiveName !== name || l.status !== 'Disbursed') return false;
        const d = new Date(l.disbursementDate || l.date);
        return d.toLocaleString('default', { month: 'short', year: 'numeric' }) === monthKey;
    });

    const rowsHtml = monthLoans.map(l => {
        const payout = l.payoutPercent || 0;
        const revenue = (l.amount * payout / 100);
        return `
        <tr style="border-bottom:1px solid #f1f5f9;">
            <td style="padding:10px; font-weight:600;">${sanitizeHTML(l.customerName)}</td>
            <td style="padding:10px; text-align:center; color:#64748b;">${sanitizeHTML(l.bankName || '-')}</td>
            <td style="padding:10px; text-align:center; color:#64748b;">${formatDate(l.disbursementDate || l.date)}</td>
            <td style="padding:10px; text-align:right; font-weight:700;">${formatCurrency(l.amount)}</td>
            <td style="padding:10px; text-align:right; font-weight:700; color:#059669;">${formatCurrency(revenue)}</td>
        </tr>
    `}).join('') || `<tr><td colspan="5" style="padding:20px; text-align:center; color:#94a3b8;">No disbursed files for this month.</td></tr>`;

    document.getElementById('slip-rows-container').innerHTML = rowsHtml;
    
    const totalVolume = monthLoans.reduce((s, l) => s + (l.amount || 0), 0);
    // Revenue logic: use bankPayoutConfig or specific loan percent
    const totalRevenue = monthLoans.reduce((s, l) => {
        const pct = l.payoutPercent || 0;
        return s + (l.amount * pct / 100);
    }, 0);

    // Adjustments
    const adj = executiveAdjustments.filter(a => {
        const d = new Date(a.date);
        const aMonth = d.toLocaleString('default', { month: 'short', year: 'numeric' });
        return a.executiveName === name && aMonth === monthKey;
    });
    const netAdj = adj.reduce((s, a) => a.type === 'Credit' ? s + a.amount : s - a.amount, 0);

    document.getElementById('slip-total-volume').textContent = formatCurrency(totalVolume);
    document.getElementById('slip-net-revenue').textContent = formatCurrency(totalRevenue + netAdj);
};

window.downloadSlipImage = (name) => {
    const slip = document.getElementById('payout-slip-capture');
    if (!slip) return;
    
    showToast("Generating image...");
    html2canvas(slip, {
        scale: 2, // Higher quality
        useCORS: true, 
        backgroundColor: '#f8fafc'
    }).then(canvas => {
        const link = document.createElement('a');
        link.download = `SBE_Payout_Slip_${name.replace(/\s+/g, '_')}_${Date.now()}.png`;
        link.href = canvas.toDataURL('image/png');
        link.click();
        showToast("Slip downloaded!");
    });
};

window.shareSlipToWhatsApp = (name) => {
    const slip = document.getElementById('payout-slip-capture');
    if (!slip) return;

    const userObj = authorizedUsers.find(u => u.assignedName === name);
    const mobile = userObj ? (userObj.mobile || '') : '';
    
    // 🔥 STEP 1: Open chat IMMEDIATELY to satisfy browser pop-up rules
    shareTextFallback(name, mobile);
    showToast("Opening WhatsApp... capturing slip inside clipboard.");

    // 🔥 STEP 2: WHILE chat opens, capture image background
    html2canvas(slip, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#f8fafc'
    }).then(canvas => {
        canvas.toBlob(async (blob) => {
            // Attempt to copy image to clipboard (works best on Desktop)
            if (navigator.clipboard && navigator.clipboard.write) {
                try {
                    const data = [new ClipboardItem({ [blob.type]: blob })];
                    await navigator.clipboard.write(data);
                    showToast("Slip Image Copied! Just PASTE (Ctrl+V) in the chat.");
                } catch (err) {
                    console.warn("Clipboard write failed", err);
                }
            }
        }, 'image/png');
    });
};

const shareTextFallback = (name, mobile) => {
    const totalVol = document.getElementById('slip-total-volume').textContent;
    const netRev = document.getElementById('slip-net-revenue').textContent;
    const month = document.getElementById('slip-month-label').textContent;

    const message = `*Monthly Payout Slip: Shri Balaji Enterprises*\n\n` +
                    `Executive: *${name}*\n` +
                    `Month: *${month}*\n\n` +
                    `✅ Total Disbursed: *${totalVol}*\n` +
                    `💰 Net Payout: *${netRev}*\n\n` +
                    `_Check the image for detailed customer & bank revenue. (PASTE IT HERE)_`;

    const encoded = encodeURIComponent(message);
    const cleanMobile = mobile ? mobile.replace(/\D/g, '') : '';
    
    // Using api.whatsapp.com for more reliable direct-to-number redirect
    const waUrl = cleanMobile 
        ? `https://api.whatsapp.com/send?phone=91${cleanMobile}&text=${encoded}` 
        : `https://api.whatsapp.com/send?text=${encoded}`;
        
    window.open(waUrl, '_blank');
};

window.openExecutiveDetails = (name, targetMonthKey = null, filterMode = 'DISBURSED_ONLY') => {
    // Determine target month (default to current if not specified)
    const now = new Date();
    const curMonthKey = now.toLocaleString('default', { month: 'short', year: 'numeric' });
    const selectedKey = targetMonthKey || curMonthKey;
    const escapedName = name.replace(/'/g, "\\'");
    document.getElementById('modal-title').textContent = 'Executive Details';

    const execLoans = loans.filter(l => {
        if (l.executiveName !== name) return false;
        
        const isDisbursed = l.status === 'Disbursed';
        const reportingDate = (filterMode === 'DISBURSED_ONLY') ? (l.disbursementDate || l.date) : l.date;
        
        // In DISBURSED_ONLY mode, we only want disbursed files
        if (filterMode === 'DISBURSED_ONLY' && !isDisbursed) return false;
        // In ALL_FILES mode, we exclude rejected
        if (filterMode === 'ALL_FILES' && l.status === 'Rejected') return false;

        const dObj = new Date(reportingDate);
        const lKey = !isNaN(dObj) ? dObj.toLocaleString('default', { month: 'short', year: 'numeric' }) : '';
        return lKey === selectedKey;
    });

    const execAdj = executiveAdjustments.filter(a => {
        const aDate = new Date(a.date);
        const aKey = !isNaN(aDate) ? aDate.toLocaleString('default', { month: 'short', year: 'numeric' }) : '';
        return a.executiveName === name && aKey === selectedKey;
    });

    // Get all unique months available in data for this executive (based on Reporting Date)
    const availableMonths = [...new Set(loans.filter(l => l.executiveName === name).map(l => {
        const isDisbursed = l.status === 'Disbursed';
        const reportingDate = (filterMode === 'DISBURSED_ONLY') ? (l.disbursementDate || l.date) : l.date;
        const d = new Date(reportingDate);
        return !isNaN(d) ? d.toLocaleString('default', { month: 'short', year: 'numeric' }) : null;
    }).filter(Boolean))];
    
    // Add current month if not in list
    if (!availableMonths.includes(curMonthKey)) availableMonths.push(curMonthKey);
    // Sort months chronologically
    availableMonths.sort((a,b) => new Date(b) - new Date(a));

    // Calculate totals for SELECTED month only
    const totalVolume = execLoans.reduce((sum, l) => l.status !== 'Rejected' ? sum + l.amount : sum, 0);
    const totalRevenue = execLoans.reduce((sum, l) => {
        if (l.status === 'Disbursed') {
            const payout = l.payoutPercent || 0;
            return sum + (l.amount * payout / 100);
        }
        return sum;
    }, 0);

    const totalDebits = execAdj.filter(a => a.type === 'Debit').reduce((s, a) => s + (a.amount || 0), 0);
    const totalCredits = execAdj.filter(a => a.type === 'Credit').reduce((s, a) => s + (a.amount || 0), 0);
    const netRevenue = totalRevenue - totalDebits + totalCredits;

    const uniqueBanks = [...new Set(execLoans.map(l => l.bankName).filter(Boolean))];
    const today = new Date().toISOString().split('T')[0];

    const content = `
        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:1.5rem; gap:1.5rem; flex-wrap:wrap;">
            <div>
                <h2 style="margin-bottom:0.25rem; color:var(--primary-color);">${sanitizeHTML(name)}</h2>
                <div style="display:flex; align-items:center; gap:8px;">
                     <span style="padding:4px 10px; background:#eef2ff; color:#4f46e5; border-radius:6px; font-size:0.85rem; font-weight:700; border:1px solid #c7d2fe;">
                        <i class="far fa-calendar-check"></i> Monthly Report
                     </span>
                </div>
            </div>

            <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
                <div style="background:white; border:1px solid #e2e8f0; padding:12px; border-radius:10px; box-shadow:var(--shadow-sm); display:flex; align-items:center; gap:12px;">
                    <label style="font-size:0.8rem; font-weight:700; color:#64748b; text-transform:uppercase;">Select Period:</label>
                    <select id="modal-period-select" onchange="openExecutiveDetails('${name.replace(/'/g, "\\'")}', this.value, '${filterMode}')" 
                        style="border:2px solid #4f46e5; color:#4f46e5; font-weight:700; padding:6px 12px; border-radius:6px; background:white; font-size:0.95rem; cursor:pointer; outline:none;">
                        ${availableMonths.map(m => `<option value="${m}" ${m === selectedKey ? 'selected' : ''}>${m}</option>`).join('')}
                    </select>
                </div>
                
                ${filterMode === 'DISBURSED_ONLY' ? `
                <button onclick="openMonthlySlipModal('${escapedName}')" class="btn" style="background:#4f46e5; color:white; padding:10px 20px; font-size:0.9rem; height:auto; border-radius:10px; box-shadow:var(--shadow-md);">
                    <i class="fas fa-file-invoice-dollar"></i> Generate Monthly Slip
                </button>` : ''}
            </div>
        </div>
        
        <div style="margin-bottom:1.5rem; padding:10px; background:#0f172a; color:white; border-radius:8px; font-size:0.85rem; display:flex; gap:15px; align-items:center;">
             <i class="fas ${filterMode === 'DISBURSED_ONLY' ? 'fa-wallet' : 'fa-folder-open'}" style="color:${filterMode === 'DISBURSED_ONLY' ? '#10b981' : '#38bdf8'};"></i>
             Showing <b>${filterMode === 'DISBURSED_ONLY' ? 'ONLY Disbursed' : 'ALL Pipeline'}</b> cases for <b>${selectedKey}</b>.
        </div>
        
        <!-- Stats Grid -->
        <div style="display:grid; grid-template-columns: repeat(4, 1fr); gap:1rem; margin-bottom:1.5rem;">
            <div style="background:#f8fafc; padding:1rem; border-radius:8px;">
                <small>Volume</small>
                <div style="font-size:1.1rem; font-weight:bold;">${formatCurrency(totalVolume)}</div>
            </div>
            <div style="background:#f8fafc; padding:1rem; border-radius:8px;">
                <small>Files</small>
                <div style="font-size:1.1rem; font-weight:bold;">${execLoans.length}</div>
            </div>
            <div style="background:#ecfdf5; padding:1rem; border-radius:8px; border:1px solid #d1fae5;">
                <small style="color:#047857;">Est. Revenue (Payout)</small>
                <div style="font-size:1.1rem; font-weight:bold; color:#047857;">${formatCurrency(totalRevenue)}</div>
            </div>
            <div style="background:${netRevenue >= 0 ? '#ecfdf5' : '#fef2f2'}; padding:1rem; border-radius:8px; border:1px solid ${netRevenue >= 0 ? '#d1fae5' : '#fecaca'};">
                <small style="color:${netRevenue >= 0 ? '#047857' : '#b91c1c'};">Net Revenue (Adjusted)</small>
                <div style="font-size:1.1rem; font-weight:bold; color:${netRevenue >= 0 ? '#047857' : '#b91c1c'};">${formatCurrency(netRevenue)}</div>
            </div>
        </div>

        <!-- Bank Payout Config & Adjustments (SHOW ONLY IN DISBURSED_ONLY MODE) -->
        ${filterMode === 'DISBURSED_ONLY' ? `
        <div style="background:#f0f9ff; padding:1rem; border-radius:8px; border:1px solid #bae6fd; margin-bottom:1.5rem;">
            <h4 style="margin-top:0; color:#0369a1; margin-bottom:0.5rem; font-size:0.9rem;">Set Payouts by Bank</h4>
            <div style="display:flex; flex-wrap:wrap; gap:1rem; align-items:end;">
                ${uniqueBanks.map((bank, index) => {
        const configKey = `${name}|||${bank}`;
        let savedPct = bankPayoutConfig[configKey];
        return `
                    <div style="display:flex; align-items:center; gap:4px;">
                        <div>
                            <label style="display:block; font-size:0.75rem; color:#0369a1; margin-bottom:2px;">
                                ${bank}${savedPct !== undefined ? ` <span style="color:#059669; font-size:0.7rem;">(${savedPct}% saved ✓)</span>` : ''}
                            </label>
                            <input type="number" step="0.01" id="payout-bank-${index}" placeholder="%" value="${savedPct !== undefined ? savedPct : ''}"
                                style="padding:4px; border:1px solid #7dd3fc; border-radius:4px; width:70px;">
                        </div>
                        <input type="hidden" id="name-bank-${index}" value="${bank}">
                    </div>
                `}).join('')}
                <button onclick="updateBankPayouts('${escapedName}', '${selectedKey}', '${filterMode}')" class="btn btn-primary" style="padding:4px 12px; font-size:0.85rem; background-color:#0284c7;">Apply to All</button>
            </div>
        </div>
        ` : ''}

        <!-- ===== DEBIT / CREDIT ADJUSTMENTS PANEL (ONLY IN DISBURSED MODE) ===== -->
        ${filterMode === 'DISBURSED_ONLY' ? `
        <div style="background:#fffbeb; border:1px solid #fde68a; border-radius:8px; padding:1rem; margin-bottom:1.5rem;">
            <h4 style="margin-top:0; margin-bottom:1rem; color:#92400e; font-size:0.95rem; display:flex; align-items:center; gap:6px;">
                <i class="fas fa-exchange-alt"></i> Debit / Credit Adjustments
            </h4>

            <!-- Summary pills -->
            <div style="display:flex; gap:10px; margin-bottom:1rem; flex-wrap:wrap;">
                <span style="background:#fee2e2; color:#991b1b; padding:4px 12px; border-radius:999px; font-size:0.8rem; font-weight:700;">
                    <i class="fas fa-arrow-up"></i> Total Debit: ${formatCurrency(totalDebits)}
                </span>
                <span style="background:#dcfce7; color:#166534; padding:4px 12px; border-radius:999px; font-size:0.8rem; font-weight:700;">
                    <i class="fas fa-arrow-down"></i> Total Credit: ${formatCurrency(totalCredits)}
                </span>
                <span style="background:${netRevenue >= 0 ? '#d1fae5' : '#fee2e2'}; color:${netRevenue >= 0 ? '#065f46' : '#991b1b'}; padding:4px 12px; border-radius:999px; font-size:0.8rem; font-weight:700;">
                    <i class="fas fa-balance-scale"></i> Net: ${formatCurrency(netRevenue)}
                </span>
            </div>

            <!-- Add Adjustment Form -->
            <div style="display:grid; grid-template-columns: auto 1fr 1fr 2fr auto; gap:8px; align-items:end; margin-bottom:1rem; flex-wrap:wrap;">
                <div>
                    <label style="display:block; font-size:0.75rem; color:#78350f; margin-bottom:3px; font-weight:600;">Date</label>
                    <input type="date" id="adj-date" value="${today}" style="padding:6px 8px; border:1px solid #fcd34d; border-radius:6px; font-size:0.85rem; width:130px;">
                </div>
                <div>
                    <label style="display:block; font-size:0.75rem; color:#78350f; margin-bottom:3px; font-weight:600;">Type</label>
                    <select id="adj-type" style="padding:6px 8px; border:1px solid #fcd34d; border-radius:6px; font-size:0.85rem; width:100%;">
                        <option value="Debit">💸 Debit</option>
                        <option value="Credit">💰 Credit</option>
                    </select>
                </div>
                <div>
                    <label style="display:block; font-size:0.75rem; color:#78350f; margin-bottom:3px; font-weight:600;">Amount (₹)</label>
                    <input type="number" id="adj-amount" placeholder="0" min="0" step="0.01" style="padding:6px 8px; border:1px solid #fcd34d; border-radius:6px; font-size:0.85rem; width:100%;">
                </div>
                <div>
                    <label style="display:block; font-size:0.75rem; color:#78350f; margin-bottom:3px; font-weight:600;">Note</label>
                    <input type="text" id="adj-reason" placeholder="Reason..." style="padding:6px 8px; border:1px solid #fcd34d; border-radius:6px; font-size:0.85rem; width:100%;">
                </div>
                <div>
                    <button onclick="addExecutiveAdjustment('${escapedName}', '${selectedKey}', '${filterMode}')" class="btn btn-primary" style="padding:6px 16px; font-size:0.85rem; background:#d97706; border:none; height:36px;">
                        <i class="fas fa-plus"></i>
                    </button>
                </div>
            </div>

            <!-- Adjustments List -->
            ${execAdj.length > 0 ? `
            <div style="overflow-x:auto;">
                <table style="font-size:0.82rem; width:100%; border-collapse:collapse;">
                    <thead>
                        <tr style="background:#fef3c7;">
                            <th style="padding:6px 10px; text-align:left;">Date</th>
                            <th style="padding:6px 10px; text-align:left;">Type</th>
                            <th style="padding:6px 10px; text-align:left;">Amt</th>
                            <th style="padding:6px 10px; text-align:left;">Reason</th>
                            <th style="padding:6px 10px; text-align:center;">Action</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${execAdj.map(adj => `
                        <tr style="border-bottom:1px solid #fef3c7;">
                            <td style="padding:6px 10px;">${formatDate(adj.date)}</td>
                            <td><span class="adj-badge adj-badge-${adj.type.toLowerCase()}">${adj.type}</span></td>
                            <td style="font-weight:600; color:${adj.type === 'Debit' ? '#b91c1c' : '#065f46'};">${formatCurrency(adj.amount)}</td>
                            <td style="color:#6b7280;">${adj.reason || '-'}</td>
                            <td style="text-align:center;">
                                <button onclick="deleteExecutiveAdjustment('${adj.id}', '${escapedName}', '${selectedKey}', '${filterMode}')" style="background:#fee2e2; color:#b91c1c; border:none; border-radius:4px; padding:3px 8px;">
                                    <i class="fas fa-trash"></i>
                                </button>
                            </td>
                        </tr>`).join('')}
                    </tbody>
                </table>
            </div>
            ` : ''}
        </div>
        ` : ''}
        <!-- ===== END ADJUSTMENTS PANEL ===== -->

        <div style="overflow-x:auto;">
            <table style="font-size:0.85rem; width:100%;">
                <thead>
                    <tr>
                        <th>Entry Date</th>
                        ${filterMode === 'DISBURSED_ONLY' ? `<th>Disb. Date</th>` : ''}
                        <th>Customer</th>
                        <th>Bank</th>
                        <th>Status</th>
                        <th>Amount</th>
                        ${filterMode === 'DISBURSED_ONLY' ? `<th>Payout %</th><th>Revenue</th>` : ''}
                    </tr>
                </thead>
                <tbody>
                    ${execLoans.map(l => {
            const payout = l.payoutPercent || 0;
            const revenue = l.status === 'Disbursed' ? (l.amount * payout / 100) : 0;
            return `
                        <tr>
                            <td>${formatDate(l.date)}</td>
                            ${filterMode === 'DISBURSED_ONLY' ? `<td><span style="color:#059669; font-weight:bold;">${formatDate(l.disbursementDate) || '-'}</span></td>` : ''}
                            <td>${sanitizeHTML(l.customerName)}</td>
                            <td>${l.bankName || '-'}</td>
                            <td><span class="status-badge ${getStatusClass(l.status)}">${l.status}</span></td>
                            <td class="amount">${formatCurrency(l.amount)}</td>
                            ${filterMode === 'DISBURSED_ONLY' ? `
                            <td>
                                <input type="number" step="0.01" min="0" value="${payout || ''}" style="width:50px;" onchange="updateLoanPayout('${l.id}', this.value, '${escapedName}', '${selectedKey}', '${filterMode}')">
                            </td>
                            <td style="font-weight:bold; color:#059669;">${formatCurrency(revenue)}</td>
                            ` : ''}
                        </tr>`;
        }).join('')}
                </tbody>
            </table>
        </div>
    `;

    document.getElementById('modal-content').innerHTML = content;
    document.getElementById('exec-modal').classList.add('open');
};

window.updateBankPayouts = (execName, selectedKey, filterMode) => {
    const execLoans = loans.filter(l => l.executiveName === execName);
    const uniqueBanks = [...new Set(execLoans.map(l => l.bankName).filter(Boolean))];
    let updatedCount = 0;

    // Read and save input values into bankPayoutConfig (persistent across re-renders)
    const bankValues = {};
    uniqueBanks.forEach((bank, index) => {
        const inputVal = document.getElementById(`payout-bank-${index}`)?.value;
        if (inputVal !== '' && inputVal !== null && inputVal !== undefined) {
            const percent = parseFloat(inputVal);
            if (!isNaN(percent)) {
                bankValues[bank] = percent;
                // ✅ Save to global config so inputs stay filled and new loans auto-inherit
                bankPayoutConfig[`${execName}|||${bank}`] = percent;

                // Update in-memory loans for this exec+bank (Disbursed only)
                loans.forEach(l => {
                    if (l.executiveName === execName && l.bankName === bank && l.status === 'Disbursed') {
                        l.payoutPercent = percent;
                        updatedCount++;
                    }
                });
            }
        }
    });

    if (updatedCount > 0) {
        // Push all updates to Firestore
        const updatePromises = [];
        loans.forEach(l => {
            if (l.executiveName === execName && bankValues[l.bankName] !== undefined && l.status === 'Disbursed') {
                const loanRef = doc(db, COLLECTION_NAME, l.id);
                updatePromises.push(updateDoc(loanRef, { payoutPercent: bankValues[l.bankName] }));
            }
        });

        Promise.all(updatePromises).then(() => {
            showToast(`Updated ${updatedCount} payout(s) for ${execName}.`);
            openExecutiveDetails(execName, selectedKey, filterMode); 
        }).catch(error => {
            console.error("Error updating bank payouts: ", error);
            alert("Failed to update payouts in Firebase. Error: " + error.message);
        });
    } else {
        alert("Please enter a percentage for at least one bank.");
    }
};


window.updateLoanPayout = (loanId, percent, execName, selectedKey, filterMode) => {
    const loanRef = doc(db, COLLECTION_NAME, loanId);
    updateDoc(loanRef, { payoutPercent: parseFloat(percent) || 0 })
        .then(() => {
            openExecutiveDetails(execName, selectedKey, filterMode);
        })
        .catch(error => {
            console.error("Error updating payout: ", error);
        });
};

window.closeModal = () => {
    document.getElementById('exec-modal').classList.remove('open');
};

// Add a Debit or Credit adjustment for an executive
window.addExecutiveAdjustment = async (execName, selectedKey, filterMode) => {
    const date = document.getElementById('adj-date')?.value;
    const type = document.getElementById('adj-type')?.value;
    const amount = parseFloat(document.getElementById('adj-amount')?.value);
    const reason = document.getElementById('adj-reason')?.value?.trim();

    if (!date) { showToast('Please select a date.'); return; }
    if (!amount || amount <= 0) { showToast('Please enter a valid amount.'); return; }

    const adjData = {
        executiveName: execName,
        type,         // 'Debit' or 'Credit'
        amount,
        reason: reason || '',
        date,
        createdAt: new Date().toISOString()
    };

    try {
        await addDoc(collection(db, 'ExecutiveAdjustments'), adjData);
        showToast(`${type} of ${formatCurrency(amount)} added for ${execName}.`);
        openExecutiveDetails(execName, selectedKey, filterMode);
    } catch (error) {
        console.error('Error adding adjustment:', error);
        showToast('Error saving adjustment. Check permissions.');
    }
};

// Delete an adjustment entry
window.deleteExecutiveAdjustment = (adjId, execName, selectedKey, filterMode) => {
    if (!confirm('Delete this adjustment?')) return;
    
    const delTask = async () => {
        try {
            await deleteDoc(doc(db, 'ExecutiveAdjustments', adjId));
            showToast('Adjustment deleted.');
            openExecutiveDetails(execName, selectedKey, filterMode);
        } catch (error) {
            console.error('Error deleting adjustment:', error);
            showToast('Error deleting adjustment. Check permissions.');
        }
    };
    delTask();
};

// --- AUTHENTICATION LOGIC ---

if (loginForm) {
    console.log("Login form found");
    loginForm.addEventListener('submit', (e) => {
        e.preventDefault();
        const email = document.getElementById('login-email').value;
        const password = document.getElementById('login-password').value;
        console.log("Attempting email login...");

        signInWithEmailAndPassword(auth, email, password)
            .then(() => {
                console.log("Email login successful");
                authError.style.display = 'none';
            })
            .catch((error) => {
                console.error("Email login error:", error);
                authError.textContent = "Invalid email or password. " + error.message;
                authError.style.display = 'block';
            });
    });
}


if (googleLoginBtn) {
    console.log("Google login button found");
    googleLoginBtn.addEventListener('click', () => {
        console.log("Attempting Google login...");
        signInWithPopup(auth, googleProvider)
            .then(() => {
                console.log("Google login successful");
                authError.style.display = 'none';
            })
            .catch((error) => {
                console.error("Google login error:", error);
                authError.textContent = "Google sign-in failed. " + error.message;
                authError.style.display = 'block';
            });
    });
} else {
    console.error("Google login button NOT found!");
}

if (logoutBtn) {
    logoutBtn.addEventListener('click', () => {
        signOut(auth).then(() => {
            location.reload(); // Refresh to clear state
        });
    });
}

// Global listeners for Firestore subscriptions to close them on logout
let unsubscribeData = null;
let unsubscribeUsers = null;

// Monitor Auth State
onAuthStateChanged(auth, async (user) => {
    if (user) {
        console.log("Checking authorization for:", user.email);

        // ADMIN HARDCODE: Ensure you are always an admin
        const isAdmin = user.email === 'sharmashubham22657@gmail.com';

        try {
            // Check Users collection
            const userDoc = await getDocs(query(collection(db, 'Users')));
            authorizedUsers = userDoc.docs.map(d => ({ ...d.data(), id: d.id }));

            const userData = authorizedUsers.find(u => u.email.toLowerCase() === user.email.toLowerCase());

            if (!userData && !isAdmin) {
                alert("ACCESS DENIED: You are not authorized to access this system. Contact the administrator.");
                signOut(auth);
                return;
            }

            currentUserRole = isAdmin ? 'ADMIN' : (userData?.role || 'EXECUTIVE');
            assignedName = userData?.assignedName || (isAdmin ? 'Admin' : null);

            proceedWithLogin(user, isAdmin);

        } catch (error) {
            console.error("Auth Verification Error:", error);
            // Even if User collection check fails (e.g. permission error), the Hardcoded Admin must still work
            if (isAdmin) {
                currentUserRole = 'ADMIN';
                assignedName = 'Admin';
                proceedWithLogin(user, isAdmin);
            } else {
                alert("CRITICAL AUTH ERROR: " + error.message);
                signOut(auth);
            }
        }
    } else {
        document.body.classList.add('not-logged-in');
        document.body.classList.remove('logged-in');
        currentUserRole = null;
        assignedName = null;
        executiveAdjustments = [];
        if (unsubscribeData) {
            unsubscribeData();
            unsubscribeData = null;
        }
        if (unsubscribeUsers) {
            unsubscribeUsers();
            unsubscribeUsers = null;
        }
        if (unsubscribeAdjustments) {
            unsubscribeAdjustments();
            unsubscribeAdjustments = null;
        }
    }
});

const proceedWithLogin = (user, isAdmin) => {
    // Update Header Profile
    const profileName = document.getElementById('user-display-name');
    const profileRole = document.getElementById('user-display-role');
    const profileIcon = document.getElementById('user-display-icon');

    if (profileName) profileName.textContent = user.displayName || user.email.split('@')[0];
    if (profileRole) {
        if (currentUserRole === 'ADMIN') profileRole.textContent = 'Owner / Admin';
        else if (currentUserRole === 'BACK_OFFICE') profileRole.textContent = `Back Office (${assignedName || 'Main'})`;
        else profileRole.textContent = `Executive (${assignedName || 'Unassigned'})`;
    }
    if (profileIcon) {
        profileIcon.innerHTML = `<i class="fas ${currentUserRole === 'ADMIN' ? 'fa-user-shield' : 'fa-user'}"></i>`;
        profileIcon.style.background = currentUserRole === 'ADMIN'
            ? 'linear-gradient(135deg, var(--accent-color), var(--accent-hover))'
            : 'linear-gradient(135deg, var(--primary-color), var(--primary-light))';
    }

    document.body.classList.remove('not-logged-in');
    document.body.classList.add('logged-in');

    // UI Adjustments based on role
    const dashboardTab = document.querySelector('.nav-item[data-view="dashboard"]');
    if (currentUserRole !== 'ADMIN') {
        if (dashboardTab) dashboardTab.style.display = 'none';
    } else {
        if (dashboardTab) dashboardTab.style.display = 'flex';
        // Add Users tab if not exists
        if (!document.querySelector('.nav-item[data-view="users"]')) {
            const navLinks = document.querySelector('.nav-links');
            const usersTab = document.createElement('li');
            usersTab.className = 'nav-item';
            usersTab.setAttribute('data-view', 'users');
            usersTab.innerHTML = '<i class="fas fa-users-cog"></i> <span>Manage Users</span>';
            usersTab.addEventListener('click', () => {
                loadView('users');
            });
            navLinks.appendChild(usersTab);
        }
    }

    loadView('entry');

    // Start Firestore Subscription - ROLE BASED QUERY
    const colRef = collection(db, COLLECTION_NAME);
    let q = query(colRef, orderBy('date', 'desc'));

    if (currentUserRole === 'EXECUTIVE' && assignedName) {
        console.log("FIRESTORE: Applying filter for", assignedName);
        q = query(colRef, where('executiveName', '==', assignedName));
    }

    // Update Sidebar visibility for Admin tools
    const adminSidebar = document.getElementById('admin-actions-sidebar');
    if (adminSidebar) adminSidebar.style.display = (currentUserRole === 'ADMIN') ? 'block' : 'none';

    // Show Settings tab for Admin
    const settingsTab = document.getElementById('settings-tab');
    if (settingsTab) settingsTab.style.display = (currentUserRole === 'ADMIN') ? 'flex' : 'none';

    // Show Diagnostic button for Admin
    const adminDebugBtn = document.getElementById('admin-debug-btn');
    if (adminDebugBtn) adminDebugBtn.style.display = (currentUserRole === 'ADMIN') ? 'block' : 'none';

    // Cleanup previous subscription if any
    if (unsubscribeData) unsubscribeData();

    // Custom Columns & Order Subscription
    onSnapshot(doc(db, 'Settings', 'columns'), (docSnapshot) => {
        if (docSnapshot.exists()) {
            const data = docSnapshot.data();
            customColumns = data.customColumns || [];
            columnOrder = data.columnOrder || [];

            // Initialization: If columnOrder is empty, build it from system columns
            if (columnOrder.length === 0) {
                columnOrder = SYSTEM_COLUMNS.map(c => c.id);
                // Also set it in Firestore so it's persisted
                setDoc(doc(db, 'Settings', 'columns'), {
                    customColumns: customColumns,
                    columnOrder: columnOrder
                }, { merge: true });
            }

            console.log("Column configuration synced:", { customColumns, columnOrder });

            // Re-render certain views if active
            const activeNav = document.querySelector('.nav-item.active');
            if (activeNav) {
                const currentView = activeNav.getAttribute('data-view');
                if (currentView === 'entry') renderEntryPage();
                if (currentView === 'settings') renderSettingsPage();
            }
        } else {
            // Document doesn't exist, initialize it
            columnOrder = SYSTEM_COLUMNS.map(c => c.id);
            setDoc(doc(db, 'Settings', 'columns'), {
                customColumns: [],
                columnOrder: columnOrder
            });
        }
    });

    unsubscribeData = onSnapshot(q, (snapshot) => {
        loans = snapshot.docs.map(doc => ({
            ...doc.data(),
            id: doc.id
        }));

        loans.sort((a, b) => new Date(b.date) - new Date(a.date));
        renderTableRows();

        const activeNav = document.querySelector('.nav-item.active');
        if (activeNav) {
            const view = activeNav.getAttribute('data-view');
            if (view === 'dashboard' && currentUserRole === 'ADMIN') renderDashboardPage();
            if (view === 'users' && currentUserRole === 'ADMIN') renderUserManagement();
        }
    }, (error) => {
        console.error("Firestore Snapshot Error:", error.code, error.message);
        alert("REAL-TIME SYNC ERROR: " + error.message + "\nCode: " + error.code);
    });

    // Users Real-Time Sync
    if (currentUserRole === 'ADMIN') {
        if (unsubscribeUsers) unsubscribeUsers();
        unsubscribeUsers = onSnapshot(collection(db, 'Users'), (snapshot) => {
            authorizedUsers = snapshot.docs.map(doc => ({
                ...doc.data(),
                id: doc.id
            }));
            console.log("Users configuration synced:", authorizedUsers.length, "users");

            const activeNav = document.querySelector('.nav-item.active');
            if (activeNav) {
                const currentView = activeNav.getAttribute('data-view');
                updateNavActive(currentView); // Sync sidebar texts
                if (currentView === 'users') renderUserManagement();
                if (currentView === 'entry') renderEntryPage();
            }
        });

        // Executive Adjustments (Debit/Credit) - Admin only
        if (unsubscribeAdjustments) unsubscribeAdjustments();
        unsubscribeAdjustments = onSnapshot(
            query(collection(db, 'ExecutiveAdjustments'), orderBy('date', 'desc')),
            (snapshot) => {
                executiveAdjustments = snapshot.docs.map(d => ({ ...d.data(), id: d.id }));
                // Refresh dashboard if active
                const activeNav = document.querySelector('.nav-item.active');
                if (activeNav && activeNav.getAttribute('data-view') === 'dashboard') {
                    renderDashboardPage();
                }
            }
        );
    }
};

// User Management View
const renderUserManagement = () => {
    pageTitle.textContent = 'User Management';
    contentArea.innerHTML = `
        <div class="card">
            <h3>Authorize New User</h3>
            <form id="add-user-form" style="margin-top:1.5rem;">
                <div class="form-row">
                    <div>
                        <label>Google Email</label>
                        <input type="email" id="newUserEmail" class="form-control" placeholder="user@gmail.com" required>
                    </div>
                    <div>
                        <label>Assigned Name (Must match Daily Entry exactly)</label>
                        <input type="text" id="newUserAssignedName" class="form-control" placeholder="Executive Name" required>
                    </div>
                    <div>
                        <label>Role</label>
                        <select id="newUserRole" class="form-control">
                            <option value="EXECUTIVE">Executive (Own files, read only)</option>
                            <option value="BACK_OFFICE">Back Office (All files, can add)</option>
                            <option value="ADMIN">Admin (Sees and does everything)</option>
                        </select>
                    </div>
                    <div>
                        <label>Mobile Number (WhatsApp)</label>
                        <input type="tel" id="newUserPhone" class="form-control" placeholder="e.g. 9876543210" required>
                    </div>
                </div>
                <button type="submit" class="btn btn-primary" style="margin-top:1rem;"><i class="fas fa-user-plus"></i> Grant Access</button>
            </form>
        </div>

        <div class="card">
            <h3 style="margin-bottom:1.5rem;">Authorized Users</h3>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Email</th>
                            <th>Assigned Name</th>
                            <th>Mobile</th>
                            <th>Role</th>
                            <th>Action</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${authorizedUsers.map(user => {
        const rowBg = user.role === 'ADMIN' ? '#f5f3ff' : (user.role === 'BACK_OFFICE' ? '#fffbeb' : 'transparent');
        return `
                            <tr style="background-color: ${rowBg}; transition: background-color 0.3s ease;">
                                <td>${sanitizeHTML(user.email)}</td>
                                <td>${sanitizeHTML(user.assignedName)}</td>
                                <td style="font-family:monospace; color:#4f46e5; font-weight:600;">${sanitizeHTML(user.mobile || '-')}</td>
                                <td>
                                    ${user.email === 'sharmashubham22657@gmail.com' ? `
                                        <span class="status-badge status-approved" style="background:#eef2ff; color:#4f46e5; border:1px solid #c7d2fe;">ADMIN</span>
                                    ` : `
                                        <select class="form-control" 
                                            style="padding:0.25rem; font-size:0.85rem; font-weight:600; 
                                            border:1px solid ${user.role === 'ADMIN' ? '#c7d2fe' : (user.role === 'BACK_OFFICE' ? '#fde68a' : '#e2e8f0')};
                                            background-color: white;
                                            color: ${user.role === 'ADMIN' ? '#4f46e5' : (user.role === 'BACK_OFFICE' ? '#b45309' : '#475569')};" 
                                            onchange="updateUserRole('${user.id}', this.value)">
                                            <option value="EXECUTIVE" ${user.role === 'EXECUTIVE' ? 'selected' : ''}>EXECUTIVE</option>
                                            <option value="BACK_OFFICE" ${user.role === 'BACK_OFFICE' ? 'selected' : ''}>BACK_OFFICE</option>
                                            <option value="ADMIN" ${user.role === 'ADMIN' ? 'selected' : ''}>ADMIN</option>
                                        </select>
                                    `}
                                </td>
                                <td>
                                    ${user.email === 'sharmashubham22657@gmail.com' ? '<span style="color:var(--text-secondary)">System Owner</span>' : `
                                        <button onclick="deleteUser('${user.id}')" class="btn btn-danger" style="padding:0.4rem 0.8rem; font-size:0.8rem;">
                                            <i class="fas fa-trash"></i> Revoke Access
                                        </button>
                                    `}
                                </td>
                            </tr>
                        `;
    }).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `;

    document.getElementById('add-user-form').addEventListener('submit', async (e) => {
        e.preventDefault();
        const email = document.getElementById('newUserEmail').value.trim().toLowerCase();
        const assignedName = document.getElementById('newUserAssignedName').value.trim();
        const role = document.getElementById('newUserRole').value;

        try {
            // Use setDoc with email as ID for better rules management
            const phone = document.getElementById('newUserPhone').value.trim();
            const userRef = doc(db, 'Users', email);
            await setDoc(userRef, {
                email,
                assignedName,
                role,
                mobile: phone,
                createdAt: new Date().toISOString()
            });
            showToast("Access granted successfully.");
            // Refresh users list
            const userDoc = await getDocs(query(collection(db, 'Users')));
            authorizedUsers = userDoc.docs.map(d => ({ ...d.data(), id: d.id }));
            renderUserManagement();
        } catch (error) {
            alert("Error granting access: " + error.message);
        }
    });
};

window.updateUserRole = async (userId, newRole) => {
    try {
        const userRef = doc(db, 'Users', userId);
        await updateDoc(userRef, { role: newRole });
        showToast(`User role updated to ${newRole}`);

        // Update local state and re-render only if needed (usually handled by next view load)
        const user = authorizedUsers.find(u => u.id === userId);
        if (user) user.role = newRole;
    } catch (error) {
        console.error("Update Role Error:", error);
        alert("Failed to update user role: " + error.message);
        renderUserManagement(); // Revert UI on error
    }
};

window.deleteUser = async (id) => {
    if (!id) {
        alert("Error: No User ID provided for deletion.");
        return;
    }

    showConfirm(`Revoke access and delete permissions for ${id}?`, async () => {
        try {
            // Optimistic Update
            authorizedUsers = authorizedUsers.filter(u => u.id !== id);
            renderUserManagement();

            await deleteDoc(doc(db, 'Users', id));
            showToast("User access revoked.");
        } catch (error) {
            console.error("Delete User Error:", error);
            // Re-fetch to revert correctly
            const userDoc = await getDocs(query(collection(db, 'Users')));
            authorizedUsers = userDoc.docs.map(d => ({ ...d.data(), id: d.id }));
            renderUserManagement();
            alert(`SERVER REJECTION: ${error.message}. You might not have permission to modify users.`);
        }
    });
};

// --- VIEW: SETTINGS ---
const renderSettingsPage = () => {
    pageTitle.textContent = 'System Settings';

    // Helper to get Label
    const getLabel = (key) => {
        const sys = SYSTEM_COLUMNS.find(c => c.id === key);
        return sys ? sys.label : key;
    };

    contentArea.innerHTML = `
        <div class="card">
            <h3>Table Column Management</h3>
            <p style="color:var(--text-secondary); margin-top:0.5rem; margin-bottom:1.5rem; font-size:0.9rem;">
                Manage all table columns. <strong>Drag rows</strong> to reorder them, or add new custom columns below.
            </p>
            
            <form id="add-column-form" style="margin-bottom:2rem; display:flex; gap:1rem; align-items:flex-end;">
                <div style="flex:1;">
                    <label style="display:block; margin-bottom:0.25rem; font-weight:600; font-size:0.85rem; color:#475569;">New Custom Column Name</label>
                    <input type="text" id="newColumnName" class="form-control" placeholder="e.g. Reference No, Aadhar, etc." required>
                </div>
                <button type="submit" class="btn btn-primary">
                    <i class="fas fa-plus"></i> Add Column
                </button>
            </form>

            <div class="table-container">
                <table id="settings-columns-table" style="font-size: 0.9rem;">
                    <thead>
                        <tr>
                            <th style="width: 40px;"></th>
                            <th style="width: 50px;">No.</th>
                            <th>Column Name</th>
                            <th>Type</th>
                            <th>Action</th>
                        </tr>
                    </thead>
                    <tbody id="draggable-tbody">
                        ${columnOrder.map((key, index) => {
        const isSystem = SYSTEM_COLUMNS.some(c => c.id === key);
        return `
                            <tr draggable="false" data-index="${index}" class="draggable-row">
                                <td class="drag-handle" style="color: #94a3b8; text-align: center; cursor: grab;"><i class="fas fa-grip-vertical"></i></td>
                                <td style="text-align:center; color:#64748b; font-weight:600;">${index + 1}</td>
                                <td style="font-weight:600;">${getLabel(key)}</td>
                                <td><span style="font-size:0.75rem; color:#64748b; font-weight:600;">${isSystem ? 'SYSTEM' : 'CUSTOM'}</span></td>
                                <td>
                                    ${!isSystem ? `
                                    <button onclick="deleteCustomColumn('${key}')" class="btn btn-danger" style="padding:0.4rem 0.8rem; font-size:0.8rem;">
                                        <i class="fas fa-trash"></i> Remove
                                    </button>
                                    ` : '<span style="font-size:0.7rem; color:#94a3b8;">System Field</span>'}
                                </td>
                            </tr>
                            `;
    }).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `;

    // Drag and Drop implementation with Handle restriction
    const tbody = document.getElementById('draggable-tbody');
    let draggedItem = null;

    // Only enable draggable when clicking the handle
    tbody.addEventListener('mousedown', (e) => {
        const handle = e.target.closest('.drag-handle');
        if (handle) {
            const row = handle.closest('tr');
            row.draggable = true;
        }
    });

    tbody.addEventListener('dragstart', (e) => {
        draggedItem = e.target.closest('tr');
        if (!draggedItem.draggable) {
            e.preventDefault();
            return;
        }
        e.dataTransfer.effectAllowed = 'move';
        draggedItem.classList.add('dragging');
        draggedItem.style.opacity = '0.4';
    });

    tbody.addEventListener('dragend', (e) => {
        if (draggedItem) {
            draggedItem.style.opacity = '1';
            draggedItem.classList.remove('dragging');
            draggedItem.draggable = false; // Disable until next handle mousedown

            // Final save after reorder
            const rows = Array.from(tbody.querySelectorAll('tr'));
            const newOrder = rows.map(tr => columnOrder[parseInt(tr.dataset.index)]);

            if (JSON.stringify(newOrder) !== JSON.stringify(columnOrder)) {
                saveNewColumnOrder(newOrder);
            }
        }
        draggedItem = null;
    });

    tbody.addEventListener('dragover', (e) => {
        e.preventDefault();
        const target = e.target.closest('tr');
        if (target && target !== draggedItem) {
            const rect = target.getBoundingClientRect();
            const next = (e.clientY - rect.top) / (rect.bottom - rect.top) > 0.5;
            tbody.insertBefore(draggedItem, next ? target.nextSibling : target);
        }
    });

    const saveNewColumnOrder = async (newOrder) => {
        try {
            await setDoc(doc(db, 'Settings', 'columns'), {
                customColumns: customColumns,
                columnOrder: newOrder
            }, { merge: true });
            showToast("Column order updated!");
        } catch (error) {
            console.error("Save Order Error:", error);
            alert("Failed to save new order: " + error.message);
        }
    };

    document.getElementById('add-column-form').addEventListener('submit', async (e) => {
        e.preventDefault();
        const input = document.getElementById('newColumnName');
        const colName = input.value.trim();

        if (!colName) return;
        if (customColumns.includes(colName) || SYSTEM_COLUMNS.some(c => c.id === colName)) {
            alert("This column name already exists.");
            return;
        }

        const newCustom = [...customColumns, colName];
        const newOrder = [...columnOrder, colName];

        try {
            await setDoc(doc(db, 'Settings', 'columns'), {
                customColumns: newCustom,
                columnOrder: newOrder
            });
            showToast("Column added successfully!");
        } catch (error) {
            console.error("Error adding column:", error);
            alert("Failed to add column: " + error.message);
        }
    });
};

window.moveColumn = async (index, direction) => {
    const newOrder = [...columnOrder];
    const newIndex = index + direction;

    if (newIndex < 0 || newIndex >= newOrder.length) return;

    // Swap columns
    [newOrder[index], newOrder[newIndex]] = [newOrder[newIndex], newOrder[index]];

    try {
        await setDoc(doc(db, 'Settings', 'columns'), {
            customColumns: customColumns,
            columnOrder: newOrder
        }, { merge: true });
        showToast("Column order updated!");
    } catch (error) {
        console.error("Error moving column:", error);
        alert("Failed to update column order: " + error.message);
    }
};

window.deleteCustomColumn = async (colName) => {
    showConfirm(`Remove the column "${colName}"? Data stored in this column will not be deleted but will be hidden.`, async () => {
        const newCustom = customColumns.filter(c => c !== colName);
        const newOrder = columnOrder.filter(c => c !== colName);
        try {
            await setDoc(doc(db, 'Settings', 'columns'), {
                customColumns: newCustom,
                columnOrder: newOrder
            });
            showToast("Column removed.");
        } catch (error) {
            console.error("Error removing column:", error);
            alert(`SERVER REJECTION: ${error.message}. Permissions check failed for Settings collection.`);
        }
    });
};




// Emergency Admin Switch (Use only if database permissions fail)
window.debugAppStatus = () => {
    console.log("--- APP STATUS ---");
    console.log("Logged In:", !!auth.currentUser);
    console.log("User Email:", auth.currentUser?.email);
    console.log("Role:", currentUserRole);
    console.log("Assigned Name:", assignedName);
    console.log("Loans Count:", loans.length);
    console.log("Authorized Users:", authorizedUsers);

    alert(`Status Check: \nEmail: ${auth.currentUser?.email} \nRole: ${currentUserRole} \nFiles Loaded: ${loans.length} `);
};

// Tool to test deletion permission specifically
window.testDeletePermission = async (loanId) => {
    if (!loanId) {
        alert("Please provide a Loan ID to test.");
        return;
    }
    const loanRef = doc(db, COLLECTION_NAME, loanId);
    try {
        // We try to update a dummy field to test write permission without deleting
        await updateDoc(loanRef, { _lastTest: new Date().toISOString() });
        alert("PERMISSION CHECK: You HAVE write/delete access for this file.");
    } catch (error) {
        alert("PERMISSION CHECK: ACCESS DENIED.\nReason: " + error.message + "\n\nThis confirms your Firestore Security Rules are blocking you.");
    }
};
