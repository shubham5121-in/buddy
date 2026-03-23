// Mock Sanitizer Function (Node compatible for testing)
const sanitizeHTML = (str) => {
    if (str === null || str === undefined) return '';
    return str.toString()
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
};

// Mock Data including a "Malicious" entry
const loans = [
    { id: 1, customerName: "Rahul Sharma", bankName: "HDFC Bank", amount: 500000, status: "Disbursed" },
    { id: 99, customerName: "<script>alert('XSS')</script>", bankName: "Hacker Bank", amount: 0, status: "Rejected" }
];

// The Exact Search Logic from app.js
const performSearch = (searchTerm) => {
    searchTerm = searchTerm.toString().toLowerCase().trim();

    return loans.filter(loan => {
        const safeStr = (val) => String(val || '').toLowerCase();
        return safeStr(loan.customerName).includes(searchTerm) ||
               safeStr(loan.bankName).includes(searchTerm);
    });
};

// Run Security Tests
console.log("--- Starting Security & Search Test ---");

// Test 1: Identify the malicious entry
const malicious = loans.find(l => l.id === 99);
console.log("Original Malicious Name:", malicious.customerName);

// Test 2: Sanitize the malicious entry
const sanitized = sanitizeHTML(malicious.customerName);
console.log("Sanitized Name (SAFE):", sanitized);

if (sanitized.includes("<script>")) {
    console.error("❌ SECURITY FAILED: Script tag was not escaped!");
} else {
    console.log("✅ SECURITY PASSED: Script tags are now harmless text.");
}

// Test 3: Search still works
const searchRes = performSearch("script");
console.log(`Search result for 'script': Found ${searchRes.length} matches.`);

console.log("--- Test Complete ---");
