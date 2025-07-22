// =============================
// 1. GLOBAL VARIABLES
// =============================
let workbook = null;               // Workbook loaded from Excel
let sheetName = '';                // Active sheet name (usually first)
let templateFileName = 'template.xlsx';  // Default export name

// =============================
// 2. IMPORT EXCEL TEMPLATE FILE
// =============================
document.querySelector('.import-file').addEventListener('change', function (e) {
    const file = e.target.files[0];
    templateFileName = file.name;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = e.target.result;
        workbook = XLSX.read(data, { type: 'binary' });

        sheetName = workbook.SheetNames[0]; // Capture first sheet name

        alert("Template loaded successfully.");
    };

    reader.readAsBinaryString(file);
});

// =============================
// 3. DOWNLOAD UPDATED WORKBOOK
// =============================
document.getElementById('downloadBtn').addEventListener('click', function () {
    if (!workbook) {
        alert("No workbook loaded.");
        return;
    }

    // Export as new file
    XLSX.writeFile(workbook, 'LCR_Export.xlsx');
});

// =============================
// 4. ICON NAVBAR ACTIVE STATES
// =============================
document.addEventListener('DOMContentLoaded', function () {
    const sectionContainer = document.querySelector('.section-content');
    const sections = document.querySelectorAll('section');
    const navLinks = document.querySelectorAll('.icon-navbar a');

    sectionContainer.addEventListener('scroll', () => {
        let current = '';

        sections.forEach(section => {
            const sectionTop = section.offsetTop;
            const sectionHeight = section.offsetHeight;

            if (
                sectionContainer.scrollTop >= sectionTop - 100 &&
                sectionContainer.scrollTop < sectionTop + sectionHeight - 100
            ) {
                current = section.getAttribute('id');
            }
        });

        navLinks.forEach(link => {
            link.classList.remove('active', 'active-home', 'active-import', 'active-form', 'active-export');

            const href = link.getAttribute('href').substring(1);
            if (href === current) {
                link.classList.add('active');

                switch (href) {
                    case 'home':
                        link.classList.add('active-home');
                        break;
                    case 'import':
                        link.classList.add('active-import');
                        break;
                    case 'form':
                        link.classList.add('active-form');
                        break;
                    case 'export':
                        link.classList.add('active-export');
                        break;
                }
            }
        });
    });
});

// =============================
// 5. CLEAR FORM BUTTON
// =============================
document.getElementById('clearBtn').addEventListener('click', () => {
    const inputs = document.querySelectorAll('.form-input');
    inputs.forEach(input => {
        if (!input.hasAttribute('readonly')) {
            input.value = '';
        }
    });

    // Reset defaults
    document.getElementById('documentNo').value = 'LCR-2025-001';
    document.getElementById('registerNo').value = 'REG-145-A';
    document.getElementById('categorySelector').selectedIndex = 0;
    document.getElementById('timeSpent').value = '00:35:12 Automatically Calculated';
});

// =============================
// 6. REMOVE FILE BUTTON
// =============================
const fileInput = document.querySelector('.import-file');
const removeFileBtn = document.getElementById('removeFileBtn');

removeFileBtn.addEventListener('click', function () {
    fileInput.value = ''; // Reset file input
    workbook = null;      // Clear workbook reference
    alert("File removed.");
});

// =============================
// 7. CALCULATE TIME SPENT
// =============================
function calculateTimeSpent() {
    const timeReceived = document.getElementById('timeReceived').value;
    const actedUpon = document.getElementById('actedUpon').value;

    if (timeReceived && actedUpon) {
        const [hr1, min1, sec1] = timeReceived.split(':').map(Number);
        const [hr2, min2, sec2] = actedUpon.split(':').map(Number);

        const start = new Date(0, 0, 0, hr1, min1, sec1);
        const end = new Date(0, 0, 0, hr2, min2, sec2);

        const diffMs = Math.abs(end - start);
        const diff = new Date(diffMs);

        const hh = String(diff.getUTCHours()).padStart(2, '0');
        const mm = String(diff.getUTCMinutes()).padStart(2, '0');
        const ss = String(diff.getUTCSeconds()).padStart(2, '0');

        document.getElementById('timeSpent').value = `${hh}:${mm}:${ss}`;
    }
}

document.getElementById('timeReceived').addEventListener('input', calculateTimeSpent);
document.getElementById('actedUpon').addEventListener('input', calculateTimeSpent);

// =============================
// 8. SAVE FORM DATA INTO WORKBOOK
// =============================
document.getElementById('saveBtn').addEventListener('click', () => {
    if (!workbook) {
        alert("No spreadsheet file imported.");
        return;
    }

    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Get form values
    const docNo = document.getElementById('documentNo').value;
    const regNo = document.getElementById('registerNo').value;
    const phase = document.getElementById('categorySelector').value;
    const total = document.getElementById('totalReceived').value;
    const time = document.getElementById('timeReceived').value;
    const acted = document.getElementById('actedUpon').value;
    const spent = document.getElementById('timeSpent').value;

    // Validation (optional)
    if (!docNo || !regNo || !total || !time || !acted || !spent) {
        alert("Please complete all form fields before saving.");
        return;
    }

    // Create row template
    const newRow = Array(16).fill('');
    newRow[0] = docNo;
    newRow[1] = regNo;

    const phaseIndex = {
        received: 2,
        researched: 6,
        recording: 10,
        release: 14
    };

    const colOffset = phaseIndex[phase];
    newRow[colOffset] = total;
    newRow[colOffset + 1] = time;
    newRow[colOffset + 2] = acted;
    newRow[colOffset + 3] = spent;

    // Append row after headers
    jsonData.push(newRow);

    // WARNING: This line resets formatting
    workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(jsonData);

    alert("Form data saved to workbook.");
});