// Import Excel File and Parse It

let workbook = null;

document.querySelector('.import-file').addEventListener('change', function(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = e.target.result;
        workbook = XLSX.read(data, { type: 'binary' });
    };

    reader.readAsBinaryString(file);
});


// Export to Excel 
document.getElementById('downloadBtn').addEventListener('click', function () {
    if (!workbook) {
        alert("No workbook loaded.");
        return;
    }

    XLSX.writeFile(workbook, 'LCR_Export.xlsx');
});


// Function for Icon Active State 
document.addEventListener('DOMContentLoaded', function () {
    const sectionContainer = document.querySelector('.section-content');
    const sections = document.querySelectorAll('section');
    const navLinks = document.querySelectorAll('.icon-navbar a');

    sectionContainer.addEventListener('scroll', () => {
        let current = '';

        sections.forEach(section => {
            const sectionTop = section.offsetTop;
            const sectionHeight = section.offsetHeight;

            if (sectionContainer.scrollTop >= sectionTop - 100 &&
                sectionContainer.scrollTop < sectionTop + sectionHeight - 100) {
                current = section.getAttribute('id');
            }
        });

        navLinks.forEach(link => {
            link.classList.remove('active', 'active-home', 'active-import', 'active-form', 'active-export');

            const href = link.getAttribute('href').substring(1); // remove #
            if (href === current) {
                link.classList.add('active');

                // Custom active color class
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


// Clear Button Behavior 
document.getElementById('clearBtn').addEventListener('click', () => {
    const inputs = document.querySelectorAll('.form-input');
    inputs.forEach(input => {
        if (!input.hasAttribute('readonly')) {
            input.value = '';
        }
    });

    // Reset to defaults if needed
    document.getElementById('documentNo').value = 'LCR-2025-001';
    document.getElementById('registerNo').value = 'REG-145-A';
    document.getElementById('categorySelector').selectedIndex = 0;
    document.getElementById('timeSpent').value = '00:35:12 Automatically Calculated';
});


// Remove File Logic 
const fileInput = document.querySelector('.import-file');
const removeFileBtn = document.getElementById('removeFileBtn');

removeFileBtn.addEventListener('click', function () {
    fileInput.value = ''; // Clears the selected file
});


// Time Spent Automatic Calculation logic 
function calculateTimeSpent() {
    const timeReceived = document.getElementById('timeReceived').value;
    const actedUpon = document.getElementById('actedUpon').value;

    if (timeReceived && actedUpon) {
        const [hr1, min1, sec1] = timeReceived.split(':').map(Number);
        const [hr2, min2, sec2] = actedUpon.split(':').map(Number);

        const start = new Date(0, 0, 0, hr1, min1, sec1);
        const end = new Date(0, 0, 0, hr2, min2, sec2);

        let diff = new Date(end - start);
        if (end < start) diff = new Date(start - end); // handle negative time

        const hh = String(diff.getUTCHours()).padStart(2, '0');
        const mm = String(diff.getUTCMinutes()).padStart(2, '0');
        const ss = String(diff.getUTCSeconds()).padStart(2, '0');

        document.getElementById('timeSpent').value = `${hh}:${mm}:${ss}`;
    }
}

document.getElementById('timeReceived').addEventListener('input', calculateTimeSpent);
document.getElementById('actedUpon').addEventListener('input', calculateTimeSpent);


// Save Values Logic 
document.getElementById('saveBtn').addEventListener('click', () => {
    if (!workbook) {
        alert("No spreadsheet file imported.");
        return;
    }

    const sheetName = workbook.SheetNames[0]; // Save to first sheet
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Prepare data
    const docNo = document.getElementById('documentNo').value;
    const regNo = document.getElementById('registerNo').value;
    const phase = document.getElementById('categorySelector').value;
    const total = document.getElementById('totalReceived').value;
    const time = document.getElementById('timeReceived').value;
    const acted = document.getElementById('actedUpon').value;
    const spent = document.getElementById('timeSpent').value;

    // Define where to save in sheet
    const headers = jsonData[0]; // assume first row = headers
    let newRow = new Array(16).fill('');

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

    // Add to sheet
    jsonData.push(newRow);

    // Write back to worksheet
    const newSheet = XLSX.utils.aoa_to_sheet(jsonData);
    workbook.Sheets[sheetName] = newSheet;

    alert("Data saved to workbook in memory.");
});
