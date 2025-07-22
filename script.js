// Import Excel File and Parse It

let workbook = null;

document.querySelector('.import-file').addEventListener('change', function(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = e.target.result;
        workbook = XLSX.read(data, { type: 'binary' });

        // Example: Show first sheet data in form
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        populateForm(jsonData);
    };

    reader.readAsBinaryString(file);
});


// Export to Excel 
document.getElementById('downloadBtn').addEventListener('click', function () {
    // Get input values
    const documentNo = document.getElementById('documentNo').value;
    const registerNo = document.getElementById('registerNo').value;
    const phase = document.getElementById('categorySelector').value;
    const totalReceived = document.getElementById('totalReceived').value;
    const timeReceived = document.getElementById('timeReceived').value;
    const actedUpon = document.getElementById('actedUpon').value;
    const timeSpent = document.getElementById('timeSpent').value;

    // Create sheet data (row-by-row)
    const sheetData = [
        ['Document No.', 'Register No.', 'Phase', 'Total Received', 'Time Received', 'Acted Upon', 'Time Spent'],
        [documentNo, registerNo, phase, totalReceived, timeReceived, actedUpon, timeSpent]
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'FormData');
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
// Select the file input and the remove button
const fileInput = document.querySelector('.import-file');
const removeFileBtn = document.getElementById('removeFileBtn');

// Add a click event listener to the Remove File button
removeFileBtn.addEventListener('click', function () {
  fileInput.value = ''; // Clears the selected file
});

// Replacement for populateForm() 
reader.onload = function(e) {
    const data = e.target.result;
    workbook = XLSX.read(data, { type: 'binary' });
};
