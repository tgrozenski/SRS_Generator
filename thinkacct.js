document.addEventListener('DOMContentLoaded', () => {
    const balancesFileInput = document.getElementById('balancesFileInput');
    const sarFileInput = document.getElementById('sarFileInput');
    const balancesFileNameElement = document.getElementById('balancesFileName');
    const sarFileNameElement = document.getElementById('sarFileName');
    const processBtn = document.getElementById('processBtn');
    const statusElement = document.getElementById('status');
    const reportListElement = document.getElementById('reportList');
    const errorsSection = document.getElementById('errors-section');
    const errorList = document.getElementById('errorList');

    let selectedBalancesFile = null;
    let selectedSarFile = null;

    function updateButtonState() {
        processBtn.disabled = !selectedSarFile;
    }

    balancesFileInput.addEventListener('change', (event) => {
        if (event.target.files.length > 0) {
            selectedBalancesFile = event.target.files[0];
            balancesFileNameElement.textContent = `Selected file: ${selectedBalancesFile.name}`;
        } else {
            selectedBalancesFile = null;
            balancesFileNameElement.textContent = '';
        }
        updateButtonState();
    });

    sarFileInput.addEventListener('change', (event) => {
        if (event.target.files.length > 0) {
            selectedSarFile = event.target.files[0];
            sarFileNameElement.textContent = `Selected file: ${selectedSarFile.name}`;
        } else {
            selectedSarFile = null;
            sarFileNameElement.textContent = '';
        }
        updateButtonState();
    });

    processBtn.addEventListener('click', () => {
        if (!selectedSarFile) {
            alert("Please select the Attendance SAR file.");
            return;
        }

        // Reset UI
        processBtn.disabled = true;
        statusElement.textContent = 'Processing...';
        statusElement.classList.remove('text-red-600');
        reportListElement.innerHTML = '<p class="placeholder text-center text-gray-500">Generating report...</p>';
        errorList.innerHTML = '';
        errorsSection.classList.add('hidden');

        const balancesPromise = selectedBalancesFile
            ? new Promise((resolve, reject) => {
                Papa.parse(selectedBalancesFile, {
                    header: true,
                    skipEmptyLines: true,
                    complete: resolve,
                    error: reject
                });
              })
            : Promise.resolve({ data: [], errors: [], meta: { fields: [] } });

        const sarPromise = new Promise((resolve, reject) => {
            Papa.parse(selectedSarFile, {
                header: true,
                skipEmptyLines: true,
                complete: resolve,
                error: reject
            });
        });

        Promise.all([balancesPromise, sarPromise])
            .then(([balancesResults, sarResults]) => {
                // Validate headers
                validateInputs(balancesResults, sarResults);

                const balances = processBalances(balancesResults.data);
                const attendanceData = countAttendance(sarResults.data);
                const finalBalances = mergeData(balances, attendanceData);

                generateFinalCsv(finalBalances);

                statusElement.textContent = 'Processing complete!';
            })
            .catch(error => {
                console.error("Processing failed:", error);
                statusElement.textContent = `Error: ${error.message}`;
                statusElement.classList.add('text-red-600');
                errorsSection.classList.remove('hidden');
                const errorItem = document.createElement('div');
                errorItem.className = 'p-2';
                errorItem.textContent = error.message;
                errorList.appendChild(errorItem);
            })
            .finally(() => {
                processBtn.disabled = false;
            });
    });
    
    function validateInputs(balancesResults, sarResults) {
        const sarHeaders = sarResults.meta.fields;
        if (!sarHeaders.includes('Student ID')) {
            throw new Error("SAR file is missing the required 'Student ID' column.");
        }
        if (!sarHeaders.includes('Student: Full Name') && !sarHeaders.includes('Student: Last, First')) {
            throw new Error("SAR file must include at least one of 'Student: Full Name' or 'Student: Last, First' columns.");
        }

        if (selectedBalancesFile) {
            const balanceHeaders = balancesResults.meta.fields;
            if (!balanceHeaders.includes('Student ID')) {
                throw new Error("Balances file is missing the required 'Student ID' column.");
            }
             if (!balanceHeaders.includes('Student: Full Name')) {
                throw new Error("Balances file is missing the required 'Student: Full Name' column.");
            }
            if (!balanceHeaders.includes('Balance')) {
                throw new Error("Balances file is missing the required 'Balance' column.");
            }
        }
    }

    function processBalances(data) {
        const balances = new Map();
        if (!data) return balances;
        
        data.forEach(row => {
            const id = row['Student ID']?.trim();
            const name = row['Student: Full Name']?.trim();
            const balance = parseFloat(row['Balance']);

            if (id && name && !isNaN(balance)) {
                balances.set(id, { name, balance });
            }
        });
        return balances;
    }

    function countAttendance(data) {
        const attendance = new Map();
        data.forEach(row => {
            const id = row['Student ID']?.trim();
            if (id) {
                const existing = attendance.get(id) || { name: '', count: 0 };
                
                // Prioritize 'Student: Full Name' but fall back to 'Student: Last, First'
                const rawName = row['Student: Full Name'] || row['Student: Last, First'];
                const name = normalizeName(rawName);

                const updatedCount = existing.count + 1;
                
                // Store the name associated with the ID, preferably the one that isn't just a reversed version
                const storedName = existing.name && !existing.name.includes(',') ? existing.name : name;

                attendance.set(id, { name: storedName, count: updatedCount });
            }
        });
        return attendance;
    }
    
    function normalizeName(name) {
        if (!name) return '';
        let normalized = name.trim().toLowerCase();
        if (normalized.includes(',')) {
            const parts = normalized.split(',').map(p => p.trim());
            normalized = `${parts[1] || ''} ${parts[0] || ''}`.trim();
        }
        return normalized
            .split(' ')
            .map(word => word.charAt(0).toUpperCase() + word.slice(1))
            .join(' ');
    }

    function mergeData(balances, attendance) {
        const merged = new Map(balances);

        for (const [id, attendanceData] of attendance.entries()) {
            const { name: sarName, count } = attendanceData;
            const existingBalanceData = merged.get(id);

            if (existingBalanceData) {
                merged.set(id, {
                    name: existingBalanceData.name,
                    balance: existingBalanceData.balance + count
                });
            } else {
                merged.set(id, {
                    name: sarName,
                    balance: count
                });
            }
        }
        return merged;
    }

    function generateFinalCsv(finalBalances) {
        const dataForCsv = [];
        for (const [id, data] of finalBalances.entries()) {
            dataForCsv.push({
                'Student ID': id,
                'Student: Full Name': data.name,
                'Balance': data.balance
            });
        }

        dataForCsv.sort((a, b) => a['Student: Full Name'].localeCompare(b['Student: Full Name']));

        const csv = Papa.unparse(dataForCsv);
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        
        const today = new Date();
        const dateStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
        const filename = `Balances-${dateStr}.csv`;
        
        const reportItem = document.createElement('div');
        reportItem.className = 'report-item flex justify-between items-center p-3 border-b border-gray-200 last:border-b-0';
        reportItem.innerHTML = `
            <span class="font-medium text-gray-800">${filename}</span>
            <a href="${URL.createObjectURL(blob)}" download="${filename}" class="bg-green-500 text-white text-sm font-semibold py-1 px-3 rounded-md hover:bg-green-600 transition-colors">Download</a>
        `;
        
        reportListElement.innerHTML = '';
        reportListElement.appendChild(reportItem);
    }
});