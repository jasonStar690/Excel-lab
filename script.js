// DOM Elements
const uploadZone = document.getElementById('upload-zone');
const fileInput = document.getElementById('file-input');
const editorZone = document.getElementById('editor-zone');
const dataTable = document.getElementById('data-table');
const closeBtn = document.getElementById('close-btn');
const exportBtn = document.getElementById('export-btn');
const currentFilename = document.getElementById('current-filename');

// State
let workbook = null;
let originalFilename = 'data.xlsx';
let tableData = [];
let activeSchema = null;

// --- 1. Schema Definition System (Soft Rules for Logic) ---
const SCHEMAS = {
    'salary': {
        name: '工资表模版',
        columns: [
            { key: 'name', label: '姓名', type: 'text' },
            { key: 'base', label: '基本工资', type: 'number' },
            { key: 'bonus', label: '奖金', type: 'number' },
            {
                key: 'tax',
                label: '个税 (10%)',
                type: 'number',
                readOnly: true,
                formula: (row) => (parseNum(row.base) + parseNum(row.bonus)) * 0.1
            },
            {
                key: 'net',
                label: '实发工资',
                type: 'number',
                readOnly: true,
                formula: (row) => (parseNum(row.base) + parseNum(row.bonus)) - parseNum(row.tax)
            }
        ]
    }
};

const parseNum = (val) => {
    const n = parseFloat(val);
    return isNaN(n) ? 0 : n;
};

// --- Event Listeners ---
uploadZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', handleFileSelect);

// Drag & Drop
uploadZone.addEventListener('dragover', (e) => { e.preventDefault(); uploadZone.classList.add('dragover'); });
uploadZone.addEventListener('dragleave', () => { uploadZone.classList.remove('dragover'); });
uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadZone.classList.remove('dragover');
    if (e.dataTransfer.files.length) processFile(e.dataTransfer.files[0]);
});

// Demo Button (Restored for showcasing)
const demoBtn = document.createElement('button');
demoBtn.textContent = '加载工资演示模版';
demoBtn.className = 'btn btn-secondary';
demoBtn.style.marginTop = '1rem';
demoBtn.onclick = loadDemoTemplate;
uploadZone.querySelector('.upload-content').appendChild(demoBtn);

closeBtn.addEventListener('click', resetApp);
exportBtn.addEventListener('click', exportData);

function handleFileSelect(e) {
    if (e.target.files[0]) processFile(e.target.files[0]);
}

function loadDemoTemplate(e) {
    e.stopPropagation();
    originalFilename = 'salary_demo.xlsx';
    currentFilename.textContent = originalFilename;

    const dummyData = [
        { name: '张三', base: 5000, bonus: 1000 },
        { name: '李四', base: 8000, bonus: 2000 },
        { name: '王五', base: 12000, bonus: 500 }
    ];

    activeSchema = SCHEMAS['salary'];
    tableData = dummyData.map(row => recalculateRow(row));
    renderTable();
    showEditor();
}

async function processFile(file) {
    originalFilename = file.name;
    currentFilename.textContent = originalFilename;
    const data = await file.arrayBuffer();
    workbook = XLSX.read(data);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

    if (!jsonData || jsonData.length < 2) {
        alert("文件内容为空或格式不正确");
        return;
    }

    const headers = jsonData[0];
    const isSalary = headers.includes('姓名') && headers.includes('基本工资');

    if (isSalary) {
        activeSchema = SCHEMAS['salary'];
        tableData = jsonData.slice(1).map(rowArr => {
            const rowObj = {};
            activeSchema.columns.forEach((col, index) => {
                const headerIdx = headers.indexOf(col.label);
                if (headerIdx !== -1) {
                    rowObj[col.key] = rowArr[headerIdx];
                }
            });
            return recalculateRow(rowObj);
        });
    } else {
        // Raw Mode
        activeSchema = null;
        tableData = jsonData;
    }

    renderTable();
    showEditor();
}

function recalculateRow(row) {
    if (!activeSchema) return row;
    activeSchema.columns.forEach(col => {
        if (col.formula) {
            row[col.key] = col.formula(row);
        }
    });
    return row;
}

function renderTable() {
    dataTable.innerHTML = '';
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');
    const headerRow = document.createElement('tr');

    // Headers
    if (activeSchema) {
        activeSchema.columns.forEach(col => {
            const th = document.createElement('th');
            th.textContent = col.label;
            headerRow.appendChild(th);
        });
    } else {
        if (tableData.length > 0) {
            tableData[0].forEach(cell => {
                const th = document.createElement('th');
                th.textContent = cell;
                headerRow.appendChild(th);
            });
        }
    }
    thead.appendChild(headerRow);

    // Body
    const rowsToRender = activeSchema ? tableData : tableData.slice(1);

    rowsToRender.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');

        if (activeSchema) {
            activeSchema.columns.forEach(col => {
                const td = document.createElement('td');
                const val = row[col.key] !== undefined ? row[col.key] : '';

                if (typeof val === 'number' && !Number.isInteger(val)) {
                    td.innerText = val.toFixed(2);
                } else {
                    td.innerText = val;
                }

                if (!col.readOnly) {
                    td.contentEditable = true;
                    td.onblur = (e) => handleCellEdit(rowIndex, col.key, e.target.innerText);
                } else {
                    // Logic: Keep read-only, but strict validation colors removed
                    td.style.color = '#64748b'; // Muted text
                    td.style.backgroundColor = 'transparent'; // Clean look
                }
                tr.appendChild(td);
            });
        } else {
            // Raw Mode
            row.forEach((cell, colIndex) => {
                const td = document.createElement('td');
                td.contentEditable = true;
                td.innerText = cell;
                tr.appendChild(td);
            });
        }
        tbody.appendChild(tr);
    });

    dataTable.appendChild(thead);
    dataTable.appendChild(tbody);
}

function handleCellEdit(rowIndex, key, value) {
    if (!activeSchema) return;

    const row = tableData[rowIndex];
    const colDef = activeSchema.columns.find(c => c.key === key);

    if (colDef.type === 'number') {
        row[key] = parseFloat(value);
    } else {
        row[key] = value;
    }

    tableData[rowIndex] = recalculateRow(row);
    renderTable();
}

/**
 * UI Transitions
 */
function showEditor() {
    uploadZone.style.display = 'none';
    editorZone.classList.remove('hidden');
}

function resetApp() {
    fileInput.value = '';
    workbook = null;
    activeSchema = null;
    tableData = [];
    uploadZone.style.display = 'flex';
    editorZone.classList.add('hidden');
    dataTable.innerHTML = '';
}

function exportData() {
    // Same export logic...
    let ws;
    if (activeSchema) {
        const headers = activeSchema.columns.map(c => c.label);
        const rows = tableData.map(row => activeSchema.columns.map(c => row[c.key]));
        ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    } else {
        ws = XLSX.utils.aoa_to_sheet(tableData);
    }

    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, ws, "Sheet1");
    XLSX.writeFile(newWb, "modified_" + originalFilename);
}
