// Global state
const state = {
    sourceData: null,
    sourceHeaders: [],
    targetTemplate: [],
    mappings: {},
    convertedData: null
};

// Software templates configuration
const softwareTemplates = {
    kemai: {
        name: '科脉云帆',
        fields: [
            '商品编号', '商品名称', '商品条码', '规格', '单位',
            '进货价', '销售价1', '销售价2', '批发价', '会员价',
            '库存数量', '分类编码', '分类名称', '品牌', '供应商',
            '商品简称', '助记码', '保质期', '商品状态'
        ]
    },
    jushang: {
        name: '聚商荟',
        fields: [
            '商品ID', '商品名称', '条形码', '规格型号', '计量单位',
            '进价', '原价', '现价', '库存', '分类',
            '品牌', '供应商', '是否上架', '创建时间', '更新时间'
        ]
    },
    sixun: {
        name: '思迅天店',
        fields: [
            '商品代码', '商品名称', '条码', '规格', '单位',
            '进价', '单价', '折前价', '会员价', '库存',
            '分类号', '分类名', '品牌', '供应商', '拼音码'
        ]
    },
    yinbao: {
        name: '银豹',
        fields: [
            '商品条码', '商品名称', '简码', '单位', '进价',
            '售价', '会员价', '批发价', '库存', '分类',
            '品牌', '供应商', '规格', '商品属性'
        ]
    },
    fuzhanggui: {
        name: '富掌柜',
        fields: [
            '商品编码', '商品名称', '条形码', '规格', '单位',
            '成本价', '原价', '现价', '促销价', '库存',
            '分类', '品牌', '供应商', '备注', '状态'
        ]
    },
    higuanjia: {
        name: '惠管家',
        fields: [
            '商品编号', '商品名称', '条码', '规格', '单位',
            '进价', '销售价', '会员价', '库存量', '分类',
            '品牌', '供应商', '商品简称', '助记符'
        ]
    },
    mingxian: {
        name: '明献零售',
        fields: [
            '商品代码', '商品名称', '自编码', '规格', '单位',
            '进货价', '零售价', '批发价', '库存数', '类别',
            '品牌', '供货商', '备注', '是否启用'
        ]
    }
};

// Common field aliases for auto-mapping
const fieldAliases = {
    '商品编号': ['商品编号', '商品ID', '商品代码', '商品编码', '自编码', '编号', 'ID', 'code'],
    '商品名称': ['商品名称', '名称', '品名', '商品', 'name'],
    '商品条码': ['商品条码', '条形码', '条码', 'barcode'],
    '规格': ['规格', '规格型号', 'model'],
    '单位': ['单位', '计量单位', 'unit'],
    '进货价': ['进货价', '进价', '成本价', '采购价', '入库价'],
    '销售价': ['销售价', '现价', '售价', '单价', '原价', '零售价', 'price'],
    '会员价': ['会员价', 'VIP价'],
    '批发价': ['批发价'],
    '库存': ['库存', '库存数量', '库存量', '库存数', 'stock'],
    '分类': ['分类', '分类编码', '分类号', '类别', 'category'],
    '分类名称': ['分类名称', '分类名'],
    '品牌': ['品牌', 'brand'],
    '供应商': ['供应商', '供货商', 'supplier']
};

// Initialize event listeners
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
});

function initializeEventListeners() {
    // Source software selection
    document.getElementById('sourceSoftware').addEventListener('change', handleSourceSoftwareChange);

    // Target software selection
    document.getElementById('targetSoftware').addEventListener('change', handleTargetSoftwareChange);

    // File upload
    const uploadArea = document.getElementById('sourceUpload');
    const fileInput = document.getElementById('sourceFile');

    uploadArea.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);

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
        const files = e.dataTransfer.files;
        if (files.length) {
            handleFile(files[0]);
        }
    });
}

function handleSourceSoftwareChange(e) {
    const software = e.target.value;
    // Can pre-fill based on software selection if needed
    console.log('Source software selected:', software);
}

function handleTargetSoftwareChange(e) {
    const software = e.target.value;

    if (software && softwareTemplates[software]) {
        state.targetTemplate = softwareTemplates[software].fields;
        displayTemplateFields();
        showMappingSection();
    } else {
        document.getElementById('templateInfo').classList.add('hidden');
        document.getElementById('mappingSection').classList.add('hidden');
    }
}

function displayTemplateFields() {
    const container = document.getElementById('templateFields');
    container.innerHTML = state.targetTemplate.map(field =>
        `<span class="field-tag">${field}</span>`
    ).join('');
    document.getElementById('templateInfo').classList.remove('hidden');
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
}

function handleFile(file) {
    // Validate file type
    const validTypes = ['.xlsx', '.xls', '.csv'];
    const fileExt = '.' + file.name.split('.').pop().toLowerCase();

    if (!validTypes.includes(fileExt)) {
        alert('请上传有效的文件格式 (.xlsx, .xls, .csv)');
        return;
    }

    // Show file info
    const fileInfo = document.getElementById('sourceFileInfo');
    fileInfo.querySelector('.file-name').textContent = file.name;
    fileInfo.classList.remove('hidden');
    document.getElementById('sourceUpload').classList.add('hidden');

    // Parse file
    parseFile(file);
}

function removeSourceFile() {
    document.getElementById('sourceFile').value = '';
    document.getElementById('sourceFileInfo').classList.add('hidden');
    document.getElementById('sourceUpload').classList.remove('hidden');
    document.getElementById('sourcePreview').classList.add('hidden');
    state.sourceData = null;
    state.sourceHeaders = [];
    updateConvertButton();
}

function parseFile(file) {
    const reader = new FileReader();

    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Get first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (jsonData.length < 2) {
            alert('文件数据为空或格式不正确');
            return;
        }

        // Extract headers and data
        state.sourceHeaders = jsonData[0];
        state.sourceData = jsonData.slice(1).filter(row => row.some(cell => cell !== undefined && cell !== ''));

        // Display preview
        displayDataPreview();
        updateMappings();
        updateConvertButton();
    };

    reader.readAsArrayBuffer(file);
}

function displayDataPreview() {
    const preview = document.getElementById('sourcePreview');
    const table = document.getElementById('sourceTable');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    const recordCount = preview.querySelector('.record-count');

    // Build headers
    thead.innerHTML = `<tr>${state.sourceHeaders.map(h => `<th>${h || ''}</th>`).join('')}</tr>`;

    // Build preview data (max 5 rows)
    const previewData = state.sourceData.slice(0, 5);
    tbody.innerHTML = previewData.map(row =>
        `<tr>${state.sourceHeaders.map((_, i) => `<td>${row[i] !== undefined ? row[i] : ''}</td>`).join('')}</tr>`
    ).join('');

    recordCount.textContent = `共 ${state.sourceData.length} 条记录`;
    preview.classList.remove('hidden');
}

function showMappingSection() {
    if (state.sourceHeaders.length > 0 && state.targetTemplate.length > 0) {
        document.getElementById('mappingSection').classList.remove('hidden');
        updateMappings();
    }
}

function updateMappings() {
    if (state.sourceHeaders.length === 0 || state.targetTemplate.length === 0) return;

    const container = document.getElementById('mappingList');
    container.innerHTML = '';

    state.targetTemplate.forEach((targetField, index) => {
        const item = document.createElement('div');
        item.className = 'mapping-item';

        const sourceOptions = state.sourceHeaders.map(header =>
            `<option value="${header}">${header}</option>`
        ).join('');

        item.innerHTML = `
            <div class="mapping-field">
                <label>目标字段: ${targetField}</label>
                <select class="mapping-select" data-target="${targetField}" onchange="updateMapping('${targetField}', this.value)">
                    <option value="">-- 选择源字段 --</option>
                    ${sourceOptions}
                </select>
            </div>
            <div class="mapping-arrow">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
                    <line x1="5" y1="12" x2="19" y2="12"/>
                    <polyline points="12 5 19 12 12 19"/>
                </svg>
            </div>
            <div class="mapping-field">
                <label>源字段</label>
                <input type="text" readonly class="mapping-source" placeholder="未选择" id="source-${index}">
            </div>
        `;

        container.appendChild(item);
    });

    // Auto-map on first load
    autoMapFields();
}

function updateMapping(targetField, sourceField) {
    if (sourceField) {
        state.mappings[targetField] = sourceField;
    } else {
        delete state.mappings[targetField];
    }
    updateConvertButton();
}

function autoMapFields() {
    const selects = document.querySelectorAll('.mapping-select');

    selects.forEach(select => {
        const targetField = select.dataset.target;
        let bestMatch = null;
        let bestScore = 0;

        // Check for direct match or alias match
        state.sourceHeaders.forEach(sourceHeader => {
            let score = 0;

            // Direct match
            if (sourceHeader === targetField) {
                score = 100;
            }
            // Check aliases
            else if (fieldAliases[targetField]) {
                const aliases = fieldAliases[targetField];
                const sourceLower = sourceHeader.toLowerCase();
                if (aliases.some(alias => alias.toLowerCase() === sourceLower)) {
                    score = 80;
                }
                // Partial match
                else if (aliases.some(alias => sourceLower.includes(alias.toLowerCase()) ||
                                              alias.toLowerCase().includes(sourceLower))) {
                    score = 60;
                }
            }

            if (score > bestScore) {
                bestScore = score;
                bestMatch = sourceHeader;
            }
        });

        if (bestMatch && bestScore >= 60) {
            select.value = bestMatch;
            updateMapping(targetField, bestMatch);

            // Update display
            const index = Array.from(selects).indexOf(select);
            document.getElementById(`source-${index}`).value = bestMatch;
        }
    });

    updateConvertButton();
}

function resetMappings() {
    const selects = document.querySelectorAll('.mapping-select');
    selects.forEach((select, index) => {
        select.value = '';
        document.getElementById(`source-${index}`).value = '';
    });
    state.mappings = {};
    updateConvertButton();
}

function updateConvertButton() {
    const btn = document.getElementById('convertBtn');
    const hasSourceData = state.sourceData && state.sourceData.length > 0;
    const hasMappings = Object.keys(state.mappings).length > 0;
    const hasTarget = document.getElementById('targetSoftware').value !== '';

    btn.disabled = !(hasSourceData && hasMappings && hasTarget);
}

function convertData() {
    if (!state.sourceData || !state.sourceData.length) {
        alert('请先上传源数据文件');
        return;
    }

    if (Object.keys(state.mappings).length === 0) {
        alert('请先配置字段映射');
        return;
    }

    // Get source header indices
    const sourceIndices = {};
    state.sourceHeaders.forEach((header, index) => {
        sourceIndices[header] = index;
    });

    // Convert data
    const convertedData = state.sourceData.map(row => {
        const newRow = {};
        state.targetTemplate.forEach(targetField => {
            const sourceField = state.mappings[targetField];
            if (sourceField && sourceIndices[sourceField] !== undefined) {
                newRow[targetField] = row[sourceIndices[sourceField]] || '';
            } else {
                newRow[targetField] = '';
            }
        });
        return newRow;
    });

    state.convertedData = convertedData;
    displayResult();
}

function displayResult() {
    const resultSection = document.getElementById('resultSection');
    const resultTable = document.getElementById('resultTable');
    const thead = resultTable.querySelector('thead');
    const tbody = resultTable.querySelector('tbody');

    // Update stats
    document.getElementById('convertedCount').textContent = state.convertedData.length;
    document.getElementById('targetFieldCount').textContent = state.targetTemplate.length;

    // Build table headers
    thead.innerHTML = `<tr>${state.targetTemplate.map(h => `<th>${h}</th>`).join('')}</tr>`;

    // Build table body (preview first 5 rows)
    const previewData = state.convertedData.slice(0, 5);
    tbody.innerHTML = previewData.map(row =>
        `<tr>${state.targetTemplate.map(field => `<td>${row[field] || ''}</td>`).join('')}</tr>`
    ).join('');

    // Show result section
    resultSection.classList.remove('hidden');

    // Scroll to result
    resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function toggleFullPreview() {
    const tbody = document.querySelector('#resultTable tbody');
    const btnText = document.getElementById('previewToggleText');

    if (btnText.textContent === '查看更多') {
        // Show all data (limit to 100 for performance)
        const previewData = state.convertedData.slice(0, 100);
        tbody.innerHTML = previewData.map(row =>
            `<tr>${state.targetTemplate.map(field => `<td>${row[field] || ''}</td>`).join('')}</tr>`
        ).join('');
        btnText.textContent = '收起';
    } else {
        // Show first 5 rows
        const previewData = state.convertedData.slice(0, 5);
        tbody.innerHTML = previewData.map(row =>
            `<tr>${state.targetTemplate.map(field => `<td>${row[field] || ''}</td>`).join('')}</tr>`
        ).join('');
        btnText.textContent = '查看更多';
    }
}

function downloadExcel() {
    if (!state.convertedData || !state.convertedData.length) {
        alert('没有可下载的数据');
        return;
    }

    const targetSoftware = document.getElementById('targetSoftware');
    const softwareName = softwareTemplates[targetSoftware.value]?.name || '导出';

    // Create worksheet
    const ws = XLSX.utils.json_to_sheet(state.convertedData);

    // Set column widths
    const colWidths = state.targetTemplate.map(() => ({ wch: 15 }));
    ws['!cols'] = colWidths;

    // Create workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '商品资料');

    // Generate filename with timestamp
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '').replace('T', '_');
    const filename = `${softwareName}_商品资料_${timestamp}.xlsx`;

    // Download
    XLSX.writeFile(wb, filename);
}

function downloadCSV() {
    if (!state.convertedData || !state.convertedData.length) {
        alert('没有可下载的数据');
        return;
    }

    const targetSoftware = document.getElementById('targetSoftware');
    const softwareName = softwareTemplates[targetSoftware.value]?.name || '导出';

    // Build CSV content with BOM for Excel Chinese support
    let csvContent = '\uFEFF'; // BOM

    // Add headers
    csvContent += state.targetTemplate.join(',') + '\n';

    // Add data rows
    state.convertedData.forEach(row => {
        const rowData = state.targetTemplate.map(field => {
            let value = row[field] || '';
            // Escape quotes and wrap in quotes if contains comma or quote
            if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) {
                value = '"' + value.replace(/"/g, '""') + '"';
            }
            return value;
        });
        csvContent += rowData.join(',') + '\n';
    });

    // Create blob and download
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '').replace('T', '_');
    link.href = URL.createObjectURL(blob);
    link.download = `${softwareName}_商品资料_${timestamp}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
}

function resetAll() {
    // Reset state
    state.sourceData = null;
    state.sourceHeaders = [];
    state.targetTemplate = [];
    state.mappings = {};
    state.convertedData = null;

    // Reset UI
    document.getElementById('sourceSoftware').value = '';
    document.getElementById('targetSoftware').value = '';
    document.getElementById('sourceFile').value = '';
    document.getElementById('sourceFileInfo').classList.add('hidden');
    document.getElementById('sourceUpload').classList.remove('hidden');
    document.getElementById('sourcePreview').classList.add('hidden');
    document.getElementById('templateInfo').classList.add('hidden');
    document.getElementById('mappingSection').classList.add('hidden');
    document.getElementById('resultSection').classList.add('hidden');
    document.getElementById('convertBtn').disabled = true;

    // Scroll to top
    window.scrollTo({ top: 0, behavior: 'smooth' });
}
