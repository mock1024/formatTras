// Global state
const state = {
    sourceFile: null,
    sourceData: null,
    sourceHeaders: [],
    sourceHeadersWithSample: [], // Store headers with sample data
    targetTemplate: [],
    mappings: {}, // Key: targetField, Value: sourceField
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
    '商品编号': ['商品编号', '商品ID', '商品代码', '商品编码', '自编码', '编号', 'ID', 'code', 'goods_id', 'product_id', 'product_code'],
    '商品名称': ['商品名称', '名称', '品名', '商品', 'name', 'goods_name', 'product_name', 'title'],
    '商品条码': [
        '商品条码', '条形码', '条码', 'barcode', 'bar_code',
        'barcode1', 'barcode2', 'barcode_1', 'barcode_2',
        '条码1', '条码2', '条码_1', '条码_2',
        '商品条形码', '条码号', '条形码号', '码',
        'barcode_no', 'barcode_num', 'bar_code_no',
        '条形码编号', '条码编号', '商品码',
        '国际条码', 'ean', 'upc', 'ean13', 'ean_13', 'upc_a',
        '商品条形码1', '商品条形码2', '商品条码1', '商品条码2'
    ],
    '规格': ['规格', '规格型号', 'model', 'spec', 'specification', 'specs', '规格说明'],
    '单位': ['单位', '计量单位', 'unit', 'uom', 'unit_name'],
    '进货价': ['进货价', '进价', '成本价', '采购价', '入库价', 'buy_price', 'cost_price', 'purchase_price', 'in_price'],
    '销售价': ['销售价', '现价', '售价', '单价', '原价', '零售价', 'price', 'sell_price', 'sale_price', 'retail_price', 'unit_price'],
    '会员价': ['会员价', 'VIP价', 'vip_price', 'member_price', '会员价1', '会员价2'],
    '批发价': ['批发价', 'wholesale_price', '批发价1', '批发价2'],
    '库存': ['库存', '库存数量', '库存量', '库存数', 'stock', 'quantity', 'qty', 'inventory', 'stock_num', 'amount'],
    '分类': ['分类', '分类编码', '分类号', '类别', 'category', 'category_code', 'category_id', 'class', 'type'],
    '分类名称': ['分类名称', '分类名', 'category_name', 'category'],
    '品牌': ['品牌', 'brand', 'brand_name', 'brand_id'],
    '供应商': ['供应商', '供货商', 'supplier', 'supplier_name', 'supplier_id', 'vendor']
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
        checkAndShowMapping();
    } else {
        document.getElementById('templateInfo').classList.add('hidden');
        document.getElementById('targetFieldsSection').classList.add('hidden');
        checkAndShowMapping();
    }
}

function displayTemplateFields() {
    const container = document.getElementById('templateFields');
    if (!container) return;

    container.innerHTML = state.targetTemplate.map(field =>
        `<span class="field-tag">${field}</span>`
    ).join('');

    document.getElementById('templateInfo').classList.remove('hidden');
    document.getElementById('targetFieldCount').textContent = `${state.targetTemplate.length} 个字段`;
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
    document.getElementById('sourceFieldsSection').classList.add('hidden');
    state.sourceData = null;
    state.sourceHeaders = [];
    state.sourceHeadersWithSample = [];
    state.sourceFile = null;
    updateConvertButton();
    checkAndShowMapping();
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

        // Store headers with sample data from first 3 rows
        state.sourceHeadersWithSample = state.sourceHeaders.map((header, index) => {
            const samples = [];
            for (let i = 1; i < Math.min(4, jsonData.length); i++) {
                if (jsonData[i][index] !== undefined && jsonData[i][index] !== '') {
                    samples.push(jsonData[i][index]);
                    if (samples.length >= 3) break;
                }
            }
            return {
                name: header,
                samples: samples,
                index: index
            };
        });

        state.sourceData = jsonData.slice(1).filter(row => row.some(cell => cell !== undefined && cell !== ''));
        state.sourceFile = file;

        // Display source fields
        displaySourceFields();

        // Check if we can show mapping section
        checkAndShowMapping();
    };

    reader.readAsArrayBuffer(file);
}

function displaySourceFields() {
    const container = document.getElementById('sourceFieldsDisplay');
    if (!container) return;

    container.innerHTML = state.sourceHeadersWithSample.map(field => `
        <div class="field-item source-field-item" draggable="true" data-field="${field.name}">
            <div class="field-header">
                <svg class="drag-handle" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                    <circle cx="9" cy="12" r="1"/>
                    <circle cx="9" cy="5" r="1"/>
                    <circle cx="9" cy="19" r="1"/>
                    <circle cx="15" cy="12" r="1"/>
                    <circle cx="15" cy="5" r="1"/>
                    <circle cx="15" cy="19" r="1"/>
                </svg>
                <span class="field-name">${field.name}</span>
            </div>
            <div class="field-samples">
                ${field.samples.slice(0, 2).map(s => `<span class="sample-value">${String(s).substring(0, 20)}${String(s).length > 20 ? '...' : ''}</span>`).join('')}
            </div>
        </div>
    `).join('');

    document.getElementById('sourceFieldsSection').classList.remove('hidden');
    document.getElementById('sourceFieldCount').textContent = `${state.sourceHeaders.length} 个字段`;

    // Add drag event listeners
    addDragEventListeners();
}

function displayTargetFields() {
    const container = document.getElementById('targetFieldsDisplay');
    if (!container) return;

    container.innerHTML = state.targetTemplate.map(field => `
        <div class="field-item target-field-item" data-field="${field}" data-mapped="">
            <div class="field-header">
                <svg class="drag-handle" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                    <circle cx="9" cy="12" r="1"/>
                    <circle cx="9" cy="5" r="1"/>
                    <circle cx="9" cy="19" r="1"/>
                    <circle cx="15" cy="12" r="1"/>
                    <circle cx="15" cy="5" r="1"/>
                    <circle cx="15" cy="19" r="1"/>
                </svg>
                <span class="field-name">${field}</span>
            </div>
            <div class="mapped-source-field" id="mapped-${field.replace(/\s+/g, '-')}">
                <span class="empty-hint">拖拽左侧字段到此处</span>
            </div>
        </div>
    `).join('');

    document.getElementById('targetFieldsSection').classList.remove('hidden');

    // Add drag event listeners
    addDragEventListeners();
}

function showMappingSection() {
    // Check if both source and target are ready
    checkAndShowMapping();
}

function checkAndShowMapping() {
    const hasSource = state.sourceHeaders && state.sourceHeaders.length > 0;
    const hasTarget = state.targetTemplate && state.targetTemplate.length > 0;

    if (hasSource && hasTarget) {
        document.getElementById('mappingSection').classList.remove('hidden');
        displayTargetFields();
        updateMappingList();
    } else {
        document.getElementById('mappingSection').classList.add('hidden');
    }
    updateConvertButton();
}

function addDragEventListeners() {
    const sourceFields = document.querySelectorAll('.source-field-item');
    const targetFields = document.querySelectorAll('.target-field-item');

    sourceFields.forEach(field => {
        field.addEventListener('dragstart', handleDragStart);
        field.addEventListener('dragend', handleDragEnd);
    });

    targetFields.forEach(field => {
        field.addEventListener('dragover', handleDragOver);
        field.addEventListener('drop', handleDrop);
        field.addEventListener('dragenter', handleDragEnter);
        field.addEventListener('dragleave', handleDragLeave);
    });
}

let draggedElement = null;

function handleDragStart(e) {
    draggedElement = e.target.closest('.field-item');
    e.target.style.opacity = '0.4';
    e.dataTransfer.effectAllowed = 'move';
}

function handleDragEnd(e) {
    e.target.style.opacity = '1';
    document.querySelectorAll('.target-field-item').forEach(item => {
        item.classList.remove('drag-over');
    });
}

function handleDragOver(e) {
    if (e.preventDefault) {
        e.preventDefault();
    }
    e.dataTransfer.dropEffect = 'move';
    return false;
}

function handleDragEnter(e) {
    const targetField = e.target.closest('.target-field-item');
    if (targetField) {
        targetField.classList.add('drag-over');
    }
}

function handleDragLeave(e) {
    const targetField = e.target.closest('.target-field-item');
    if (targetField && !targetField.contains(e.relatedTarget)) {
        targetField.classList.remove('drag-over');
    }
}

function handleDrop(e) {
    if (e.stopPropagation) {
        e.stopPropagation();
    }

    const targetField = e.target.closest('.target-field-item');
    if (draggedElement && targetField) {
        const sourceFieldName = draggedElement.dataset.field;
        const targetFieldName = targetField.dataset.field;

        // Create mapping
        createMapping(sourceFieldName, targetFieldName);

        // Visual feedback
        targetField.classList.remove('drag-over');

        // Update mapped field display
        updateMappedFieldDisplay(targetFieldName, sourceFieldName);
    }

    return false;
}

function createMapping(sourceField, targetField) {
    // Remove any existing mapping for this target field
    Object.keys(state.mappings).forEach(key => {
        if (state.mappings[key] === sourceField) {
            delete state.mappings[key];
            // Clear old display
            const oldTarget = document.querySelector(`.target-field-item[data-field="${key}"]`);
            if (oldTarget) {
                const mappedDiv = oldTarget.querySelector('.mapped-source-field');
                mappedDiv.innerHTML = '<span class="empty-hint">拖拽左侧字段到此处</span>';
                oldTarget.dataset.mapped = '';
            }
        }
    });

    // Create new mapping
    state.mappings[targetField] = sourceField;

    // Update mapping list
    updateMappingList();
    updateConvertButton();
}

function updateMappedFieldDisplay(targetField, sourceField) {
    const targetElement = document.querySelector(`.target-field-item[data-field="${targetField}"]`);
    if (targetElement) {
        const mappedDiv = targetElement.querySelector('.mapped-source-field');
        const sourceData = state.sourceHeadersWithSample.find(h => h.name === sourceField);

        mappedDiv.innerHTML = `
            <div class="mapped-field-info">
                <svg class="mapped-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor">
                    <polyline points="20 6 9 17 4 12"/>
                </svg>
                <span class="mapped-field-name">${sourceField}</span>
                <button class="btn-remove-mapping" onclick="removeMapping('${targetField}')">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
                        <line x1="18" y1="6" x2="6" y2="18"/>
                        <line x1="6" y1="6" x2="18" y2="18"/>
                    </svg>
                </button>
            </div>
            ${sourceData && sourceData.samples.length > 0 ? `
                <div class="mapped-field-samples">
                    ${sourceData.samples.slice(0, 2).map(s => `<span>${String(s).substring(0, 15)}${String(s).length > 15 ? '...' : ''}</span>`).join('')}
                </div>
            ` : ''}
        `;

        targetElement.dataset.mapped = sourceField;
        targetElement.classList.add('has-mapping');
    }
}

function removeMapping(targetField) {
    delete state.mappings[targetField];

    const targetElement = document.querySelector(`.target-field-item[data-field="${targetField}"]`);
    if (targetElement) {
        const mappedDiv = targetElement.querySelector('.mapped-source-field');
        mappedDiv.innerHTML = '<span class="empty-hint">拖拽左侧字段到此处</span>';
        targetElement.dataset.mapped = '';
        targetElement.classList.remove('has-mapping');
    }

    updateMappingList();
    updateConvertButton();
}

function updateMappingList() {
    const container = document.getElementById('mappingList');
    if (!container) return;

    container.innerHTML = '';

    // Get all mappings
    const mappings = Object.entries(state.mappings);

    // Update mapping count
    document.getElementById('mappingCount').textContent = mappings.length;

    if (mappings.length === 0) {
        container.innerHTML = '<div class="no-mappings">暂无字段映射，请从左侧拖拽字段到右侧</div>';
        return;
    }

    mappings.forEach(([targetField, sourceField]) => {
        const item = document.createElement('div');
        item.className = 'mapping-item-list';

        const sourceData = state.sourceHeadersWithSample.find(h => h.name === sourceField);

        item.innerHTML = `
            <div class="mapping-source">
                <span class="field-label">${sourceField}</span>
                ${sourceData && sourceData.samples.length > 0 ? `
                    <span class="field-sample">${String(sourceData.samples[0]).substring(0, 20)}</span>
                ` : ''}
            </div>
            <div class="mapping-arrow">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
                    <line x1="5" y1="12" x2="19" y2="12"/>
                    <polyline points="12 5 19 12 12 19"/>
                </svg>
            </div>
            <div class="mapping-target">
                <span class="field-label">${targetField}</span>
            </div>
            <button class="btn-remove-mapping-small" onclick="removeMapping('${targetField}')" title="删除映射">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
                    <line x1="18" y1="6" x2="6" y2="18"/>
                    <line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
            </button>
        `;

        container.appendChild(item);
    });
}

function updateMapping(targetField, sourceField) {
    if (sourceField) {
        createMapping(sourceField, targetField);
        updateMappedFieldDisplay(targetField, sourceField);
    } else {
        removeMapping(targetField);
    }
}

function autoMapFields() {
    if (!state.sourceHeaders || state.sourceHeaders.length === 0) return;
    if (!state.targetTemplate || state.targetTemplate.length === 0) return;

    // Build content-based patterns for barcode detection
    const barcodeContentScores = {};
    if (state.sourceData && state.sourceData.length > 0) {
        state.sourceHeaders.forEach((header, index) => {
            // Check first 10 rows for barcode-like patterns
            let barcodePatternCount = 0;
            let checkRows = Math.min(10, state.sourceData.length);

            for (let i = 0; i < checkRows; i++) {
                const value = state.sourceData[i][index];
                if (value && typeof value === 'string') {
                    const trimmed = value.trim();
                    // Check if it looks like a barcode (8, 12, 13 digits)
                    if (/^\d{8}$|^\d{12}$|^\d{13}$/.test(trimmed)) {
                        barcodePatternCount++;
                    }
                }
            }

            // Calculate score based on how many rows have barcode patterns
            barcodeContentScores[header] = (barcodePatternCount / checkRows) * 50;
        });
    }

    // Auto-map each target field
    state.targetTemplate.forEach(targetField => {
        let bestMatch = null;
        let bestScore = 0;

        state.sourceHeaders.forEach(sourceHeader => {
            let score = 0;
            const sourceLower = sourceHeader.toLowerCase();

            // Direct match
            if (sourceHeader === targetField) {
                score = 100;
            }
            // Check aliases
            else if (fieldAliases[targetField]) {
                const aliases = fieldAliases[targetField];

                // Exact alias match
                if (aliases.some(alias => alias.toLowerCase() === sourceLower)) {
                    score = 80;
                }
                // Partial match - check if source contains alias or vice versa
                else if (aliases.some(alias => {
                    const aliasLower = alias.toLowerCase();
                    return sourceLower.includes(aliasLower) || aliasLower.includes(sourceLower);
                })) {
                    score = 60;
                }
            }

            // Special handling for barcode fields - boost score if content looks like barcodes
            if (targetField === '商品条码' || targetField === '条形码' || targetField === '条码') {
                if (barcodeContentScores[sourceHeader]) {
                    score += barcodeContentScores[sourceHeader];
                }
            }

            // Prefer exact field name match for barcodes
            if ((targetField === '商品条码' || targetField === '条形码' || targetField === '条码') &&
                (sourceHeader === '商品条码' || sourceHeader === '条形码' || sourceHeader === '条码' ||
                 sourceLower === 'barcode' || sourceLower === 'bar_code')) {
                score = Math.max(score, 95);
            }

            if (score > bestScore) {
                bestScore = score;
                bestMatch = sourceHeader;
            }
        });

        if (bestMatch && bestScore >= 60) {
            createMapping(bestMatch, targetField);
            updateMappedFieldDisplay(targetField, bestMatch);
        }
    });

    updateConvertButton();
}

function resetMappings() {
    state.mappings = {};

    // Clear all target field displays
    document.querySelectorAll('.target-field-item').forEach(item => {
        const mappedDiv = item.querySelector('.mapped-source-field');
        if (mappedDiv) {
            mappedDiv.innerHTML = '<span class="empty-hint">拖拽左侧字段到此处</span>';
        }
        item.dataset.mapped = '';
        item.classList.remove('has-mapping');
    });

    updateMappingList();
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

    // Update stats
    document.getElementById('convertedCount').textContent = state.convertedData.length;
    document.getElementById('targetFieldCount').textContent = state.targetTemplate.length;

    // Show result section
    resultSection.classList.remove('hidden');

    // Scroll to result
    resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
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
    state.sourceFile = null;
    state.sourceData = null;
    state.sourceHeaders = [];
    state.sourceHeadersWithSample = [];
    state.targetTemplate = [];
    state.mappings = {};
    state.convertedData = null;

    // Reset UI
    document.getElementById('sourceSoftware').value = '';
    document.getElementById('targetSoftware').value = '';
    document.getElementById('sourceFile').value = '';
    document.getElementById('sourceFileInfo').classList.add('hidden');
    document.getElementById('sourceUpload').classList.remove('hidden');
    document.getElementById('templateInfo').classList.add('hidden');
    document.getElementById('sourceFieldsSection').classList.add('hidden');
    document.getElementById('targetFieldsSection').classList.add('hidden');
    document.getElementById('mappingSection').classList.add('hidden');
    document.getElementById('resultSection').classList.add('hidden');
    document.getElementById('convertBtn').disabled = true;

    // Scroll to top
    window.scrollTo({ top: 0, behavior: 'smooth' });
}
