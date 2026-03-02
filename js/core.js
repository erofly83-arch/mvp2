"use strict";
// ════════════════════════════════════════════════════════════════════════════
// core.js — App Shell: Monitor, Navigation, OBR, JSON Editor, Brand, Sidebar
// ════════════════════════════════════════════════════════════════════════════

// ── Brand config (populated by server/deployment if needed) ──
const BRAND_CONFIG_SAVED = {};

// ════════════════════════════════════════════════════════════════════════════
// MAIN TABS
// ════════════════════════════════════════════════════════════════════════════

// FIX: declare _bcAddState at top-level to avoid TDZ when openAddToDB is called from inline onclick
let _bcAddState = null; // { mainBC, synonyms: [{bc, fileName, checked}] }

// switchMainTab is now an alias for the combined app's switchMainPane
function switchMainTab(name) { switchMainPane(name); }

// ════════════════════════════════════════════════════════════════════════════
// TOAST
// ════════════════════════════════════════════════════════════════════════════
function showToast(msg, type, ms) {
  type = type || 'info'; ms = ms || 3200;
  const rack = document.getElementById('toastRack');
  const el = document.createElement('div');
  el.className = 'je-toast ' + type;
  el.textContent = ({ok:'✅',err:'❌',warn:'⚠️',info:'ℹ️'}[type]||'') + ' ' + msg;
  rack.appendChild(el);
  requestAnimationFrame(() => el.classList.add('show'));
  setTimeout(() => { el.classList.remove('show'); setTimeout(() => el.remove(), 250); }, ms);
}

// ════════════════════════════════════════════════════════════════════════════
// CONFIRM DIALOG
// ════════════════════════════════════════════════════════════════════════════
let _jeConfirmResolve = null;
function jeConfirmDialog(msg, title) {
  return new Promise(resolve => {
    _jeConfirmResolve = resolve;
    document.getElementById('jeConfirmTitle').textContent = title || 'Подтверждение';
    document.getElementById('jeConfirmMsg').textContent = msg;
    document.getElementById('jeConfirmModal').style.display = 'flex';
  });
}
function jeConfirmClose(result) {
  document.getElementById('jeConfirmModal').style.display = 'none';
  if (_jeConfirmResolve) { _jeConfirmResolve(result); _jeConfirmResolve = null; }
}

// ════════════════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════════════════
// MONITOR APP (PM)
// ════════════════════════════════════════════════════════════════════════════
let barcodeAliasMap=new Map(),synonymsLoaded=false;
// По умолчанию синонимы не загружены

function resetBarcodeAliases(){
    barcodeAliasMap=new Map();synonymsLoaded=false;
    const s=document.getElementById('synonymsStatus');
    if(s){s.className='upload-status upload-status--idle';s.textContent='Не загружены';}
}

function canonicalizeBarcode(rawBarcode) {
    if (rawBarcode === undefined || rawBarcode === null) return { canonical: rawBarcode, wasSynonym: false };
    const b = String(rawBarcode).trim().replace(/\.0+$/, '');
    if (!synonymsLoaded) return { canonical: b, wasSynonym: false };
    const canon = barcodeAliasMap.get(b);
    if (canon && canon !== b) {
        return { canonical: canon, wasSynonym: true };
    }
    return { canonical: b, wasSynonym: false };
}

// ======== ОСНОВНОЙ КОД ПРИЛОЖЕНИЯ ========

let myPriceData = null;
let competitorFilesData = [];
let allFilesData = [];
let groupedData = [];
let allColumns = [];
let visibleColumns = new Set();
let barcodeColumn = null;
let nameColumn = null;
let stockColumn = null; // колонка остатка в МОЁМ прайсе (если есть)
let transitColumn = null; // колонка "в пути" в МОЁМ прайсе (если есть)
let showTransitColumn = false; // пользовательский тоггл «В пути»
let customColumns = []; // [{key, displayName}] — пользовательские колонки
let customColData = {}; // {colKey: {barcode: value}}
let _customColCounter = 0;
let _transitDisplayName = 'В пути'; // сохраняется при переименовании колонки «В пути»
let sortMode = 'default';

// ── Virtual scroll for main monitoring table ──────────────────────────
const MVS = { ROW_H: 42, OVERSCAN: 30, start: 0, end: 0, ticking: false };
let _vsData = [];      // current filtered+sorted data array
// Cached per-render values (reset each renderTable call)
let _vsVisibleCols = [];
let _vsSupplierPriceCols = [];
let _vsColPayGroupMap = new Map();
// Cached select options string (99 options, generated once)
const _DIV_OPTIONS = Array.from({length:99}, (_,i)=>i+2).map(n=>`<option value="${n}">${n}</option>`).join('');
let _searchDebounceTimer = null;
let compactMatches = false;
let searchQuery = '';
let showFileBarcodes = false; // UI: отображать ли колонки ШК по файлам
let filterNewItems = false;   // Фильтр новинок: показывать только то, чего нет в моём прайсе

const myPriceInput = document.getElementById('myPriceInput');
const competitorInput = document.getElementById('competitorInput');
const synonymsInput = document.getElementById('synonymsInput');
const searchInput = document.getElementById('searchInput');
const sortMatchesBtn = document.getElementById('sortMatchesBtn');
const bigDiffBtn = document.getElementById('bigDiffBtn');
const showMyPriceBtn = document.getElementById('showMyPriceBtn');
const maxCoverageBtn = document.getElementById('maxCoverageBtn');
const compactMatchesBtn     = document.getElementById('compactMatchesBtn');
// Кнопки удалены из UI, но переменные нужны для null-guard в остальном коде
// toggleTransitBtn removed from UI — kept via toggleTransitColumn() function
const exportMyPriceBtn = document.getElementById('exportMyPriceBtn');
const exportAllBtn = document.getElementById('exportAllBtn');
const exportCurrentBtn = document.getElementById('exportCurrentBtn');
const clearBtn=document.getElementById('clearBtn');
const infoPanel = document.getElementById('infoPanel');
const hiddenColumnsPanel = document.getElementById('hiddenColumnsPanel');
const hiddenColumnsList = document.getElementById('hiddenColumnsList');
const tableContainer = document.getElementById('tableContainer');

const BARCODE_SYNONYMS = [
    'штрих-код', 'штрихкод', 'barcode', 'Штрих-код', 'Штрихкод', 'Barcode',
    'код', 'Код', 'ean', 'EAN', 'ean13', 'EAN13', 'штрих код', 'Штрих код',
    'bar_code', 'bar-code', 'product_code', 'sku', 'SKU', 'артикул', 'Артикул'
];

const NAME_SYNONYMS = [
    'название', 'name', 'Название', 'Name', 'наименование', 'Наименование',
    'товар', 'Товар', 'product', 'Product', 'описание', 'Описание',
    'product_name', 'title', 'Title', 'имя', 'Имя'
];

myPriceInput.addEventListener('change', handleMyPriceUpload);
competitorInput.addEventListener('change', handleCompetitorUpload);
// synonymsInput loading is handled by the jeDB listener below (with capture=true)
// which correctly parses both combined {barcodes,brands} and legacy flat formats.
searchInput.addEventListener('input', handleSearch);
sortMatchesBtn.addEventListener('click', toggleSortMatches);
bigDiffBtn.addEventListener('click', toggleBigDiff);
showMyPriceBtn.addEventListener('click', toggleMyPriceView);
compactMatchesBtn.addEventListener('click', toggleCompactMatches);
maxCoverageBtn.addEventListener('click', toggleMaxCoverage);
exportMyPriceBtn.addEventListener('click', async () => await generateExcel('myprice'));
exportAllBtn.addEventListener('click', async () => await generateExcel('all'));
exportCurrentBtn.addEventListener('click', async () => await generateExcel('current'));

clearBtn.addEventListener('click', clearAll);

// --- Deduplicate same-price duplicates inside one supplier file ---
const PRICE_COL_SYNONYMS = [
'цена', 'price', 'cost', 'стоимость', 'прайс', 'опт', 'оптов', 'розн', 'ррц', 'рц',
'retail', 'wholesale'
];

const MY_PRICE_FILE_NAME = 'Мой прайс';
const META_STOCK_KEY = '__meta_stock';
const META_TRANSIT_KEY = '__meta_transit';
const STOCK_COL_SYNONYMS = ['остаток', 'остатки', 'наличие', 'склад'];
const TRANSIT_COL_SYNONYMS = ['в пути', 'впути', 'транзит', 'transit', 'in transit'];


const PRICE_DECIMALS = 1;

function roundPrice(n) {
const m = 10 ** PRICE_DECIMALS;
return Math.round(n * m) / m;
}


function isPriceLikeColumn(colName) {
const s = String(colName || '').toLowerCase();
// Исключаем колонки с остатками
const isStock = STOCK_COL_SYNONYMS.some(k => s.includes(k));
if (isStock) return false;
return PRICE_COL_SYNONYMS.some(k => s.includes(k));
}


function parsePriceNumber(val) {
const s = String(val ?? '').trim();
if (!s) return null;
const n = parseFloat(s.replace(/[^0-9.,-]/g, '').replace(',', '.'));
return Number.isFinite(n) ? n : null;
}

function getColPayGroup(col){const n=(col.displayName||col.columnName||"").toLowerCase();if(n.includes("нал"))return"нал";if(n.includes("бн"))return"бн";return"other";}
function extractPackQtyFromName(name) {
const s = String(name ?? '');
// Любое число перед "шт"/"штук": 30шт/24бл, 30 шт, 30 шт.
const m = s.match(/(\d{1,6})\s*(?:шт|штук)(?=[^0-9A-Za-zА-Яа-я]|$)/i);
if (!m) return null;
const q = parseInt(m[1], 10);
if (!Number.isFinite(q) || q <= 1) return null;
return q;
}


function samePrice(a, b) {
const na = parsePriceNumber(a);
const nb = parsePriceNumber(b);
if (na !== null && nb !== null) return roundPrice(na) === roundPrice(nb);
return String(a ?? '').trim() === String(b ?? '').trim();
}
function removeFileExtension(fileName) {
    return fileName.replace(/\.(csv|xlsx|xls)$/i, '');
}

function handleSearch(e) {
    clearTimeout(_searchDebounceTimer);
    _searchDebounceTimer = setTimeout(() => {
        searchQuery = e.target.value.toLowerCase().trim();
        renderTable();
    }, 180); // debounce 180ms — prevents re-render on every keystroke
}

async function handleMyPriceUpload(e) {
    try {
        const file = e.target.files[0];
        if (!file) return;
        myPriceData = await parseFile(file, 'Мой прайс');
    const _mpSt=document.getElementById('myPriceStatus');if(_mpSt){_mpSt.className='upload-status upload-status--ok';_mpSt.textContent='✅ '+file.name;}
        processAllData();
    } catch (error) {
        console.error("Ошибка загрузки файла:", error);
        const _mpSt=document.getElementById('myPriceStatus');
        if(_mpSt){_mpSt.className='upload-status upload-status--error';_mpSt.textContent='❌ '+error.message;}
        showToast('Ошибка загрузки: ' + error.message, 'err');
    }
}

async function handleCompetitorUpload(e) {
    try {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;
        // Additive loading: keep existing, confirm duplicates
        for (const file of files) {
            const fn = removeFileExtension(file.name);
            const dup = competitorFilesData.findIndex(f => f.fileName === fn);
            if (dup !== -1) {
                if (!confirm('Файл «' + fn + '» уже загружен.\nЗаменить его новой версией?')) continue;
                competitorFilesData.splice(dup, 1);
            }
            const fd = await parseFile(file, fn);
            competitorFilesData.push(fd);
        }
        const n = competitorFilesData.length;
        const _cSt = document.getElementById('competitorStatus');
        if (_cSt) { _cSt.className='upload-status upload-status--ok'; _cSt.textContent='✅ '+n+' файл'+(n===1?'':'а'+(n<5?'':'ов')); }
        if (typeof _sfUpdateSuppliers==='function') _sfUpdateSuppliers(competitorFilesData.map(f=>({name:f.fileName,rows:f.data?f.data.length:null})));
        processAllData();
    } catch (error) {
        console.error('Ошибка загрузки файлов:', error);
        const _cSt = document.getElementById('competitorStatus');
        if (_cSt) { _cSt.className='upload-status upload-status--error'; _cSt.textContent='❌ '+error.message; }
        showToast('Ошибка загрузки: ' + error.message, 'err');
    }
}

// Remove a supplier file by name and refresh everything
window.removeSupplierFile = function(fileName) {
    const idx = competitorFilesData.findIndex(f => f.fileName === fileName);
    if (idx === -1) return;
    if (!confirm('Удалить файл поставщика «' + fileName + '» из мониторинга?')) return;
    competitorFilesData.splice(idx, 1);
    const n = competitorFilesData.length;
    const _cSt = document.getElementById('competitorStatus');
    if (_cSt) {
        if (n === 0) { _cSt.className='upload-status upload-status--idle'; _cSt.textContent='Не загружены'; }
        else { _cSt.className='upload-status upload-status--ok'; _cSt.textContent='✅ '+n+' файл'+(n===1?'':'а'+(n<5?'':'ов')); }
    }
    if (typeof _sfUpdateSuppliers==='function') _sfUpdateSuppliers(competitorFilesData.map(f=>({name:f.fileName,rows:f.data?f.data.length:null})));
    if (competitorFilesData.length === 0 && !myPriceData) { clearAll && clearAll(); }
    else processAllData();
    showToast('Файл «' + fileName + '» удалён', 'ok');
};
async function parseFile(file, fileName) {
    try {
        if (file.name.endsWith('.csv')) {
            return await parseCSV(file, fileName);
        } else {
            return await parseExcel(file, fileName);
        }
    } catch (error) {
        console.error("Ошибка парсинга:", error);
        throw new Error("Не удалось прочитать файл");
    }
}

function parseCSV(file, fileName) {
    return new Promise((resolve, reject) => {
        Papa.parse(file, {
            header: true,
            encoding: 'UTF-8',
            skipEmptyLines: true,
            complete: (results) => {
                resolve({
                    fileName,
                    data: results.data,
                    isMyPrice: fileName === 'Мой прайс'
                });
            },
            error: (err) => {
                // Retry with Windows-1251 encoding (common in Russia)
                Papa.parse(file, {
                    header: true,
                    encoding: 'windows-1251',
                    skipEmptyLines: true,
                    complete: (results2) => resolve({
                        fileName,
                        data: results2.data,
                        isMyPrice: fileName === 'Мой прайс'
                    }),
                    error: (err2) => reject(new Error('CSV parse error: ' + err2.message))
                });
            }
        });
    });
}

function parseExcel(file, fileName) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { raw: false, defval: '' });
                resolve({fileName, data: jsonData, isMyPrice: fileName === 'Мой прайс'});
            } catch (err) {
                reject(new Error('Excel parse error: ' + err.message));
            }
        };
        reader.onerror = () => reject(new Error('Не удалось прочитать файл'));
        reader.readAsArrayBuffer(file);
    });
}

function processAllData() {
    if (!myPriceData && competitorFilesData.length === 0) return;
    allFilesData = [];
    _matcherDisabledFiles = new Set(); // сброс при полной перезагрузке
    if (typeof matcherFileChipsRender === 'function') matcherFileChipsRender();
    if (myPriceData) allFilesData.push(myPriceData);
    allFilesData = allFilesData.concat(competitorFilesData);
    // Убираем из disabled-сета файлы, которые уже не загружены
    const _currentFileNames = new Set(allFilesData.map(f => f.fileName));
    _matcherDisabledFiles.forEach(n => { if (!_currentFileNames.has(n)) _matcherDisabledFiles.delete(n); });

    autoDetectColumns();
    processData();
    updateHiddenColumnsPanel();
    renderTable();
    updateUI();
    showCompletionToast();
}

// FIX: detect barcode/name column per-file (different files can have different column names)
function detectFileCols(fileData) {
    if (!fileData.data || !fileData.data.length) return;
    const cols = Object.keys(fileData.data[0]);
    fileData._bcCol = cols.find(c => BARCODE_SYNONYMS.some(s => c.toLowerCase().includes(s.toLowerCase()))) || cols[0];
    fileData._nameCol = cols.find(c => c !== fileData._bcCol && NAME_SYNONYMS.some(s => c.toLowerCase().includes(s.toLowerCase())))
        || (cols.length > 1 ? cols[1] : cols[0]);
}

function autoDetectColumns() {
    allColumns = [];
    visibleColumns.clear();
    stockColumn = null;
    transitColumn = null;
    barcodeColumn = null;
    nameColumn = null;

    allFilesData.forEach(detectFileCols); // detect per-file columns

    if (allFilesData.length > 0 && allFilesData[0].data.length > 0) {
        // display columns (for table header) use file-0 names
        barcodeColumn = allFilesData[0]._bcCol;
        nameColumn    = allFilesData[0]._nameCol;

        // stock and transit columns only in myPrice
        if (myPriceData && myPriceData.data && myPriceData.data.length > 0) {
            const myCols = Object.keys(myPriceData.data[0]);
            stockColumn = myCols.find(c => STOCK_COL_SYNONYMS.some(s => String(c).toLowerCase().includes(s))) || null;
            transitColumn = myCols.find(c => TRANSIT_COL_SYNONYMS.some(s => String(c).toLowerCase().includes(s))) || null;
        }
    }

    allFilesData.forEach((fd) => {
        const { fileName, data } = fd;
        if (data.length > 0) {
            const fileColumns = Object.keys(data[0]);
            fileColumns.forEach(colName => {
                // FIX: skip per-file bc/name columns (not just file-0 names)
                if (colName === fd._bcCol || colName === fd._nameCol) return;
                if (fileName === MY_PRICE_FILE_NAME && stockColumn && colName === stockColumn) return;
                if (fileName === MY_PRICE_FILE_NAME && transitColumn && colName === transitColumn) return;
                const colKey = `${fileName}|${colName}`;
                allColumns.push({
                    fileName,
                    columnName: colName,
                    displayName: `${fileName} - ${colName}`,
                    key: colKey
                });
                visibleColumns.add(colKey);
            });
        }
    });


    // Колонка «Остаток» — если найдена в файле
    if (stockColumn) {
        const stockCol = { fileName: MY_PRICE_FILE_NAME, columnName: 'Остаток', displayName: 'Остаток', key: META_STOCK_KEY, metaType: 'stock' };
        allColumns.unshift(stockCol);
        visibleColumns.add(META_STOCK_KEY);
    }
    // Колонка «В пути» — если пользователь включил её (независимо от наличия в файле)
    if (showTransitColumn) {
        // Используем сохранённое имя (из _transitDisplayName при переименовании)
        const transitCol = { fileName: MY_PRICE_FILE_NAME, columnName: 'В пути', displayName: _transitDisplayName, key: META_TRANSIT_KEY, metaType: 'transit' };
        allColumns.unshift(transitCol);
        visibleColumns.add(META_TRANSIT_KEY);
    }
    // Сортируем и перестраиваем: мета → мой прайс → поставщики → кастомные
    {
        const meta = allColumns.filter(c => c.metaType && c.metaType !== 'custom');
        const myP  = allColumns.filter(c => !c.metaType && c.fileName === MY_PRICE_FILE_NAME);
        const sup  = allColumns.filter(c => !c.metaType && c.fileName !== MY_PRICE_FILE_NAME);
        const cust = customColumns.map(cc => ({ fileName: MY_PRICE_FILE_NAME, columnName: cc.displayName, displayName: cc.displayName, key: cc.key, metaType: 'custom' }));
        sup.sort((a,b)=>{const o={нал:0,бн:1,other:2};return(o[getColPayGroup(a)]??2)-(o[getColPayGroup(b)]??2);});
        allColumns = [...meta, ...myP, ...sup, ...cust];
    }
    customColumns.forEach(cc => visibleColumns.add(cc.key));
}

function toggleColumn(colKey) {
    if (visibleColumns.has(colKey)) visibleColumns.delete(colKey);
    else visibleColumns.add(colKey);
    updateHiddenColumnsPanel();
    processData();
    renderTable();
}

function updateHiddenColumnsPanel() {
    const hiddenCols = allColumns.filter(col => !visibleColumns.has(col.key));
    if (hiddenCols.length > 0) {
        let html = '';
        hiddenCols.forEach(col => {
            html += `<button class="restore-column-btn" onclick="toggleColumn('${col.key}')" title="Показать колонку ${col.displayName}">
                    ↩️ ${col.displayName}
                </button>`;
        });
        hiddenColumnsList.innerHTML = html;
        hiddenColumnsPanel.style.display = 'flex';
    } else {
        hiddenColumnsPanel.style.display = 'none';
    }
}

function processData() {
    const barcodeMap = new Map();

    allFilesData.forEach((fd) => {
        const { fileName, data, isMyPrice } = fd;
        // FIX: use per-file column names (different files may have different barcode/name column names)
        const fileBcCol   = fd._bcCol   || barcodeColumn;
        const fileNameCol = fd._nameCol || nameColumn;
        data.forEach((row, index) => {
            let rawBarcode = row[fileBcCol];
            if (!rawBarcode) return;

            const { canonical, wasSynonym } = canonicalizeBarcode(rawBarcode);
            const barcode = canonical;

            if (!barcodeMap.has(barcode)) {
                barcodeMap.set(barcode, {
                    barcode,
                    names: [],
                    values: new Map(),
                    isInMyPrice: false,
                    myPriceOrder: -1,
                    filesWithBarcode: new Set(),
                    namesByFile: new Map(),
                    originalBarcodesByFile: new Map(),
                    isSynonym: false
                });
            }

            const item = barcodeMap.get(barcode);
            item.filesWithBarcode.add(fileName);
            item.originalBarcodesByFile.set(fileName, rawBarcode);

            if (wasSynonym) {
                item.isSynonym = true;
            }

            if (isMyPrice) {
                item.isInMyPrice = true;
                item.myPriceOrder = index;
            }

            const currentRowName = row[fileNameCol];

                // Остаток только из моего прайса (если колонка найдена)
                if (isMyPrice && stockColumn) {
                    const stockVal = row[stockColumn];
                    if (!item.values.has(META_STOCK_KEY)) item.values.set(META_STOCK_KEY, []);
                    const arrStock = item.values.get(META_STOCK_KEY);
                    arrStock.length = 0;
                    arrStock.push({ val: (stockVal === undefined || stockVal === null) ? '' : stockVal, rowName: currentRowName, originalBarcode: rawBarcode, meta: true });
                }
                // В пути только из моего прайса (если колонка найдена)
                if (isMyPrice && transitColumn) {
                    const transitVal = row[transitColumn];
                    if (!item.values.has(META_TRANSIT_KEY)) item.values.set(META_TRANSIT_KEY, []);
                    const arrTransit = item.values.get(META_TRANSIT_KEY);
                    arrTransit.length = 0;
                    arrTransit.push({ val: (transitVal === undefined || transitVal === null) ? '' : transitVal, rowName: currentRowName, originalBarcode: rawBarcode, meta: true });
                }
            if (currentRowName) {
                const nameObj = {fileName, name: currentRowName};
                if (!item.names.some(n => n.fileName === fileName && n.name === currentRowName)) {
                    item.names.push(nameObj);
                }
                if (!item.namesByFile.has(fileName)) {
                    item.namesByFile.set(fileName, currentRowName);
                }

                if (isMyPrice) {
                    // Берём количество только из моего прайса, но ищем не только в nameColumn
                    const vals = Object.values(row).map(v => (v === undefined || v === null) ? '' : String(v));
                    for (const t of vals) {
                        const q = extractPackQtyFromName(t);
                        if (q) { item.packQty = q; break; }
                    }
                }
            }

            Object.keys(row).forEach(colName => {
                if (colName !== fileBcCol && colName !== fileNameCol) {
                    const key = `${fileName}|${colName}`;
                    const value = row[colName];
                    if (value !== undefined && value !== null && value !== '') {
                        if (!item.values.has(key)) {
                            item.values.set(key, []);
                        }
                        const arr = item.values.get(key);

                        // Для колонок с ценой: одинаковые цены не создают новую вариацию
                        if (isPriceLikeColumn(colName)) {
                            const exists = arr.some(v => samePrice(v.val, value));
                            if (!exists) {
                                arr.push({val: value, rowName: currentRowName, originalBarcode: rawBarcode});
                            }
                        } else {
                            arr.push({val: value, rowName: currentRowName, originalBarcode: rawBarcode});
                        }
                    }
                }
            });
        });
    });

    groupedData = Array.from(barcodeMap.values()).map(item => {
        const visibleCols = allColumns.filter(col => visibleColumns.has(col.key));
        const numericValues = [];
        visibleCols.forEach(col => {
            const valuesArr = item.values.get(col.key);
            if (valuesArr && valuesArr.length > 0) {
                valuesArr.forEach(vObj => {
                    const numValue = parseFloat(String(vObj.val).replace(/[^0-9.,]/g, '').replace(',', '.'));
                    if (!isNaN(numValue) && numValue > 0) {
                        numericValues.push(numValue);
                    }
                });
            }
        });



        // Авто-деление: если загружен "Мой прайс" и в названии найдено количество (например 30шт/24бл),
        // то делим все цены в строке, которые выше минимальной в 3 раза, на это количество.
        const hasMyPriceLoaded = !!myPriceData;
        const packQty = (hasMyPriceLoaded && item.packQty) ? item.packQty : null;
        let autoDivFactor = null;

        if (packQty) {
            // Автологика не должна зависеть от скрытия колонок в UI
            const cols2 = allColumns;

            // Колонки цен поставщиков (мой прайс исключён, meta-колонки исключены)
            const supplierPriceCols2 = cols2.filter(col => !col.metaType && col.fileName !== MY_PRICE_FILE_NAME && isPriceLikeColumn(col.columnName));

            // Колонки цен в моём прайсе (для исключения)
            const myPriceCols2 = cols2.filter(col => !col.metaType && col.fileName === MY_PRICE_FILE_NAME && isPriceLikeColumn(col.columnName));

            const myNums2 = [];
            myPriceCols2.forEach(col => {
                const arr = item.values.get(col.key);
                if (!arr || arr.length === 0) return;
                arr.forEach(v => {
                    const n = parsePriceNumber(v.val);
                    if (n !== null && n > 0) myNums2.push(n);
                });
            });
            const myMin2 = myNums2.length ? Math.min(...myNums2) : null;

            const supplierNums2 = [];
            supplierPriceCols2.forEach(col => {
                const arr = item.values.get(col.key);
                if (!arr || arr.length === 0) return;
                arr.forEach(v => {
                    const n = parsePriceNumber(v.val);
                    if (n !== null && n > 0) supplierNums2.push(n);
                });
            });

            if (supplierNums2.length > 0) {
                const minSupplier = Math.min(...supplierNums2);
                const thresholdSupplier = minSupplier * 3;

                // Исключение: если абсолютно все цены поставщиков >= (моя цена * 3),
                // то делим ВСЕ цены поставщиков на packQty.
                const allSuppliers3xAboveMy = (myMin2 !== null) && supplierNums2.every(n => n >= myMin2 * 3);

                let changed2 = false;
                supplierPriceCols2.forEach(col => {
                    const arr = item.values.get(col.key);
                    if (!arr || arr.length === 0) return;
                    arr.forEach(vObj => {
                        const n = parsePriceNumber(vObj.val);
                        if (n === null || n <= 0) return;

                        if (allSuppliers3xAboveMy || n >= thresholdSupplier) {
                            if (vObj._autoDiv) return;
                            vObj._origVal = vObj.val;
                            vObj.val = roundPrice(n / packQty);
                            vObj._autoDiv = true;
                            vObj._autoDivFactor = packQty;
                            changed2 = true;
                        }
                    });
                });

                if (changed2) autoDivFactor = packQty;
            }
        }
        let priceDiffPercent = 0;
        if (numericValues.length > 1) {
            const minPrice = Math.min(...numericValues);
            const maxPrice = Math.max(...numericValues);
            if (minPrice > 0) {
                priceDiffPercent = ((maxPrice - minPrice) / minPrice) * 100;
            }
        }

        const filesWithPrices = new Set();
        for (const [key, valuesArr] of item.values.entries()) {
            if (key.startsWith('__')) continue; // skip meta/custom keys
            if (valuesArr && valuesArr.length > 0) {
                const fileName = key.split('|')[0];
                if (fileName && fileName !== MY_PRICE_FILE_NAME) {
                    filesWithPrices.add(fileName);
                }
            }
        }
        const coverageCount = filesWithPrices.size;

                    // Гарантируем служебные колонки (пустые, если нет значения)
        if (stockColumn) {
            if (!item.values.has(META_STOCK_KEY)) {
                item.values.set(META_STOCK_KEY, [{ val: '', rowName: '', originalBarcode: item.barcode, meta: true }]);
            }
        }
        if (showTransitColumn) {
            // Restore user-edited transit value if present in customColData, else file value or empty
            const transitUserVal = customColData[META_TRANSIT_KEY] && customColData[META_TRANSIT_KEY][item.barcode] !== undefined
                ? customColData[META_TRANSIT_KEY][item.barcode] : null;
            if (transitUserVal !== null) {
                item.values.set(META_TRANSIT_KEY, [{ val: transitUserVal, rowName: '', originalBarcode: item.barcode, meta: true }]);
            } else if (!item.values.has(META_TRANSIT_KEY)) {
                item.values.set(META_TRANSIT_KEY, [{ val: '', rowName: '', originalBarcode: item.barcode, meta: true }]);
            }
        }
        // Гарантируем кастомные колонки: берём из customColData или пусто
        customColumns.forEach(cc => {
            const savedVal = customColData[cc.key] && customColData[cc.key][item.barcode] !== undefined
                ? customColData[cc.key][item.barcode] : '';
            item.values.set(cc.key, [{ val: savedVal, rowName: '', originalBarcode: item.barcode, meta: true }]);
        });

return { barcode: item.barcode, packQty, autoDivFactor,
            names: item.names,
            namesByFile: item.namesByFile,
            values: item.values,
            isInMyPrice: item.isInMyPrice,
            myPriceOrder: item.myPriceOrder,
            originalFileCount: item.filesWithBarcode.size,
            priceDiffPercent,
            coverageCount,
            isSynonym: item.isSynonym,
            originalBarcodesByFile: item.originalBarcodesByFile
        };
    });
}

// Сбросить все эксклюзивные кнопки фильтра (без рендера)
function _clearAllFilterBtns() {
    sortMode = 'default';
    filterNewItems = false;
    sortMatchesBtn.classList.remove('active');
    bigDiffBtn.classList.remove('active');
    showMyPriceBtn.classList.remove('active');
    maxCoverageBtn.classList.remove('active');
}

function toggleSortMatches() {
    if (sortMode === 'matches') {
        _clearAllFilterBtns();
    } else {
        _clearAllFilterBtns();
        sortMode = 'matches';
        sortMatchesBtn.classList.add('active');
    }
    renderTable();
}

function toggleBigDiff() {
    if (sortMode === 'bigdiff') {
        _clearAllFilterBtns();
    } else {
        _clearAllFilterBtns();
        sortMode = 'bigdiff';
        bigDiffBtn.classList.add('active');
    }
    renderTable();
}

function toggleMyPriceView() {
    if (sortMode === 'myprice') {
        _clearAllFilterBtns();
    } else {
        _clearAllFilterBtns();
        sortMode = 'myprice';
        showMyPriceBtn.classList.add('active');
    }
    renderTable();
}

function toggleMaxCoverage() {
    if (filterNewItems) {
        _clearAllFilterBtns();
    } else {
        if (!myPriceData) {
            showToast('Загрузите свой прайс — иначе нет смысла искать новинки', 'warn');
            return;
        }
        _clearAllFilterBtns();
        filterNewItems = true;
        maxCoverageBtn.classList.add('active');
    }
    renderTable();
}

function toggleCompactMatches() {
    compactMatches = !compactMatches;
    if (compactMatches) compactMatchesBtn.classList.add('active');
    else compactMatchesBtn.classList.remove('active');
    renderTable();

}

function toggleTransitColumn() {
    showTransitColumn = !showTransitColumn;
    const btn = document.getElementById('toggleTransitBtn');
    if (btn) btn.classList.toggle('active', showTransitColumn);
    autoDetectColumns();
    processData();
    renderTable();
    updateUI();
    showToast(showTransitColumn ? '«В пути» добавлена' : '«В пути» скрыта', 'ok');
}

function deleteCustomColumn(colKey) {
    if (!confirm('Удалить колонку? Все введённые данные будут потеряны.')) return;
    customColumns = customColumns.filter(c => c.key !== colKey);
    delete customColData[colKey];
    visibleColumns.delete(colKey);
    autoDetectColumns();
    processData();
    renderTable();
    updateUI();
    showToast('Колонка удалена', 'ok');
}

// Inline edit for custom/transit cells
function editCustomCell(barcode, colKey, cellDiv) {
    if (cellDiv.querySelector('.custom-cell-input')) return; // already editing
    const currentVal = (customColData[colKey] && customColData[colKey][barcode] !== undefined)
        ? customColData[colKey][barcode] : '';
    const spanVal = cellDiv.querySelector('.custom-cell-val');
    const spanBtn = cellDiv.querySelector('.custom-cell-edit-btn');
    if (spanVal) spanVal.style.display = 'none';
    if (spanBtn) spanBtn.style.display = 'none';
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.className = 'custom-cell-input';
    inp.value = currentVal;
    cellDiv.appendChild(inp);
    inp.focus();
    inp.select();
    const commit = () => {
        const newVal = inp.value.trim();
        if (!customColData[colKey]) customColData[colKey] = {};
        customColData[colKey][barcode] = newVal;
        // Update in groupedData to avoid full rerender
        const item = groupedData.find(i => i.barcode === barcode);
        if (item) item.values.set(colKey, [{ val: newVal, rowName: '', originalBarcode: barcode, meta: true }]);
        inp.remove();
        if (spanVal) {
            spanVal.style.display = '';
            const isEmpty = !newVal;
            spanVal.textContent = isEmpty ? '—' : newVal;
            spanVal.className = 'custom-cell-val' + (isEmpty ? ' empty' : '');
        }
        if (spanBtn) spanBtn.style.display = '';
    };
    inp.addEventListener('blur', commit);
    inp.addEventListener('keydown', e => {
        if (e.key === 'Enter') { e.preventDefault(); inp.blur(); }
        if (e.key === 'Escape') { inp.value = currentVal; inp.blur(); }
    });
}

function getSortedData() {
    let data = [...groupedData];

    if (searchQuery) {
        data = data.filter(item =>
            item.names.some(n => n.name.toLowerCase().includes(searchQuery))
        );
    }

    // Фильтр новинок: скрываем "мой прайс" и сортируем по охвату поставщиков (coverageCount desc)
    if (filterNewItems) {
        data = data.filter(item => !item.isInMyPrice);
        data.sort((a, b) => {
            if (b.coverageCount !== a.coverageCount) return b.coverageCount - a.coverageCount;
            if (b.originalFileCount !== a.originalFileCount) return b.originalFileCount - a.originalFileCount;
            const nameA = (a.names[0]?.name || '').toLowerCase();
            const nameB = (b.names[0]?.name || '').toLowerCase();
            return nameA.localeCompare(nameB, 'ru');
        });
    }

    if (sortMode === 'matches') {
        data.sort((a, b) => {
            if (a.originalFileCount > 1 && b.originalFileCount === 1) return -1;
            if (a.originalFileCount === 1 && b.originalFileCount > 1) return 1;
            return 0;
        });
    } else if (sortMode === 'bigdiff') {
        data = data.filter(item => item.originalFileCount > 1 && item.priceDiffPercent > 10);
        data.sort((a, b) => b.priceDiffPercent - a.priceDiffPercent);
    } else if (sortMode === 'myprice') {
        data = data.filter(item => item.isInMyPrice);
        data.sort((a, b) => a.myPriceOrder - b.myPriceOrder);
    } else if (sortMode === 'maxcoverage') {
        data.sort((a, b) => b.coverageCount - a.coverageCount);
    }

    return data;
}

function copyBarcode(barcode, btn) {
    if (!navigator.clipboard) return;
    navigator.clipboard.writeText(String(barcode)).then(() => {
        const orig = btn.textContent;
        btn.textContent = '✓';
        setTimeout(() => {
            btn.textContent = orig;
        }, 600);
    }).catch(() => {
    });
}

function copyBarcodeFromPrice(barcode) {
    if (!navigator.clipboard) return;
    navigator.clipboard.writeText(String(barcode)).catch(() => {});
}

function editColumnName(colKey) {
    const col = allColumns.find(c => c.key === colKey);
    if (!col) return;
    const newName = prompt('Новое название колонки:', col.displayName);
    if (newName && newName.trim() !== '' && newName.trim() !== col.displayName) {
        col.displayName = newName.trim();
        // Sync into customColumns array if it is a custom column
        const cc = customColumns.find(c => c.key === colKey);
        if (cc) cc.displayName = newName.trim();
        // Sync transit display name
        if (colKey === META_TRANSIT_KEY) _transitDisplayName = newName.trim();
        updateHiddenColumnsPanel();
        renderTable();
    }
}

// ── Virtual scroll helpers ───────────────────────────────────────────
function _mvsBuildHeader(visibleCols) {
    let h = `<tr>`;
    h += `<th class="col-barcode">Штрихкод</th>`;
    allFilesData.forEach(({fileName}, idx) => {
        const ec = showFileBarcodes ? '' : 'hidden-barcode-col';
        h += `<th class="col-barcode file-barcode-col ${ec}" data-file-index="${idx}">Штрихкод (${fileName})</th>`;
    });
    h += `<th class="col-name">${nameColumn}</th>`;
    visibleCols.forEach(col => {
        const _isMyP = !col.metaType && col.fileName === MY_PRICE_FILE_NAME;
        const _isMeta = !!col.metaType;
        const _isCustom = col.metaType === 'custom';
        const _isTransit = col.key === META_TRANSIT_KEY;
        const _cL = col.metaType ? col.displayName : col.columnName;
        const _fL = col.metaType ? null : col.fileName;
        const _ck = col.key.replace(/'/g, "\\'");
        let _actions = '';
        if (_isMeta || _isMyP) {
            _actions += `<div class="col-header-actions">`;
            _actions += `<button class="col-header-btn" onclick="event.stopPropagation();editColumnName('${_ck}')" title="Переименовать">✏️</button>`;
            if (_isTransit) _actions += `<button class="col-header-btn col-header-btn--del" onclick="event.stopPropagation();toggleTransitColumn()" title="Скрыть В пути">✕</button>`;
            if (_isCustom) _actions += `<button class="col-header-btn col-header-btn--del" onclick="event.stopPropagation();deleteCustomColumn('${_ck}')" title="Удалить колонку">✕</button>`;
            _actions += `</div>`;
        }
        if (_isMyP || _isMeta) {
            h += `<th class="${_isMyP ? 'col-my-price' : 'col-meta'}" title="${MY_PRICE_FILE_NAME} — ${_cL}"><div class="column-header"><div class="column-file-name column-file-name--my-price">${MY_PRICE_FILE_NAME}</div><div class="column-header-title"><span class="column-name-text">${_cL}</span></div>${_actions}</div></th>`;
        } else {
            const _hideAct = `<div class="col-header-actions"><button class="col-header-btn hide-col-btn col-header-btn--del" onclick="event.stopPropagation();toggleColumn('${_ck}')" title="Скрыть колонку">✕</button></div>`;
            h += `<th title="${_fL} — ${_cL}"><div class="column-header"><div class="column-file-name" title="${_fL}">${_fL}</div><div class="column-header-title"><span class="column-name-text">${_cL}</span></div>${_hideAct}</div></th>`;
        }
    });
    h += `</tr>`;
    return h;
}
function _mvsRenderRow(item, visibleCols, supplierPriceCols, colPayGroupMap) {
    let rowClass = '';
    if (item.isSynonym) rowClass = 'synonym-row';
    else if (item.isInMyPrice) rowClass = 'my-price-row';
    rowClass += (rowClass ? ' ' : '') + 'group-border-top group-border-bottom';

    let html = `<tr class="${rowClass}" data-barcode="${item.barcode}" data-in-my-price="${item.isInMyPrice?'1':'0'}" data-is-synonym="${item.isSynonym?'1':'0'}">`;
    // Barcode cell: show DB badge if already in jeDB, otherwise show quick-add button
    const _bcInDB = typeof jeDB !== 'undefined' && (jeDB[item.barcode] !== undefined);
    const _bcBadge = _bcInDB
      ? `<span class="bc-in-db-badge" title="Штрихкод уже есть в базе синонимов">📚</span>`
      : `<button class="bc-add-db-btn" title="Добавить в базу синонимов" onclick="openAddToDB('${item.barcode.replace(/'/g,"\\'").replace(/"/g,'&quot;')}',this)">+</button>`;
    html += `<td class="col-barcode"><div class="barcode-cell"><span class="barcode-text" title="${item.barcode}">${item.barcode}</span><button class="copy-btn" onclick="copyBarcode('${item.barcode}',this)">📋</button>${_bcBadge}</div></td>`;

    allFilesData.forEach(({fileName}, idx) => {
        const ec = showFileBarcodes ? '' : 'hidden-barcode-col';
        const ob = item.originalBarcodesByFile.get(fileName) || '—';
        html += `<td class="col-barcode file-barcode-col ${ec}" data-file-index="${idx}"><div class="barcode-cell"><span class="barcode-text">${ob}</span>${ob!=='—'?`<button class="copy-btn" onclick="copyBarcode('${ob}',this)">📋</button>`:''}</div></td>`;
    });

    // Name cell
    if (compactMatches && item.names.length > 1) {
        // Compact: show only the first unique name; tooltip shows all
        const _nmC = new Map();
        item.names.forEach(n => { if (!_nmC.has(n.name)) _nmC.set(n.name, n.fileName); });
        const _firstName = [..._nmC.keys()][0];
        const _allNames = [..._nmC.keys()].join(' | ');
        const _extraCount = _nmC.size - 1;
        html += `<td class="col-name"><div class="name-compact" title="${esc(_allNames)}">${esc(_firstName)}<span style="color:#999;font-size:10px;margin-left:4px;">(+${_extraCount})</span></div></td>`;
    } else if (item.names.length > 0) {
        html += `<td class="col-name"><div class="name-cell">`;
        const _nm = new Map();
        item.names.forEach(n => { if (!_nm.has(n.name)) _nm.set(n.name, n.fileName); });
        _nm.forEach((fn, name) => { html += `<div class="name-item" title="📁 ${esc(fn)}">${esc(name)}</div>`; });
        html += `</div></td>`;
    } else {
        html += `<td class="col-name">Без названия</td>`;
    }

    // Price computation — FIX: supplierPriceCols computed once per render, not per row
    const numericValues = [];
    const _gn = {'нал':[], 'бн':[], 'other':[]};
    // Track unique supplier files contributing prices (to avoid false min highlighting with 1 supplier)
    const _supFilesWithPrice = new Set();
    supplierPriceCols.forEach(col => {
        const _g = colPayGroupMap.get(col.key) || 'other';
        const valuesArr = item.values.get(col.key);
        if (valuesArr) valuesArr.forEach(vObj => {
            const n = parsePriceNumber(vObj.val);
            if (n !== null && n > 0) { numericValues.push(n); _gn[_g].push(n); _supFilesWithPrice.add(col.fileName); }
        });
    });
    // Only highlight min if prices from 2+ distinct supplier files exist
    const _multiSuppliers = _supFilesWithPrice.size > 1;
    const _gMin = {
        'нал': _gn['нал'].length > 0 ? Math.min(..._gn['нал']) : null,
        'бн':  _gn['бн'].length  > 0 ? Math.min(..._gn['бн'])  : null,
        'other': _gn['other'].length > 0 ? Math.min(..._gn['other']) : null
    };
    // _gM: true only if multiple distinct supplier files in that pay group
    const _gFilesPerGroup = {'нал': new Set(), 'бн': new Set(), 'other': new Set()};
    supplierPriceCols.forEach(col => {
        const _g = colPayGroupMap.get(col.key) || 'other';
        const valuesArr = item.values.get(col.key);
        if (valuesArr && valuesArr.some(v => { const n = parsePriceNumber(v.val); return n !== null && n > 0; })) {
            _gFilesPerGroup[_g].add(col.fileName);
        }
    });
    const _gM = { 'нал': _gFilesPerGroup['нал'].size > 1, 'бн': _gFilesPerGroup['бн'].size > 1, 'other': _gFilesPerGroup['other'].size > 1 };
    const globalMin = numericValues.length > 0 ? Math.min(...numericValues) : null;
    const globalMax = numericValues.length > 0 ? Math.max(...numericValues) : null;
    const hasMultipleGlobals = _multiSuppliers && numericValues.length > 1;

    visibleCols.forEach(col => {
        const valuesArr = item.values.get(col.key);
        let cellContent = '—';
        if (col.metaType) {
            const _isEditable = col.metaType === 'custom' || col.key === META_TRANSIT_KEY;
            const _mv = (valuesArr && valuesArr.length > 0) ? valuesArr[0].val : '';
            const _mvStr = (_mv === undefined || _mv === null) ? '' : String(_mv).trim();
            if (_isEditable) {
                // Editable custom/transit cell
                const _ck = col.key.replace(/'/g, "\'");
                const _bc = item.barcode.replace(/'/g, "\'");
                const _display = _mvStr || '';
                const _isEmpty = !_display;
                html += `<td><div class="custom-cell" onclick="editCustomCell('${_bc}','${_ck}',this)"><span class="custom-cell-val${_isEmpty?' empty':''}">${_isEmpty?'—':_display}</span><span class="custom-cell-edit-btn">✎</span></div></td>`;
                return;
            }
            if (!_mvStr) {
                cellContent = '—';
            } else {
                const _mn = parseFloat(_mvStr.replace(',', '.'));
                cellContent = !isNaN(_mn) ? String(Math.round(_mn)) : _mvStr;
            }
            html += `<td>${cellContent}</td>`; return;
        }
        if (valuesArr && valuesArr.length > 0) {
            cellContent = '<div class="multi-value-container">';
            valuesArr.forEach((vObj, vIndex) => {
                let displayValue, isMin = false, isMax = false, numValue = null;
                const parsed = parseFloat(String(vObj.val).replace(/[^0-9.,]/g, '').replace(',', '.'));
                if (!isNaN(parsed) && parsed > 0) {
                    numValue = parsed;
                    displayValue = parsed.toFixed(PRICE_DECIMALS).replace(/\.0+$/, '');
                    const _pg = colPayGroupMap.get(col.key) || 'other';
                    if (_gM[_pg] && _gMin[_pg] !== null && parsed === _gMin[_pg]) isMin = true;
                    if (hasMultipleGlobals && parsed === globalMax && globalMax >= 3 * globalMin) isMax = true;
                } else { displayValue = vObj.val; }
                const barcodeForCopy = vObj.originalBarcode || item.barcode;
                const autoBadge = vObj._autoDiv ? `<span class="auto-div-badge" title="Автоделение /${vObj._autoDivFactor || item.packQty}">÷</span>` : '';
                let innerHtml;
                if (isMin) {
                    innerHtml = `<span class="price-val is-min price-clickable" onclick="copyBarcodeFromPrice('${barcodeForCopy}')">${displayValue}</span>${autoBadge}`;
                } else if (isMax && numValue) {
                    // FIX: _DIV_OPTIONS cached once, not regenerated per cell
                    innerHtml = `<span class="price-clickable" onclick="copyBarcodeFromPrice('${barcodeForCopy}')">${displayValue}</span>${autoBadge}<div class="div-wrapper" title="Цена указана за блок?"><div class="div-icon">÷</div><select class="div-select" onchange="dividePrice('${item.barcode}','${col.key}',${vIndex},this.value);this.value=''"><option value="" disabled selected>÷</option>${_DIV_OPTIONS}</select></div>`;
                } else {
                    innerHtml = `<span class="price-clickable" onclick="copyBarcodeFromPrice('${barcodeForCopy}')">${displayValue}</span>${autoBadge}`;
                }
                cellContent += `<div class="value-variant">${innerHtml}</div>`;
            });
            cellContent += '</div>';
        }
        html += `<td>${cellContent}</td>`;
    });
    html += '</tr>';
    return html;
}

function _mvsRenderVisible() {
    const wrap = document.getElementById('mainTableWrap');
    if (!wrap) return;
    const total = _vsData.length;
    if (!total) return;
    const scrollTop = wrap.scrollTop;
    const viewH = wrap.clientHeight || 600;
    MVS.start = Math.max(0, Math.floor(scrollTop / MVS.ROW_H) - MVS.OVERSCAN);
    MVS.end = Math.min(total, Math.ceil((scrollTop + viewH) / MVS.ROW_H) + MVS.OVERSCAN);
    const topPad = MVS.start * MVS.ROW_H;
    const botPad = Math.max(0, total - MVS.end) * MVS.ROW_H;
    const tbody = document.getElementById('mainTbody');
    if (!tbody) return;
    const colSpan = 3 + allFilesData.length + _vsVisibleCols.length;
    let rows = '';
    for (let i = MVS.start; i < MVS.end; i++) {
        rows += _mvsRenderRow(_vsData[i], _vsVisibleCols, _vsSupplierPriceCols, _vsColPayGroupMap);
    }
    tbody.innerHTML =
        (topPad > 0 ? `<tr style="height:${topPad}px;border:none;pointer-events:none;"><td colspan="${colSpan}" style="padding:0;border:none;"></td></tr>` : '') +
        rows +
        (botPad > 0 ? `<tr style="height:${botPad}px;border:none;pointer-events:none;"><td colspan="${colSpan}" style="padding:0;border:none;"></td></tr>` : '');
}

function renderTable(preserveScroll = false) {
    const dataToShow = getSortedData();
    _vsData = dataToShow;

    if (dataToShow.length === 0) {
        tableContainer.innerHTML = `<div class="empty-state"><div class="empty-state-icon">⚠️</div><h3>Нет данных для отображения</h3><p>Проверьте содержимое загруженных файлов или измените фильтры</p></div>`;
        return;
    }

    // Pre-compute per-render caches — FIX: computed once, not per row
    _vsVisibleCols = allColumns.filter(col => visibleColumns.has(col.key));
    _vsSupplierPriceCols = _vsVisibleCols.filter(col => !col.metaType && col.fileName !== MY_PRICE_FILE_NAME && isPriceLikeColumn(col.columnName));
    _vsColPayGroupMap = new Map();
    _vsVisibleCols.forEach(col => _vsColPayGroupMap.set(col.key, getColPayGroup(col)));

    // Build or update table structure
    let wrap = document.getElementById('mainTableWrap');
    if (!wrap) {
        // First render: create persistent structure
        tableContainer.innerHTML = `
            <div id="mainTableWrap" style="overflow-y:auto;overflow-x:auto;max-height:75vh;border:1px solid #d0d0d0;border-radius:4px;" class="table-wrapper">
                <table id="mainTable" style="width:100%;border-collapse:collapse;min-width:700px;">
                    <thead id="mainThead"></thead>
                    <tbody id="mainTbody"></tbody>
                </table>
            </div>`;
        wrap = document.getElementById('mainTableWrap');
        // Attach scroll handler for virtual scroll
        wrap.addEventListener('scroll', () => {
            if (!MVS.ticking) {
                MVS.ticking = true;
                requestAnimationFrame(() => { _mvsRenderVisible(); MVS.ticking = false; });
            }
        }, { passive: true });
    }

    // Always rebuild header (sort icon, columns may change)
    document.getElementById('mainThead').innerHTML = _mvsBuildHeader(_vsVisibleCols);

    // Reset scroll to top on filter/sort changes; preserve on in-place edits
    const _prevScroll = preserveScroll ? wrap.scrollTop : 0;
    wrap.scrollTop = 0;
    MVS.start = 0; MVS.end = 0;

    _mvsRenderVisible();

    if (preserveScroll && _prevScroll > 0) {
        requestAnimationFrame(() => { wrap.scrollTop = _prevScroll; });
    }
}

    function dividePrice(barcode, colKey, valueIndex, factorStr) {
    const factor = parseFloat(factorStr);
    if (!factor || factor <= 0) return;

    const item = groupedData.find(x => String(x.barcode) === String(barcode));
    if (!item) return;

    const visibleCols = allColumns.filter(col => visibleColumns.has(col.key));
const priceCols = visibleCols.filter(col => !col.metaType && isPriceLikeColumn(col.columnName));
const nums = [];
    priceCols.forEach(col => {
        const arr = item.values.get(col.key);
        if (!arr || arr.length === 0) return;
        arr.forEach(v => {
            const n = parsePriceNumber(v.val);
            if (n !== null && n > 0) nums.push(n);
        });
    });
    if (nums.length === 0) return;

    const minValue = Math.min(...nums);
    const threshold = minValue * 3;

    let changed = false;
    priceCols.forEach(col => {
        const arr = item.values.get(col.key);
        if (!arr || arr.length === 0) return;
        arr.forEach(vObj => {
            const n = parsePriceNumber(vObj.val);
            if (n === null || n <= 0) return;
            if (n > threshold) {
                vObj.val = roundPrice(n / factor);
                changed = true;
            }
        });
    });

    if (!changed) return;

    renderTable(true); // preserve scroll: editing in-place

}

// ── helpers ──────────────────────────────────────────────────────────────

// Проверить разрешение на запись в сохранённую директорию
// Простое сохранение через стандартный диалог браузера
async function saveBlobWithDialogOrDownload(blob, fileName) {
    // Пробуем showSaveFilePicker (Chrome 86+, Edge 86+) — даёт выбор папки
    if (window.showSaveFilePicker) {
        const ext = fileName.split('.').pop().toLowerCase();
        const mimeMap = {
            xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            json: 'application/json',
            csv:  'text/csv',
        };
        try {
            const handle = await window.showSaveFilePicker({
                suggestedName: fileName,
                types: [{ description: fileName, accept: { [mimeMap[ext] || 'application/octet-stream']: ['.' + ext] } }],
            });
            const writable = await handle.createWritable();
            await writable.write(blob);
            await writable.close();
            return;
        } catch (e) {
            if (e.name === 'AbortError') return; // пользователь нажал «Отмена»
            // Иначе — фолбэк на обычное скачивание
        }
    }
    // Фолбэк: скачивание в папку «Загрузки» (Firefox, Safari, старые браузеры)
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function generateExcel(mode) {
try {
const _exMeta=allColumns.filter(c=>c.metaType&&c.metaType!=='custom'&&visibleColumns.has(c.key));
const _exMyP =allColumns.filter(c=>!c.metaType&&c.fileName===MY_PRICE_FILE_NAME&&visibleColumns.has(c.key));
const _exSup =allColumns.filter(c=>!c.metaType&&c.fileName!==MY_PRICE_FILE_NAME&&visibleColumns.has(c.key));
const _exCust=allColumns.filter(c=>c.metaType==='custom'&&visibleColumns.has(c.key));
const excelCols=[..._exMeta,..._exMyP,..._exSup,..._exCust];
const fileNames=allFilesData.map(f=>f.fileName);
const hasMyPrice=!!myPriceData;
const myPriceFileName=hasMyPrice?MY_PRICE_FILE_NAME:null;
const nameFileOrder=[];
if(hasMyPrice) nameFileOrder.push(myPriceFileName);
fileNames.forEach(fn=>{if(!hasMyPrice||fn!==myPriceFileName)nameFileOrder.push(fn);});
const priceStartColBase=1+fileNames.length+nameFileOrder.length;

// Жирные разделители: мой прайс→поставщики + нал↔бн
const thickLeftAt=new Set();
const _fsi=excelCols.findIndex(c=>!c.metaType&&c.fileName!==MY_PRICE_FILE_NAME);
if(_fsi>0) thickLeftAt.add(_fsi);
for(let i=1;i<excelCols.length;i++){
    const p=excelCols[i-1],c=excelCols[i];
    if(!c.metaType&&c.fileName!==MY_PRICE_FILE_NAME&&!p.metaType&&p.fileName!==MY_PRICE_FILE_NAME)
        if(getColPayGroup(p)!==getColPayGroup(c)) thickLeftAt.add(i);
}

let dataToExport=[];
if(mode==='myprice') dataToExport=groupedData.filter(item=>item.isInMyPrice).sort((a,b)=>a.myPriceOrder-b.myPriceOrder);
else if(mode==='current') dataToExport=getSortedData();
else                      dataToExport=groupedData;

const workbook=new ExcelJS.Workbook();
const worksheet=workbook.addWorksheet('Сравнение');
const totalCols=1+fileNames.length+nameFileOrder.length+excelCols.length;

// Заголовки: мой прайс — только колонка, поставщик — файл+колонка
const headers=['Штрихкод'];
fileNames.forEach(fn=>headers.push(`Штрихкод (${fn})`));
nameFileOrder.forEach(fn=>headers.push('Наименование'));
excelCols.forEach(col=>{
    const _isMyP=!col.metaType&&col.fileName===MY_PRICE_FILE_NAME;
    if(_isMyP||col.metaType){
        headers.push(col.metaType?col.displayName:col.columnName);
    } else {
        headers.push(col.fileName+'\n'+col.columnName);
    }
});
const _xlH=worksheet.addRow(headers);
_xlH.alignment={vertical:'middle',horizontal:'center',wrapText:true};
_xlH.height=45;

// Данные
dataToExport.forEach(item=>{
    {
        const row=[item.barcode];
        fileNames.forEach(fn=>row.push(item.originalBarcodesByFile.get(fn)||''));
        if(mode==='current'){
            nameFileOrder.forEach(fn=>row.push((item.namesByFile&&item.namesByFile.get(fn))||''));
        } else {
            nameFileOrder.forEach(fn=>row.push((item.namesByFile&&item.namesByFile.get(fn))||''));
        }
        const priceStartCol=row.length;
        const numericColsInRow=[];
        const _eg={'нал':[],'бн':[],'other':[]};
        const _ei={'нал':[],'бн':[],'other':[]};
        // FIX: track distinct supplier files per pay group, highlight min only if 2+ files
        const _egFiles={'нал':new Set(),'бн':new Set(),'other':new Set()};
        const _multiValCols=new Set();
        excelCols.forEach((col,idx)=>{
            const va=item.values.get(col.key);let cellValue='';
            if(va&&va.length>0){
                if(!col.metaType&&isPriceLikeColumn(col.columnName)){
                    const uniquePrices=[];
                    va.forEach(vObj=>{
                        const num=parseFloat(String(vObj.val).replace(/[^0-9.,]/g,'').replace(',','.'));
                        if(!isNaN(num)&&num>0){const r=roundPrice(num);if(!uniquePrices.includes(r))uniquePrices.push(r);}
                    });
                    if(uniquePrices.length===1){
                        cellValue=uniquePrices[0];
                        numericColsInRow.push(priceStartCol+idx);
                        if(col.fileName!==MY_PRICE_FILE_NAME){
                            const _g=getColPayGroup(col);
                            _eg[_g].push(uniquePrices[0]);
                            _ei[_g].push({ci:priceStartCol+idx,vi:_eg[_g].length-1});
                            _egFiles[_g].add(col.fileName); // track supplier file
                        }
                    } else if(uniquePrices.length>1){
                        cellValue=uniquePrices.map(p=>String(p).replace('.',',')).join(' / ');
                        _multiValCols.add(priceStartCol+idx);
                        if(col.fileName!==MY_PRICE_FILE_NAME){
                            const _g=getColPayGroup(col);
                            // FIX: add all prices to comparison pool for correct global min detection.
                            // Track the cell once with its minimum price as representative:
                            // the cell highlights only if its best (minimum) price equals the global minimum.
                            const cellMin=Math.min(...uniquePrices);
                            const cellMinIdx=_eg[_g].length; // index of cellMin in _eg
                            uniquePrices.forEach(p=>{ _eg[_g].push(p); });
                            // _ei entry uses the position where cellMin was just pushed
                            _ei[_g].push({ci:priceStartCol+idx,vi:cellMinIdx});
                            _egFiles[_g].add(col.fileName);
                        }
                    } else {
                        cellValue=va[0]?.val||'';
                    }
                } else if(col.metaType){
                    const num=parseFloat(String(va[0].val).replace(/[^0-9.,]/g,'').replace(',','.'));
                    cellValue=(!isNaN(num)&&num>=0)?num:(va[0].val??'');
                } else {
                    cellValue=va[0]?.val||'';
                }
            }
            row.push(cellValue);
        });
        const excelRow=worksheet.addRow(row);
        numericColsInRow.forEach(ci=>{const c=excelRow.getCell(ci+1);if(typeof c.value==='number')c.numFmt='0.0';});
        if(item.isSynonym){excelRow.getCell(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE6F7FF'}};}
        ['нал','бн','other'].forEach(g=>{
            if(_eg[g].length>1 && _egFiles[g].size>1){
                const mn=Math.min(..._eg[g]);
                const minCells=new Set();
                _ei[g].forEach(({ci,vi})=>{if(_eg[g][vi]===mn)minCells.add(ci);});
                minCells.forEach(ci=>{
                    const cell=excelRow.getCell(ci+1);
                    cell.font={color:{argb:'FFDC2626'},bold:true};
                });
            }
        });
    }
});

// Ширины + группировка
const fbS=2,fbE=1+fileNames.length,nsS=fbE+1,nsE=nsS+nameFileOrder.length-1;
worksheet.columns.forEach((col,idx)=>{
    const ci=idx+1;
    if(ci===1)col.width=15;
    else if(ci>=fbS&&ci<=fbE)col.width=13;
    else if(ci>=nsS&&ci<=nsE)col.width=(hasMyPrice&&nameFileOrder[0]===myPriceFileName&&ci===nsS)?50:20;
    else col.width=10;
});
if(fileNames.length>0){worksheet.properties.outlineProperties={summaryBelow:true};for(let c=fbS;c<=fbE;c++){worksheet.getColumn(c).outlineLevel=1;worksheet.getColumn(c).hidden=true;}}
if(nameFileOrder.length>1){worksheet.properties.outlineProperties={summaryBelow:true};const st=hasMyPrice?nsS+1:nsS;for(let c=st;c<=nsE;c++){worksheet.getColumn(c).outlineLevel=1;worksheet.getColumn(c).hidden=true;}}
worksheet.views=[{state:'frozen',xSplit:0,ySplit:1}];

// Границы: тонкие + жирные разделители + жирная внешняя рамка
const totalRows=worksheet.rowCount;
worksheet.eachRow((row,rowNum)=>{
    row.eachCell({includeEmpty:true},(cell,colNum)=>{
        const _eci=colNum-1-priceStartColBase;
        cell.border={
            top:   {style:rowNum===1?'medium':'thin'},
            left:  {style:(colNum===1||(_eci>=0&&thickLeftAt.has(_eci)))?'medium':'thin'},
            bottom:{style:rowNum===totalRows?'medium':'thin'},
            right: {style:colNum===totalCols?'medium':'thin'}
        };
    });
});

// Заголовок оформление
const headerRow=worksheet.getRow(1);
headerRow.font={bold:true};
headerRow.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFE0E0E0'}};
headerRow.alignment={vertical:'middle',horizontal:'center',wrapText:true};
headerRow.height=45;
if(fbS<=headers.length)headerRow.getCell(fbS).note='Штрихкоды по файлам';
if(nsS<=headers.length)headerRow.getCell(nsS).note='Наименования по файлам';

const buffer=await workbook.xlsx.writeBuffer();
const now=new Date(),yyyy=now.getFullYear(),mm=String(now.getMonth()+1).padStart(2,'0'),dd=String(now.getDate()).padStart(2,'0');
let fileName=`monitoring-${yyyy}-${mm}-${dd}.xlsx`;
if(mode==='myprice') fileName=`monitoring-myprice-${yyyy}-${mm}-${dd}.xlsx`;
if(mode==='current') fileName=`monitoring-current-${yyyy}-${mm}-${dd}.xlsx`;
const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
await saveBlobWithDialogOrDownload(blob,fileName);
} catch(err) {
    console.error('generateExcel error:', err);
    alert('Ошибка при создании Excel-файла:\n' + (err.message || err));
}
}


function updateUI() {
    const hasData = groupedData.length > 0;
    const hasMyPrice = myPriceData !== null;
    exportAllBtn.disabled = !hasData;
    exportCurrentBtn.disabled = !hasData;
    exportMyPriceBtn.disabled = !hasMyPrice || !hasData;
    clearBtn.disabled = !hasData;
    sortMatchesBtn.disabled = !hasData;
    bigDiffBtn.disabled = !hasData;
    showMyPriceBtn.disabled = !hasMyPrice || !hasData;
    compactMatchesBtn.disabled = !hasData;
    maxCoverageBtn.disabled = !hasData;
    searchInput.disabled = !hasData;

    // Info panel always visible — update values
    if (hasData) {
        const matchCount = groupedData.filter(item => item.originalFileCount > 1).length;
        document.getElementById('productCount').textContent = groupedData.length;
        document.getElementById('fileCount').textContent = allFilesData.length;
        document.getElementById('columnCount').textContent = allColumns.length;
        document.getElementById('matchCount').textContent = matchCount;
        if (typeof matcherFileChipsRender === 'function') matcherFileChipsRender();
    } else {
        document.getElementById('productCount').textContent = '—';
        document.getElementById('fileCount').textContent = '—';
        document.getElementById('columnCount').textContent = '—';
        document.getElementById('matchCount').textContent = '—';
        const _lp2=document.getElementById('legendPanel');if(_lp2)_lp2.style.display='none';
    }
}

function clearAll() {
    myPriceData = null;
    competitorFilesData = [];
    allFilesData = [];
    groupedData = [];
    allColumns = [];
    visibleColumns.clear();
    barcodeColumn = null;
    nameColumn = null;
    stockColumn = null;
    transitColumn = null;
    showTransitColumn = false;
    customColumns = [];
    customColData = {};
    if (document.getElementById('toggleTransitBtn')) document.getElementById('toggleTransitBtn').classList.remove('active');
    _transitDisplayName = 'В пути';
    sortMode = 'default';
    compactMatches = false;
    searchQuery = '';
    showFileBarcodes = false;
    filterNewItems = false;

    myPriceInput.value = '';
    competitorInput.value = '';
    synonymsInput.value = '';
    const _mpSt2=document.getElementById('myPriceStatus');if(_mpSt2){_mpSt2.className='upload-status upload-status--idle';_mpSt2.textContent='Не загружен';}
    const _cSt2=document.getElementById('competitorStatus');if(_cSt2){_cSt2.className='upload-status upload-status--idle';_cSt2.textContent='Не загружены';}
    const _snSt2=document.getElementById('synonymsStatus');if(_snSt2){_snSt2.className='upload-status upload-status--idle';_snSt2.textContent='Не загружены';}
    // Полная очистка базы синонимов и брендов
    jeDB = {}; _jeDupsCache = null; jeChanges = 0;
    jeUndoStack = []; jeRedoStack = [];
    if (typeof jeUpdateUndoUI === 'function') jeUpdateUndoUI();
    if (typeof jeUpdateStatus === 'function') jeUpdateStatus();
    if (typeof jeRenderEditor === 'function') jeRenderEditor();
    if (typeof _brandDB !== 'undefined') { _brandDB = {}; } // window._brandDB было ошибкой — let-переменная не на window
    if (typeof brandRender === 'function') brandRender();
    if (typeof unifiedMarkUnsaved === 'function') unifiedMarkUnsaved(false);
    resetBarcodeAliases();
    // Сбрасываем bcCountBadge и синоним-статус до 0 (rebuildBarcodeAliasFromJeDB не вызывался при clearAll)
    if (typeof rebuildBarcodeAliasFromJeDB === 'function') rebuildBarcodeAliasFromJeDB();
    searchInput.value = '';
    sortMatchesBtn.classList.remove('active');
    showMyPriceBtn.classList.remove('active');
    compactMatchesBtn.classList.remove('active');
    maxCoverageBtn.classList.remove('active');
    // toggleFileBarcodesBtn удалён из UI — строка убрана (была ReferenceError)
    hiddenColumnsPanel.style.display = 'none';

    tableContainer.innerHTML = `
        <div class="empty-state">
            <div class="empty-state-icon">📦</div>
            <h3>Загрузите прайс-листы</h3>
            <p>Загрузите ваш прайс и файлы конкурентов для автоматического сравнения</p>
        </div>`;
    const _lp=document.getElementById('obr-loaded-files');
    if(_lp){_lp.style.display='none';}
    const _ll=document.getElementById('obr-loaded-list');
    if(_ll){_ll.innerHTML='';}
    if(typeof _sfUpdateMyPrice==='function')_sfUpdateMyPrice(null,null);
    if(typeof _sfUpdateSuppliers==='function')_sfUpdateSuppliers([]);
    if(typeof _sfUpdateJson==='function')_sfUpdateJson(null,null);
    updateUI();

    // Сброс результатов матчинга
    if (typeof _matchActivePairs !== 'undefined') _matchActivePairs = [];
    if (typeof _matchKnownPairs !== 'undefined') _matchKnownPairs = [];
    if (typeof _matchAllItems !== 'undefined') _matchAllItems = [];
    if (typeof _matchRenderedPairs !== 'undefined') _matchRenderedPairs = [];
    if (typeof _matchCurrentView !== 'undefined') _matchCurrentView = 'all';
    const _mWrap = document.getElementById('matcherTableWrap');
    if (_mWrap) { _mWrap.style.display = 'none'; if (_mWrap._mvsRender) _mWrap._mvsRender = null; }
    const _mEmpty = document.getElementById('matcherEmpty');
    if (_mEmpty) { _mEmpty.style.display = ''; _mEmpty.querySelector('h3').textContent = 'Запустите матчинг'; _mEmpty.querySelector('p').textContent = 'Загрузите прайсы на вкладке «Мониторинг», затем нажмите «Запустить матчинг»'; }
    const _mTbody = document.getElementById('matcherTbody');
    if (_mTbody) _mTbody.innerHTML = '';
    const _mStats = document.getElementById('matcherStats');
    if (_mStats) { _mStats.style.display = 'none'; ['ms-all','ms-high','ms-mid'].forEach(id => { const el = document.getElementById(id); if (el) el.textContent = '0'; }); }
    const _mProg = document.getElementById('matcherProgress');
    if (_mProg) _mProg.style.display = 'none';
    const _mBtn = document.getElementById('matcherRunBtn');
    if (_mBtn) { _mBtn.disabled = false; _mBtn.textContent = '▶ Запустить матчинг'; }
    const _mSearch = document.getElementById('matcherSearchInp');
    if (_mSearch) _mSearch.value = '';
    document.querySelectorAll('.mstat[data-mv]').forEach(s => s.classList.toggle('active', s.dataset.mv === 'all'));
}
async function downloadCurrentSynonyms(){
  // Export directly from jeDB — the single source of truth
  // Also include OBR columnSettings if available
  const combined = {
    barcodes: jeDB,
    brands: typeof _brandDB !== 'undefined' ? _brandDB : {},
    columnSettings: (typeof columnTemplates !== 'undefined' && typeof columnSynonyms !== 'undefined') ? {
      templates: columnTemplates,
      synonyms: columnSynonyms
    } : undefined
  };
  const blob = new Blob([JSON.stringify(combined, null, 2)], { type: 'application/json' });
  const now = new Date();
  const fname = `settings_${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}.json`;
  await saveBlobWithDialogOrDownload(blob, fname);
}
window.toggleColumn = toggleColumn;
window.copyBarcode = copyBarcode;
window.copyBarcodeFromPrice = copyBarcodeFromPrice;
window.dividePrice = dividePrice;
window.editColumnName = editColumnName;
window.editCustomCell = editCustomCell;
window.toggleTransitColumn = toggleTransitColumn;
window.deleteCustomColumn = deleteCustomColumn;

// ── Expose core monitor state/functions for AppBridge ─────────────────
window._pmApp = {
  get myPriceData() { return myPriceData; },
  set myPriceData(v) { myPriceData = v; },
  get competitorFilesData() { return competitorFilesData; },
  addCompetitorFile(fd) {
    // Check for duplicate file name and confirm replacement
    const dup = competitorFilesData.findIndex(f => f.fileName === fd.fileName);
    if (dup !== -1) {
      if (!confirm('Файл «' + fd.fileName + '» уже загружен в мониторинг.\nЗаменить его новой версией?')) return false;
      competitorFilesData.splice(dup, 1);
    }
    competitorFilesData.push(fd);
    return true;
  },
  parseFile,
  processAllData,
  removeFileExtension,
  renderTable,
  get myPriceInput() { return myPriceInput; },
  get competitorInput() { return competitorInput; },
  updateMyPriceStatus(name) {
    const el = document.getElementById('myPriceStatus');
    if (el) { el.className = 'upload-status upload-status--ok'; el.textContent = '✅ ' + name; }
  },
  updateCompetitorStatus() {
    const el = document.getElementById('competitorStatus');
    if (el) {
      const n = competitorFilesData.length;
      el.className = 'upload-status upload-status--ok';
      el.textContent = '✅ ' + n + ' файл' + (n===1?'':'а'+(n<5?'':'ов'));
    }
    if(typeof _sfUpdateSuppliers==='function')_sfUpdateSuppliers(competitorFilesData.map(f=>({name:f.fileName,rows:f.data?f.data.length:null})));
  }
};

// Re-render table when monitor pane becomes visible (virtual scroll fix)
window._pmAppOnMonitorShow = function() {
  if (groupedData.length > 0) {
    // Defer so pane is fully painted and clientHeight is correct
    setTimeout(() => { renderTable(); }, 30);
  }
};

// ════════════════════════════════════════════════════════════════════════
// РЕЖИМ ОБУЧЕНИЯ
// ════════════════════════════════════════════════════════════════════════


// ── Тост об успешном завершении мониторинга ───────────────────────────────
function showCompletionToast() {
    // Показываем только если загружены прайсы поставщиков
    if (competitorFilesData.length === 0) return;
    const total = groupedData.length;
    const matched = groupedData.filter(i => i.isInMyPrice).length;
    const suppliers = competitorFilesData.length;
    if (total === 0) return;

    let toast = document.getElementById('_completionToast');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = '_completionToast';
        Object.assign(toast.style, {
            position: 'fixed', bottom: '24px', right: '24px',
            background: '#0f172a', color: '#f1f5f9',
            border: '1px solid #22c55e',
            padding: '14px 18px', borderRadius: '10px',
            fontSize: '13px', lineHeight: '1.6',
            boxShadow: '0 6px 28px rgba(0,0,0,.45)',
            zIndex: '99998', maxWidth: '340px',
            transition: 'opacity .4s, transform .4s',
            transform: 'translateY(20px)', opacity: '0',
            cursor: 'pointer'
        });
        toast.title = 'Перейти к мониторингу';
        toast.addEventListener('click', () => switchMainPane('monitor'));
        document.body.appendChild(toast);
    }
    toast.innerHTML =
        `<div style="font-size:15px;font-weight:700;color:#22c55e;margin-bottom:6px;">✅ Мониторинг готов!</div>` +
        `<div>📦 Товаров: <b style="color:#7dd3fc">${total.toLocaleString('ru')}</b></div>` +
        (matched ? `<div>🏷️ Совпало с прайсом: <b style="color:#7dd3fc">${matched.toLocaleString('ru')}</b></div>` : '') +
        `<div>📂 Поставщиков: <b style="color:#7dd3fc">${suppliers}</b></div>` +
        `<div style="margin-top:8px;font-size:11px;color:#94a3b8;">Нажмите, чтобы открыть таблицу →</div>`;

    // Animate in
    requestAnimationFrame(() => {
        toast.style.opacity = '1';
        toast.style.transform = 'translateY(0)';
    });
    clearTimeout(toast._timer);
    toast._timer = setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateY(20px)';
    }, 12000);
}


// ════════════════════════════════════════════════════════════════════════════
// NAVIGATION + APP BRIDGE + OBR (PREPARE PANE)
// ════════════════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════════════════
// COMBINED TAB NAVIGATION
// ════════════════════════════════════════════════════════════════════════════
function switchMainPane(name) {
  const prev = document.querySelector('.main-pane.active');
  const prevId = prev ? prev.id : '';

  // ── Clearing prepare pane when navigating away ──
  if (prevId === 'pane-prepare' && name !== 'prepare') {
    obrClearTable();
  }

  document.querySelectorAll('.nav-tab').forEach(t =>
    t.classList.toggle('active', t.dataset.pane === name));
  document.querySelectorAll('.main-pane').forEach(p =>
    p.classList.toggle('active', p.id === 'pane-' + name));

  // Re-render monitor table after pane becomes visible (virtual scroll fix)
  if (name === 'monitor' && typeof window._pmAppOnMonitorShow === 'function') {
    window._pmAppOnMonitorShow();
  }
  // Re-render matcher virtual scroll after pane becomes visible
  if (name === 'matcher') {
    setTimeout(function() {
      const mWrap = document.getElementById('matcherTableWrap');
      if (mWrap && typeof mWrap._mvsRender === 'function') mWrap._mvsRender();
    }, 30);
  }
}
// Init: activate prepare tab
document.querySelectorAll('.nav-tab[data-pane]').forEach(t =>
  t.addEventListener('click', () => {
    const _sb = document.querySelector('.app-sidebar');
    if (_sb && _sb.classList.contains('collapsed')) {
      _sb.classList.remove('collapsed');
      localStorage.setItem('sidebarCollapsed', '0');
      const _tbtn = document.getElementById('sidebarToggle');
      if (_tbtn) _tbtn.title = 'Свернуть меню';
    }
    switchMainPane(t.dataset.pane);
  }));

// ════════════════════════════════════════════════════════════════════════════
// APP BRIDGE — event bus between OBR and PriceMatcher
// ════════════════════════════════════════════════════════════════════════════
const AppBridge = {
  _handlers: {},
  on(event, fn) { (this._handlers[event] = this._handlers[event] || []).push(fn); },
  emit(event, data) { (this._handlers[event] || []).forEach(fn => fn(data)); }
};

// ════════════════════════════════════════════════════════════════════════════
// OBR TYPE MANAGEMENT
// ════════════════════════════════════════════════════════════════════════════
let obrCurrentType = 'supplier';
function obrSetType(type) {
  obrCurrentType = type;
  const _tSup = document.getElementById('obrTypeSupplier');
  const _tMyP = document.getElementById('obrTypeMyPrice');
  if (_tSup) _tSup.classList.toggle('active', type === 'supplier');
  if (_tMyP) _tMyP.classList.toggle('active', type === 'myprice');
  const badge = document.getElementById('obrTypeBadge');
  if (badge) {
    badge.textContent = type === 'myprice' ? 'Мой прайс' : 'Поставщик';
    badge.style.background = type === 'myprice' ? '#2d9a5f' : '#4a90d9';
    badge.style.display = '';
  }
}

// userConfig stub for OBR (uses localStorage in merged app)
(function() {
  const configEl = document.createElement('script');
  configEl.id = 'obrUserConfig';
  configEl.textContent = '{}';
  document.head.appendChild(configEl);
})();

// ===== CONFIG =====
function getUserConfig() {
  const el = document.getElementById("obrUserConfig");
  if (!el) return {};
  try { return JSON.parse(el.textContent || "{}"); } catch { return {}; }
}
function setUserConfig(cfg) {
  const el = document.getElementById("obrUserConfig");
  if (el) el.textContent = JSON.stringify(cfg, null, 2);
}

const DEFAULT_COLUMN_TEMPLATES = [
  "Штрихкод","EAN","Артикул","Наименование","Бренд","Категория",
  "Единица","Количество","Цена","Цена опт","Цена РРЦ","Сумма",
  "Остаток","В пути","Склад"
];
const DEFAULT_COLUMN_SYNONYMS = {
  "Штрихкод": ["штрихкод штука","штрих-код","штрихкод","barcode","gtin","шк","код товара"],
  "EAN":       ["ean13","ean","barcode","штрихкод"],
  "Артикул":   ["артикул поставщика","артикул","арт","art","sku","vendor code","code"],
  "Наименование": ["наименование товара","наименование товаров","номенклатура",
                   "наименование","название товара","название","name","товар","продукт"],
  "Бренд":     ["торговая марка","бренд","тм","марка","производитель","brand","trademark"],
  "Категория": ["подгруппа","категория","группа товаров","раздел","группа","тип","вид"],
  "Единица":   ["единица измерения","ед.изм","единица","упак","фасовка","тип упаковки","unit","ед"],
  "Количество":["количество в упаковке","количество","кол-во","кол","qty","count","шт"],
  "Цена":      ["входящая цена","закупочная цена","цена закупки","цена входящая",
                "входящая","закупочная","цена","price","стоимость","прайс"],
  "Цена опт":  ["оптовая цена","цена оптовая","оптовая","опт","wholesale","opt"],
  "Цена РРЦ":  ["рекомендованная розничная","рекомендуемая цена","розничная цена",
                "цена розничная","ррц","розничная","розница","retail","рц"],
  "Сумма":     ["итоговая сумма","сумма итого","итого","сумма","total","amount"],
  "Остаток":   ["свободный остаток","остаток на складе","остатки","наличие","остаток",
                "available","stock","доступно"],
  "В пути":    ["количество в пути","в пути","транзит","впути","transit","in transit"],
  "Склад":     ["место хранения","warehouse","склад","storage","хранение"]
};

function loadColumnTemplates() {
  const cfg = getUserConfig();
  if (Array.isArray(cfg.columnTemplates) && cfg.columnTemplates.length) return cfg.columnTemplates;
  try { const a = JSON.parse(localStorage.getItem("columnTemplates")||"[]"); if (a.length) return a; } catch {}
  return DEFAULT_COLUMN_TEMPLATES.slice();
}
function loadColumnSynonyms() {
  const cfg = getUserConfig();
  if (cfg.columnSynonyms && typeof cfg.columnSynonyms === "object") return cfg.columnSynonyms;
  try { const o = JSON.parse(localStorage.getItem("columnSynonyms")||"{}"); if (Object.keys(o).length) return o; } catch {}
  return JSON.parse(JSON.stringify(DEFAULT_COLUMN_SYNONYMS));
}
function persistAll(markDirty = true) {
  const cfg = getUserConfig();
  cfg.columnTemplates = columnTemplates.slice();
  cfg.columnSynonyms = columnSynonyms;
  setUserConfig(cfg);
  localStorage.setItem("columnTemplates", JSON.stringify(columnTemplates));
  localStorage.setItem("columnSynonyms", JSON.stringify(columnSynonyms));
  // Отмечаем несохранённые изменения только при пользовательских действиях,
  // не при первичной инициализации
  if (markDirty && typeof unifiedMarkUnsaved === 'function') unifiedMarkUnsaved(true);
}

let columnTemplates = loadColumnTemplates();
let columnSynonyms  = loadColumnSynonyms();
// true — настройки пришли из загруженного пользователем JSON-файла
// false — используются демо-данные (DEFAULT_*)
let _columnSettingsFromFile = (() => {
  const cfg = getUserConfig();
  return !!(Array.isArray(cfg.columnTemplates) && cfg.columnTemplates.length
         && JSON.stringify(cfg.columnTemplates) !== JSON.stringify(DEFAULT_COLUMN_TEMPLATES));
})();
persistAll(false); // инициализация — не ставим «несохранено»

// ===== STATE =====
let tableData = null;
let selectedColumns = new Map();
let startRowIndex = 0;
let currentWorkbook = null;
let displayedRows = 50;
let activeDropdown = null;
let originalFileName = "export";
let fileQueue = [];
let _queueTotal = 0;     // total files in batch (for progress bar)
let _queueDone  = 0;     // files already saved in current batch
let pendingCsvContent = null;
let pendingCsvFileName = null;
let pendingSkippedRows = [];

// ===== COMPLEX MODE STATE =====
let complexModeEnabled = false;
let complexDetected = false;
let currentWs = null; // current worksheet (for merge detection)
// subheaderGroups: [{key, rows:[rowIndex], rawText, tokens, selectedTokens:[]}]
let subheaderGroups = [];

// ===== DOM =====
const fileInput      = document.getElementById("obrFileInput");
const fileInputMyPrice = document.getElementById("obrFileInputMyPrice");
const obrTableContainer = document.getElementById("obrTableContainer");
const dataTable      = document.getElementById("obrDataTable");
const downloadBtn    = document.getElementById("obrDownloadBtn");
const resetBtn       = document.getElementById("obrResetBtn");
const manageTemplatesBtn = document.getElementById("obrManageTemplatesBtn");
const fileNameDisplay = document.getElementById("obrFileNameDisplay");
const sheetSelector  = document.getElementById("obrSheetSelector");
const sheetSelect    = document.getElementById("obrSheetSelect");
const loadMoreBtn    = document.getElementById("obrLoadMoreBtn");
const loadMoreContainer = document.getElementById("obrLoadMoreContainer");
const templatesModal = document.getElementById("obrTemplatesModal");
const closeTemplatesModal = document.getElementById("obrCloseTemplatesModal");
const newTemplateInput = document.getElementById("obrNewTemplateInput");
const addTemplateBtn = document.getElementById("obrAddTemplateBtn");
const templatesList  = document.getElementById("obrTemplatesList");
const skippedModal   = document.getElementById("obrSkippedModal");
const closeSkippedModal = document.getElementById("obrCloseSkippedModal");
const skippedSummary = document.getElementById("obrSkippedSummary");
const skippedTable   = document.getElementById("obrSkippedTable");
const confirmDownloadCsvBtn = document.getElementById("obrConfirmDownloadCsvBtn");
const downloadSkippedBtn = document.getElementById("obrDownloadSkippedBtn");
const complexBanner    = document.getElementById("obrComplexBanner");
const complexEnableBtn = document.getElementById("obrComplexEnableBtn");
const complexDismiss   = document.getElementById("obrComplexDismiss");
const complexBtnGroup  = document.getElementById("obrComplexBtnGroup");
const complexSep       = document.getElementById("obrComplexSep");
const complexConfigBtn = document.getElementById("obrComplexConfigBtn");
const complexModal     = document.getElementById("obrComplexModal");
const closeComplexModal= document.getElementById("obrCloseComplexModal");
const complexSubheaderList = document.getElementById("obrComplexSubheaderList");
const complexNoData    = document.getElementById("obrComplexNoData");
const complexSummaryDiv= document.getElementById("obrComplexSummaryDiv");
const complexApplyBtn  = document.getElementById("obrComplexApplyBtn");
const complexResetBtn  = document.getElementById("obrComplexResetBtn");

// ===== STUDY MODE =====

// ===== CLOSE DROPDOWN =====
document.addEventListener("click", function(e) {
  if (!e.target.closest("#pane-prepare .rename-wrapper") && !e.target.closest(".modal-box") && activeDropdown) {
    activeDropdown.classList.remove("show");
    activeDropdown = null;
  }
});

// ===== QUEUE =====
function renderQueuePanel() {
  const panel       = document.getElementById("obrQueuePanel");
  const queueList   = document.getElementById("obrQueueList");
  const queueCurrent= document.getElementById("obrQueueCurrent");
  const fillEl      = document.getElementById("obrQueueProgressFill");
  const labelEl     = document.getElementById("obrQueueProgressLabel");
  const statusEl    = document.getElementById("obrQueueStatus");

  // Hide legacy status
  if (statusEl) statusEl.style.display = "none";

  if (_queueTotal <= 1) {
    panel.style.display = "none";
    return;
  }

  panel.style.display = "block";

  // Current file name
  if (queueCurrent) queueCurrent.textContent = originalFileName;

  // Progress
  const done    = _queueDone;
  const total   = _queueTotal;
  const pct     = Math.round((done / total) * 100);
  if (fillEl)  fillEl.style.width = Math.max(4, pct) + "%";
  if (labelEl) labelEl.textContent = `${done + 1} / ${total}`;

  // Chips — all files, marking done and active
  if (queueList) {
    // rebuild each call (queue shrinks as files are processed)
    queueList.innerHTML = "";
    // We don't have the full original list after shift(), so just show remaining + current
    const activeChip = document.createElement("span");
    activeChip.className = "obr-queue-chip active";
    activeChip.textContent = originalFileName;
    activeChip.title = originalFileName;
    queueList.appendChild(activeChip);
    fileQueue.forEach((f, i) => {
      const chip = document.createElement("span");
      chip.className = "obr-queue-chip";
      chip.textContent = f.name.replace(/\.[^.]+$/, "");
      chip.title = f.name;
      queueList.appendChild(chip);
    });
  }
}

function loadFileObject(file) {
  originalFileName = (file.name || "export").replace(/\.[^.]+$/, "");
  fileNameDisplay.textContent = file.name || "";

  // Show cleared hint = false, show table area
  const clearedHint = document.getElementById("obrClearedHint");
  const mainHint    = document.getElementById("obrMainHint");
  const tableWrap   = document.getElementById("obrTableWrap");
  if (clearedHint) clearedHint.style.display = "none";

  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    try {
      currentWorkbook = XLSX.read(data, { type: "array" });
      if (currentWorkbook.SheetNames.length > 1) {
        sheetSelect.innerHTML = "";
        currentWorkbook.SheetNames.forEach((name, idx) => {
          const o = document.createElement("option");
          o.value = String(idx); o.textContent = name;
          sheetSelect.appendChild(o);
        });
        sheetSelector.style.display = "flex";
      } else {
        sheetSelector.style.display = "none";
      }
      loadSheet(0);
      renderQueuePanel();
    } catch(err) {
      alert(err && err.message ? err.message : String(err));
      loadNextFromQueue();
    }
  };
  reader.readAsArrayBuffer(file);
}

function loadNextFromQueue() {
  _queueDone++;
  if (fileQueue.length === 0) {
    // Batch complete — reset state but keep table clear
    tableData = null; selectedColumns.clear(); startRowIndex = 0; currentWorkbook = null;
    _queueTotal = 0; _queueDone = 0;
    document.getElementById("obrQueuePanel").style.display = "none";
    return;
  }
  const next = fileQueue.shift();
  selectedColumns.clear(); startRowIndex = 0; displayedRows = 50;
  loadFileObject(next);
}

// ===== MEMORY CLEAR (when navigating away from prepare pane) =====
function obrClearTable() {
  // Drop all large data refs
  tableData        = null;
  currentWorkbook  = null;
  selectedColumns.clear();
  startRowIndex    = 0;
  displayedRows    = 50;
  fileQueue        = [];
  _queueTotal      = 0;
  _queueDone       = 0;

  // Clear DOM
  if (dataTable) dataTable.innerHTML = "";
  const tableWrap = document.getElementById("obrTableWrap");
  if (tableWrap) tableWrap.style.display = "none";
  const queuePanel = document.getElementById("obrQueuePanel");
  if (queuePanel) queuePanel.style.display = "none";
  const complexBannerEl = document.getElementById("obrComplexBanner");
  if (complexBannerEl) complexBannerEl.style.display = "none";
  const sheetSel = document.getElementById("obrSheetSelector");
  if (sheetSel) sheetSel.style.display = "none";
  const loadMoreEl = document.getElementById("obrLoadMoreContainer");
  if (loadMoreEl) loadMoreEl.style.display = "none";
  const fileNameEl = document.getElementById("obrFileNameDisplay");
  if (fileNameEl) fileNameEl.textContent = "";
  if (fileInput)       fileInput.value       = "";
  if (fileInputMyPrice) fileInputMyPrice.value = "";

  // Show cleared state (navigated away)
  const clearedHint = document.getElementById("obrClearedHint");
  if (clearedHint) {
    clearedHint.style.display = "flex";
    // Update text to "data cleared" variant
    const icon = document.getElementById("obrClearedIcon");
    const title = document.getElementById("obrClearedTitle");
    const desc = document.getElementById("obrClearedDesc");
    if (icon) icon.textContent = "📋";
    if (title) title.textContent = "Данные очищены";
    if (desc) desc.textContent = "При переходе в другой раздел таблица была выгружена из памяти. Загрузите файл снова для продолжения.";
  }
  const mainHint = document.getElementById("obrMainHint");
  if (mainHint)  mainHint.style.display = "none";

  if (typeof updateStats === 'function') updateStats();
}

// Кнопки в тайтлбаре — устанавливают тип при клике (file input открывается автоматически)
const _obrTypeSupplierBtn = document.getElementById("obrTypeSupplier");
if (_obrTypeSupplierBtn) _obrTypeSupplierBtn.addEventListener("click", function() {
  obrSetType('supplier');
});
const _obrTypeMyPriceBtn = document.getElementById("obrTypeMyPrice");
if (_obrTypeMyPriceBtn) _obrTypeMyPriceBtn.addEventListener("click", function() {
  obrSetType('myprice');
});
// File inputs в тайтлбаре
fileInput.addEventListener("change", function(e) { handleFileUpload(e); });
if (fileInputMyPrice) fileInputMyPrice.addEventListener("change", function(e) {
  obrSetType('myprice');
  handleFileUpload(e);
});
// JSON upload on upload screen
const obrJsonUploadInput = document.getElementById("obrJsonUploadInput");
if (obrJsonUploadInput) obrJsonUploadInput.addEventListener("change", function(e) {
  const file = e.target.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = function(ev) {
    try {
      const json = JSON.parse(ev.target.result);
      AppBridge.emit('settingsLoaded', json);
      const synInp = document.getElementById('synonymsInput');
      if (synInp) {
        const dt = new DataTransfer();
        dt.items.add(file);
        synInp.files = dt.files;
        synInp.dispatchEvent(new Event('change', { bubbles: true }));
      }
      showToast('JSON загружен — настройки и синонимы применены', 'ok');
      setTimeout(function(){if(typeof obrShowNextStep==='function')obrShowNextStep('json');},400);
    } catch(err) { alert('Ошибка чтения JSON: ' + err.message); }
  };
  reader.readAsText(file, 'utf-8');
  e.target.value = "";
});


// Cleared-hint file inputs
const obrClearedFileInput = document.getElementById("obrClearedFileInput");
const obrClearedFileInputMyPrice = document.getElementById("obrClearedFileInputMyPrice");
if (obrClearedFileInput) obrClearedFileInput.addEventListener("change", function(e) {
  const clearedHint = document.getElementById("obrClearedHint");
  if (clearedHint) clearedHint.style.display = "none";
  const mainHint = document.getElementById("obrMainHint");
  if (mainHint) mainHint.style.display = "";
  const tableWrap = document.getElementById("obrTableWrap");
  if (tableWrap) tableWrap.style.display = "";
  obrSetType('supplier');
  handleFileUpload(e);
});
if (obrClearedFileInputMyPrice) obrClearedFileInputMyPrice.addEventListener("change", function(e) {
  const clearedHint = document.getElementById("obrClearedHint");
  if (clearedHint) clearedHint.style.display = "none";
  const mainHint = document.getElementById("obrMainHint");
  if (mainHint) mainHint.style.display = "";
  const tableWrap = document.getElementById("obrTableWrap");
  if (tableWrap) tableWrap.style.display = "";
  obrSetType('myprice');
  handleFileUpload(e);
});

// JSON upload on cleared screen
const obrClearedJsonInput = document.getElementById("obrClearedJsonInput");
if (obrClearedJsonInput) obrClearedJsonInput.addEventListener("change", function(e) {
  const file = e.target.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = function(ev) {
    try {
      const json = JSON.parse(ev.target.result);
      AppBridge.emit('settingsLoaded', json);
      const synInp = document.getElementById('synonymsInput');
      if (synInp) {
        const dt = new DataTransfer();
        dt.items.add(file);
        synInp.files = dt.files;
        synInp.dispatchEvent(new Event('change', { bubbles: true }));
      }
      showToast('JSON загружен — настройки и синонимы применены', 'ok');
    } catch(err) { alert('Ошибка чтения JSON: ' + err.message); }
  };
  reader.readAsText(file, 'utf-8');
  e.target.value = "";
});

// Skip button
const obrQueueSkipBtn = document.getElementById("obrQueueSkipBtn");
if (obrQueueSkipBtn) {
  obrQueueSkipBtn.addEventListener("click", function() {
    if (fileQueue.length === 0) return;
    showToast(`⏭ Файл «${originalFileName}» пропущен`, 'warn');
    loadNextFromQueue();
  });
}

function handleFileUpload(e) {
  const files = e.target.files;
  if (!files || !files.length) return;
  const arr = Array.from(files);
  // Reset progress counters for new batch
  _queueDone  = 0;
  _queueTotal = arr.length;
  if (arr.length === 1) {
    fileQueue = [];
    loadFileObject(arr[0]);
  } else {
    fileQueue = arr.slice(1);
    loadFileObject(arr[0]);
  }
  e.target.value = "";
}

sheetSelect.addEventListener("change", function() {
  loadSheet(parseInt(sheetSelect.value, 10) || 0);
});

// ===== AUTO-DETECT =====
function obrAutoDetectColumns() {
  if (!tableData || !tableData.length) return;
  const SCAN = 15;
  const maxCols = Math.max(0, ...tableData.map(r => r ? r.length : 0));
  for (let col = 0; col < maxCols; col++) {
    if (selectedColumns.has(col)) continue;
    for (let row = 0; row < Math.min(SCAN, tableData.length); row++) {
      const cell = (tableData[row] || [])[col];
      if (cell == null) continue;
      const norm = String(cell).toLowerCase().replace(/\s+/g, " ").trim();
      if (!norm) continue;
      let matched = false;
      for (const tpl of columnTemplates) {
        for (const syn of (columnSynonyms[tpl] || []).filter(Boolean)) {
          if (norm === syn.toLowerCase().replace(/\s+/g, " ").trim()) {
            selectedColumns.set(col, tpl); matched = true; break;
          }
        }
        if (matched) break;
      }
      if (matched) break;
    }
  }
}

function loadSheet(idx) {
  if (!currentWorkbook) return;
  const ws = currentWorkbook.Sheets[currentWorkbook.SheetNames[idx]];
  currentWs = ws;
  tableData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
  startRowIndex = 0; selectedColumns.clear(); displayedRows = Math.min(50, tableData.length);
  obrAutoDetectColumns(); obrRenderTable(); updateLoadMore();
  // upload screen removed

  // Complex mode detection
  complexDetected = detectComplexPricelist(ws);
  subheaderGroups = []; // reset on new sheet
  if (!complexDetected) {
    showComplexBanner(false);
    setComplexMode(false);
  } else {
    showComplexBanner(true);
  }
}

// ===== RENDER =====
function obrEsc(t) {
  return String(t).replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;").replaceAll('"',"&quot;").replaceAll("'","&#039;");
}

function applyColWidths() {
  const MAX_W = 150;
  const PAD   = 18; // horizontal padding per cell (2 * 8px + border)
  if (!tableData || !tableData.length) return;
  const maxCols = Math.max(0, ...tableData.map(r => r ? r.length : 0));
  const rowsToSample = Math.min(tableData.length, 60);

  // Use canvas to measure text widths (fast, no layout thrash)
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  ctx.font = '13px "Segoe UI", Arial, sans-serif';
  const ctxH = document.createElement('canvas').getContext('2d');
  ctxH.font = 'bold 12px "Segoe UI", Arial, sans-serif';

  const widths = new Array(maxCols).fill(0);
  // measure header: column numbers or rename inputs (just use col index label)
  for (let i = 0; i < maxCols; i++) {
    widths[i] = Math.max(widths[i], ctxH.measureText(String(i + 1)).width + PAD + 24);
  }
  // measure cell content
  for (let ri = 0; ri < rowsToSample; ri++) {
    const row = tableData[ri] || [];
    for (let ci = 0; ci < maxCols; ci++) {
      const v = row[ci] != null ? String(row[ci]) : '';
      const w = ctx.measureText(v).width + PAD;
      if (w > widths[ci]) widths[ci] = w;
    }
  }

  // build colgroup with min-widths (table-layout: auto will grow beyond these if needed)
  const table = dataTable.closest('table');
  if (!table) return;
  let old = table.querySelector('colgroup');
  if (old) old.remove();
  const cg = document.createElement('colgroup');
  // first col: row-number
  const cNum = document.createElement('col');
  cNum.style.minWidth = '36px';
  cNum.style.width = '36px';
  cg.appendChild(cNum);
  for (let i = 0; i < maxCols; i++) {
    const c = document.createElement('col');
    const w = Math.min(Math.ceil(widths[i]), MAX_W);
    c.style.minWidth = w + 'px';
    cg.appendChild(c);
  }
  table.insertBefore(cg, table.firstChild);
}

function obrRenderTable() {
  // Ensure the table wrapper is always visible
  const _tw = document.getElementById('obrTableWrap');
  if (_tw) _tw.style.display = '';
  if (!tableData || !tableData.length) { dataTable.innerHTML = "<tr><td>Нет данных</td></tr>"; updateStats(); return; }
  const maxCols = Math.max(0, ...tableData.map(r => r ? r.length : 0));
  const rowsToShow = Math.min(displayedRows, tableData.length);

  let html = "<thead><tr>";
  html += `<th class="xl-row-num" title="">#</th>`;
  for (let i = 0; i < maxCols; i++) {
    const sel = selectedColumns.has(i);
    const colName = sel ? selectedColumns.get(i) : '';
    html += `<th class="${sel ? "col-selected" : ""}" data-col="${i}" style="white-space:nowrap;min-width:${sel?'130px':'40px'};">`;
    if (sel) {
      html += `<div style="display:flex;flex-direction:column;gap:3px;padding:3px 0 0;">`;
      html += `<div style="font-size:9px;color:#6B7280;font-weight:600;text-transform:uppercase;letter-spacing:0.4px;margin-bottom:1px;">Колонка ${i+1}</div>`;
      html += createRenameInput(i, colName);
      html += `</div>`;
    } else {
      html += `<div style="padding:4px 2px;display:flex;flex-direction:column;align-items:center;gap:2px;">`;
      html += `<span style="font-size:11px;font-weight:700;color:#374151;">${i+1}</span>`;
      html += `</div>`;
    }
    html += "</th>";
  }
  html += "</tr></thead><tbody>";

  for (let ri = 0; ri < rowsToShow; ri++) {
    const row = tableData[ri] || [];
    const hidden = ri < startRowIndex;
    html += `<tr class="${hidden ? "row-hidden" : ""}" data-row-index="${ri}">`;
    html += `<td class="xl-row-num">${ri + 1}</td>`;
    for (let i = 0; i < maxCols; i++) {
      const v = row[i] != null ? row[i] : "";
      const selClass = selectedColumns.has(i) ? ' style="background:#ebf7ed;"' : '';
      html += `<td data-row="${ri}" data-col="${i}"${selClass}>${obrEsc(v)}</td>`;
    }
    html += "</tr>";
  }
  html += "</tbody>";
  dataTable.innerHTML = html;
  applyColWidths();
  attachEvents();
  updateStats();
}

function createRenameInput(colIndex, value) {
  const ev = String(value || "").replaceAll('"', "&quot;");
  const items = columnTemplates.filter(Boolean).map(t =>
    `<div class="dropdown-item" data-value="${String(t).replaceAll('"','&quot;')}">${obrEsc(t)}</div>`
  ).join("");
  return `<div class="rename-wrapper" data-col="${colIndex}">
    <input class="rename-input" type="text" value="${ev}" data-col="${colIndex}" placeholder="Название колонки">
    <div class="dropdown" data-col="${colIndex}">${items}</div>
  </div>`;
}

function openDropdown(colIndex, doFocus) {
  const input = document.querySelector(`#pane-prepare .rename-input[data-col="${colIndex}"]`);
  const dd = document.querySelector(`#pane-prepare .dropdown[data-col="${colIndex}"]`);
  if (!input || !dd) return;
  if (activeDropdown && activeDropdown !== dd) activeDropdown.classList.remove("show");
  dd.classList.add("show"); activeDropdown = dd;
  if (doFocus) { input.focus(); input.select(); }
}

function attachEvents() {
  document.querySelectorAll("#pane-prepare th[data-col]").forEach(th => {
    th.addEventListener("click", function(e) {
      if (e.target.closest(".rename-wrapper")) return;
      const ci = parseInt(th.dataset.col, 10);
      if (selectedColumns.has(ci)) { selectedColumns.delete(ci); obrRenderTable(); return; }
      selectedColumns.set(ci, ""); obrRenderTable();
      requestAnimationFrame(() => openDropdown(ci, true));
    });
  });
  document.querySelectorAll("#pane-prepare .rename-input").forEach(inp => {
    inp.addEventListener("click", e => { e.stopPropagation(); openDropdown(parseInt(inp.dataset.col,10), false); });
    inp.addEventListener("input", e => { selectedColumns.set(parseInt(inp.dataset.col,10), inp.value); updateStats(); });
    inp.addEventListener("focus", e => e.target.select());
  });
  document.querySelectorAll("#pane-prepare .dropdown-item").forEach(item => {
    item.addEventListener("click", e => {
      e.stopPropagation();
      const dd = e.target.closest(".dropdown");
      const ci = parseInt(dd.dataset.col, 10);
      const inp = document.querySelector(`#pane-prepare .rename-input[data-col="${ci}"]`);
      if (!inp) return;
      inp.value = e.target.dataset.value || "";
      selectedColumns.set(ci, inp.value);
      dd.classList.remove("show"); activeDropdown = null; updateStats();
    });
  });
  // row click removed
}

function updateLoadMore() {
  const rem = (tableData ? tableData.length : 0) - displayedRows;
  loadMoreContainer.style.display = rem > 0 ? "block" : "none";
  if (rem > 0) document.getElementById("obrRemainingRows").textContent = String(rem);
}

loadMoreBtn.addEventListener("click", function() {
  displayedRows = tableData ? tableData.length : displayedRows;
  obrRenderTable(); updateLoadMore();
});

function updateStats() {
  const maxCols = tableData && tableData.length ? Math.max(0, ...tableData.map(r => r ? r.length : 0)) : 0;
  document.getElementById("obrTotalColumns").textContent  = String(maxCols);
  document.getElementById("obrSelectedColumns").textContent = String(selectedColumns.size);
  document.getElementById("obrTotalRows").textContent = String(Math.max(0, (tableData ? tableData.length : 0) - startRowIndex));
  downloadBtn.disabled = selectedColumns.size === 0;
  if (typeof _obrUpdateSkippedBtn === 'function') _obrUpdateSkippedBtn();
}

// ===== CSV =====
function esc_csv(v) {
  if (v == null) return "";
  const s = String(v);
  if (s.includes('"') || s.includes(",") || s.includes("\n") || s.includes("\r"))
    return `"${s.replaceAll('"', '""')}"`;
  return s;
}
function normHeader(n) {
  return String(n||"").toLowerCase().trim().replaceAll("ё","е").replaceAll(/\s+/g," ").replaceAll(/[^\p{L}\p{N} ]/gu,"");
}
function normalizeBarcode(raw) {
  let s = String(raw ?? "").trim().replace(/\s+/g, "");
  while (s.endsWith(".")) s = s.slice(0,-1);
  if (/^\d+$/.test(s)) return s;
  if (/^\d+\.0$/.test(s)) return s.split(".0")[0];
  const m = s.replace(",",".").match(/^(\d+)(?:\.(\d+))?e\+?(\d+)$/i);
  if (!m) return "";
  const digits = m[1]+(m[2]||""), exp = parseInt(m[3],10), shift = exp-(m[2]||"").length;
  if (shift >= 0) return digits + "0".repeat(shift);
  const cut = digits.length + shift;
  if (cut <= 0 || !/^0+$/.test(digits.slice(cut))) return "";
  return digits.slice(0, cut);
}
function findBarcodeCol(indices) {
  for (const ci of indices) {
    const n = normHeader(selectedColumns.get(ci)||"");
    if (/штрихкод/.test(n) || /barcode/.test(n) || /\bean\b/.test(n)) return ci;
  }
  return -1;
}

function buildCsvAndSkipped() {
  if (!selectedColumns.size) return { ok: false, error: "Выберите колонки." };
  const indices = Array.from(selectedColumns.keys()).sort((a,b)=>a-b);
  const bcCol = findBarcodeCol(indices);
  if (bcCol === -1) return { ok: false, error: "Не найдена колонка штрихкода (название должно содержать «штрихкод» / barcode / ean)." };

  // Find name column for complex mode prefix injection
  let nameCol = -1;
  if (complexModeEnabled) {
    for (const ci of indices) {
      const n = normHeader(selectedColumns.get(ci) || '');
      if (/наименован/.test(n) || /номенклатур/.test(n)) { nameCol = ci; break; }
    }
  }
  const prefixMap = complexModeEnabled ? buildPrefixMap() : null;

  let csv = "\uFEFF" + indices.map(i => esc_csv(selectedColumns.get(i)||"")).join(",") + "\n";
  const skipped = [];

  for (let ri = startRowIndex; ri < tableData.length; ri++) {
    const row = tableData[ri] || [];

    // Skip subheader rows in complex mode
    if (prefixMap !== null) {
      const { bcCol: bc2 } = findSpecialCols();
      if (isSubheaderRow(row, bc2)) continue;
    }

    const rawBC = row[bcCol];
    const rawBCS = rawBC == null ? "" : String(rawBC);
    const normBC = normalizeBarcode(rawBC);
    if (!normBC || !/^\d+$/.test(normBC)) {
      skipped.push({ rowIndex: ri, rowNumber: ri+1, rawBarcode: rawBCS, normalizedBarcode: normBC||"", reason: !rawBCS.trim() ? "Пустой штрихкод" : "Некорректный штрихкод" });
      continue;
    }
    const vals = indices.map(ci => {
      if (ci === bcCol) return normBC;
      let v = row[ci] != null ? String(row[ci]).trim() : "";
      if (/^\d+,\d{2}$/.test(v)) v = v.replace(",", ".");
      // Apply prefix to name column in complex mode
      if (prefixMap && ci === nameCol && v) {
        const prefix = prefixMap.get(ri) || '';
        if (prefix && !prefixContainedInName(v, prefix)) {
          v = prefix + ' ' + v;
        }
      }
      return v;
    });
    if (vals.every(v => !v)) continue;
    csv += vals.map(esc_csv).join(",") + "\n";
  }
  return { ok: true, csvContent: csv, skipped };
}

// ===== SKIPPED MODAL =====
function openSkippedModal(skipped, fn) {
  pendingSkippedRows = skipped.slice();
  pendingCsvFileName = fn;
  const preview = skipped.filter(s => s.reason !== "Пустой штрихкод");
  const hidden = skipped.length - preview.length;
  const toShow = preview.slice(0, 500);
  skippedSummary.textContent = `Всего пропусков: ${skipped.length}. Пустых скрыто: ${hidden}. Показано: ${toShow.length}${preview.length>500?" (из "+preview.length+")":""}.`;
  let h = "<thead><tr><th style='min-width:80px'>Строка</th><th style='min-width:220px'>Штрихкод в файле</th><th style='min-width:220px'>Нормализованный</th><th style='min-width:220px'>Причина</th></tr></thead><tbody>";
  toShow.forEach(s => { h += `<tr><td>${obrEsc(s.rowNumber)}</td><td>${obrEsc(s.rawBarcode)}</td><td>${obrEsc(s.normalizedBarcode)}</td><td>${obrEsc(s.reason)}</td></tr>`; });
  h += "</tbody>";
  skippedTable.innerHTML = h;
  skippedModal.style.display = "flex";
}
function hideSkippedModal() { skippedModal.style.display = "none"; }
closeSkippedModal.addEventListener("click", hideSkippedModal);
skippedModal.addEventListener("click", e => { if (e.target === skippedModal) hideSkippedModal(); });

confirmDownloadCsvBtn.addEventListener("click", async function() {
  if (!pendingCsvContent) return;
  const fn = pendingCsvFileName || originalFileName+".csv";
  const blob = new Blob([pendingCsvContent], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = fn;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  hideSkippedModal();
  const savedType = obrCurrentType;
  AppBridge.emit('csvReady', { csvText: pendingCsvContent, fileName: fn, isMyPrice: savedType === 'myprice' });
  if (fileQueue.length) {
    showToast(`✅ «${originalFileName}» сохранён → открывается следующий (${_queueDone + 2}/${_queueTotal})…`, 'ok');
    setTimeout(loadNextFromQueue, 400);
  } else {
    setTimeout(() => obrShowNextStep(savedType), 400);
  }
});
downloadSkippedBtn.addEventListener("click", function() {
  if (!pendingSkippedRows || !pendingSkippedRows.length) { alert("Нет пропусков."); return; }
  // Show full list (including empty barcodes) in expanded table
  const allRows = pendingSkippedRows;
  let h = "<thead><tr><th style='min-width:80px'>Строка</th><th style='min-width:220px'>Штрихкод в файле</th><th style='min-width:220px'>Нормализованный</th><th style='min-width:220px'>Причина</th></tr></thead><tbody>";
  allRows.forEach(s => { h += `<tr><td>${obrEsc(s.rowNumber)}</td><td>${obrEsc(s.rawBarcode)}</td><td>${obrEsc(s.normalizedBarcode)}</td><td>${obrEsc(s.reason)}</td></tr>`; });
  h += "</tbody>";
  skippedTable.innerHTML = h;
  skippedSummary.textContent = `Показаны все пропуски: ${allRows.length}`;
});

// ===== DOWNLOAD CSV =====
// Хранит строки пропусков после последнего экспорта
let _lastSkippedRows = [];

downloadBtn.addEventListener("click", async function() {
  const res = buildCsvAndSkipped();
  if (!res.ok) { showToast(res.error, 'err'); return; }

  pendingCsvContent = res.csvContent;
  pendingCsvFileName = originalFileName + ".csv";
  _lastSkippedRows = res.skipped || [];
  _obrUpdateSkippedBtn();

  const fn = originalFileName + ".csv";
  const savedType = obrCurrentType;

  // 1. Скачиваем файл
  const blob = new Blob([res.csvContent], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = fn;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);

  // 2. Передаём данные в мониторинг
  AppBridge.emit('csvReady', { csvText: res.csvContent, fileName: fn, isMyPrice: savedType === 'myprice' });

  const skippedMeaningful = _lastSkippedRows.filter(s => s.reason !== "Пустой штрихкод");

  // 3. Следующий файл из очереди или финальный шаг
  if (fileQueue.length) {
    const remaining = fileQueue.length;
    if (skippedMeaningful.length) {
      showToast(`✅ «${originalFileName}» сохранён. Пропусков: ${skippedMeaningful.length}. Открывается следующий файл (${_queueDone + 2}/${_queueTotal})…`, 'warn');
    } else {
      showToast(`✅ «${originalFileName}» сохранён → открывается следующий (${_queueDone + 2}/${_queueTotal})…`, 'ok');
    }
    setTimeout(loadNextFromQueue, 400);
  } else {
    if (skippedMeaningful.length) {
      showToast('✅ CSV сохранён и передан в мониторинг. Пропущено строк: ' + skippedMeaningful.length + ' — нажмите «Проверить пропущенные строки».', 'warn');
    } else {
      const batchMsg = _queueTotal > 1 ? ` Все ${_queueTotal} файла обработаны!` : '';
      showToast('✅ CSV сохранён и передан в мониторинг!' + batchMsg, 'ok');
    }
    setTimeout(() => obrShowNextStep(savedType), 400);
  }
});

// Кнопка «Пропущенные строки» — показывает предпросмотр пропусков (до или после сохранения)
const obrShowSkippedBtn = document.getElementById('obrShowSkippedBtn');
function _obrUpdateSkippedBtn() {
  if (!obrShowSkippedBtn) return;
  // Кнопка активна если выбраны колонки (можно сделать предпросмотр)
  obrShowSkippedBtn.disabled = !selectedColumns.size;
}
if (obrShowSkippedBtn) {
  obrShowSkippedBtn.addEventListener('click', function() {
    if (!selectedColumns.size) return;
    // Строим предпросмотр пропусков на лету из текущих данных
    const res = buildCsvAndSkipped();
    if (!res.ok) { showToast(res.error, 'err'); return; }
    const rows = res.skipped || [];
    if (!rows.length) { showToast('Пропущенных строк нет — все штрихкоды корректны ✅', 'ok'); return; }
    openSkippedModal(rows, (pendingCsvFileName || (originalFileName || 'файл') + '.csv'));
  });
}

// ===== NEXT-STEP MODAL =====
function obrShowNextStep(savedType) {
  const modal = document.getElementById('obrNextStepModal');
  const title = document.getElementById('obrNextStepTitle');
  const sub   = document.getElementById('obrNextStepSub');
  const btns  = document.getElementById('obrNextStepBtns');

  // Detect what's already loaded
  const jsonAlreadyLoaded = !!_columnSettingsFromFile
    || (typeof jeDB !== 'undefined' && Object.keys(jeDB).length > 0)
    || (document.getElementById('sfJsonName') && document.getElementById('sfJsonName').textContent !== 'JSON не загружен');
  const myPriceAlreadyLoaded = !!(window._pmApp && window._pmApp.myPriceData);
  const suppliersAlreadyLoaded = !!(window._pmApp && window._pmApp.competitorFilesData && window._pmApp.competitorFilesData.length > 0);

  if (savedType === 'json') {
    title.textContent = '✅ JSON загружен!';
    sub.textContent = 'Настройки столбцов и синонимы применены. Теперь откройте прайс.';
  } else {
    title.textContent = '✅ Файл сохранён!';
    sub.textContent = 'Что делаем дальше?';
  }

  const rows = [];

  // Offer JSON only if not yet loaded
  if (!jsonAlreadyLoaded && savedType !== 'json') {
    rows.push({
      cls: 'btn-json', icon: '📋',
      label: 'Загрузить JSON',
      hint: 'Применит настройки столбцов и базу синонимов',
      action: () => { obrCloseNextStep(); setTimeout(() => { const j = document.getElementById('obrJsonUploadInput'); if (j) j.click(); }, 80); }
    });
  }

  // Supplier — always offered (multiple files), label changes based on loaded state
  rows.push({
    cls: 'btn-supplier', icon: '📦',
    label: suppliersAlreadyLoaded ? 'Ещё один прайс поставщика' : 'Загрузить прайс поставщика',
    hint: suppliersAlreadyLoaded ? 'Добавить ещё один файл поставщика' : 'Откройте файл поставщика для подготовки столбцов',
    action: () => { obrCloseNextStep(); obrSetType('supplier'); fileInput.click(); }
  });

  // My price — only if not yet loaded
  if (!myPriceAlreadyLoaded) {
    rows.push({
      cls: 'btn-myprice', icon: '🏷️',
      label: 'Загрузить мой прайс',
      hint: 'Откройте свой прайс-лист для подготовки',
      action: () => { obrCloseNextStep(); obrSetType('myprice'); if (fileInputMyPrice) fileInputMyPrice.click(); }
    });
  }

  // Go to monitor — always last
  rows.push({
    cls: 'btn-monitor', icon: '📊',
    label: 'Перейти к мониторингу',
    hint: 'Открыть таблицу мониторинга цен',
    action: () => { obrCloseNextStep(); switchMainPane('monitor'); }
  });

  btns.innerHTML = '';
  rows.forEach(r => {
    const btn = document.createElement('button');
    btn.className = 'obr-nextstep-btn ' + r.cls;
    btn.innerHTML = `<span class="obr-nextstep-btn-icon">${r.icon}</span>
      <span class="obr-nextstep-btn-text">
        <span class="obr-nextstep-btn-label">${r.label}</span>
        <span class="obr-nextstep-btn-hint">${r.hint}</span>
      </span>`;
    btn.addEventListener('click', r.action);
    btns.appendChild(btn);
  });

  modal.classList.add('visible');
}

function obrCloseNextStep() {
  const m = document.getElementById('obrNextStepModal');
  if (m) m.classList.remove('visible');
}
document.addEventListener('DOMContentLoaded', function() {
  const m = document.getElementById('obrNextStepModal');
  if (m) m.addEventListener('click', function(e) { if (e.target === this) obrCloseNextStep(); });
});
// Also safe-bind immediately in case DOM is already ready
(function() {
  const m = document.getElementById('obrNextStepModal');
  if (m) m.addEventListener('click', function(e) { if (e.target === this) obrCloseNextStep(); });
})();

// ===== RESET =====
resetBtn.addEventListener("click", function() {
  selectedColumns.clear(); startRowIndex = 0; obrRenderTable();
});

// ===== TEMPLATES MODAL =====
function renderSynPanel(panel, tplName) {
  panel.innerHTML = "";
  const syns = columnSynonyms[tplName] || [];
  const lbl = document.createElement("div"); lbl.className = "syn-label";
  lbl.textContent = "Синонимы для автораспознавания:"; panel.appendChild(lbl);
  const chips = document.createElement("div"); chips.className = "syn-chips";
  syns.forEach((s, i) => {
    const chip = document.createElement("div"); chip.className = "syn-chip";
    const inp = document.createElement("input"); inp.className = "syn-input"; inp.type = "text"; inp.value = s;
    inp.addEventListener("change", () => { columnSynonyms[tplName][i] = inp.value.trim(); persistAll(); });
    const rm = document.createElement("button"); rm.className = "syn-remove"; rm.textContent = "×";
    rm.addEventListener("click", () => { columnSynonyms[tplName].splice(i,1); persistAll(); renderSynPanel(panel,tplName); });
    chip.appendChild(inp); chip.appendChild(rm); chips.appendChild(chip);
  });
  panel.appendChild(chips);
  const addRow = document.createElement("div"); addRow.className = "syn-add-row";
  const addInp = document.createElement("input"); addInp.className = "syn-new-input"; addInp.type = "text"; addInp.placeholder = "Новый синоним…";
  const addBtn = document.createElement("button"); addBtn.className = "btn btn-small btn-success"; addBtn.textContent = "+ Добавить";
  addBtn.addEventListener("click", () => {
    const v = addInp.value.trim(); if (!v) return;
    if (!columnSynonyms[tplName]) columnSynonyms[tplName] = [];
    columnSynonyms[tplName].push(v); persistAll(); addInp.value = ""; renderSynPanel(panel, tplName);
  });
  addInp.addEventListener("keydown", e => { if (e.key === "Enter") addBtn.click(); });
  addRow.appendChild(addInp); addRow.appendChild(addBtn); panel.appendChild(addRow);
}

function renderTemplatesList() {
  templatesList.innerHTML = "";
  const total = columnTemplates.length;
  columnTemplates.forEach((t, idx) => {
    const block = document.createElement("div"); block.className = "tpl-block";
    const row = document.createElement("div"); row.className = "tpl-row";

    const upBtn = document.createElement("button"); upBtn.className = "btn btn-small"; upBtn.textContent = "↑"; upBtn.disabled = idx === 0;
    upBtn.addEventListener("click", () => { if (!idx) return; [columnTemplates[idx-1],columnTemplates[idx]]=[columnTemplates[idx],columnTemplates[idx-1]]; persistAll(); renderTemplatesList(); obrRenderTable(); });

    const dnBtn = document.createElement("button"); dnBtn.className = "btn btn-small"; dnBtn.textContent = "↓"; dnBtn.disabled = idx === total-1;
    dnBtn.addEventListener("click", () => { if (idx===total-1) return; [columnTemplates[idx],columnTemplates[idx+1]]=[columnTemplates[idx+1],columnTemplates[idx]]; persistAll(); renderTemplatesList(); obrRenderTable(); });

    const inp = document.createElement("input"); inp.type = "text"; inp.value = t;
    const oldName = t;
    inp.addEventListener("change", () => {
      const n = inp.value.trim(); if (!n || n === oldName) return;
      if (columnSynonyms[oldName] !== undefined) { columnSynonyms[n] = columnSynonyms[oldName]; delete columnSynonyms[oldName]; }
      columnTemplates[idx] = n; persistAll(); renderTemplatesList(); obrRenderTable();
    });

    const synBtn = document.createElement("button"); synBtn.className = "btn btn-small"; synBtn.textContent = "🔤 Синонимы";
    synBtn.addEventListener("click", () => {
      const p = block.querySelector(".syn-panel"); if (!p) return;
      const vis = p.style.display !== "none";
      p.style.display = vis ? "none" : "block";
      synBtn.textContent = vis ? "🔤 Синонимы" : "🔤 Синонимы ▲";
    });

    const delBtn = document.createElement("button"); delBtn.className = "btn btn-small"; delBtn.textContent = "Удалить";
    delBtn.style.color = "#c00";
    delBtn.addEventListener("click", () => { columnTemplates.splice(idx,1); persistAll(); renderTemplatesList(); obrRenderTable(); });

    row.appendChild(upBtn); row.appendChild(dnBtn); row.appendChild(inp);
    row.appendChild(synBtn); row.appendChild(delBtn); block.appendChild(row);

    const synPanel = document.createElement("div"); synPanel.className = "syn-panel"; synPanel.style.display = "none";
    renderSynPanel(synPanel, t); block.appendChild(synPanel);
    templatesList.appendChild(block);
  });
}

manageTemplatesBtn.addEventListener("click", e => { e.stopPropagation(); renderTemplatesList(); _updateColSettingsBadge(); templatesModal.style.display = "flex"; newTemplateInput.value = ""; newTemplateInput.focus(); });
closeTemplatesModal.addEventListener("click", () => { templatesModal.style.display = "none"; });
templatesModal.addEventListener("click", e => { if (e.target === templatesModal) templatesModal.style.display = "none"; });

// Мануал-тоггл
(function() {
  const btn = document.getElementById('colDetectManualToggle');
  const body = document.getElementById('colDetectManualBody');
  const arrow = document.getElementById('colDetectManualArrow');
  if (btn && body) {
    btn.addEventListener('click', function() {
      const open = body.classList.toggle('open');
      if (arrow) arrow.textContent = open ? '▼' : '▶';
    });
  }
})();

// Обновляет бейдж «демо / из файла» в шапке модалки
function _updateColSettingsBadge() {
  const badge = document.getElementById('colSettingsSourceBadge');
  const demoBanner = document.getElementById('colSettingsDemoBanner');
  const fileBanner = document.getElementById('colSettingsFileBanner');
  if (!badge) return;
  if (_columnSettingsFromFile) {
    badge.className = 'col-settings-source-badge col-settings-source-badge--file';
    badge.textContent = '✅ из файла';
    if (demoBanner) demoBanner.style.display = 'none';
    if (fileBanner) fileBanner.style.display = '';
  } else {
    badge.className = 'col-settings-source-badge col-settings-source-badge--demo';
    badge.textContent = '📋 демо-данные';
    if (demoBanner) demoBanner.style.display = '';
    if (fileBanner) fileBanner.style.display = 'none';
  }
}

addTemplateBtn.addEventListener("click", () => {
  const v = newTemplateInput.value.trim(); if (!v) return;
  columnTemplates.push(v); persistAll(); renderTemplatesList(); obrRenderTable();
  newTemplateInput.value = ""; newTemplateInput.focus();
});
newTemplateInput.addEventListener("keydown", e => { if (e.key === "Enter") addTemplateBtn.click(); });

// ===== COMPLEX MODE LOGIC =====

function detectComplexPricelist(ws) {
  // Primary: merged cells are the clearest signal
  const merges = ws['!merges'] || [];
  if (merges.length >= 3) return true;
  // Secondary: rows with backslash separator (category\brand format)
  if (tableData) {
    let slashRows = 0;
    const sample = Math.min(tableData.length, 150);
    for (let ri = 0; ri < sample; ri++) {
      const row = tableData[ri] || [];
      const rowText = row.map(c => String(c || '').trim()).join(' ');
      if (/\\/.test(rowText)) { slashRows++; if (slashRows >= 2) return true; }
    }
  }
  return false;
}

function showComplexBanner(show) {
  complexBanner.classList.toggle('visible', show);
}

function setComplexMode(enabled) {
  complexModeEnabled = enabled;
  if (enabled) {
    complexEnableBtn.textContent = '✓ Режим активен — настроить';
    complexEnableBtn.classList.add('active');
    // Immediately open modal so user sees what was found
    renderComplexModal();
    complexModal.style.display = 'flex';
  } else {
    complexEnableBtn.textContent = 'Включить режим подзаголовков';
    complexEnableBtn.classList.remove('active');
    subheaderGroups = [];
  }
}

// Find barcode and name column indices from selectedColumns
function findSpecialCols() {
  let bcCol = -1, nameCol = -1;
  selectedColumns.forEach((label, ci) => {
    const n = normHeader(label || '');
    if (bcCol < 0 && (/штрихкод/.test(n) || /barcode/.test(n) || /\bean\b/.test(n))) bcCol = ci;
    if (nameCol < 0 && (/наименован/.test(n) || /номенклатур/.test(n))) nameCol = ci;
  });
  return { bcCol, nameCol };
}

// Detect the first "data" row index (after title and column header rows)
function findDataStartRow() {
  if (!tableData) return 0;
  const maxCols = Math.max(0, ...tableData.map(r => r ? r.length : 0));
  // Find row where most columns were auto-detected (header row)
  // Heuristic: first row that has ≥ 3 non-empty cells and doesn't have a backslash
  for (let ri = 0; ri < Math.min(tableData.length, 20); ri++) {
    const row = tableData[ri] || [];
    const nonEmpty = row.filter(c => c != null && String(c).trim() !== '');
    const rowText = nonEmpty.map(c => String(c).trim()).join(' ');
    if (nonEmpty.length >= 3 && !/\\/.test(rowText)) {
      return ri + 1; // data starts after this header row
    }
  }
  return 0;
}

function isSubheaderRow(row, bcCol) {
  // Fast path: if barcode column has a valid barcode → definitely product row
  if (bcCol >= 0) {
    const bc = normalizeBarcode(row[bcCol]);
    if (bc && /^\d{6,}$/.test(bc)) return false;
  }
  // Count non-empty cells
  const nonEmpty = row.filter(c => c != null && String(c).trim() !== '');
  if (nonEmpty.length === 0) return false;

  // If only 1-2 cells filled → strong subheader signal
  if (nonEmpty.length <= 2) {
    const text = String(nonEmpty[0]).trim();
    // Exclude pure numbers (totals, page numbers etc.)
    if (/^[\d\s.,]+$/.test(text)) return false;
    // Exclude very short texts (single letter)
    if (text.length < 2) return false;
    return true;
  }
  return false;
}

function parseTokens(rawText) {
  // Split by single backslash (the separator in this supplier's format)
  // Also handle forward slash and multiple backslashes
  const titleCase = s => s.length === 0 ? s : s[0].toUpperCase() + s.slice(1).toLowerCase();
  const seen = new Set();
  return rawText.split(/[\\\/]+/)
    .map(t => t.replace(/^[\s"«»'\u00ab\u00bb]+|[\s"«»'\u00ab\u00bb]+$/g, ''))
    .filter(t => t.length > 1)
    .map(t => titleCase(t))
    .filter(t => { const key = t.toLowerCase(); if (seen.has(key)) return false; seen.add(key); return true; });
}

// Get sample product names for a subheader group
function getSampleProducts(subheaderRow, nameCol, count) {
  if (nameCol < 0 || !tableData) return [];
  const { bcCol } = findSpecialCols();
  const samples = [];
  for (let ri = subheaderRow + 1; ri < tableData.length && samples.length < count; ri++) {
    const row = tableData[ri] || [];
    // Stop at next subheader
    if (isSubheaderRow(row, bcCol)) break;
    const name = String(row[nameCol] || '').trim();
    if (name) samples.push(name);
  }
  return samples;
}

function autoSelectTokens(group) {
  // Auto-select tokens that are NOT already present in sample product names.
  // The first token is typically a broad product group name (e.g. "Бытовая техника")
  // and should never be auto-selected — the user picks it manually if needed.
  if (!group.tokens.length) return;
  const samples = group.samples;

  // Candidates: all tokens except the first one
  const candidates = group.tokens.slice(1);
  if (!candidates.length) {
    // Only one token total — leave empty, let user decide
    group.selectedTokens = [];
    return;
  }

  if (!samples.length) {
    // No samples to analyse — pick last token as specific guess (skip first)
    group.selectedTokens = [candidates[candidates.length - 1]];
    return;
  }

  const missing = candidates.filter(tok => {
    const t = tok.toLowerCase().replace(/\s+/g, ' ');
    return !samples.some(name => name.toLowerCase().includes(t));
  });
  // Use missing tokens from candidates; if all already present — select nothing
  group.selectedTokens = missing.length > 0 ? missing : [];
}

function buildSubheaderGroups() {
  if (!tableData) return [];
  const { bcCol, nameCol } = findSpecialCols();
  const dataStart = findDataStartRow();
  const groups = new Map();

  for (let ri = Math.max(startRowIndex, dataStart); ri < tableData.length; ri++) {
    const row = tableData[ri] || [];
    if (!isSubheaderRow(row, bcCol)) continue;

    // Get text: join all non-empty cells
    const cells = row.filter(c => c != null && String(c).trim() !== '');
    const rawText = cells.map(c => String(c).trim()).join(' \\ ');
    const key = rawText.toLowerCase().replace(/\s+/g, ' ').trim();
    if (!key) continue;

    if (!groups.has(key)) {
      const tokens = parseTokens(rawText);
      // Restore previously saved selection AND skipped state
      const prev = subheaderGroups.find(g => g.key === key) || {};
      const samples = getSampleProducts(ri, nameCol, 5);
      const newGroup = {
        key, rows: [], rawText, tokens,
        selectedTokens: (prev.selectedTokens || []).slice(),
        skipped: prev.skipped || false,
        samples
      };
      // Auto-select tokens only if there's no prior user selection
      if (!prev.selectedTokens) autoSelectTokens(newGroup);
      groups.set(key, newGroup);
    }
    groups.get(key).rows.push(ri);
  }

  return Array.from(groups.values());
}

function renderComplexModal() {
  subheaderGroups = buildSubheaderGroups();
  complexSubheaderList.innerHTML = '';

  if (subheaderGroups.length === 0) {
    complexNoData.style.display = '';
    complexSummaryDiv.style.display = 'none';
    return;
  }
  complexNoData.style.display = 'none';

  const updateSummary = () => {
    const cfg = subheaderGroups.filter(g => g.selectedTokens.length > 0).length;
    const skipped = subheaderGroups.filter(g => g.skipped).length;
    complexSummaryDiv.style.display = '';
    complexSummaryDiv.textContent =
      `Найдено групп: ${subheaderGroups.length}. ` +
      `Настроено (добавят префикс): ${cfg}. ` +
      (skipped ? `Скрыто: ${skipped}.` : '');
  };
  updateSummary();

  // ---- Quick-action bar ----
  let quickBar = complexSubheaderList.previousElementSibling;
  if (!quickBar || !quickBar.classList.contains('complex-quick-bar')) {
    quickBar = document.createElement('div');
    quickBar.className = 'complex-quick-bar';
    quickBar.style.cssText = 'display:flex;gap:8px;align-items:center;margin-bottom:10px;flex-wrap:wrap;';
    complexSubheaderList.parentNode.insertBefore(quickBar, complexSubheaderList);
  }
  quickBar.innerHTML = '';

  const makeQBtn = (label, title, fn) => {
    const b = document.createElement('button');
    b.className = 'btn btn-small'; b.textContent = label; b.title = title;
    b.addEventListener('click', () => { fn(); renderComplexModal(); });
    return b;
  };

  quickBar.appendChild(makeQBtn('🤖 Авто-выбор', 'Автоматически выбрать токены по наличию/отсутствию в наименованиях товаров', () => {
    subheaderGroups.forEach(g => { g.skipped = false; autoSelectTokens(g); });
  }));
  quickBar.appendChild(makeQBtn('☑ Выбрать все первые', 'Выбрать первый токен каждой группы', () => {
    subheaderGroups.forEach(g => { g.skipped = false; if (g.tokens.length) g.selectedTokens = [g.tokens[0]]; });
  }));
  quickBar.appendChild(makeQBtn('☑ Выбрать все токены', 'Выбрать все токены каждой группы', () => {
    subheaderGroups.forEach(g => { g.skipped = false; g.selectedTokens = g.tokens.slice(); });
  }));
  quickBar.appendChild(makeQBtn('✕ Снять всё', 'Сбросить выбор во всех группах', () => {
    subheaderGroups.forEach(g => { g.selectedTokens = []; g.skipped = false; });
  }));

  subheaderGroups.forEach((group) => {
    // ---- card wrapper ----
    const card = document.createElement('div');
    card.className = 'subheader-group' + (group.skipped ? ' is-skipped' : '');

    // ---- header row (always visible, click to toggle body) ----
    const head = document.createElement('div');
    head.className = 'subheader-group-head';

    const headTitle = document.createElement('span');
    headTitle.className = 'subheader-group-head-title';
    headTitle.textContent = group.rawText;
    head.appendChild(headTitle);

    const badge = document.createElement('span');
    badge.className = 'subheader-group-head-badge ' + (group.selectedTokens.length > 0 ? 'badge-ok' : 'badge-skip');
    badge.textContent = group.skipped
      ? '⊘ скрыта'
      : group.selectedTokens.length > 0
        ? '+ ' + group.selectedTokens.join(' ')
        : '— без префикса';
    head.appendChild(badge);

    // toggle body on header click
    head.addEventListener('click', () => {
      if (group.skipped) {
        group.skipped = false;
        card.classList.remove('is-skipped');
        badge.className = 'subheader-group-head-badge badge-skip';
        badge.textContent = '— без префикса';
        updateSummary();
      }
    });
    card.appendChild(head);

    // ---- body ----
    const body = document.createElement('div');
    body.className = 'subheader-group-body';

    // Samples (product names)
    if (group.samples.length > 0) {
      const samplesDiv = document.createElement('div');
      samplesDiv.className = 'sg-samples';
      group.samples.forEach(s => {
        const row = document.createElement('div');
        row.className = 'sg-sample';
        row.textContent = s;
        samplesDiv.appendChild(row);
      });
      body.appendChild(samplesDiv);
    }

    // Instruction
    const instr = document.createElement('div');
    instr.style.cssText = 'font-size:12px;color:#888;margin-bottom:6px;';
    instr.textContent = 'Выберите токены для добавления в начало наименования:';
    body.appendChild(instr);

    // Token chips
    const chipsRow = document.createElement('div');
    chipsRow.className = 'token-chips';

    const refreshPreview = () => {
      previewList.innerHTML = '';
      const prefix = group.selectedTokens.join(' ');
      group.samples.forEach(name => {
        const item = document.createElement('div');
        item.className = 'sg-preview-item';
        if (prefix === '') {
          item.className += ' preview-none';
          item.textContent = name.length > 80 ? name.slice(0,78)+'…' : name;
        } else if (prefixContainedInName(name, prefix)) {
          item.className += ' preview-skip';
          item.textContent = '⚠ уже есть: ' + (name.length > 75 ? name.slice(0,73)+'…' : name);
        } else {
          item.className += ' preview-added';
          const full = prefix + ' ' + name;
          item.textContent = '→ ' + (full.length > 80 ? full.slice(0,78)+'…' : full);
        }
        previewList.appendChild(item);
      });
      // update badge
      badge.className = 'subheader-group-head-badge ' + (group.selectedTokens.length > 0 ? 'badge-ok' : 'badge-skip');
      badge.textContent = group.skipped
        ? '⊘ скрыта'
        : group.selectedTokens.length > 0
          ? '+ ' + group.selectedTokens.join(' ')
          : '— без префикса';
      updateSummary();
    };

    group.tokens.forEach(token => {
      const chip = document.createElement('span');
      chip.className = 'token-chip' + (group.selectedTokens.includes(token) ? ' selected' : '');
      chip.textContent = token;
      chip.addEventListener('click', () => {
        const idx = group.selectedTokens.indexOf(token);
        if (idx >= 0) group.selectedTokens.splice(idx, 1);
        else group.selectedTokens.push(token);
        chip.classList.toggle('selected', group.selectedTokens.includes(token));
        refreshPreview();
      });
      chipsRow.appendChild(chip);
    });
    body.appendChild(chipsRow);

    // Skip button
    const skipBtn = document.createElement('button');
    skipBtn.className = 'sg-skip-btn';
    skipBtn.textContent = '⊘ скрыть группу (не добавлять префикс)';
    skipBtn.addEventListener('click', () => {
      group.skipped = true;
      group.selectedTokens = [];
      card.classList.add('is-skipped');
      badge.className = 'subheader-group-head-badge badge-skip';
      badge.textContent = '⊘ скрыта';
      updateSummary();
    });
    body.appendChild(skipBtn);

    // Preview list
    const previewLabel = document.createElement('div');
    previewLabel.style.cssText = 'font-size:11px;color:#888;margin-top:10px;margin-bottom:3px;';
    previewLabel.textContent = 'Предпросмотр наименований:';
    body.appendChild(previewLabel);

    const previewList = document.createElement('div');
    previewList.className = 'sg-preview-list';
    body.appendChild(previewList);
    refreshPreview();

    card.appendChild(body);
    complexSubheaderList.appendChild(card);
  });
}

// Build rowIndex → prefix map for CSV export
// FIXED: iterate ALL rows linearly, tracking current prefix.
// Reset prefix at ANY subheader row (even those with no selected tokens).
function buildPrefixMap() {
  if (!complexModeEnabled || subheaderGroups.length === 0) return null;
  const { bcCol } = findSpecialCols();

  // Build lookup: subheaderGroupsByKey for O(1) access
  const groupByKey = new Map();
  subheaderGroups.forEach(g => groupByKey.set(g.key, g));

  const dataStart = findDataStartRow();
  const prefixMap = new Map();
  let currentPrefix = '';

  for (let ri = Math.max(startRowIndex, dataStart); ri < tableData.length; ri++) {
    const row = tableData[ri] || [];

    if (isSubheaderRow(row, bcCol)) {
      // Find which group this row belongs to
      const cells = row.filter(c => c != null && String(c).trim() !== '');
      const rawText = cells.map(c => String(c).trim()).join(' \\ ');
      const key = rawText.toLowerCase().replace(/\s+/g, ' ').trim();
      const group = groupByKey.get(key);
      // Always reset prefix (even if group has no tokens — stops prev group leaking)
      currentPrefix = (group && !group.skipped && group.selectedTokens.length > 0)
        ? group.selectedTokens.join(' ')
        : '';
      continue;
    }

    if (currentPrefix) prefixMap.set(ri, currentPrefix);
  }

  return prefixMap.size > 0 ? prefixMap : null;
}

function prefixContainedInName(name, prefix) {
  const nameLow = name.toLowerCase();
  return prefix.toLowerCase().split(/\s+/).every(tok => tok.length > 1 && nameLow.includes(tok));
}

// ===== COMPLEX MODE EVENTS =====
complexEnableBtn.addEventListener('click', () => {
  setComplexMode(!complexModeEnabled);
});
complexDismiss.addEventListener('click', () => {
  showComplexBanner(false);
});
// "Подзаголовки" button in toolbar (still kept for re-opening)
complexConfigBtn.addEventListener('click', () => {
  renderComplexModal();
  complexModal.style.display = 'flex';
});
closeComplexModal.addEventListener('click', () => { complexModal.style.display = 'none'; });
complexModal.addEventListener('click', e => { if (e.target === complexModal) complexModal.style.display = 'none'; });
complexApplyBtn.addEventListener('click', () => {
  complexModal.style.display = 'none';
  const configured = subheaderGroups.filter(g => g.selectedTokens.length > 0).length;
  complexEnableBtn.textContent = configured > 0
    ? `✓ Режим активен (${configured} групп) — изменить`
    : '✓ Режим активен — настроить';
});
complexResetBtn.addEventListener('click', () => {
  subheaderGroups.forEach(g => { g.selectedTokens = []; g.skipped = false; autoSelectTokens(g); });
  renderComplexModal();
});


// OBR file input handlers are already attached in the OBR script above

// ════════════════════════════════════════════════════════════════════════════
// APPBRIDGE: CSV injection from OBR into PriceMatcher
// ════════════════════════════════════════════════════════════════════════════
AppBridge.on('csvReady', async function(data) {
  // data = { csvText, fileName, isMyPrice }
  const { csvText, fileName, isMyPrice } = data;
  
  // Update loaded-files list
  const loadedPanel = document.getElementById('obr-loaded-files');
  const loadedList  = document.getElementById('obr-loaded-list');
  if (loadedPanel && loadedList) {
    loadedPanel.style.display = 'flex';
    const chip = document.createElement('span');
    chip.style.cssText = 'background:#c6efce;border:1px solid #70ad47;border-radius:3px;padding:1px 8px;font-weight:600;white-space:nowrap;';
    chip.textContent = (isMyPrice ? '🏷️ ' : '📦 ') + fileName;
    loadedList.appendChild(chip);
  }

  // _pmApp is set synchronously by the monitor script block - should always be ready
  if (!window._pmApp) {
    console.error('AppBridge csvReady: _pmApp not available');
    showToast('Ошибка: модуль мониторинга не инициализирован', 'err');
    return;
  }

  const pm = window._pmApp;

  // Convert CSV text to File object and parse it
  const blob = new Blob([csvText], { type: 'text/csv;charset=utf-8' });
  const file = new File([blob], fileName, { type: 'text/csv' });
  const displayName = isMyPrice ? 'Мой прайс' : pm.removeFileExtension(fileName);

  let fileData;
  try {
    fileData = await pm.parseFile(file, displayName);
  } catch(parseErr) {
    console.error('AppBridge: CSV parse error:', parseErr);
    showToast('Ошибка разбора CSV: ' + parseErr.message, 'err');
    return;
  }

  try {
    if (isMyPrice) {
      pm.myPriceData = fileData;
      pm.updateMyPriceStatus(fileName);
    } else {
      const added = pm.addCompetitorFile(fileData);
      if (added === false) return; // user cancelled replacement
      pm.updateCompetitorStatus();
    }
  } catch(stateErr) {
    console.error('AppBridge: state update error:', stateErr);
  }

  try {
    pm.processAllData();
  } catch(procErr) {
    console.error('AppBridge: processAllData error:', procErr);
    showToast('Ошибка формирования таблицы: ' + procErr.message, 'err');
  }
});

// ════════════════════════════════════════════════════════════════════════════
// APPBRIDGE: Column settings sync (OBR ↔ JSON)
// ════════════════════════════════════════════════════════════════════════════
// When unified JSON is loaded, update OBR column templates/synonyms
AppBridge.on('settingsLoaded', function(data) {
  if (data && data.columnSettings) {
    const cs = data.columnSettings;
    if (Array.isArray(cs.templates) && cs.templates.length) {
      columnTemplates = cs.templates.slice();
    }
    if (cs.synonyms && typeof cs.synonyms === 'object') {
      columnSynonyms = JSON.parse(JSON.stringify(cs.synonyms));
    }
    // Данные пришли из файла — снимаем демо-флаг
    _columnSettingsFromFile = true;
    persistAll(false); // загрузка из файла — не помечаем как «есть несохранённые изменения»
    _updateColSettingsBadge();
    if (typeof renderTemplatesList === 'function') renderTemplatesList();
  }
});

// CSV injection is handled directly in the OBR download handlers above via AppBridge.emit

// ════════════════════════════════════════════════════════════════════════════
// OBR JSON SETTINGS: load from / save to unified JSON
// ════════════════════════════════════════════════════════════════════════════
(function() {
  const settingsInput = document.getElementById('obrSettingsJsonInput');
  if (settingsInput) {
    settingsInput.addEventListener('change', function(e) {
      const file = e.target.files[0]; if (!file) return;
      const reader = new FileReader();
      reader.onload = function(ev) {
        try {
          const json = JSON.parse(ev.target.result);
          AppBridge.emit('settingsLoaded', json);
          // Also pass to PriceMatcher
          if (json.barcodes || json.brands) {
            // Simulate loading into PM synonymsInput
            const blob = new Blob([ev.target.result], { type: 'application/json' });
            const f = new File([blob], file.name, { type: 'application/json' });
            const dt = new DataTransfer(); dt.items.add(f);
            const inp = document.getElementById('synonymsInput');
            if (inp) { inp.files = dt.files; inp.dispatchEvent(new Event('change', { bubbles: true })); }
          }
          // Show success toast if available
          if (typeof showToast === 'function') showToast('Настройки загружены', 'ok');
        } catch(err) {
          alert('Ошибка чтения JSON: ' + err.message);
        }
      };
      reader.readAsText(file, 'utf-8');
      e.target.value = '';
    });
  }
})();

// downloadCurrentSynonyms is already extended in PM script1 to include columnSettings


// ════════════════════════════════════════════════════════════════════════════
// JSON EDITOR + BRAND DATABASE + SIDEBAR HELPERS
// ════════════════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════════════════
// JSON EDITOR (jeDB)
// ════════════════════════════════════════════════════════════════════════════
let jeDB = {};
let jeChanges = 0;
const JE_HISTORY_LIMIT = 50;
let jeUndoStack = [], jeRedoStack = [];
let _jeDupsCache = null;
let _jeVsKeys = [];
let _jeVsKeyIndex = new Map();

// Virtual scroll
const JE_VS = { ROW_H: 40, OVERSCAN: 20, start: 0, end: 0, ticking: false };
const JE_VS_THRESHOLD = 100; // render all rows without virtual scroll if total <= this

function jeDBSaveHistory() {
  jeUndoStack.push(JSON.stringify(jeDB));
  if (jeUndoStack.length > JE_HISTORY_LIMIT) jeUndoStack.shift();
  jeRedoStack = []; jeUpdateUndoUI();
}
function jeUndo() {
  if (!jeUndoStack.length) return;
  jeRedoStack.push(JSON.stringify(jeDB));
  jeDB = JSON.parse(jeUndoStack.pop());
  _jeDupsCache = null; jeChanges = Math.max(0, jeChanges - 1);
  jeUpdateUndoUI(); jeDBNotifyChange(false); jeRenderEditor(true); // preserve scroll on undo
}
function jeRedo() {
  if (!jeRedoStack.length) return;
  jeUndoStack.push(JSON.stringify(jeDB));
  jeDB = JSON.parse(jeRedoStack.pop());
  _jeDupsCache = null; jeChanges++;
  jeUpdateUndoUI(); jeDBNotifyChange(false); jeRenderEditor(true); // preserve scroll on redo
}
function jeUpdateUndoUI() {
  document.getElementById('jeUndoBtn').disabled = !jeUndoStack.length;
  document.getElementById('jeRedoBtn').disabled = !jeRedoStack.length;
}
function jeDBNotifyChange(bump) {
  _jeDupsCache = null;
  if (bump !== false) jeChanges++;
  unifiedMarkUnsaved(true);
  jeUpdateStatus();
  rebuildBarcodeAliasFromJeDB(true);
  updateMatchPairTags();
  if (typeof window._matcherUpdateJsonInfo === 'function') window._matcherUpdateJsonInfo();
}
function jeUpdateStatus() {
  const n = Object.keys(jeDB).length;
  let synTotal = 0;
  for (const v of Object.values(jeDB)) if (Array.isArray(v)) synTotal += Math.max(0, v.length - 1);
  document.getElementById('jeStatus').textContent = `Наименований: ${n}  |  Синонимов: ${synTotal}  |  Показано: ${_jeVsKeys.length}`;
  document.getElementById('jeExportXlsxBtn').disabled = !n;
  document.getElementById('jeClearBtn').disabled = !n;
  document.getElementById('jeSearchInp').disabled = !n;
  const dups = jeFindDuplicates();
  const dc = dups.size;
  const dupEl = document.getElementById('jeDupStatus');
  if (dc > 0) { dupEl.textContent = `⚠️ Дублей: ${dc}`; dupEl.style.display = ''; }
  else dupEl.style.display = 'none';
  // show/hide table
  const hasData = n > 0;
  document.getElementById('jeTable').style.display = hasData ? '' : 'none';
  document.getElementById('jeEmpty').style.display = hasData ? 'none' : '';
}
function jeFindDuplicates() {
  if (_jeDupsCache) return _jeDupsCache;
  const seen = new Map();
  for (const [key, val] of Object.entries(jeDB)) {
    seen.set(key, (seen.get(key)||0)+1);
    if (Array.isArray(val)) for (let i = 1; i < val.length; i++) {
      const s = String(val[i]).trim(); if (!s) continue;
      seen.set(s, (seen.get(s)||0)+1);
    }
  }
  const dups = new Set();
  for (const [bc, cnt] of seen) if (cnt > 1) dups.add(bc);
  _jeDupsCache = dups; return dups;
}
function jeGetAllBarcodes() {
  const all = new Set();
  for (const [key, val] of Object.entries(jeDB)) {
    all.add(key);
    if (Array.isArray(val)) val.slice(1).forEach(s => { s = String(s).trim(); if (s) all.add(s); });
  }
  return all;
}

function esc(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function escv(s) { return esc(s).replace(/"/g,'&quot;').replace(/'/g,'&#039;'); }
function jeSafeId(s) { try { return btoa(unescape(encodeURIComponent(s))).replace(/=/g,''); } catch { return String(s).replace(/\W/g,'_'); } }

function jeBuildEditorRow(key, absIdx, dups) {
  const val = jeDB[key] || [];
  const name = val[0] || '';
  // Build synonym pills — data-si stores the REAL 1-based index in val[] (val[0]=name)
  const pillsHtml = val.slice(1).reduce((acc, syn, i) => {
    const s = String(syn).trim();
    if (!s) return acc;
    const isDup = dups.has(s);
    const wCls = isDup ? 'syn-wrap dup' : 'syn-wrap';
    return acc + `<span class="${wCls}"><span class="syn-pill" title="${escv(s)}">${esc(s)}</span><span class="syn-x" data-key="${escv(key)}" data-si="${i + 1}" role="button" title="Удалить синоним">×</span></span>`;
  }, '');
  return `<tr id="jer-${jeSafeId(key)}">
    <td style="text-align:center;color:#999;font-size:11px;width:36px;">${absIdx+1}</td>
    <td><input class="je-inp-cell" value="${escv(name)}" data-namekey="${escv(key)}" style="min-width:140px;width:100%;"></td>
    <td><input class="je-inp-cell mono${dups.has(key)?' dup-inp':''}" value="${escv(key)}" data-origkey="${escv(key)}" style="font-family:monospace;width:100%;"></td>
    <td><div class="syn-cell">${pillsHtml}<input class="inp-add-syn" placeholder="+ ШК" data-key="${escv(key)}" title="Enter для добавления"></div></td>
    <td><button class="je-del-btn" data-delkey="${escv(key)}" title="Удалить группу">🗑</button></td>
  </tr>`;
}

function jeRenderVS() {
  const wrap = document.getElementById('jeTableWrap');
  if (!wrap) return;
  const total = _jeVsKeys.length;
  const scrollTop = wrap.scrollTop;
  const viewH = wrap.clientHeight || 500;
  // FIX: if all rows fit in view, render all without virtual scroll to avoid click-blocking issues
  if (total * JE_VS.ROW_H <= viewH + JE_VS.ROW_H * 2 || total <= JE_VS_THRESHOLD) {
    JE_VS.start = 0;
    JE_VS.end = total;
  } else {
    JE_VS.start = Math.max(0, Math.floor(scrollTop / JE_VS.ROW_H) - JE_VS.OVERSCAN);
    JE_VS.end = Math.min(total, Math.ceil((scrollTop + viewH) / JE_VS.ROW_H) + JE_VS.OVERSCAN);
  }
  const topPad = JE_VS.start * JE_VS.ROW_H;
  const botPad = Math.max(0, total - JE_VS.end) * JE_VS.ROW_H;
  const dups = jeFindDuplicates();
  const rows = _jeVsKeys.slice(JE_VS.start, JE_VS.end).map((key, rel) => jeBuildEditorRow(key, JE_VS.start + rel, dups)).join('');
  document.getElementById('jeTbody').innerHTML =
    (topPad > 0 ? `<tr style="pointer-events:none;"><td colspan="5" style="height:${topPad}px;padding:0;border:none;"></td></tr>` : '') +
    rows +
    (botPad > 0 ? `<tr style="pointer-events:none;"><td colspan="5" style="height:${botPad}px;padding:0;border:none;"></td></tr>` : '');
}

function jePatchRow(key) { jePatchRowSafe(key); }

function jeRenderEditor(preserveScroll = false) {
  const query = (document.getElementById('jeSearchInp').value||'').toLowerCase().trim();
  const keys = Object.keys(jeDB);
  _jeVsKeys = keys.filter(k => {
    if (!query) return true;
    const v = jeDB[k] || [];
    return k.includes(query) || (v[0]||'').toLowerCase().includes(query) || v.slice(1).join(' ').toLowerCase().includes(query);
  });
  _jeVsKeyIndex.clear();
  _jeVsKeys.forEach((k, i) => _jeVsKeyIndex.set(k, i));
  const wrap = document.getElementById('jeTableWrap');
  const _savedJeScroll = (preserveScroll && wrap) ? wrap.scrollTop : 0;
  if (wrap && !preserveScroll) wrap.scrollTop = 0;
  jeRenderVS();
  jeUpdateStatus();
  if (preserveScroll && _savedJeScroll > 0 && wrap) {
    requestAnimationFrame(() => { wrap.scrollTop = _savedJeScroll; });
  }
}

// ── Scroll editor to a key row after creation ──────────────────────────────
function jeScrollToKey(key) {
  const idx = _jeVsKeyIndex.get(String(key));
  if (idx === undefined) return;
  const wrap = document.getElementById('jeTableWrap');
  if (!wrap) return;
  const targetTop = Math.max(0, idx * JE_VS.ROW_H - wrap.clientHeight / 2);
  requestAnimationFrame(() => {
    wrap.scrollTop = targetTop;
    jeRenderVS();
    setTimeout(() => {
      const rowEl = document.getElementById('jer-' + jeSafeId(String(key)));
      if (rowEl) { rowEl.style.transition = 'background 0.5s'; rowEl.style.background = '#fffde7'; setTimeout(() => { rowEl.style.background = ''; }, 1200); }
    }, 80);
  });
}

// Virtual scroll event
const jeWrap = document.getElementById('jeTableWrap');
if (jeWrap) {
  jeWrap.addEventListener('scroll', () => {
    if (!JE_VS.ticking) {
      JE_VS.ticking = true;
      requestAnimationFrame(() => { jeRenderVS(); JE_VS.ticking = false; });
    }
  }, { passive: true });
}




// ── Надёжное обновление строки редактора (patch DOM или full re-render) ──
function jeForceUpdateRow(key) {
  try {
    const rowEl = document.getElementById('jer-' + jeSafeId(key));
    if (rowEl) {
      const dups = jeFindDuplicates();
      const idx = _jeVsKeyIndex.get(key) ?? 0;
      const html = jeBuildEditorRow(key, idx, dups);
      const tmp = document.createElement('tbody');
      tmp.innerHTML = html;
      const newRow = tmp.firstElementChild;
      if (newRow) { rowEl.replaceWith(newRow); return; }
    }
  } catch(err) { /* fall through */ }
  const wrap = document.getElementById('jeTableWrap');
  const scroll = wrap ? wrap.scrollTop : 0;
  jeRenderVS();
  if (wrap) requestAnimationFrame(() => { wrap.scrollTop = scroll; });
}

// ── Safe patch: falls back to full re-render if row not found ──────────
function jePatchRowSafe(key) {
  const rowEl = document.getElementById('jer-' + jeSafeId(key));
  const idx = _jeVsKeyIndex.get(key);
  if (!rowEl || idx === undefined) {
    // Fallback: full re-render but preserve scroll
    const wrap = document.getElementById('jeTableWrap');
    const scroll = wrap ? wrap.scrollTop : 0;
    jeRenderVS();
    if (wrap) wrap.scrollTop = scroll;
    return;
  }
  const dups = jeFindDuplicates();
  const tmp = document.createElement('tbody');
  tmp.innerHTML = jeBuildEditorRow(key, idx, dups);
  rowEl.replaceWith(tmp.firstElementChild);
}
// Table events (delegation)
const jeTbody = document.getElementById('jeTbody');
jeTbody.addEventListener('keydown', function(e) {
  if (e.target.classList.contains('inp-add-syn') && e.key === 'Enter') {
    e.preventDefault(); jeSaveSynInput(e.target);
  } else if (e.target.classList.contains('je-inp-cell') && e.target.dataset.origkey !== undefined && e.key === 'Enter') {
    e.preventDefault(); e.target.blur();
  }
});
jeTbody.addEventListener('focusout', function(e) {
  if (e.target.classList.contains('inp-add-syn')) jeSaveSynInput(e.target);
  else if (e.target.classList.contains('je-inp-cell') && e.target.dataset.origkey !== undefined) jeRenameMainBC(e.target);
});
jeTbody.addEventListener('input', function(e) {
  if (e.target.classList.contains('inp-add-syn')) jeCheckSynDup(e.target);
  else if (e.target.classList.contains('je-inp-cell') && e.target.dataset.namekey) {
    const k = e.target.dataset.namekey;
    if (jeDB[k]) { jeDB[k][0] = e.target.value; jeDBNotifyChange(); }
  }
});
jeTbody.addEventListener('click', function(e) {
  const x = e.target.closest('.syn-x[data-key]');
  if (x) {
    const key = x.dataset.key, si = parseInt(x.dataset.si, 10);
    // si is real 1-based index in val[]; val[0] is the name
    if (!jeDB[key] || isNaN(si) || si < 1 || si >= jeDB[key].length) return;
    jeDBSaveHistory();
    jeDB[key].splice(si, 1);
    jeDBNotifyChange();
    jeForceUpdateRow(key);
    return;
  }
  const d = e.target.closest('[data-delkey]');
  if (d) {
    const key = d.dataset.delkey;
    if (!jeDB[key]) return;
    jeConfirmDialog('Удалить группу «' + key + '»?', '🗑 Удаление').then(function(ok) {
      if (!ok) return;
      jeDBSaveHistory(); delete jeDB[key];
      jeDBNotifyChange(); jeRenderEditor(true); // preserve scroll after deletion
    });
  }
});

function jeCheckSynDup(input) {
  const val = (input.value||'').trim();
  if (!val) { input.style.borderColor = ''; return; }
  const all = jeGetAllBarcodes();
  if (all.has(val)) input.style.borderColor = '#e8a000';
  else input.style.borderColor = '#217346';
}
function jeSaveSynInput(input) {
  const key = input.dataset.key, val = (input.value||'').trim();
  input.style.borderColor = '';
  if (!val || !jeDB[key]) return;
  if (jeDB[key].includes(val)) { showToast(`ШК «${val}» уже есть в группе`, 'warn'); input.style.borderColor = '#d93025'; setTimeout(() => { input.style.borderColor = ''; input.select(); }, 1400); return; }
  const all = jeGetAllBarcodes();
  if (all.has(val)) { showToast(`ШК «${val}» уже существует в другой группе`, 'warn'); input.style.borderColor = '#e8a000'; setTimeout(() => { input.style.borderColor = ''; input.select(); }, 1400); return; }
  jeDBSaveHistory(); jeDB[key].push(val); input.value = '';
  jeDBNotifyChange(); jePatchRow(key);
}
function jeRenameMainBC(input) {
  const oldKey = input.dataset.origkey, newKey = (input.value||'').trim();
  if (!newKey || newKey === oldKey) return;
  if (jeDB[newKey]) { showToast(`ШК «${newKey}» уже существует`, 'warn'); input.value = oldKey; return; }
  if (!jeDB[oldKey]) { input.value = ''; return; }
  jeDBSaveHistory(); jeDB[newKey] = jeDB[oldKey]; delete jeDB[oldKey];
  jeDBNotifyChange(); jeRenderEditor(true); // preserve scroll on rename
}

// New group form
document.getElementById('jeCreateBtn').addEventListener('click', function() {
  const name = document.getElementById('jeNName').value.trim();
  const mainBC = document.getElementById('jeNMainBC').value.trim();
  const synsRaw = document.getElementById('jeNSyns').value.split(',').map(s=>s.trim()).filter(Boolean);

  if (!mainBC) { showToast('Введите главный штрихкод', 'warn'); document.getElementById('jeNMainBC').focus(); return; }

  const all = jeGetAllBarcodes(); // Set of strings
  if (all.has(mainBC)) {
    showToast(`ШК «${mainBC}» уже существует в базе как ${jeDB[mainBC] ? 'главный' : 'синоним'}`, 'warn');
    document.getElementById('jeNMainBC').focus(); return;
  }

  const dupSyn = synsRaw.find(s => all.has(s));
  const doCreate = () => {
    jeDBSaveHistory();
    jeDB[mainBC] = [name || mainBC, ...synsRaw];
    jeDBNotifyChange();
    jeRenderEditor(true); // preserve scroll: don't jump to top
    jeClearForm();
    // Switch to editor tab so user can see the new entry
    switchMainPane('jsoneditor');
    setTimeout(() => jeScrollToKey(mainBC), 60); // scroll to new entry with highlight
    showToast(`Группа «${mainBC}» создана`, 'ok');
  };

  if (dupSyn) {
    jeConfirmDialog(`Синоним «${dupSyn}» уже есть в базе. Всё равно добавить?`, '⚠️ Дубль').then(ok => {
      if (!ok) return;
      doCreate();
    });
    return;
  }
  doCreate();
});
document.getElementById('jeClearFormBtn').addEventListener('click', jeClearForm);
function jeClearForm() {
  ['jeNName','jeNMainBC','jeNSyns'].forEach(id => document.getElementById(id).value = '');
}


// ════════════════════════════════════════════════════════════════════════════
// UNIFIED SAVE BAR — sub-tab switching
// ════════════════════════════════════════════════════════════════════════════
document.querySelectorAll('.syn-subtab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.syn-subtab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.syn-subtab-pane').forEach(p => p.classList.remove('active'));
    tab.classList.add('active');
    const pane = document.getElementById('subpane-' + tab.dataset.subtab);
    if (pane) pane.classList.add('active');
  });
});

// ── jeDB undo/redo
document.getElementById('jeUndoBtn').addEventListener('click', jeUndo);
document.getElementById('jeRedoBtn').addEventListener('click', jeRedo);
document.getElementById('jeSearchInp').addEventListener('input', jeRenderEditor);

// ── jeDB clear
document.getElementById('jeClearBtn').addEventListener('click', function() {
  jeConfirmDialog('Очистить всю базу штрихкодов?', '🗑 Очистка').then(ok => {
    if (!ok) return;
    jeDBSaveHistory(); jeDB = {}; _jeDupsCache = null; jeChanges = 0;
    document.getElementById('jeSearchInp').value = '';
    jeUpdateStatus(); jeRenderEditor(); rebuildBarcodeAliasFromJeDB();
    unifiedMarkUnsaved();
  });
});

// ── Excel export/import for barcodes
document.getElementById('jeExportXlsxBtn').addEventListener('click', async function() {
  const rows = [['Штрихкод','Наименование','Синонимы ШК']];
  for (const [k,v] of Object.entries(jeDB)) rows.push([k, Array.isArray(v)?(v[0]||''):'', Array.isArray(v)?v.slice(1).join(', '):'']);
  const ws = XLSX.utils.aoa_to_sheet(rows); ws['!cols'] = [{wch:20},{wch:44},{wch:60}];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Синонимы');
  const buf = XLSX.write(wb, {bookType:'xlsx',type:'array'});
  const _sxBlob = new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const _sxFname = `synonyms_${new Date().toISOString().slice(0,10)}.xlsx`;
  await saveBlobWithDialogOrDownload(_sxBlob, _sxFname);
});

let _jeXlsResolve = null;
function jeXlsModalClose(mode) {
  document.getElementById('jeXlsModal').style.display = 'none';
  if (_jeXlsResolve) { _jeXlsResolve(mode); _jeXlsResolve = null; }
}
document.getElementById('jeImportXlsxBtn').addEventListener('click', () => document.getElementById('jeXlsxFileIn').click());
document.getElementById('jeXlsxFileIn').addEventListener('change', async function(e) {
  const file = e.target.files[0]; if (!file) return; e.target.value = '';
  try {
    const data = await new Promise((res,rej) => {
      const r = new FileReader(); r.onload = ev => {
        try { const wb=XLSX.read(new Uint8Array(ev.target.result),{type:'array'}); res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]])); } catch(err){rej(err);}
      }; r.readAsArrayBuffer(file);
    });
    const bcCols=['штрихкод','штрих-код','barcode','шк','ean','код'];
    const nmCols=['название','наименование','name','товар','продукт','наим'];
    const snCols=['синонимы','synonyms','синоним','synonym'];
    const cols = data.length ? Object.keys(data[0]) : [];
    const bcCol = cols.find(c=>bcCols.some(v=>c.toLowerCase().includes(v)))||cols[0];
    const nmCol = cols.find(c=>nmCols.some(v=>c.toLowerCase().includes(v)))||cols[1];
    const snCol = cols.find(c=>snCols.some(v=>c.toLowerCase().includes(v)))||null;
    if (!bcCol) { showToast('Не найдена колонка штрихкода', 'warn'); return; }
    const validRows = data.filter(r=>String(r[bcCol]||'').trim());
    const conflictCnt = validRows.filter(r=>jeDB[String(r[bcCol]).trim()]).length;
    let mode = 'overwrite';
    if (conflictCnt > 0) {
      document.getElementById('jeXlsModalMsg').textContent = `${conflictCnt} из ${validRows.length} записей уже есть в базе. Как поступить?`;
      document.getElementById('jeXlsModal').style.display = 'flex';
      mode = await new Promise(res => { _jeXlsResolve = res; });
      if (!mode) return;
    }
    jeDBSaveHistory();
    let added = 0, skipped = 0;
    for (const row of validRows) {
      const bc = String(row[bcCol]).trim();
      const name = nmCol ? String(row[nmCol]||'').trim() : bc;
      const syns = snCol ? String(row[snCol]||'').split(',').map(s=>s.trim()).filter(Boolean) : [];
      if (jeDB[bc] && mode==='skip') { skipped++; continue; }
      if (jeDB[bc] && mode==='merge') {
        const ex=jeDB[bc];const exSet=new Set(ex.slice(1).map(s=>String(s).trim()));
        syns.forEach(s=>{if(s&&!exSet.has(s))ex.push(s);});
        if(name&&!ex[0])ex[0]=name; added++; continue;
      }
      jeDB[bc]=[name,...syns]; added++;
    }
    jeDBNotifyChange(); jeRenderEditor();
    unifiedMarkUnsaved();
    showToast(mode==='skip'?`Добавлено: ${added}, пропущено: ${skipped}`:`Обработано: ${added}`, 'ok');
  } catch(err) { showToast('Ошибка импорта: '+err.message, 'warn'); }
});

// ════════════════════════════════════════════════════════════════════════════
// SYNC: jeDB ↔ barcodeAliasMap
// ════════════════════════════════════════════════════════════════════════════
function rebuildBarcodeAliasFromJeDB(_skip) {
  if (typeof barcodeAliasMap === 'undefined') return;
  const m = new Map();
  for (const [key, val] of Object.entries(jeDB)) {
    m.set(key, key);
    if (Array.isArray(val)) val.slice(1).forEach(s => { s = String(s).trim(); if (s) m.set(s, key); });
  }
  barcodeAliasMap.clear();
  m.forEach((v, k) => barcodeAliasMap.set(k, v));
  synonymsLoaded = m.size > 0;
  const st = document.getElementById('synonymsStatus');
  if (st && m.size > 0) {
    st.className = 'upload-status upload-status--ok';
    st.textContent = '✅ Групп: ' + Object.keys(jeDB).length;
  }
  // Update barcode count badge
  const bcBadge = document.getElementById('bcCountBadge');
  if (bcBadge) bcBadge.textContent = Object.keys(jeDB).length;
  if (!_skip && typeof allFilesData !== 'undefined' && allFilesData.length > 0) {
    clearTimeout(rebuildBarcodeAliasFromJeDB._t);
    rebuildBarcodeAliasFromJeDB._t = setTimeout(function() { processData(); renderTable(true); updateUI(); }, 80);
  }
}

// When synonyms.json is loaded — single source of truth handler
document.getElementById('synonymsInput').addEventListener('change', function(e) {
  const file = e.target.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const json = JSON.parse(ev.target.result);
      if (!json || typeof json !== 'object' || Array.isArray(json)) return;
      // Support both combined format { barcodes, brands } and legacy flat format
      const hasBarcodes = 'barcodes' in json;
      const barcodes = hasBarcodes ? json.barcodes : json;
      jeDB = JSON.parse(JSON.stringify(barcodes));
      if (json.brands) {
        _brandDB = JSON.parse(JSON.stringify(json.brands));
        brandRender();
        // Считаем конфликтные записи — они будут подсвечены в редакторе брендов
        const conflicted = Object.entries(_brandDB).filter(([k, v]) => {
          const c = brandCheckConflicts(k, v.synonyms || [], v.antonyms || [], k);
          return c.conflicts.length > 0;
        });
        if (conflicted.length > 0) {
          showToast(`⚠️ ${conflicted.length} бренд(ов) с конфликтами — откройте «База синонимов → Бренды» для исправления`, 'warn');
        }
      }
      // Загружаем настройки колонок если они есть в JSON
      if (json.columnSettings) {
        AppBridge.emit('settingsLoaded', json);
      }
      _jeDupsCache = null; jeChanges = 0;
      // ① Rebuild alias map from jeDB (fixes "Групп: 2" on monitoring tab)
      rebuildBarcodeAliasFromJeDB();
      // ② Update editor table + status badge (fixes "Штрихкоды 0" on synonyms tab)
      jeUpdateStatus(); jeRenderEditor();
      unifiedMarkUnsaved(false);
      // ③ Re-process price data so grouping reflects new synonyms
      if (allFilesData.length > 0) { processData(); renderTable(); updateUI(); }
      const bcCount = Object.keys(jeDB).length;
      const brCount = Object.keys(_brandDB).length;
      const hasColSettings = !!json.columnSettings;
      showToast(`Загружено: ${bcCount} ШК-групп, ${brCount} брендов${hasColSettings ? ', настройки колонок' : ''}`, 'ok');
    } catch(err) { showToast('Ошибка JSON: ' + err.message, 'warn'); }
  };
  reader.readAsText(file, 'utf-8');
}, true);

// Keyboard shortcuts
document.addEventListener('keydown', function(e) {
  const inInput = e.target.matches('input,textarea,select');
  if ((e.ctrlKey||e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 'z' && !inInput) {
    e.preventDefault(); jeUndo();
  }
  if ((e.ctrlKey||e.metaKey) && (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z')) && !inInput) {
    e.preventDefault(); jeRedo();
  }
  if (e.key === 'Escape') {
    closeMatchModal();
    jeConfirmClose(false);
    jeXlsModalClose(null);
    brandCloseEdit();
  }
});

// ── Init jeDB
jeUpdateStatus();
jeRenderEditor();

// ════════════════════════════════════════════════════════════════════════════
// UNIFIED SAVE: combined JSON { barcodes: jeDB, brands: _brandDB }
// ════════════════════════════════════════════════════════════════════════════
let _unifiedUnsaved = false;

function unifiedMarkUnsaved(dirty = true) {
  _unifiedUnsaved = dirty;
  const badge = document.getElementById('unifiedUnsaved');
  if (badge) badge.style.display = dirty ? '' : 'none';
}

// jeDBNotifyChange already calls rebuildBarcodeAliasFromJeDB and updateMatchPairTags

// JSON загружается через synonymsInput (карточка «Синонимы» на вкладке Мониторинг)

// ── SAVE unified JSON (includes barcodes, brands, and OBR column settings)
document.getElementById('unifiedSaveJsonBtn').addEventListener('click', async function() {
  const combined = {
    barcodes: jeDB,
    brands: _brandDB,
    columnSettings: (typeof columnTemplates !== 'undefined' && typeof columnSynonyms !== 'undefined') ? {
      templates: columnTemplates,
      synonyms: columnSynonyms
    } : undefined
  };
  const blob = new Blob([JSON.stringify(combined, null, 2)], { type: 'application/json' });
  const _sjFname = `settings_${new Date().toISOString().slice(0,10)}.json`;
  await saveBlobWithDialogOrDownload(blob, _sjFname);
  unifiedMarkUnsaved(false);
  showToast(`JSON сохранён: ${Object.keys(jeDB).length} ШК + ${Object.keys(_brandDB).length} брендов + настройки колонок`, 'ok');
});

// ════════════════════════════════════════════════════════════════════════════
// BRAND DICTIONARY
// { "tchibo": { synonyms: ["чибо","тибио"], antonyms: ["нескафе"] }, ... }
// ════════════════════════════════════════════════════════════════════════════
let _brandDB = (() => {
  if (typeof BRAND_CONFIG_SAVED === 'undefined') return {};
  // Support both old (just brands obj) and new combined format
  const s = BRAND_CONFIG_SAVED;
  return (s && typeof s === 'object' && !Array.isArray(s) && ('barcodes' in s || 'brands' in s))
    ? JSON.parse(JSON.stringify(s.brands || {}))
    : JSON.parse(JSON.stringify(s));
})();

// Also load jeDB from combined config if available
if (typeof BRAND_CONFIG_SAVED !== 'undefined' && BRAND_CONFIG_SAVED && BRAND_CONFIG_SAVED.barcodes) {
  jeDB = JSON.parse(JSON.stringify(BRAND_CONFIG_SAVED.barcodes));
  _jeDupsCache = null;
  jeUpdateStatus();
  jeRenderEditor();
  rebuildBarcodeAliasFromJeDB();
}

function brandParseCsv(raw) {
  return (raw || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
}
function brandNormKey(s) {
  return String(s || '').trim().toLowerCase();
}

function brandMarkUnsaved() {
  unifiedMarkUnsaved(true);
  const badge = document.getElementById('brandCountBadge');
  if (badge) badge.textContent = Object.keys(_brandDB).length;
}

function brandRender() {
  const q = (document.getElementById('brandSearchInp').value || '').toLowerCase().trim();
  const keys = Object.keys(_brandDB).sort((a, b) => a.localeCompare(b, 'ru'));
  const filtered = q ? keys.filter(k => {
    const v = _brandDB[k];
    if (k.includes(q)) return true;
    if ((v.synonyms || []).some(s => s.includes(q))) return true;
    if ((v.antonyms || []).some(s => s.includes(q))) return true;
    return false;
  }) : keys;

  const list   = document.getElementById('brandList');
  const empty  = document.getElementById('brandEmpty');
  const badge  = document.getElementById('brandCountBadge');
  const tableWrap = document.getElementById('brandTableWrap');
  if (badge) badge.textContent = keys.length;

  if (!filtered.length) {
    list.innerHTML = '';
    if (tableWrap) tableWrap.style.display = 'none';
    empty.style.display = '';
    empty.innerHTML = q
      ? `<div style="font-size:28px;margin-bottom:8px;">🔍</div><div>По запросу «${q}» ничего не найдено</div>`
      : `<div style="font-size:32px;margin-bottom:8px;">🏷️</div><div>Словарь брендов пуст.<br>Добавьте бренд вручную.</div>`;
    return;
  }
  empty.style.display = 'none';
  if (tableWrap) tableWrap.style.display = '';
  list.innerHTML = filtered.map((k, idx) => {
    const v = _brandDB[k];
    const syns = (v.synonyms || []).join(', ') || '<span style="color:var(--text-muted)">—</span>';
    const anti = (v.antonyms || []).join(', ') || '<span style="color:var(--text-muted)">—</span>';
    const check = brandCheckConflicts(k, v.synonyms || [], v.antonyms || [], k);
    const hasConflict = check.conflicts.length > 0;
    const conflictTip = hasConflict
      ? `<span class="brand-conflict-badge" onclick="brandOpenEdit(decodeURIComponent('${encodeURIComponent(k)}'))" title="${check.conflicts.join('; ')}">⚠️ Конфликт</span>`
      : '';
    return `<tr class="${hasConflict ? 'brand-row--conflict' : ''}" data-key="${encodeURIComponent(k)}">
      <td style="text-align:center;color:var(--text-muted);font-size:10px;">${idx+1}</td>
      <td><div class="brand-canonical">${k}</div>${conflictTip}</td>
      <td><div class="brand-syns">${syns}</div></td>
      <td><div class="brand-antonyms">${anti}</div></td>
      <td style="text-align:center;white-space:nowrap;">
        <button class="btn btn-xs" onclick="brandOpenEdit(decodeURIComponent('${encodeURIComponent(k)}'))" title="Редактировать">✏️</button>
        <button class="btn btn-xs btn-danger" onclick="brandDelete(decodeURIComponent('${encodeURIComponent(k)}'))" title="Удалить" style="margin-left:3px;">🗑️</button>
      </td>
    </tr>`;
  }).join('');
}

function brandDelete(key) {
  if (!_brandDB[key]) return;
  jeConfirmDialog('Удалить бренд «' + key + '»?', '🗑 Удаление').then(function(ok) {
    if (!ok) return;
    delete _brandDB[key];
    brandRender(); brandMarkUnsaved();
    showToast('Бренд «' + key + '» удалён', 'ok');
  });
}

function brandOpenEdit(key) {
  const v = _brandDB[key] || {};
  document.getElementById('beEditKey').value = key;
  document.getElementById('beCanon').value = key;
  document.getElementById('beSyns').value = (v.synonyms || []).join(', ');
  document.getElementById('beAnti').value = (v.antonyms || []).join(', ');
  document.getElementById('brandEditModal').classList.add('open');
}
function brandCloseEdit() {
  document.getElementById('brandEditModal').classList.remove('open');
}
function brandSaveEdit() {
  const oldKey = document.getElementById('beEditKey').value;
  const newKey = brandNormKey(document.getElementById('beCanon').value);
  if (!newKey) { showToast('Укажите канонический бренд', 'warn'); return; }
  const syns = brandParseCsv(document.getElementById('beSyns').value);
  const anti = brandParseCsv(document.getElementById('beAnti').value);

  // Проверяем конфликты, исключая сам редактируемый ключ
  const check = brandCheckConflicts(newKey, syns, anti, oldKey);

  if (check.conflicts.length) {
    const msg = `<b>Обнаружены противоречия, исправьте перед сохранением:</b><br><br>${check.conflicts.map(c=>'• '+c).join('<br>')}`;
    jeConfirmDialog(msg, '⚠️ Противоречия в бренде').then(()=>{});
    return;
  }

  if (check.existingKey && check.existingKey !== oldKey) {
    const ex = _brandDB[check.existingKey];
    const mergedSyns = [...new Set([...(ex.synonyms||[]), ...syns])];
    const mergedAnti = [...new Set([...(ex.antonyms||[]), ...anti])];
    const warnHtml = check.warnings.length ? brandConflictHtml({ conflicts:[], warnings: check.warnings }) : '';
    const msg = [
      `Бренд <b>«${newKey}»</b> уже существует. ${warnHtml}`,
      `Объединить синонимы/антонимы с существующим?`
    ].join('<br>');
    jeConfirmDialog(msg, '🔀 Бренд существует').then(function(ok) {
      if (!ok) return;
      if (oldKey && oldKey !== newKey) delete _brandDB[oldKey];
      _brandDB[newKey] = { synonyms: mergedSyns, antonyms: mergedAnti };
      brandCloseEdit(); brandRender(); brandMarkUnsaved();
      showToast(`Бренд «${newKey}» объединён и сохранён`, 'ok');
    });
    return;
  }

  // Предупреждения (без ошибок) — показываем тост, но сохраняем
  if (check.warnings.length) {
    showToast(`⚠ ${check.warnings[0]}`, 'warn');
  }

  if (oldKey && oldKey !== newKey) delete _brandDB[oldKey];
  _brandDB[newKey] = { synonyms: syns, antonyms: anti };
  brandCloseEdit(); brandRender(); brandMarkUnsaved();
  showToast(`Бренд «${newKey}» сохранён`, 'ok');
}

document.getElementById('brandSearchInp').addEventListener('input', brandRender);

// Click any data row in brand table to open edit
document.getElementById('brandList').addEventListener('click', function(e) {
  // Ignore clicks on action buttons
  if (e.target.closest('button')) return;
  const tr = e.target.closest('tr[data-key]');
  if (!tr) return;
  const key = decodeURIComponent(tr.dataset.key);
  brandOpenEdit(key);
});

document.getElementById('brandClearAllBtn').addEventListener('click', function () {
  if (!Object.keys(_brandDB).length) return;
  if (!confirm('Очистить весь словарь брендов?')) return;
  _brandDB = {};
  brandRender(); brandMarkUnsaved();
  showToast('Словарь брендов очищен', 'ok');
});

// ── Bulk text import


// ════════════════════════════════════════════════════════════════════════════
// BRAND ADD MODAL
// ════════════════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════════════════
// BRAND CONFLICT CHECKER
// Проверяет противоречия и дубли перед сохранением бренда.
// skipKey — ключ редактируемого бренда (чтобы не сравнивать с самим собой).
// Возвращает объект:
//   { existingKey, conflicts: string[], warnings: string[] }
// existingKey — если canonical уже есть в _brandDB (предложить объединить)
// conflicts   — критические противоречия (синонимы vs антонимы)
// warnings    — мягкие предупреждения (пересечения с другими брендами)
// ════════════════════════════════════════════════════════════════════════════
function brandCheckConflicts(newCanon, newSyns, newAnti, skipKey) {
  const result = { existingKey: null, conflicts: [], warnings: [] };
  if (!newCanon) return result;

  const newSynSet  = new Set(newSyns.map(s => brandNormKey(s)));
  const newAntiSet = new Set(newAnti.map(s => brandNormKey(s)));

  // 1. Внутреннее противоречие: слово одновременно в синонимах и антонимах
  const innerConflict = [...newSynSet].filter(s => newAntiSet.has(s));
  if (innerConflict.length) {
    result.conflicts.push(`Слова одновременно в синонимах и антонимах: «${innerConflict.join('», «')}»`);
  }

  // 2. Canonical сам является антонимом — добавлен в антонимы
  if (newAntiSet.has(newCanon)) {
    result.conflicts.push(`Канонический бренд «${newCanon}» указан в собственных антонимах`);
  }

  // 3. Проверка по всему _brandDB
  for (const [key, val] of Object.entries(_brandDB)) {
    if (key === skipKey) continue; // пропускаем себя при редактировании
    const exSynSet  = new Set(val.synonyms  || []);
    const exAntiSet = new Set(val.antonyms  || []);

    // 3a. Canonical уже существует — предложить объединить
    if (key === newCanon) {
      result.existingKey = key;
      // дополнительно проверяем противоречия со старыми антонимами
      const antiConflicts = [...newSynSet].filter(s => exAntiSet.has(s));
      if (antiConflicts.length) {
        result.conflicts.push(`В бренде «${key}» слова «${antiConflicts.join('», «')}» уже в антонимах, а вы добавляете их как синонимы`);
      }
      const synConflicts = [...newAntiSet].filter(s => exSynSet.has(s));
      if (synConflicts.length) {
        result.conflicts.push(`В бренде «${key}» слова «${synConflicts.join('», «')}» уже в синонимах, а вы добавляете их как антонимы`);
      }
      continue;
    }

    // 3b. Canonical совпадает с синонимом другого бренда
    if (exSynSet.has(newCanon)) {
      result.warnings.push(`«${newCanon}» уже является синонимом бренда «${key}»`);
    }

    // 3c. Один из новых синонимов — это канонический другого бренда
    for (const s of newSynSet) {
      if (s === key) {
        result.warnings.push(`Синоним «${s}» уже является каноническим брендом`);
      }
      // 3d. Синоним уже принадлежит другому бренду
      if (exSynSet.has(s)) {
        result.warnings.push(`Синоним «${s}» уже есть у бренда «${key}»`);
      }
      // 3e. Новый синоним находится в антонимах другого бренда (противоречие)
      if (exAntiSet.has(s)) {
        result.conflicts.push(`Синоним «${s}» находится в антонимах бренда «${key}»`);
      }
    }

    // 3f. Новый антоним — это синоним другого бренда (перекрёстное противоречие, только предупреждение)
    for (const a of newAntiSet) {
      if (exSynSet.has(a) && key !== newCanon) {
        result.warnings.push(`Антоним «${a}» является синонимом бренда «${key}»`);
      }
    }
  }

  // Дедупликация
  result.conflicts = [...new Set(result.conflicts)];
  result.warnings  = [...new Set(result.warnings)];
  return result;
}

// Строит HTML-блок предупреждений для отображения пользователю
function brandConflictHtml(check) {
  let html = '';
  if (check.conflicts.length) {
    html += `<div style="background:#fff3cd;border:1px solid #ffc107;border-radius:4px;padding:8px 10px;margin-bottom:6px;font-size:12px;">
      <b>⚠️ Противоречия (требуют исправления):</b><br>
      ${check.conflicts.map(c => `• ${c}`).join('<br>')}
    </div>`;
  }
  if (check.warnings.length) {
    html += `<div style="background:#e8f4ff;border:1px solid #90c8f0;border-radius:4px;padding:8px 10px;font-size:12px;">
      <b>ℹ️ Предупреждения:</b><br>
      ${check.warnings.map(w => `• ${w}`).join('<br>')}
    </div>`;
  }
  return html;
}

function brandOpenAddModal() {
  ['brNCanon','brNSyns','brNAnti'].forEach(id => { const el=document.getElementById(id); if(el) el.value=''; });
  const errEl = document.getElementById('brandFormError');
  if (errEl) { errEl.style.display='none'; errEl.textContent=''; }
  document.getElementById('brandAddModal').style.display = 'flex';
  setTimeout(() => { const el=document.getElementById('brNCanon'); if(el) el.focus(); }, 50);
}
function brandCloseAddModal() {
  document.getElementById('brandAddModal').style.display = 'none';
}

// FIX: brand modal elements are declared AFTER this script block in HTML,
// so we defer their listener registration to after DOM is fully parsed.
document.addEventListener('DOMContentLoaded', function() {
  const elMap = {
    'brandOpenAddModalBtn': () => brandOpenAddModal(),
    'brandAddModalCloseX':  () => brandCloseAddModal(),
    'brandAddModalCancel':  () => brandCloseAddModal(),
  };
  Object.entries(elMap).forEach(([id, fn]) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('click', fn);
  });

  document.getElementById('brandAddBtn') && document.getElementById('brandAddBtn').addEventListener('click', function () {
    const canon = brandNormKey(document.getElementById('brNCanon').value);
    const errEl = document.getElementById('brandFormError');
    errEl.style.display = 'none'; errEl.innerHTML = '';
    if (!canon) { errEl.textContent = 'Укажите канонический бренд'; errEl.style.display = ''; return; }

    const syns = brandParseCsv(document.getElementById('brNSyns').value);
    const anti = brandParseCsv(document.getElementById('brNAnti').value);

    // ── Проверяем конфликты и дубли
    const check = brandCheckConflicts(canon, syns, anti, null);

    // Критические противоречия — блокируем сохранение, показываем ошибку
    if (check.conflicts.length) {
      errEl.innerHTML = brandConflictHtml(check);
      errEl.style.display = '';
      return;
    }

    // Дубль — бренд уже существует: показываем диалог с выбором
    if (check.existingKey) {
      const ex = _brandDB[check.existingKey];
      const existSyns  = (ex.synonyms  || []).join(', ') || '—';
      const existAnti  = (ex.antonyms  || []).join(', ') || '—';
      const mergedSyns = [...new Set([...(ex.synonyms||[]), ...syns])];
      const mergedAnti = [...new Set([...(ex.antonyms||[]), ...anti])];

      // Предупреждения (не критические) показываем
      const warnHtml = check.warnings.length ? brandConflictHtml({ conflicts:[], warnings: check.warnings }) : '';

      const msg = [
        `Бренд <b>«${canon}»</b> уже существует в базе:`,
        `<div style="margin:6px 0;font-size:11px;background:#f5f5f5;border-radius:4px;padding:6px 8px;">`,
        `  <b>Синонимы:</b> ${existSyns}<br>`,
        `  <b>Антонимы:</b> ${existAnti}`,
        `</div>`,
        `После объединения:<br>`,
        `<div style="margin:4px 0;font-size:11px;background:#e8f7ef;border-radius:4px;padding:6px 8px;">`,
        `  <b>Синонимы:</b> ${mergedSyns.join(', ') || '—'}<br>`,
        `  <b>Антонимы:</b> ${mergedAnti.join(', ') || '—'}`,
        `</div>`,
        warnHtml,
        `Добавить в существующий бренд?`,
      ].join('');

      jeConfirmDialog(msg, '🔀 Бренд уже существует').then(function(ok) {
        if (!ok) return;
        _brandDB[canon] = { synonyms: mergedSyns, antonyms: mergedAnti };
        ['brNCanon','brNSyns','brNAnti'].forEach(id => document.getElementById(id).value = '');
        brandCloseAddModal();
        brandRender(); brandMarkUnsaved();
        showToast(`Бренд «${canon}» обновлён (объединено)`, 'ok');
      });
      return;
    }

    // Предупреждения — обязательное подтверждение через диалог, сохранение только по «Да»
    function _doSaveBrand() {
      _brandDB[canon] = { synonyms: syns, antonyms: anti };
      ['brNCanon','brNSyns','brNAnti'].forEach(id => document.getElementById(id).value = '');
      errEl.style.display = 'none'; errEl.innerHTML = '';
      brandCloseAddModal();
      brandRender(); brandMarkUnsaved();
      showToast(`Бренд «${canon}» добавлен`, 'ok');
    }

    if (check.warnings.length) {
      const msg = brandConflictHtml({ conflicts: [], warnings: check.warnings })
        + `<div style="margin-top:8px;">Всё равно создать новый бренд?</div>`;
      jeConfirmDialog(msg, '⚠️ Возможные пересечения').then(function(ok) {
        if (ok) _doSaveBrand();
      });
      return;
    }

    _doSaveBrand();
  });

  document.getElementById('brandClearFormBtn') && document.getElementById('brandClearFormBtn').addEventListener('click', function () {
    ['brNCanon','brNSyns','brNAnti'].forEach(id => document.getElementById(id).value = '');
    document.getElementById('brandFormError').style.display = 'none';
  });

  // bcAddModal listeners
  const bcModal = document.getElementById('bcAddModal');
  if (bcModal) bcModal.addEventListener('click', e => { if (e.target === bcModal) closeBcAddModal(); });
  const bcCloseX = document.getElementById('bcAddCloseX');
  if (bcCloseX) bcCloseX.addEventListener('click', closeBcAddModal);
  const bcCancel = document.getElementById('bcAddCancelBtn');
  if (bcCancel) bcCancel.addEventListener('click', closeBcAddModal);
  const bcSave = document.getElementById('bcAddSaveBtn');
  if (bcSave) bcSave.addEventListener('click', saveBcAddModal);
});

// ════════════════════════════════════════════════════════════════════════════
// BARCODE → DB QUICK ADD (from monitoring table)
// ════════════════════════════════════════════════════════════════════════════

function openAddToDB(barcode, btnEl) {
  // Find the item in groupedData
  const item = groupedData.find(d => String(d.barcode) === String(barcode));
  if (!item) return;

  // Already in DB? (check both main key and synonym)
  const _bcAlreadyInDB = jeDB[barcode] !== undefined
    || (typeof barcodeAliasMap !== 'undefined' && barcodeAliasMap.has(String(barcode)));
  if (_bcAlreadyInDB) { showToast('Штрихкод уже есть в базе', 'info'); return; }

  // Collect synonyms: other barcodes for same item from other files
  const synonymOptions = [];
  if (item.originalBarcodesByFile) {
    item.originalBarcodesByFile.forEach((bc, fileName) => {
      if (String(bc) !== String(barcode)) {
        synonymOptions.push({ bc: String(bc), fileName });
      }
    });
  }

  // Best name
  const bestName = (item.namesByFile && item.namesByFile.get(MY_PRICE_FILE_NAME))
    || (item.names && item.names[0] && item.names[0].name) || '';

  _bcAddState = { mainBC: barcode, synonyms: synonymOptions };

  document.getElementById('bcAddMainBC').value = barcode;
  document.getElementById('bcAddName').value = bestName;

  // Build synonym rows
  const list = document.getElementById('bcAddSynList');
  if (synonymOptions.length === 0) {
    list.innerHTML = '<div style="color:#999;font-size:11px;font-style:italic;">Нет синонимов от поставщиков</div>';
  } else {
    list.innerHTML = synonymOptions.map((s, i) => `
      <label class="bc-modal-syn-row">
        <input type="checkbox" data-syi="${i}" checked>
        <span class="bc-modal-syn-bc">${s.bc}</span>
        <span class="bc-modal-syn-file">${s.fileName}</span>
      </label>`).join('');
  }

  document.getElementById('bcAddModal').style.display = 'flex';
  setTimeout(() => document.getElementById('bcAddName').focus(), 50);
}

function closeBcAddModal() {
  document.getElementById('bcAddModal').style.display = 'none';
  _bcAddState = null;
}

function saveBcAddModal() {
  if (!_bcAddState) return;
  const mainBC = document.getElementById('bcAddMainBC').value.trim();
  if (!mainBC) { showToast('Штрихкод не может быть пустым', 'warn'); return; }
  if (jeDB[mainBC] !== undefined) { showToast('Штрихкод уже есть в базе', 'warn'); return; }
  const name = document.getElementById('bcAddName').value.trim() || mainBC;

  // Collect checked synonyms
  const checkedSyns = [];
  document.querySelectorAll('#bcAddSynList input[type=checkbox]:checked').forEach(cb => {
    const i = parseInt(cb.dataset.syi, 10);
    const s = _bcAddState.synonyms[i];
    if (s) checkedSyns.push(s.bc);
  });

  jeDBSaveHistory();
  jeDB[mainBC] = [name, ...checkedSyns];
  jeDBNotifyChange();
  jeRenderEditor(true); // preserve scroll: don't jump to top
  unifiedMarkUnsaved();
  closeBcAddModal();
  showToast(`Группа «${mainBC}» добавлена в базу`, 'ok');
  setTimeout(() => jeScrollToKey(mainBC), 60); // scroll to new entry with highlight
  // Re-render table to update badge
  if (typeof _mvsRenderVisible === 'function') _mvsRenderVisible();
}


// ── Expose globals
window.closeMatchModal = closeMatchModal;
window.confirmMatchAction = confirmMatchAction;
window.jeConfirmClose = jeConfirmClose;
window.jeXlsModalClose = jeXlsModalClose;
window.brandOpenEdit = brandOpenEdit;
window.brandCloseEdit = brandCloseEdit;
window.brandSaveEdit = brandSaveEdit;
window.brandDelete = brandDelete;
window.brandOpenAddModal = brandOpenAddModal;
window.brandCloseAddModal = brandCloseAddModal;
window.openAddToDB = openAddToDB;
window.closeBcAddModal = closeBcAddModal;

// ── Init brand editor
brandRender();

// ════════════════════════════════════════════════════════════════════════════
// MATCHER: JSON file info display + enable/disable toggle
// ════════════════════════════════════════════════════════════════════════════
window._matcherUpdateJsonInfo = function() {
  const jsonRow = document.getElementById('matcherJsonRow');
  const jsonLabel = document.getElementById('matcherJsonLabel');
  if (!jsonRow || !jsonLabel) return;
  const hasSynonyms = typeof jeDB !== 'undefined' && Object.keys(jeDB).length > 0;
  const sfName = document.getElementById('sfJsonName');
  const fileName = sfName && sfName.textContent !== 'JSON не загружен' ? sfName.textContent : null;
  if (fileName || hasSynonyms) {
    jsonRow.style.display = '';
    const count = Object.keys(jeDB || {}).length;
    jsonLabel.innerHTML = `<strong>${fileName || 'JSON'}</strong> — ${count} записей`;
  } else {
    jsonRow.style.display = 'none';
  }
};

// Hook JSON toggle checkbox to update row style
(function() {
  const chk = document.getElementById('matcherJsonEnabled');
  if (chk) {
    chk.addEventListener('change', function() {
      const row = document.getElementById('matcherJsonToggleRow');
      if (row) row.classList.toggle('disabled', !chk.checked);
    });
  }
})();

// ════════════════════════════════════════════════════════════════════════════
// BEFOREUNLOAD: warn if JSON has unsaved changes
// ════════════════════════════════════════════════════════════════════════════
window.addEventListener('beforeunload', function(e) {
  if (typeof jeChanges !== 'undefined' && jeChanges > 0) {
    const msg = 'В базе синонимов есть несохранённые изменения. Они будут потеряны при закрытии вкладки.';
    e.preventDefault();
    e.returnValue = msg;
    return msg;
  }
});

// ════════════════════════════════════════════════════════════════════════════
// EXTENDED JSON LOAD: also parse columnSettings when synonymsInput loads
// ════════════════════════════════════════════════════════════════════════════
// Monkey-patch the synonymsInput handler by wrapping via AppBridge
(function() {
  const origHandler = document.getElementById('synonymsInput');
  // Add a supplemental listener that fires AFTER the main one to sync columnSettings
  document.getElementById('synonymsInput').addEventListener('change', function(e) {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = function(ev) {
      try {
        const json = JSON.parse(ev.target.result);
        if (json && json.columnSettings) {
          AppBridge.emit('settingsLoaded', json);
        }
      } catch(err) {}
    };
    reader.readAsText(file, 'utf-8');
  }); // no capture = fires after main handler
})();


// ════════════════════════════════════════════════════
// SIDEBAR COLLAPSE
// ════════════════════════════════════════════════════
function toggleSidebar() {
  const sidebar = document.querySelector('.app-sidebar');
  const collapsed = sidebar.classList.toggle('collapsed');
  localStorage.setItem('sidebarCollapsed', collapsed ? '1' : '0');
  document.getElementById('sidebarToggle').title = collapsed ? 'Развернуть меню' : 'Свернуть меню';
}

// Restore on load
(function() {
  if (localStorage.getItem('sidebarCollapsed') === '1') {
    const sidebar = document.querySelector('.app-sidebar');
    if (sidebar) {
      sidebar.classList.add('collapsed');
      const btn = document.getElementById('sidebarToggle');
      if (btn) btn.title = 'Развернуть меню';
    }
  }
})();

// ════════════════════════════════════════════════════
// SIDEBAR FILE INFO PANEL
// ════════════════════════════════════════════════════
(function() {
  function sfShorten(name, max) {
    if (!name) return '';
    return name.length > max ? name.slice(0, max - 1) + '…' : name;
  }

  function sfUpdateJson(fileName, entryCount) {
    const item = document.getElementById('sfJsonItem');
    const nameEl = document.getElementById('sfJsonName');
    const badge = document.getElementById('sfJsonBadge');
    const meta = document.getElementById('sfJsonMeta');
    if (!item) return;
    if (fileName) {
      item.classList.remove('sidebar-file-item--empty');
      item.classList.add('sidebar-file-item--loaded');
      nameEl.textContent = sfShorten(fileName, 22);
      badge.style.display = '';
      if (entryCount != null) {
        meta.style.display = '';
        meta.innerHTML = '<strong>' + entryCount + '</strong> записей в базе';
      }
    } else {
      item.classList.add('sidebar-file-item--empty');
      item.classList.remove('sidebar-file-item--loaded');
      nameEl.textContent = 'JSON не загружен';
      badge.style.display = 'none';
      meta.style.display = 'none';
    }
  }

  function sfUpdateMyPrice(fileName, rows) {
    const item = document.getElementById('sfMyPriceItem');
    const nameEl = document.getElementById('sfMyPriceName');
    const badge = document.getElementById('sfMyPriceBadge');
    const meta = document.getElementById('sfMyPriceMeta');
    if (!item) return;
    if (fileName) {
      item.classList.remove('sidebar-file-item--empty');
      item.classList.add('sidebar-file-item--myprice');
      nameEl.textContent = sfShorten(fileName, 22);
      badge.style.display = '';
      if (rows != null) {
        meta.style.display = '';
        meta.innerHTML = '<strong>' + rows.toLocaleString('ru') + '</strong> строк';
      }
    } else {
      item.classList.add('sidebar-file-item--empty');
      item.classList.remove('sidebar-file-item--myprice');
      nameEl.textContent = 'Мой прайс не загружен';
      badge.style.display = 'none';
      meta.style.display = 'none';
    }
  }

  function sfUpdateSuppliers(list) {
    // list = [{name, rows}]
    const emptyEl = document.getElementById('sfSuppliersEmpty');
    const listEl = document.getElementById('sfSuppliersList');
    if (!listEl) return;
    if (!list || list.length === 0) {
      emptyEl.style.display = '';
      listEl.style.display = 'none';
      listEl.innerHTML = '';
      return;
    }
    emptyEl.style.display = 'none';
    listEl.style.display = '';
    listEl.innerHTML = list.map(f => `
      <div class="sidebar-file-item sidebar-file-item--supplier">
        <div class="sf-row sf-supplier-row">
          <span class="sf-icon">📦</span>
          <span class="sf-name" title="${f.name.replace(/"/g,'&quot;')}">${f.name.length > 18 ? f.name.slice(0,17) + '…' : f.name}</span>
          <span class="sf-type sf-type--supplier">CSV</span>
          <button class="sf-supplier-del" title="Удалить файл поставщика" onclick="removeSupplierFile('${f.name.replace(/'/g,"\\'")}')">✕</button>
        </div>
      </div>`).join('');
  }

  // Update monitor slot supplier file list
  function _monitorUpdateSupplierList(list) {
    const el = document.getElementById('monitorSupplierFileList');
    if (!el) return;
    if (!list || list.length === 0) { el.style.display = 'none'; el.innerHTML = ''; return; }
    el.style.display = '';
    el.innerHTML = list.map(f => `
      <div class="sup-file-row">
        <span class="sup-file-row-name" title="${f.name.replace(/"/g,'&quot;')}">📦 ${f.name.replace(/</g,'&lt;').replace(/>/g,'&gt;')}</span>
        <button class="sup-file-row-del" title="Удалить файл поставщика" onclick="removeSupplierFile('${f.name.replace(/'/g,"\\'")}')">✕</button>
      </div>`).join('');
  }

  // Expose for use by other code
  window._sfUpdateJson = sfUpdateJson;
  window._sfUpdateMyPrice = sfUpdateMyPrice;
  window._sfUpdateSuppliers = function(list) {
    sfUpdateSuppliers(list);
    _monitorUpdateSupplierList(list);
  };

  // Watch for status changes via MutationObserver to sync sidebar
  function watchStatus(id, cb) {
    const el = document.getElementById(id);
    if (!el) return;
    const obs = new MutationObserver(() => cb(el.textContent));
    obs.observe(el, { childList: true, subtree: true, characterData: true });
  }

  // Sync JSON status
  watchStatus('synonymsStatus', function(txt) {
    if (txt && txt !== 'Не загружены' && txt !== 'Не загружена') {
      // Extract file name from status text
      const match = txt.match(/(.+?)\s*[\(—]/);
      sfUpdateJson(match ? match[1].trim() : txt, null);
    } else {
      sfUpdateJson(null, null);
    }
  });

  // Sync my price status
  watchStatus('myPriceStatus', function(txt) {
    if (txt && txt !== 'Не загружен') {
      const rowMatch = txt.match(/(\d[\d\s]*)\s*строк/);
      const rows = rowMatch ? parseInt(rowMatch[1].replace(/\s/g, '')) : null;
      const nameMatch = txt.match(/^(.+?)\s*[\(—\|]/);
      sfUpdateMyPrice(nameMatch ? nameMatch[1].trim() : txt, rows);
    } else {
      sfUpdateMyPrice(null, null);
    }
  });

  // Note: suppliers sidebar is updated directly via window._sfUpdateSuppliers() on load/remove

  // Also hook JSON DB changes for entry count
  if (typeof AppBridge !== 'undefined') {
    AppBridge.on('settingsLoaded', function(json) {
      const count = json && json.synonyms ? Object.keys(json.synonyms).length : null;
      const sfName = document.getElementById('sfJsonName');
      if (sfName && sfName.textContent !== 'JSON не загружен') {
        const meta = document.getElementById('sfJsonMeta');
        if (meta && count != null) {
          meta.style.display = '';
          meta.innerHTML = '<strong>' + count + '</strong> записей в базе';
        }
      }
      // Update matcher JSON panel
      setTimeout(function() {
        if (typeof window._matcherUpdateJsonInfo === 'function') window._matcherUpdateJsonInfo();
      }, 200);
    });
  }

  // Update matcher JSON info when synonyms status changes
  watchStatus('synonymsStatus', function() {
    setTimeout(function() {
      if (typeof window._matcherUpdateJsonInfo === 'function') window._matcherUpdateJsonInfo();
    }, 300);
  });
})();

