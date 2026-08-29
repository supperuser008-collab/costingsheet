// Global Price Catalog loaded from uploaded Excel
let priceCatalog = [
    // Default catalog seed (can be overridden by uploading a Price List Excel)
    { brand: 'UNV', model: 'IPC3232SB- AHDZK-PI-I0', desc: 'Dome Camera - 2MP WDR LightHunter IR Network Dome Camera', unitCost: 371 },
    { brand: 'UNV', model: 'IPC2322LB- ADZK-G', desc: 'Bullet Camera - 2MP HD IR VF Bullet Network Camera', unitCost: 422 },
    { brand: 'Hikvision', model: 'DS-2CD2143G0-I', desc: 'Dome Camera - 4MP EXIR Fixed Dome Network Camera', unitCost: 390 },
    { brand: 'Hikvision', model: 'DS-2CD2043G0-I', desc: 'Bullet Camera - 4MP EXIR Fixed Bullet Network Camera', unitCost: 440 },
    { brand: 'UNV', model: 'NVR508-32B', desc: '32 Ch Network Video Recorder', unitCost: 1650 },
    { brand: 'Seagate', model: 'ST16000NM002H', desc: '16TB HDD Surveillance Grade', unitCost: 2650 },
    { brand: 'Hikvision', model: 'DS-3E0524P-E', desc: '24 Port Gigabit POE Switch', unitCost: 950 },
    { brand: 'Hikvision', model: 'DS-3E0109P-E/M', desc: '8 Port POE switch with power supply', unitCost: 650 },
    { brand: 'Netcon', model: 'NET-OD-ENC', desc: 'Outdoor IP66 Enclosure', unitCost: 750 },
    { brand: 'Dell', model: 'Pro Tower QCT 1250', desc: 'CCTV Workstation with K/M', unitCost: 3000 },
    { brand: 'Norden', model: '1311-186610BK', desc: '18U Floor Mount Rack (600X800)', unitCost: 1300 }
];

// BOQ Master Data
let boqData = [
    { cat: 'A', catTitle: 'Cameras and accessories', isCat: true },
    { id: 1, type: 'cam', subtype: 'dome', cat: 'A', desc: 'Dome Camera - 2MP WDR LightHunter IR Network Dome Camera', brand: 'UNV', model: 'IPC3232SB- AHDZK-PI-I0', uom: 'EA', qty: 0, cost: 371, isService: false },
    { id: 2, type: 'cam', subtype: 'bullet', cat: 'A', desc: 'Bullet Camera - 2MP HD IR VF Bullet Network Camera', brand: 'UNV', model: 'IPC2322LB- ADZK-G', uom: 'EA', qty: 0, cost: 422, isService: false },
    
    { cat: 'B', catTitle: 'VMS, NVR & Storage', isCat: true },
    { id: 3, type: 'item', cat: 'B', desc: '32 Ch Network Video Recorder', brand: 'UNV', model: 'NVR508-32B', uom: 'EA', qty: 1, cost: 1650, isService: false },
    { id: 4, type: 'item', cat: 'B', desc: '16TB HDD', brand: 'Seagate', model: 'ST16000NM002H', uom: 'EA', qty: 7, cost: 2650, isService: false },
    
    { cat: 'C', catTitle: 'Network Switches and accessories', isCat: true },
    { id: 5, type: 'item', cat: 'C', desc: '24 Port POE Switch', brand: 'Hikvision', model: 'DS-3E0524P-E', uom: 'EA', qty: 1, cost: 950, isService: false },
    { id: 6, type: 'item', cat: 'C', desc: '8 Port POE switch with power supply', brand: 'Hikvision', model: 'DS-3E0109P-E/M', uom: 'EA', qty: 1, cost: 650, isService: false },
    { id: 7, type: 'item', cat: 'C', desc: 'Outdoor Enclosure', brand: 'Netcon', model: 'NET-OD-ENC', uom: 'EA', qty: 1, cost: 750, isService: false },
    
    { cat: 'D', catTitle: 'Workstation and Monitor', isCat: true },
    { id: 8, type: 'item', cat: 'D', desc: '24" LED FHD Monitor_MOI Approved', brand: 'UNV', model: 'UNV-MW3224-V', uom: 'EA', qty: 1, cost: 320, isService: false },
    { id: 9, type: 'item', cat: 'D', desc: '32" LED FHD Monitor_MOI Approved', brand: 'UNV', model: 'UNV-MW3232-V-K2', uom: 'EA', qty: 1, cost: 750, isService: false },
    { id: 10, type: 'item', cat: 'D', desc: 'CCTV Workstation with K/M', brand: 'Dell', model: 'Pro Tower QCT 1250', uom: 'EA', qty: 1, cost: 3000, isService: false },

    { cat: 'J', catTitle: 'ACS System', isCat: true },
    { id: 11, type: 'item', cat: 'J', desc: 'Standalone Access control system', brand: 'Hikvision', model: 'TBD', uom: 'EA', qty: 1, cost: 500, isService: false },

    { cat: 'E', catTitle: 'CCTV Racks', isCat: true },
    { id: 12, type: 'item', cat: 'E', desc: '18U Floor Mount Rack (600X800)', brand: 'Norden', model: '1311-186610BK', uom: 'EA', qty: 1, cost: 1300, isService: false },

    { cat: 'F', catTitle: 'UPS System', isCat: true },
    { id: 13, type: 'item', cat: 'F', desc: '2KVA UPS with Battery Pack for 60 Min. Backup', brand: 'Hikvision', model: 'TBD', uom: 'EA', qty: 1, cost: 4000, isService: false },

    { cat: 'G', catTitle: 'Cabling and accessories', isCat: true },
    { id: 14, type: 'auto_cable', cat: 'G', desc: 'Cabling and accessories (Calculated Lot)', brand: 'Flora', model: 'CALC-LOT', uom: 'LOT', qty: 1, cost: 0, isService: false },
    { id: 15, type: 'auto_conduit', cat: 'G', desc: 'Conduits and accessories (Estimation Lot)', brand: 'Flora', model: 'EST-LOT', uom: 'LOT', qty: 1, cost: 0, isService: false },

    { cat: 'H', catTitle: 'Engineering & Services', isCat: true },
    { id: 16, type: 'auto_labor', cat: 'H', desc: 'Installation, configuration, Testing & Commissioning', brand: 'Flora', model: 'SVC-LOT', uom: 'LOT', qty: 1, cost: 0, isService: true }
];

/**
 * Upload & Parse Excel Master Price List (.xlsx / .xls)
 * Expected Columns: Brand | Model | Description | Price (or Cost)
 */
function handlePriceListUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet);

        if (rows.length === 0) {
            alert("No rows found in the uploaded price list.");
            return;
        }

        priceCatalog = [];
        rows.forEach(row => {
            const brand = row['Brand'] || row['BRAND'] || row['brand'] || 'Generic';
            const model = row['Model'] || row['MODEL'] || row['Model No'] || row['model'] || 'TBD';
            const desc = row['Description'] || row['DESCRIPTION'] || row['desc'] || '';
            const unitCost = parseFloat(row['Price'] || row['PRICE'] || row['Cost'] || row['COST'] || row['Unit Price'] || 0);

            if (model !== 'TBD') {
                priceCatalog.push({ brand, model, desc, unitCost });
            }
        });

        alert(`Price List Uploaded Successfully! Loaded ${priceCatalog.length} catalog items.`);
        calculateBOQ();
        event.target.value = '';
    };
    reader.readAsArrayBuffer(file);
}

/**
 * Handle Word Survey Upload (.docx)
 */
function handleWordUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    if (typeof mammoth === 'undefined') {
        alert("Mammoth library is not loaded.");
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        mammoth.convertToHtml({ arrayBuffer: e.target.result })
            .then(function(result) {
                parseWordHtml(result.value);
                event.target.value = '';
            })
            .catch(function(err) {
                console.error("Error parsing Word file:", err);
                alert("Failed to parse Word document.");
            });
    };
    reader.readAsArrayBuffer(file);
}

/**
 * Parse Extracted Word HTML Document
 */
function parseWordHtml(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');

    let domeCount = 0;
    let bulletCount = 0;
    let conduitLength = 0;

    const rows = doc.querySelectorAll('tr');
    rows.forEach(row => {
        const cells = Array.from(row.querySelectorAll('td, th')).map(c => c.textContent.trim().toLowerCase());
        cells.forEach((cellText, idx) => {
            if (cellText.includes('dome')) {
                const num = findNumberInRow(cells, idx);
                if (num > 0) domeCount = num;
            }
            if (cellText.includes('bullet')) {
                const num = findNumberInRow(cells, idx);
                if (num > 0) bulletCount = num;
            }
            if (cellText.includes('conduit')) {
                const num = findNumberInRow(cells, idx);
                if (num > 0) conduitLength = num;
            }
        });
    });

    const fullText = doc.body.textContent || '';
    if (domeCount === 0) {
        const dMatch = fullText.match(/dome[^\d]*(\d+)/i) || fullText.match(/(\d+)[^\d]*dome/i);
        if (dMatch) domeCount = parseInt(dMatch[1], 10);
    }
    if (bulletCount === 0) {
        const bMatch = fullText.match(/bullet[^\d]*(\d+)/i) || fullText.match(/(\d+)[^\d]*bullet/i);
        if (bMatch) bulletCount = parseInt(bMatch[1], 10);
    }
    if (conduitLength === 0) {
        const cMatch = fullText.match(/conduit[^\d]*(\d+)/i) || fullText.match(/(\d+)[^\d]*conduit/i);
        if (cMatch) conduitLength = parseInt(cMatch[1], 10);
    }

    const domeItem = boqData.find(i => i.subtype === 'dome');
    if (domeItem) domeItem.qty = domeCount;

    const bulletItem = boqData.find(i => i.subtype === 'bullet');
    if (bulletItem) bulletItem.qty = bulletCount;

    if (conduitLength > 0) {
        document.getElementById('conduitMeters').value = conduitLength;
    }

    calculateBOQ();
    alert(`Survey Extracted:\n- Dome Cameras: ${domeCount}\n- Bullet Cameras: ${bulletCount}\n- Conduit Length: ${conduitLength} Mtr`);
}

function findNumberInRow(cells, keywordIdx) {
    for (let i = 0; i < cells.length; i++) {
        if (i === keywordIdx) continue;
        const val = parseInt(cells[i].replace(/[^\d]/g, ''), 10);
        if (!isNaN(val) && val > 0) return val;
    }
    return 0;
}

/**
 * Handle UI Item Change Handlers
 */
function updateBrand(id, newBrand) {
    const item = boqData.find(i => i.id === id);
    if (!item) return;

    item.brand = newBrand;
    const availableModels = priceCatalog.filter(c => c.brand.toLowerCase() === newBrand.toLowerCase());
    if (availableModels.length > 0) {
        item.model = availableModels[0].model;
        item.desc = availableModels[0].desc || item.desc;
        item.cost = availableModels[0].unitCost;
    }
    calculateBOQ();
}

function updateModel(id, newModel) {
    const item = boqData.find(i => i.id === id);
    if (!item) return;

    item.model = newModel;
    const match = priceCatalog.find(c => c.model.toLowerCase() === newModel.toLowerCase());
    if (match) {
        if (match.desc) item.desc = match.desc;
        item.cost = match.unitCost;
    }
    calculateBOQ();
}

function updateDesc(id, newDesc) {
    const item = boqData.find(i => i.id === id);
    if (item) item.desc = newDesc;
}

function updateQty(id, val) {
    const item = boqData.find(i => i.id === id);
    if (item) {
        item.qty = parseFloat(val) || 0;
        calculateBOQ();
    }
}

function updateCost(id, val) {
    const item = boqData.find(i => i.id === id);
    if (item) {
        item.cost = parseFloat(val) || 0;
        calculateBOQ();
    }
}

/**
 * Recalculate BOQ Totals & Render UI Table
 */
function calculateBOQ() {
    const matMarkup = parseFloat(document.getElementById('matMarkup').value) / 100 || 0;
    const srvMarkup = parseFloat(document.getElementById('srvMarkup').value) / 100 || 0;
    const conduitMeters = parseFloat(document.getElementById('conduitMeters').value) || 0;

    let totalCams = 0;
    boqData.forEach(item => {
        if (item.type === 'cam') totalCams += parseFloat(item.qty || 0);
    });

    const cat6Coils = Math.ceil((totalCams * 100) / 305); 
    const totalCableCost = (cat6Coils * 460) + (1 * 40) + (totalCams * 9) + (1 * 60) + (1 * 35) + (totalCams * 8) + (3 * 12) + (10 * 20) + 750 + 500 + 500;
    
    const pvcPipes = (conduitMeters / 2) / 90;
    const giPipes = (conduitMeters / 2) / 90;
    const totalConduitCost = (conduitMeters > 0) ? ((pvcPipes * 280) + (giPipes * 470) + 500 + 200) : 0;

    const totalLaborCost = (totalCams * 150) + (conduitMeters * 1) + (conduitMeters * 1) + (totalCams * 25) + 2500;

    boqData.forEach(item => {
        if (item.type === 'auto_cable') item.cost = totalCableCost;
        if (item.type === 'auto_conduit') item.cost = totalConduitCost;
        if (item.type === 'auto_labor') item.cost = totalLaborCost;
    });

    const tbody = document.getElementById('boqTableBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    let grandTotalCost = 0;
    let grandTotalSales = 0;

    // Unique Brands from Catalog
    const uniqueBrands = Array.from(new Set(priceCatalog.map(c => c.brand)));

    boqData.forEach(item => {
        if (item.isCat) {
            tbody.innerHTML += `
                <tr class="bg-slate-800/90 font-bold text-sky-300">
                    <td class="p-2 border-t border-slate-700">${item.cat}</td>
                    <td colspan="9" class="p-2 border-t border-slate-700">${item.catTitle}</td>
                </tr>`;
            return;
        }

        const markup = item.isService ? srvMarkup : matMarkup;
        const unitCost = parseFloat(item.cost || 0);
        const qty = parseFloat(item.qty || 0);
        const totCost = unitCost * qty;
        const unitSell = unitCost * (1 + markup);
        const totSales = totCost * (1 + markup);

        grandTotalCost += totCost;
        grandTotalSales += totSales;

        const isAuto = item.type.startsWith('auto_');

        // Brand Selector HTML
        let brandHtml = `<span class="text-slate-400">${item.brand}</span>`;
        if (!isAuto && uniqueBrands.length > 0) {
            let options = uniqueBrands.map(b => `<option value="${b}" ${b.toLowerCase() === item.brand.toLowerCase() ? 'selected' : ''}>${b}</option>`).join('');
            brandHtml = `<select onchange="updateBrand(${item.id}, this.value)" class="bg-slate-900 border border-slate-700 rounded px-1.5 py-1 text-xs text-slate-200 focus:border-sky-500 focus:outline-none w-full">${options}</select>`;
        }

        // Model Selector HTML
        let modelHtml = `<span class="text-slate-400 font-mono text-xs">${item.model}</span>`;
        if (!isAuto) {
            const filteredModels = priceCatalog.filter(c => c.brand.toLowerCase() === item.brand.toLowerCase());
            if (filteredModels.length > 0) {
                let options = filteredModels.map(m => `<option value="${m.model}" ${m.model.toLowerCase() === item.model.toLowerCase() ? 'selected' : ''}>${m.model}</option>`).join('');
                modelHtml = `<select onchange="updateModel(${item.id}, this.value)" class="bg-slate-900 border border-slate-700 rounded px-1.5 py-1 text-xs text-slate-200 focus:border-sky-500 focus:outline-none w-full">${options}</select>`;
            }
        }

        // Description Input HTML
        const descInputHtml = isAuto 
            ? `<span class="text-slate-200">${item.desc}</span>`
            : `<input type="text" value="${item.desc.replace(/"/g, '&quot;')}" onchange="updateDesc(${item.id}, this.value)" class="w-full bg-slate-900 border border-slate-700 text-slate-200 rounded px-2 py-1 text-sm focus:border-sky-500 focus:outline-none">`;

        // Quantity Input HTML
        const qtyInputHtml = isAuto 
            ? `<span class="text-slate-400 font-mono">${qty}</span>` 
            : `<input type="number" min="0" value="${qty}" onchange="updateQty(${item.id}, this.value)" class="w-16 bg-slate-900 border border-slate-700 text-right rounded px-1.5 py-1 text-sm focus:border-sky-500 focus:outline-none">`;

        // Cost Input HTML
        const costInputHtml = isAuto
            ? `<span class="font-mono text-slate-300">${unitCost.toFixed(2)}</span>`
            : `<input type="number" min="0" step="0.01" value="${unitCost}" onchange="updateCost(${item.id}, this.value)" class="w-20 bg-slate-900 border border-slate-700 text-right rounded px-1.5 py-1 text-sm font-mono text-slate-300 focus:border-sky-500 focus:outline-none">`;

        tbody.innerHTML += `
            <tr class="hover:bg-slate-700/30 transition-colors">
                <td class="p-3 text-center text-slate-400">${item.id}</td>
                <td class="p-3">${descInputHtml}</td>
                <td class="p-3 w-28">${brandHtml}</td>
                <td class="p-3 w-36">${modelHtml}</td>
                <td class="p-3 text-center text-slate-400">${item.uom}</td>
                <td class="p-3 text-right">${qtyInputHtml}</td>
                <td class="p-3 text-right">${costInputHtml}</td>
                <td class="p-3 text-right font-mono text-slate-300">${totCost.toFixed(2)}</td>
                <td class="p-3 text-right font-mono text-emerald-300">${unitSell.toFixed(2)}</td>
                <td class="p-3 text-right font-mono font-semibold text-emerald-400">${totSales.toFixed(2)}</td>
            </tr>`;
    });

    document.getElementById('summaryTotalCams').innerText = totalCams;
    document.getElementById('summaryTotalCost').innerText = `QAR ${grandTotalCost.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}`;
    document.getElementById('summaryTotalSales').innerText = `QAR ${grandTotalSales.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}`;
}

/**
 * Export Multi-Tab Excel Spreadsheet
 */
function exportToExcel() {
    const matMarkup = parseFloat(document.getElementById('matMarkup').value) / 100 || 0;
    const srvMarkup = parseFloat(document.getElementById('srvMarkup').value) / 100 || 0;
    const projSubject = document.getElementById('projSubject').value;
    const conduitMeters = parseFloat(document.getElementById('conduitMeters').value) || 0;

    const wb = XLSX.utils.book_new();

    const boqRows = [
        ["Costing Summary", "", "", "", "", "Total Cemera", "", "Total Sales (QAR)", "", "", matMarkup, "MATERIALS"],
        ["", "", "", "", "", "=QTY!I8", "", "=J49", "", "", srvMarkup, "SERVICES"],
        [],
        [`SUBJECT : ${projSubject}`],
        [],
        ["", "", "", "", "", "", "", "All-in-Cost Price (QAR)", "", "Sell Price (QAR)"],
        ["No", "Product Description", "Brand", "Model No", "UOM", "Qty", "Unit Price", "Total Price", "Unit Price", "Total Price"]
    ];

    let rowIdx = 8;
    boqData.forEach(item => {
        if (item.isCat) {
            boqRows.push([item.cat, item.catTitle]);
        } else {
            const mCell = item.isService ? "$K$3" : "$K$2";
            boqRows.push([
                item.id,
                item.desc,
                item.brand,
                item.model,
                item.uom,
                item.qty,
                item.cost,
                `=G${rowIdx}*H${rowIdx}`,
                `=H${rowIdx}*(1+${mCell})`,
                `=G${rowIdx}*J${rowIdx}`
            ]);
        }
        rowIdx++;
    });

    boqRows.push([], ["", "", "", "", "", "", "Total Cost", `=SUM(I8:I${rowIdx-1})`, "Total Sales", `=SUM(K8:K${rowIdx-1})`]);
    const wsBOQ = XLSX.utils.aoa_to_sheet(boqRows);

    const domeQty = boqData.find(i => i.subtype === 'dome')?.qty || 0;
    const bulletQty = boqData.find(i => i.subtype === 'bullet')?.qty || 0;
    const totalCams = domeQty + bulletQty;

    const breakdownRows = [
        [],
        ["BREAK DOWN DETAILS DO NOT ATTACHED WITH THE FINAL QUOTATION"],
        ["BREAK DOWN"],
        ["", "Cables and accessories", "Brand", "Model No", "UOM", "Qty", "Unit Price", "Total Price", "Remarks"],
        ["", "Cat6 Cable 305M", "TBD", "TBD", "EA", `=CEILING(${totalCams}*100/305, 1)`, 460, "=G5*H5"],
        ["", "24 Port Patch panel", "TBD", "TBD", "EA", 1, 40, "=G6*H6"],
        ["", "Rj45 Keystone", "TBD", "TBD", "EA", totalCams, 9, "=G7*H7"],
        ["", "Rj45 Connector - Packet (50 pcs)", "TBD", "TBD", "EA", 1, 60, "=G8*H8"],
        ["", "Cable Manager", "TBD", "TBD", "EA", 1, 35, "=G9*H9"],
        ["", "Patch cord - 1M", "TBD", "TBD", "EA", totalCams, 8, "=G10*H10"],
        ["", "Patch cord - 3M", "TBD", "TBD", "EA", 3, 12, "=G11*H11"],
        ["", "CCTV Stickers", "TBD", "TBD", "LOT", 10, 20, "=G12*H12"],
        ["", "Fiber Cables and accessories", "TBD", "TBD", "LOT", 1, 750, "=G13*H13"],
        ["", "Sundries and miscellaneous", "TBD", "TBD", "LOT", 1, 500, "=G14*H14"],
        ["", "", "", "", "", "", "", "=SUM(I5:I14)"],
        ["", "Electrical cables and accessories"],
        ["", "Electrical cables", "TBD", "TBD", "EA", 1, 500, "=G17*H17"],
        ["", "Sundries and miscellaneous", "TBD", "TBD", "EA", 1, 0, "=G18*H18"],
        ["", "", "", "", "", "", "", "=SUM(I17:I18)", "=I15+I19"],
        ["", "Conduits and accessories"],
        ["", "PVC Conduit and accessories", "TBD", "TBD", "EA", (conduitMeters / 2) / 90, 280, "=G21*H21"],
        ["", "GI Conduit and accessories", "TBD", "TBD", "EA", (conduitMeters / 2) / 90, 470, "=G22*H22"],
        ["", "Flexible conduit and accessories", "TBD", "TBD", "EA", conduitMeters > 0 ? 1 : 0, 500, "=G23*H23"],
        ["", "Sundries and miscellaneous", "TBD", "TBD", "LOT", conduitMeters > 0 ? 1 : 0, 200, "=G24*H24"],
        ["", "", "", "", "", "", "", "=SUM(I21:I24)", "=I25"],
        [],
        ["", "Engineering & Services"],
        ["", "Camera Installation", "TBD", "TBD", "EA", totalCams, 150, "=G28*H28"],
        ["", "Conduit Installation", "TBD", "TBD", "MTR", conduitMeters, 1, "=G29*H29"],
        ["", "Cable pulling and installation", "TBD", "TBD", "MTR", conduitMeters, 1, "=G30*H30"],
        ["", "Cabling Termination, Testing and Labeling", "TBD", "TBD", "EA", totalCams, 25, "=G31*H31"],
        ["", "Head end Installation & Handover", "TBD", "TBD", "LOT", 1, 2500, "=G32*H32"],
        ["", "", "", "", "", "", "", "=SUM(I28:I32)", "=I33"]
    ];
    const wsBreakdown = XLSX.utils.aoa_to_sheet(breakdownRows);

    const qtyRows = [
        [],
        ["", "", "Dome", "Bullet", "Pole", "PTZ", "K-POI", "Pole", "Total Camera"],
        ["", "Site Plan", domeQty, bulletQty, 0, 0, 0, 0, "=SUM(C3:H3)"],
        [],
        [],
        ["", "", domeQty, bulletQty, 0, 0, 0, 0, "=C3+D3"]
    ];
    const wsQTY = XLSX.utils.aoa_to_sheet(qtyRows);

    XLSX.utils.book_append_sheet(wb, wsBOQ, "UNV-BOQ");
    XLSX.utils.book_append_sheet(wb, wsBreakdown, "BREAK DOWN");
    XLSX.utils.book_append_sheet(wb, wsQTY, "QTY");

    XLSX.writeFile(wb, `BOQ_${projSubject.replace(/[^a-zA-Z0-9]/g, '_')}.xlsx`);
}

window.onload = function() {
    calculateBOQ();
};
