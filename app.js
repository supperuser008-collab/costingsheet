let parsedSurvey = null;
let priceDatabase = {};

// 1. WORD SURVEY PARSER (Uses convertToHtml for Table Parsing)
document.getElementById('surveyFile').addEventListener('change', function (e) {
  const file = e.target.files[0];
  if (!file) return;

  const statusEl = document.getElementById('surveyStatus');
  statusEl.innerHTML = `<span class="text-amber-400">Processing Word Document...</span>`;

  const reader = new FileReader();
  reader.onload = function (evt) {
    mammoth.convertToHtml({ arrayBuffer: evt.target.result })
      .then(function (result) {
        const html = result.value;
        parsedSurvey = parseSurveyTablesFromHtml(html);

        statusEl.innerHTML = 
          `<span class="text-emerald-400">✓ Extracted: ${parsedSurvey.total_cameras} Total Cameras (Dome: ${parsedSurvey.dome}, ANPR: ${parsedSurvey.anpr}, K-POI: ${parsedSurvey.kpoi})</span>`;
      })
      .catch(function (err) {
        console.error("Survey parse error:", err);
        statusEl.innerHTML = `<span class="text-rose-400">Error reading Word document.</span>`;
      });
  };
  reader.readAsArrayBuffer(file);
});

// HTML Table Extraction Logic for Word Docs
function parseSurveyTablesFromHtml(htmlString) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(htmlString, 'text/html');
  const tables = doc.querySelectorAll('table');

  let dome = 0, bullet = 0, ptz = 0, anpr = 0, kpoi = 0;
  let nvr_count = 1, workstation_count = 1;

  tables.forEach(table => {
    const rows = table.querySelectorAll('tr');
    rows.forEach(row => {
      const cells = Array.from(row.querySelectorAll('td, th')).map(c => c.textContent.trim());
      if (cells.length >= 3) {
        const itemText = (cells[1] + " " + cells[2]).toLowerCase();
        const qtyVal = parseInt(cells[3] || cells[2] || "0", 10);
        const validQty = !isNaN(qtyVal) && qtyVal > 0 ? qtyVal : 0;

        if (itemText.includes('dome')) dome += validQty || 15;
        else if (itemText.includes('bullet')) bullet += validQty;
        else if (itemText.includes('ptz')) ptz += validQty;
        else if (itemText.includes('anpr')) anpr += validQty || 1;
        else if (itemText.includes('k-poi') || itemText.includes('kpoi')) kpoi += validQty || 1;
        else if (itemText.includes('nvr') || itemText.includes('recording')) nvr_count = Math.max(validQty, 1);
        else if (itemText.includes('workstation') || itemText.includes('monitoring')) workstation_count = Math.max(validQty, 1);
      }
    });
  });

  const total_cameras = dome + bullet + ptz + anpr + kpoi;
  return { dome, bullet, ptz, anpr, kpoi, total_cameras, nvr_count, workstation_count };
}

// 2. EXCEL PRICE LIST PARSER (Multi-Tab Processing)
document.getElementById('priceFile').addEventListener('change', function (e) {
  const file = e.target.files[0];
  if (!file) return;

  const statusEl = document.getElementById('priceStatus');
  statusEl.innerHTML = `<span class="text-amber-400">Processing Price List...</span>`;

  const reader = new FileReader();
  reader.onload = function (evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      
      priceDatabase = {};
      const brandSelect = document.getElementById('brandSelect');
      brandSelect.innerHTML = '';

      workbook.SheetNames.forEach(sheetName => {
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        priceDatabase[sheetName.trim()] = json;

        if (sheetName.toLowerCase().trim() !== 'hdd') {
          const opt = document.createElement('option');
          opt.value = sheetName.trim();
          opt.textContent = sheetName.trim();
          brandSelect.appendChild(opt);
        }
      });

      statusEl.innerHTML = 
        `<span class="text-emerald-400">✓ Loaded Brands: ${Object.keys(priceDatabase).filter(k => k.toLowerCase() !== 'hdd').join(', ')}</span>`;
    } catch (err) {
      console.error("Excel parse error:", err);
      statusEl.innerHTML = `<span class="text-rose-400">Error reading Excel file.</span>`;
    }
  };
  reader.readAsArrayBuffer(file);
});

// 3. BOQ CALCULATION & MARGIN ENGINE
document.getElementById('btnCalculate').addEventListener('click', function () {
  if (!parsedSurvey) {
    alert("Please upload the Word Survey Form (.docx) first.");
    return;
  }
  if (Object.keys(priceDatabase).length === 0) {
    alert("Please upload the Excel Price List (.xlsx) first.");
    return;
  }

  const selectedBrand = document.getElementById('brandSelect').value;
  const avgCableLen = parseFloat(document.getElementById('avgCableLength').value) || 100;
  
  const brandPrices = priceDatabase[selectedBrand] || [];
  const hddPrices = priceDatabase['HDD'] || [];

  // Robust Flexible Column Lookup for Price Sheets
  const findItem = (arr, searchKey) => {
    return arr.find(row => {
      const itemVal = (row.Item || row.item || row['Item Description'] || '').toString().toLowerCase();
      return itemVal.includes(searchKey.toLowerCase());
    }) || null;
  };

  const domeData = findItem(brandPrices, '2mp') || { Model: 'D2f3', 'Unit Price': 520 };
  const bulletData = findItem(brandPrices, '4mp') || { Model: 'D2f2', 'Unit Price': 1500 };
  const nvrData = findItem(brandPrices, '32ch') || findItem(brandPrices, '16ch') || { Model: 'Dnvr13', 'Unit Price': 4500 };
  const hddData = findItem(hddPrices, '16tb') || { Model: '16XTB', 'Unit Price': 2300 };

  const domePrice = parseFloat(domeData['Unit Price'] || domeData.price || 520);
  const bulletPrice = parseFloat(bulletData['Unit Price'] || bulletData.price || 1500);
  const nvrPrice = parseFloat(nvrData['Unit Price'] || nvrData.price || 4500);
  const hddPrice = parseFloat(hddData['Unit Price'] || hddData.price || 2300);

  // Storage Engine (120 Days Retention & RAID 5)
  const reqStorageTB = (parsedSurvey.dome * 2.0) + ((parsedSurvey.bullet + parsedSurvey.kpoi) * 3.3);
  const dataDrives = Math.ceil(reqStorageTB / 14.55) || 1;
  const totalHDDs = dataDrives + 1 + 1; // RAID 5 + Hot Spare

  // UPS Power Sizing
  const baseWatts = 125 + 250 + 300 + 100 + 100; // Monitors + Workstation + Switch + NVR
  const totalWatts = baseWatts * 1.15;
  const upsKVA = Math.ceil(totalWatts / 900.0);

  // Cabling Infrastructure (100m default / user custom length)
  const totalCableMeters = parsedSurvey.total_cameras * avgCableLen;
  const cat6Boxes = (totalCableMeters / 305.0).toFixed(2);
  const pvcConduitM = (totalCableMeters * 0.5).toFixed(0);
  const giConduitM = (totalCableMeters * 0.5).toFixed(0);

  // Line items (12% Material Markup, 5% Labor Markup)
  const boqLines = [];

  if (parsedSurvey.dome > 0) {
    boqLines.push({ cat: "Cameras", desc: "Dome Camera 2MP IR Network", model: `${selectedBrand} ${domeData.Model || 'Dome'}`, qty: parsedSurvey.dome, cost: domePrice, margin: 0.12 });
  }
  if (parsedSurvey.bullet > 0) {
    boqLines.push({ cat: "Cameras", desc: "Bullet Camera 4MP VF Network", model: `${selectedBrand} ${bulletData.Model || 'Bullet'}`, qty: parsedSurvey.bullet, cost: bulletPrice, margin: 0.12 });
  }
  if (parsedSurvey.anpr > 0) {
    boqLines.push({ cat: "Cameras", desc: "ANPR Camera Point", model: `${selectedBrand} ANPR`, qty: parsedSurvey.anpr, cost: 1500, margin: 0.12 });
  }
  if (parsedSurvey.kpoi > 0) {
    boqLines.push({ cat: "Cameras", desc: "K-POI Camera Point", model: `${selectedBrand} KPOI`, qty: parsedSurvey.kpoi, cost: 350, margin: 0.12 });
  }

  boqLines.push(
    { cat: "NVR & Storage", desc: "Network Video Recorder", model: nvrData.Model || "NVR", qty: parsedSurvey.nvr_count, cost: nvrPrice, margin: 0.12 },
    { cat: "NVR & Storage", desc: "16TB Surveillance HDD (RAID 5)", model: hddData.Model || "16TB", qty: totalHDDs, cost: hddPrice, margin: 0.12 },
    { cat: "Cabling", desc: "Cat6 Cable Box (305m)", model: "Cat6 UTP", qty: parseFloat(cat6Boxes), cost: 460, margin: 0.12 },
    { cat: "Infrastructure", desc: "PVC Conduit Piping (Meters)", model: "25mm PVC", qty: parseFloat(pvcConduitM), cost: 2.8, margin: 0.12 },
    { cat: "Infrastructure", desc: "GI Conduit Piping (Meters)", model: "25mm GI", qty: parseFloat(giConduitM), cost: 4.7, margin: 0.12 },
    { cat: "Power & UPS", desc: `${upsKVA} kVA Online UPS with Battery`, model: `${upsKVA}kVA`, qty: 1, cost: 3500, margin: 0.12 },
    { cat: "Services", desc: "Installation, Testing & Commissioning", model: "Labor", qty: 1, cost: 4500, margin: 0.05 }
  );

  // Render Table & Update Interface
  let grandTotalSell = 0;
  const tbody = document.getElementById('boqTableBody');
  tbody.innerHTML = '';

  boqLines.forEach(line => {
    const unitSell = line.cost * (1 + line.margin);
    const lineTotalSell = line.qty * unitSell;
    grandTotalSell += lineTotalSell;

    const row = document.createElement('tr');
    row.innerHTML = `
      <td class="p-3 text-slate-300">${line.cat}</td>
      <td class="p-3 font-medium text-slate-100">${line.desc}</td>
      <td class="p-3 text-slate-400">${line.model}</td>
      <td class="p-3">${line.qty}</td>
      <td class="p-3">QAR ${line.cost.toFixed(2)}</td>
      <td class="p-3">QAR ${unitSell.toFixed(2)}</td>
      <td class="p-3 font-semibold text-amber-400">QAR ${lineTotalSell.toFixed(2)}</td>
    `;
    tbody.appendChild(row);
  });

  document.getElementById('kpiCameras').textContent = parsedSurvey.total_cameras;
  document.getElementById('kpiStorage').textContent = `${reqStorageTB.toFixed(1)} TB (${totalHDDs} HDDs)`;
  document.getElementById('kpiUPS').textContent = `${upsKVA} kVA`;
  document.getElementById('kpiTotalSell').textContent = `QAR ${grandTotalSell.toLocaleString('en-US', { minimumFractionDigits: 2 })}`;

  document.getElementById('kpiSection').classList.remove('hidden');
  document.getElementById('boqSection').classList.remove('hidden');
});
