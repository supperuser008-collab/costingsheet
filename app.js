// 4. EXPORT TO EXCEL FUNCTIONALITY
document.getElementById('btnExport').addEventListener('click', function () {
  const table = document.querySelector("#boqSection table");
  if (!table) {
    alert("No BOQ generated yet!");
    return;
  }

  // Convert HTML Table directly into an Excel Workbook
  const wb = XLSX.utils.table_to_book(table, { sheet: "CCTV_BOQ_Proposal" });
  
  // Format column widths for a clean presentation
  const ws = wb.Sheets["CCTV_BOQ_Proposal"];
  ws['!cols'] = [
    { wch: 18 }, // Category
    { wch: 38 }, // Description
    { wch: 22 }, // Brand / Model
    { wch: 10 }, // QTY
    { wch: 15 }, // Unit Cost
    { wch: 18 }, // Unit Sell (+Margin)
    { wch: 20 }  // Total Sell (QAR)
  ];

  // Trigger browser download
  XLSX.writeFile(wb, `Flora_CCTV_BOQ_${new Date().toISOString().slice(0, 10)}.xlsx`);
});
