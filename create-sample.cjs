const XLSX_MODULE = require('xlsx');

// Create sample data
const data = [
  ['Name', 'ID Number', 'Position', 'Date', 'Company'],
  ['أحمد محمد علي حسن', '29901011234567', 'فني كهرباء', '2026-01-15', 'AFRE'],
  ['محمد إبراهيم عبدالله', '30012152345678', 'عامل لحام', '2026-01-15', 'AFRE'],
  ['عبدالرحمن سعيد محمود', '28805203456789', 'مهندس مدني', '2026-02-01', 'CEEC'],
];

const ws = XLSX_MODULE.utils.aoa_to_sheet(data);
const wb = XLSX_MODULE.utils.book_new();
XLSX_MODULE.utils.book_append_sheet(wb, ws, 'Workers');
XLSX_MODULE.writeFile(wb, 'sample_workers.xlsx');
console.log('Sample Excel file created: sample_workers.xlsx');
