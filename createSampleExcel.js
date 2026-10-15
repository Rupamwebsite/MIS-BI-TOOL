const xlsx = require('xlsx');
const path = require('path');
const { generateHospitalBillingSample, generateSalesERPSample } = require('./seedSamples');

function createSampleFiles() {
  // 1. Hospital Sample
  const hosp = generateHospitalBillingSample();
  const ws1 = xlsx.utils.json_to_sheet(hosp.rows);
  const wb1 = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb1, ws1, 'OPD_Billing_Report');
  const hospPath = path.join(__dirname, 'Sample_OPD_Billing_MIS.xlsx');
  xlsx.writeFile(wb1, hospPath);
  console.log(`[Created] ${hospPath}`);

  // 2. Sales Sample
  const sales = generateSalesERPSample();
  const ws2 = xlsx.utils.json_to_sheet(sales.rows);
  const wb2 = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb2, ws2, 'Sales_Summary');
  const salesPath = path.join(__dirname, 'Sample_Sales_ERP_MIS.xlsx');
  xlsx.writeFile(wb2, salesPath);
  console.log(`[Created] ${salesPath}`);
}

createSampleFiles();
