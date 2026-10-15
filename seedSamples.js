/**
 * Pre-configured ERP MIS Datasets Generator
 * Allows immediate evaluation with realistic Hospital Billing and Commercial ERP data.
 */

function generateHospitalBillingSample() {
  const departments = ['Cardiology', 'Orthopedics', 'Neurology', 'Pediatrics', 'Radiology', 'Gynecology', 'General Medicine', 'Oncology'];
  const doctors = {
    'Cardiology': ['Dr. S. Mukherjee', 'Dr. A. Roy', 'Dr. R. Banerjee'],
    'Orthopedics': ['Dr. P. Sen', 'Dr. K. Das'],
    'Neurology': ['Dr. M. Ghosh', 'Dr. N. Chatterjee'],
    'Pediatrics': ['Dr. T. Bose', 'Dr. S. Guha'],
    'Radiology': ['Dr. V. Sharma', 'Dr. D. Dey'],
    'Gynecology': ['Dr. R. Mitra', 'Dr. S. Bhattacharya'],
    'General Medicine': ['Dr. B. Paul', 'Dr. K. Saha'],
    'Oncology': ['Dr. P. Ganguly', 'Dr. H. Chakraborty']
  };

  const services = {
    'Cardiology': [{ name: 'Echocardiogram', price: 2500 }, { name: 'TMT Test', price: 1800 }, { name: 'Consultation', price: 800 }, { name: 'ECG', price: 350 }],
    'Orthopedics': [{ name: 'Joint Injections', price: 3200 }, { name: 'X-Ray Knee', price: 850 }, { name: 'Consultation', price: 750 }, { name: 'Cast & Splint', price: 1500 }],
    'Neurology': [{ name: 'EEG Brain Scan', price: 3800 }, { name: 'NCV Test', price: 4200 }, { name: 'Consultation', price: 1000 }, { name: 'EMG Study', price: 3500 }],
    'Pediatrics': [{ name: 'Vaccination Routine', price: 1200 }, { name: 'Consultation', price: 700 }, { name: 'Nebulization', price: 400 }],
    'Radiology': [{ name: 'MRI Brain Contrast', price: 7500 }, { name: 'CT Abdomen', price: 5500 }, { name: 'USG Whole Abdomen', price: 1600 }, { name: 'Chest X-Ray', price: 600 }],
    'Gynecology': [{ name: 'Anomaly Ultrasound', price: 2800 }, { name: 'Consultation', price: 900 }, { name: 'Pap Smear Lab', price: 1100 }],
    'General Medicine': [{ name: 'General Consultation', price: 600 }, { name: 'Blood Sugar & Lipid Profile', price: 950 }, { name: 'Comprehensive Health Checkup', price: 4500 }],
    'Oncology': [{ name: 'Chemotherapy Daycare', price: 12500 }, { name: 'Consultation', price: 1200 }, { name: 'Biopsy Histopathology', price: 3800 }]
  };

  const paymentModes = ['UPI / QR Code', 'Credit Card', 'Debit Card', 'Cash', 'TPA Insurance', 'Corporate Credit'];
  const patientTypes = ['OPD Walk-in', 'Emergency Visit', 'Corporate Client', 'Senior Citizen', 'VIP Patient'];
  const firstNames = ['Amit', 'Rahul', 'Pooja', 'Sneha', 'Sourav', 'Ananya', 'Debashis', 'Sunita', 'Tanmoy', 'Rina', 'Manoj', 'Priyanka', 'Vikram', 'Kakoli', 'Subhash'];
  const lastNames = ['Dutta', 'Bose', 'Chatterjee', 'Roy', 'Chakraborty', 'Banerjee', 'Mandal', 'Sarkar', 'Biswas', 'Mukherjee', 'Das', 'Sen'];

  const rows = [];
  const baseDate = new Date();
  baseDate.setDate(baseDate.getDate() - 45); // Past 45 days

  for (let i = 1; i <= 280; i++) {
    const curDate = new Date(baseDate.getTime() + Math.floor(Math.random() * 45 * 86400000));
    const dept = departments[Math.floor(Math.random() * departments.length)];
    const docList = doctors[dept];
    const doctor = docList[Math.floor(Math.random() * docList.length)];
    const srvList = services[dept];
    const srv = srvList[Math.floor(Math.random() * srvList.length)];

    const billNo = `OPD-2024-${String(1000 + i)}`;
    const uhid = `UHID-${String(50000 + (i % 180))}`;
    const pName = `${firstNames[Math.floor(Math.random() * firstNames.length)]} ${lastNames[Math.floor(Math.random() * lastNames.length)]}`;
    const age = Math.floor(Math.random() * 65) + 12;
    const gender = Math.random() > 0.48 ? 'Male' : 'Female';
    const payMode = paymentModes[Math.floor(Math.random() * paymentModes.length)];
    const pType = patientTypes[Math.floor(Math.random() * patientTypes.length)];

    // Pricing calculation
    const qty = Math.random() > 0.85 ? 2 : 1;
    const billAmount = srv.price * qty;
    // Occasional discount for seniors / corporate
    const discountPct = (pType === 'Senior Citizen' || pType === 'Corporate Client') ? (Math.random() > 0.5 ? 10 : 0) : 0;
    const discount = Math.round((billAmount * discountPct) / 100);
    const netAmount = billAmount - discount;
    const status = (i % 47 === 0) ? 'Refunded' : (i % 65 === 0 ? 'Pending' : 'Settled');

    rows.push({
      'Bill No': billNo,
      'Bill Date': curDate.toISOString().split('T')[0],
      'UHID': uhid,
      'Patient Name': pName,
      'Age': age,
      'Gender': gender,
      'Patient Category': pType,
      'Department': dept,
      'Consulting Doctor': doctor,
      'Service Name': srv.name,
      'Quantity': qty,
      'Gross Amount': billAmount,
      'Discount': discount,
      'Net Amount': netAmount,
      'Payment Mode': payMode,
      'Billing Status': status
    });
  }

  return {
    name: 'Hospital OPD & Clinical Billing MIS (Q3-Q4)',
    filename: 'OPD_Billing_Executive_Report_2024.xlsx',
    detectedType: 'Hospital & OPD Billing MIS',
    rows,
    columns: Object.keys(rows[0])
  };
}

function generateSalesERPSample() {
  const regions = ['North Zone', 'East Zone', 'West Zone', 'South Zone', 'Central Zone'];
  const categories = ['Electronics & Gadgets', 'Office Hardware', 'Network Equipment', 'Software Licenses', 'Industrial Consumables'];
  const products = {
    'Electronics & Gadgets': [{ name: 'Enterprise Tablet 10"', price: 18500 }, { name: 'Handheld Barcode Scanner', price: 6200 }, { name: 'Thermal Receipt Printer', price: 8400 }],
    'Office Hardware': [{ name: 'LaserJet Pro Workstation', price: 34000 }, { name: 'Ergonomic Desk Station', price: 14500 }, { name: 'Biometric Access Terminal', price: 12000 }],
    'Network Equipment': [{ name: 'Managed 24-Port Switch', price: 28000 }, { name: 'Enterprise Wi-Fi 6 AP', price: 15600 }, { name: 'Fiber Patch Panel 48-Core', price: 9200 }],
    'Software Licenses': [{ name: 'ERP Cloud User License (Annual)', price: 45000 }, { name: 'Security Suite Enterprise', price: 22000 }, { name: 'Database Maintenance Pack', price: 38000 }],
    'Industrial Consumables': [{ name: 'Thermal Paper Rolls (Box of 50)', price: 3200 }, { name: 'Barcode Labels Roll (10k)', price: 2400 }, { name: 'High-Yield Toner Cartridge', price: 5800 }]
  };

  const reps = ['Rajesh Verma', 'Anupam Sen', 'Kavita Iyer', 'Gaurav Mehta', 'Sunil Nambiar', 'Pooja Agarwal'];
  const clients = ['Tata Logistics Ltd', 'Apollo Healthcare', 'Reliance Retail', 'Wipro Digital', 'Infosys Tech Hub', 'L&T Infra Solutions', 'Titan Industries'];
  const payTerms = ['Net 30 Days', 'Immediate Online', 'Net 15 Days', 'Letter of Credit', 'Advance 50%'];

  const rows = [];
  const baseDate = new Date();
  baseDate.setDate(baseDate.getDate() - 60);

  for (let i = 1; i <= 250; i++) {
    const curDate = new Date(baseDate.getTime() + Math.floor(Math.random() * 60 * 86400000));
    const region = regions[Math.floor(Math.random() * regions.length)];
    const rep = reps[Math.floor(Math.random() * reps.length)];
    const client = clients[Math.floor(Math.random() * clients.length)];
    const category = categories[Math.floor(Math.random() * categories.length)];
    const prodList = products[category];
    const prod = prodList[Math.floor(Math.random() * prodList.length)];
    const terms = payTerms[Math.floor(Math.random() * payTerms.length)];

    const qty = Math.floor(Math.random() * 8) + 1;
    const gross = prod.price * qty;
    const taxRate = 0.18;
    const taxAmount = Math.round(gross * taxRate);
    const netRevenue = gross + taxAmount;
    const status = (i % 38 === 0) ? 'Cancelled' : (i % 12 === 0 ? 'Pending Payment' : 'Delivered & Paid');

    rows.push({
      'Invoice No': `INV-ERP-${String(2024000 + i)}`,
      'Invoice Date': curDate.toISOString().split('T')[0],
      'Customer Name': client,
      'Sales Representative': rep,
      'Territory Region': region,
      'Product Category': category,
      'Item Description': prod.name,
      'Unit Price': prod.price,
      'Order Qty': qty,
      'Gross Total': gross,
      'GST Tax (18%)': taxAmount,
      'Net Invoice Value': netRevenue,
      'Payment Terms': terms,
      'Fulfillment Status': status
    });
  }

  return {
    name: 'Enterprise B2B Sales & Distribution ERP Report',
    filename: 'B2B_Commercial_Sales_2024.xlsx',
    detectedType: 'Sales & Commercial ERP',
    rows,
    columns: Object.keys(rows[0])
  };
}

module.exports = {
  generateHospitalBillingSample,
  generateSalesERPSample
};
