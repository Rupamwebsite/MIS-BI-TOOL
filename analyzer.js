/**
 * Advanced ERP MIS Intelligence & Analytics Engine v2.0
 * Features:
 * - Smart Header Detection (handles multi-row titles, blanks, and ERP export banners)
 * - Trailing Summary / Grand Total row exclusion (prevents double-counting)
 * - Strict Identifier & Metadata exclusion (Phone, PIN, Sl No, UHID, Age are NEVER treated as metrics)
 * - Multi-format Date parsing (DD/MM/YYYY, DD-MM-YYYY, Excel serial numbers, ISO)
 * - Dynamic Metric and Dimension switching
 * - Advanced Multi-Column & Multi-Operator Filtering
 * - Pareto (80/20) and Anomaly Intelligence
 */

const xlsx = require('xlsx');

// Identifier & Metadata keywords that must NEVER be treated as financial metrics
const IDENTIFIER_KEYWORDS = [
  'id', 'uhid', 'no', 'number', 'code', 'phone', 'mobile', 'contact', 'pin', 'pincode',
  'zip', 'zipcode', 'age', 'year', 'sl', 'serial', 'sr', 'sno', 'bill_no', 'inv_no',
  'invoice_no', 'token', 'mrn', 'ipd_no', 'opd_no', 'aadhaar', 'pan', 'ref_no', 'ip_no'
];

// Summary & Total row keywords to exclude from bottom of sheet
const SUMMARY_ROW_KEYWORDS = [
  'total', 'grand total', 'sub total', 'subtotal', 'summary', 'overall', 'closing balance',
  'average', 'mean', 'count', 'maximum', 'minimum'
];

/**
 * Smart Sheet Parser
 * Automatically identifies table headers, ignores title banners, and trims total rows.
 */
function smartParseSheet(worksheet) {
  const rawMatrix = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: '', blankrows: false });
  if (!rawMatrix || rawMatrix.length === 0) {
    return { rows: [], columns: [], headerRowIndex: 0 };
  }

  // Find Header Row: scan first 15 rows for the one with maximum distinct non-empty string headers
  let bestHeaderIdx = 0;
  let maxHeaderScore = -1;

  const maxScanRows = Math.min(15, rawMatrix.length);
  for (let r = 0; r < maxScanRows; r++) {
    const row = rawMatrix[r];
    if (!Array.isArray(row)) continue;

    // Filter non-empty cells
    const textCells = row.filter(cell => cell !== null && cell !== undefined && String(cell).trim() !== '');
    if (textCells.length < 2) continue;

    // Check how many cells look like string column headers vs numbers
    let score = 0;
    const seen = new Set();
    for (const cell of textCells) {
      const s = String(cell).trim();
      if (isNaN(Number(s)) || s.length <= 4) {
        score += 2; // String header gets higher weight
      } else {
        score += 0.5;
      }
      seen.add(s.toLowerCase());
    }

    // Reward diversity of column names
    score += seen.size;

    if (score > maxHeaderScore) {
      maxHeaderScore = score;
      bestHeaderIdx = r;
    }
  }

  // Extract raw column headers
  const rawHeaders = rawMatrix[bestHeaderIdx].map((h, idx) => {
    const str = String(h || '').trim();
    return str || `Column_${idx + 1}`;
  });

  // Deduplicate headers
  const headers = [];
  const counts = {};
  for (const h of rawHeaders) {
    const base = h.replace(/[\r\n\t]/g, ' ').replace(/\s+/g, ' ').trim();
    if (!counts[base]) {
      counts[base] = 1;
      headers.push(base);
    } else {
      counts[base]++;
      headers.push(`${base}_${counts[base]}`);
    }
  }

  // Parse Data Rows starting after bestHeaderIdx
  const parsedRows = [];
  for (let r = bestHeaderIdx + 1; r < rawMatrix.length; r++) {
    const rowArr = rawMatrix[r];
    if (!Array.isArray(rowArr) || rowArr.length === 0) continue;

    // Check if entire row is empty
    const nonBlank = rowArr.some(c => c !== null && c !== undefined && String(c).trim() !== '');
    if (!nonBlank) continue;

    // Check if this is a "Grand Total" or "Summary" row at the bottom (only inspect last 5 rows)
    const isNearBottom = (rawMatrix.length - r) <= 5;
    const firstNonBlank = rowArr.find(c => c !== null && c !== undefined && String(c).trim() !== '');
    const firstStr = String(firstNonBlank || '').toLowerCase().trim();
    const isSummaryRow = isNearBottom && SUMMARY_ROW_KEYWORDS.some(kw => firstStr === kw || firstStr === kw + ':' || firstStr === 'grand total' || firstStr === 'sub total');
    if (isSummaryRow) {
      // Skip summary row to avoid double counting!
      continue;
    }

    // Map row to object
    const rowObj = {};
    let hasValidData = false;
    headers.forEach((colName, cIdx) => {
      let val = rowArr[cIdx];
      if (val === undefined || val === '') val = null;
      // Handle Date objects if xlsx parsed them
      if (val instanceof Date) {
        val = val.toISOString().split('T')[0];
      }
      rowObj[colName] = val;
      if (val !== null) hasValidData = true;
    });

    if (hasValidData) {
      parsedRows.push(rowObj);
    }
  }

  return {
    rows: parsedRows,
    columns: headers,
    headerRowIndex: bestHeaderIdx
  };
}

/**
 * Intelligent Data & Column Type Detection
 */
function isIdentifierColumn(colName) {
  const norm = colName.toLowerCase().replace(/[^a-z0-9]/g, '_');
  return IDENTIFIER_KEYWORDS.some(kw => {
    return norm === kw || norm.startsWith(kw + '_') || norm.endsWith('_' + kw) || norm.includes('_' + kw + '_');
  });
}

function parseNumber(val) {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const str = String(val).replace(/[$₹€£,\s]/g, '');
  const parsed = parseFloat(str);
  return isNaN(parsed) ? 0 : parsed;
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;

  // Handle Excel Serial Date Number (e.g. 45000)
  if (typeof val === 'number' && val > 30000 && val < 60000) {
    const epoch = new Date(1899, 11, 30);
    const d = new Date(epoch.getTime() + val * 86400000);
    return isNaN(d.getTime()) ? null : d;
  }

  const str = String(val).trim();

  // Try standard parse
  const std = new Date(str);
  if (!isNaN(std.getTime()) && str.includes('-') && std.getFullYear() > 1990 && std.getFullYear() < 2100) {
    return std;
  }

  // Handle DD/MM/YYYY or DD-MM-YYYY
  const dmyMatch = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10);
    const month = parseInt(dmyMatch[2], 10) - 1;
    let year = parseInt(dmyMatch[3], 10);
    if (year < 100) year += 2000;
    const d = new Date(year, month, day);
    if (!isNaN(d.getTime())) return d;
  }

  return isNaN(std.getTime()) ? null : std;
}

function detectColumnType(colName, values) {
  const nonNull = values.filter(v => v !== null && v !== undefined && String(v).trim() !== '');
  if (nonNull.length === 0) return 'empty';

  // If column name clearly identifies an ID or Contact, do NOT classify as financial numeric
  if (isIdentifierColumn(colName)) {
    return 'identifier';
  }

  let numCount = 0;
  let dateCount = 0;
  let boolCount = 0;

  for (const val of nonNull) {
    if (typeof val === 'boolean' || val === 'true' || val === 'false') {
      boolCount++;
      continue;
    }

    // Check date
    if (val instanceof Date || (typeof val === 'string' && val.match(/[-/]/) && parseDate(val) !== null && isNaN(Number(val)))) {
      dateCount++;
      continue;
    }

    // Check number (cleaned of currencies, commas)
    const cleanNum = String(val).replace(/[$₹€£,\s]/g, '');
    if (!isNaN(cleanNum) && cleanNum !== '') {
      numCount++;
      continue;
    }
  }

  const threshold = nonNull.length * 0.65;
  if (dateCount >= threshold) return 'date';
  if (numCount >= threshold) return 'numeric';
  if (boolCount >= threshold) return 'boolean';

  // Dimension keywords detection (Doctor, Department, Ward, Panel, Category, etc.)
  const colNorm = colName.toLowerCase().replace(/[^a-z0-9]/g, '_');
  const isExplicitDimension = [
    'doctor', 'primary_doctor', 'consulting_doctor', 'physician', 'surgeon', 'consultant',
    'department', 'dept', 'speciality', 'ward', 'discharge_ward', 'room', 'discharge_room', 'bed',
    'panel', 'panel_group', 'insurance', 'company', 'classification', 'patient_classification',
    'purpose', 'admission_purpose', 'condition', 'discharge_condition', 'category', 'referred', 'refferred',
    'status', 'bill_status', 'mode', 'payment_mode', 'type', 'organization', 'state', 'district', 'city', 'country',
    'service', 'procedure', 'test', 'item', 'customer', 'agent', 'sales_rep', 'client',
    'branch', 'territory', 'region', 'gender'
  ].some(kw => colNorm.includes(kw));

  // Check unique values ratio for category vs general text
  const uniqueVals = new Set(nonNull.map(v => String(v).toLowerCase().trim()));
  if (isExplicitDimension || uniqueVals.size <= Math.min(350, nonNull.length * 0.85)) {
    return 'category';
  }
  return 'text';
}

/**
 * Main Dataset Analysis
 * Supports custom overrides for primary metric & primary dimension.
 */
function analyzeDataset(rows, columns, options = {}) {
  if (!rows || rows.length === 0) {
    return {
      rowCount: 0,
      colCount: columns ? columns.length : 0,
      columns: [],
      summaryKPIs: {},
      insights: [],
      charts: {},
      categoryBreakdown: [],
      paymentBreakdown: [],
      trendData: [],
      anomalies: []
    };
  }

  // 1. Column Profiles
  const columnProfiles = columns.map(col => {
    const rawValues = rows.map(r => r[col]);
    const detectedType = detectColumnType(col, rawValues);
    const nonNull = rawValues.filter(v => v !== null && v !== undefined && String(v).trim() !== '');

    let stats = {
      totalCount: rows.length,
      nonNullCount: nonNull.length,
      nullCount: rows.length - nonNull.length
    };

    if (detectedType === 'numeric') {
      const nums = nonNull.map(v => parseNumber(v)).sort((a, b) => a - b);
      if (nums.length > 0) {
        const sum = nums.reduce((acc, n) => acc + n, 0);
        const avg = sum / nums.length;
        const min = nums[0];
        const max = nums[nums.length - 1];
        const median = nums[Math.floor(nums.length / 2)];
        const variance = nums.reduce((acc, n) => acc + Math.pow(n - avg, 2), 0) / nums.length;
        const stdDev = Math.sqrt(variance);

        stats = {
          ...stats,
          sum: Math.round(sum * 100) / 100,
          avg: Math.round(avg * 100) / 100,
          min,
          max,
          median,
          stdDev: Math.round(stdDev * 100) / 100
        };
      }
    } else if (detectedType === 'category' || detectedType === 'text') {
      const freqMap = {};
      nonNull.forEach(v => {
        const k = String(v).trim();
        freqMap[k] = (freqMap[k] || 0) + 1;
      });
      const topItems = Object.entries(freqMap)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15)
        .map(([val, count]) => ({
          value: val,
          count,
          pct: Math.round((count / nonNull.length) * 1000) / 10
        }));

      stats = {
        ...stats,
        uniqueCount: Object.keys(freqMap).length,
        topItems
      };
    } else if (detectedType === 'date') {
      const validDates = nonNull.map(v => parseDate(v)).filter(d => d !== null).sort((a, b) => a - b);
      if (validDates.length > 0) {
        stats = {
          ...stats,
          startDate: validDates[0].toISOString().split('T')[0],
          endDate: validDates[validDates.length - 1].toISOString().split('T')[0],
          dateCount: validDates.length
        };
      }
    }

    return {
      name: col,
      type: detectedType,
      sampleValues: nonNull.slice(0, 5),
      stats
    };
  });

  // 2. Metric Column Selection (Highest business priority)
  const metricRankings = [
    'net_amount', 'net amount', 'net_amt', 'net amt', 'net_collection', 'net collection',
    'total_amount', 'total amount', 'total_amt', 'bill_amount', 'bill amount', 'bill_amt',
    'grand_total', 'grand total', 'collection', 'revenue', 'sales', 'sales_amount',
    'gross_amount', 'gross amount', 'gross_amt', 'amount', 'total', 'fee', 'price',
    'paid', 'charge', 'cost', 'val', 'value'
  ];

  let primaryMetricCol = null;
  // If user specified override:
  if (options.metricCol && columnProfiles.some(c => c.name === options.metricCol)) {
    primaryMetricCol = columnProfiles.find(c => c.name === options.metricCol);
  } else {
    // Ranked lookup
    for (const kw of metricRankings) {
      primaryMetricCol = columnProfiles.find(c => 
        c.type === 'numeric' && c.name.toLowerCase().replace(/[^a-z0-9]/g, ' ').includes(kw)
      );
      if (primaryMetricCol) break;
    }
    // Fallback: pick numeric column that is NOT an identifier
    if (!primaryMetricCol) {
      primaryMetricCol = columnProfiles.find(c => c.type === 'numeric' && !isIdentifierColumn(c.name));
    }
  }

  // 3. Dimension Column Selection (Doctor, Department, Category, Product, etc.)
  const dimensionRankings = [
    'doctor', 'primary_doctor', 'consulting_doctor', 'physician', 'surgeon', 'consultant',
    'department', 'dept', 'speciality', 'ward', 'discharge_ward', 'panel', 'panel_group',
    'category', 'product_category', 'classification', 'patient_classification', 'purpose',
    'admission_purpose', 'product', 'service', 'item', 'customer', 'client', 'territory',
    'region', 'branch', 'type'
  ];

  let primaryDimCol = null;
  // If user specified override:
  if (options.dimCol && columnProfiles.some(c => c.name === options.dimCol)) {
    primaryDimCol = columnProfiles.find(c => c.name === options.dimCol);
  } else {
    for (const kw of dimensionRankings) {
      primaryDimCol = columnProfiles.find(c => 
        (c.type === 'category' || c.type === 'text') &&
        c.name.toLowerCase().replace(/[^a-z0-9]/g, ' ').includes(kw)
      );
      if (primaryDimCol) break;
    }
    if (!primaryDimCol) {
      primaryDimCol = columnProfiles.find(c => c.type === 'category');
    }
  }

  // Secondary Payment / Status Dimension
  const statusRankings = ['payment_mode', 'pay_mode', 'payment_method', 'payment', 'method', 'status', 'mode'];
  const paymentDimCol = columnProfiles.find(c =>
    (c.type === 'category' || c.type === 'text') &&
    c.name !== (primaryDimCol ? primaryDimCol.name : '') &&
    statusRankings.some(kw => c.name.toLowerCase().includes(kw))
  );

  // Date Column
  const dateRankings = ['bill_date', 'invoice_date', 'txn_date', 'date', 'created_at', 'timestamp', 'time'];
  let primaryDateCol = columnProfiles.find(c => c.type === 'date');
  if (!primaryDateCol) {
    primaryDateCol = columnProfiles.find(c => dateRankings.some(kw => c.name.toLowerCase().includes(kw)));
  }

  // 4. Domain Classification
  const colText = columns.join(' ').toLowerCase();
  let detectedDomain = 'Enterprise General MIS';
  if (colText.includes('doctor') || colText.includes('patient') || colText.includes('opd') || colText.includes('uhid') || colText.includes('hospital') || colText.includes('clinic')) {
    detectedDomain = 'Hospital & OPD Billing MIS';
  } else if (colText.includes('sales') || colText.includes('order') || colText.includes('customer') || colText.includes('invoice') || colText.includes('sku')) {
    detectedDomain = 'Sales & Commercial ERP';
  } else if (colText.includes('inventory') || colText.includes('stock') || colText.includes('warehouse') || colText.includes('batch')) {
    detectedDomain = 'Supply Chain & Inventory ERP';
  } else if (colText.includes('debit') || colText.includes('credit') || colText.includes('ledger') || colText.includes('voucher') || colText.includes('expense')) {
    detectedDomain = 'Financial Accounting & P&L MIS';
  }

  // 5. Executive Summary KPIs
  const summaryKPIs = {
    totalRecords: rows.length,
    activeColumns: columns.length,
    detectedDomain,
    metricColumn: primaryMetricCol ? primaryMetricCol.name : null,
    dimensionColumn: primaryDimCol ? primaryDimCol.name : null,
    dateColumn: primaryDateCol ? primaryDateCol.name : null,
    totalValue: 0,
    averageTicket: 0,
    medianValue: 0,
    highestTransaction: 0,
    lowestTransaction: 0,
    zeroOrNegativeCount: 0,
    categoryCount: primaryDimCol && primaryDimCol.stats.uniqueCount ? primaryDimCol.stats.uniqueCount : 0
  };

  if (primaryMetricCol) {
    summaryKPIs.totalValue = primaryMetricCol.stats.sum || 0;
    summaryKPIs.averageTicket = primaryMetricCol.stats.avg || 0;
    summaryKPIs.medianValue = primaryMetricCol.stats.median || 0;
    summaryKPIs.highestTransaction = primaryMetricCol.stats.max || 0;
    summaryKPIs.lowestTransaction = primaryMetricCol.stats.min || 0;

    rows.forEach(r => {
      const v = parseNumber(r[primaryMetricCol.name]);
      if (v <= 0) summaryKPIs.zeroOrNegativeCount++;
    });
  }

  // 6. Categorical Aggregations (Top dimensions by metric value & volume)
  let categoryBreakdown = [];
  if (primaryDimCol) {
    const catMap = {};
    rows.forEach(r => {
      const rawCat = r[primaryDimCol.name];
      const catKey = (rawCat !== null && rawCat !== undefined && String(rawCat).trim() !== '') ? String(rawCat).trim() : '(Unspecified)';
      const metricVal = primaryMetricCol ? parseNumber(r[primaryMetricCol.name]) : 1;

      if (!catMap[catKey]) {
        catMap[catKey] = { category: catKey, volume: 0, totalValue: 0 };
      }
      catMap[catKey].volume++;
      catMap[catKey].totalValue += metricVal;
    });

    categoryBreakdown = Object.values(catMap)
      .map(item => ({
        ...item,
        totalValue: Math.round(item.totalValue * 100) / 100,
        avgValue: Math.round((item.totalValue / item.volume) * 100) / 100,
        sharePct: summaryKPIs.totalValue > 0 ? Math.round((item.totalValue / summaryKPIs.totalValue) * 1000) / 10 : 0
      }))
      .sort((a, b) => b.totalValue - a.totalValue);
  }

  // 7. Payment / Status Breakdown
  let paymentBreakdown = [];
  if (paymentDimCol) {
    const payMap = {};
    rows.forEach(r => {
      const rawPay = r[paymentDimCol.name];
      const payKey = (rawPay !== null && rawPay !== undefined && String(rawPay).trim() !== '') ? String(rawPay).trim() : 'Unknown';
      const metricVal = primaryMetricCol ? parseNumber(r[primaryMetricCol.name]) : 1;

      if (!payMap[payKey]) {
        payMap[payKey] = { method: payKey, count: 0, totalAmount: 0 };
      }
      payMap[payKey].count++;
      payMap[payKey].totalAmount += metricVal;
    });

    paymentBreakdown = Object.values(payMap)
      .map(p => ({
        ...p,
        totalAmount: Math.round(p.totalAmount * 100) / 100,
        sharePct: summaryKPIs.totalValue > 0 ? Math.round((p.totalAmount / summaryKPIs.totalValue) * 1000) / 10 : 0
      }))
      .sort((a, b) => b.totalAmount - a.totalAmount);
  }

  // 8. Time Trend Aggregations (Chronological)
  let trendData = [];
  if (primaryDateCol) {
    const dateMap = {};
    rows.forEach(r => {
      const d = parseDate(r[primaryDateCol.name]);
      if (!d) return;
      const dateKey = d.toISOString().split('T')[0];
      const metricVal = primaryMetricCol ? parseNumber(r[primaryMetricCol.name]) : 1;

      if (!dateMap[dateKey]) {
        dateMap[dateKey] = { date: dateKey, volume: 0, totalValue: 0 };
      }
      dateMap[dateKey].volume++;
      dateMap[dateKey].totalValue += metricVal;
    });

    trendData = Object.values(dateMap)
      .map(t => ({
        ...t,
        totalValue: Math.round(t.totalValue * 100) / 100,
        volume: t.volume,
        avgTicket: t.volume > 0 ? Math.round((t.totalValue / t.volume) * 100) / 100 : 0
      }))
      .sort((a, b) => a.date.localeCompare(b.date));
  }

  // 9. Pareto 80/20 Distribution Calculation
  let cumulativeSum = 0;
  const paretoData = categoryBreakdown.map(c => {
    cumulativeSum += c.totalValue;
    const cumPct = summaryKPIs.totalValue > 0 ? Math.round((cumulativeSum / summaryKPIs.totalValue) * 1000) / 10 : 0;
    return {
      category: c.category,
      totalValue: c.totalValue,
      volume: c.volume,
      avgValue: c.avgValue,
      cumPct
    };
  });

  // 10. Anomaly & Outlier Watch
  const anomalies = [];
  if (primaryMetricCol) {
    const stdDev = primaryMetricCol.stats.stdDev || 0;
    const avg = primaryMetricCol.stats.avg || 0;
    const upperLimit = avg + (2.5 * stdDev);

    rows.forEach((r, idx) => {
      const val = parseNumber(r[primaryMetricCol.name]);
      if (val < 0) {
        anomalies.push({
          rowIndex: idx + 1,
          type: 'Negative Value',
          severity: 'Critical',
          description: `Row #${idx + 1} has negative value (${val}) in ${primaryMetricCol.name}`,
          data: r
        });
      } else if (val === 0 && anomalies.filter(a => a.type === 'Zero Amount').length < 5) {
        anomalies.push({
          rowIndex: idx + 1,
          type: 'Zero Amount',
          severity: 'Warning',
          description: `Row #${idx + 1} has zero ${primaryMetricCol.name}`,
          data: r
        });
      } else if (stdDev > 0 && val > upperLimit && anomalies.filter(a => a.type === 'Statistical Outlier').length < 6) {
        anomalies.push({
          rowIndex: idx + 1,
          type: 'Statistical Outlier',
          severity: 'Info',
          description: `High spike (${val}) exceeds 2.5x standard deviation limit (${Math.round(upperLimit)})`,
          data: r
        });
      }
    });
  }

  // 11. AI Executive Findings
  const insights = [];
  insights.push(`Analyzed **${detectedDomain}** with **${rows.length.toLocaleString()} valid records** and **${columns.length} columns**.`);

  if (primaryMetricCol && summaryKPIs.totalValue > 0) {
    insights.push(`Active Financial Metric: **${primaryMetricCol.name}** totaling **${formatCurrency(summaryKPIs.totalValue)}** (Average: **${formatCurrency(summaryKPIs.averageTicket)}**).`);
  }

  if (categoryBreakdown.length > 0) {
    const topCat = categoryBreakdown[0];
    insights.push(`Top **${primaryDimCol ? primaryDimCol.name : 'Category'}**: **${topCat.category}** leads with **${topCat.sharePct}%** market share (${formatCurrency(topCat.totalValue)} across ${topCat.volume.toLocaleString()} entries).`);

    let runningSum = 0;
    let paretoCount = 0;
    for (const cat of categoryBreakdown) {
      runningSum += cat.totalValue;
      paretoCount++;
      if (summaryKPIs.totalValue > 0 && runningSum >= summaryKPIs.totalValue * 0.8) break;
    }
    if (categoryBreakdown.length >= 3) {
      insights.push(`Pareto Rule (80/20): Top **${paretoCount} of ${categoryBreakdown.length}** categories drive **80%** of total business impact.`);
    }
  }

  if (paymentBreakdown.length > 0) {
    const topPay = paymentBreakdown[0];
    insights.push(`Highest Settlement Mode: **${topPay.method}** accounts for **${topPay.sharePct}%** of total volume.`);
  }

  if (summaryKPIs.zeroOrNegativeCount > 0) {
    insights.push(`⚠️ Audit Alert: Detected **${summaryKPIs.zeroOrNegativeCount} entries** with zero or negative amounts.`);
  }

  // 12. Specialized Doctor & Provider MIS Performance
  const doctorColProfile = columnProfiles.find(c => 
    ['doctor', 'primary_doctor', 'consulting_doctor', 'physician', 'surgeon', 'consultant'].some(kw => 
      c.name.toLowerCase().replace(/[^a-z0-9]/g, '_').includes(kw)
    )
  );

  const deptColProfile = columnProfiles.find(c => 
    ['department', 'dept', 'speciality'].some(kw => 
      c.name.toLowerCase().replace(/[^a-z0-9]/g, '_').includes(kw)
    )
  );

  const billedCol = columnProfiles.find(c => c.type === 'numeric' && ['billed_amount', 'billed', 'gross_amount', 'gross'].some(kw => c.name.toLowerCase().replace(/[^a-z0-9]/g, '_').includes(kw)));
  const discountCol = columnProfiles.find(c => c.type === 'numeric' && ['discount_amount', 'discount'].some(kw => c.name.toLowerCase().replace(/[^a-z0-9]/g, '_').includes(kw)));
  const paidCol = columnProfiles.find(c => c.type === 'numeric' && ['amount_paid', 'paid', 'patient_amount_paid', 'collected'].some(kw => c.name.toLowerCase().replace(/[^a-z0-9]/g, '_').includes(kw)));
  const pendingCol = columnProfiles.find(c => c.type === 'numeric' && ['amount_pending', 'pending', 'patient_amount_pending', 'insurance_pending'].some(kw => c.name.toLowerCase().replace(/[^a-z0-9]/g, '_').includes(kw)));

  let doctorPerformance = [];
  let doctorKPIs = null;

  if (doctorColProfile) {
    const docMap = {};
    rows.forEach(r => {
      const rawDoc = r[doctorColProfile.name];
      const docName = (rawDoc !== null && rawDoc !== undefined && String(rawDoc).trim() !== '') ? String(rawDoc).trim() : '(Unspecified Doctor)';
      const rawDept = deptColProfile ? r[deptColProfile.name] : '';
      const deptName = (rawDept !== null && rawDept !== undefined && String(rawDept).trim() !== '') ? String(rawDept).trim() : 'General';
      const netVal = primaryMetricCol ? parseNumber(r[primaryMetricCol.name]) : 1;
      const billedVal = billedCol ? parseNumber(r[billedCol.name]) : netVal;
      const discountVal = discountCol ? parseNumber(r[discountCol.name]) : 0;
      const paidVal = paidCol ? parseNumber(r[paidCol.name]) : 0;
      const pendingVal = pendingCol ? parseNumber(r[pendingCol.name]) : 0;

      if (!docMap[docName]) {
        docMap[docName] = {
          doctor: docName,
          deptCounts: {},
          patientVolume: 0,
          totalNet: 0,
          totalBilled: 0,
          totalDiscount: 0,
          totalPaid: 0,
          totalPending: 0
        };
      }
      docMap[docName].patientVolume++;
      docMap[docName].totalNet += netVal;
      docMap[docName].totalBilled += billedVal;
      docMap[docName].totalDiscount += discountVal;
      docMap[docName].totalPaid += paidVal;
      docMap[docName].totalPending += pendingVal;
      docMap[docName].deptCounts[deptName] = (docMap[docName].deptCounts[deptName] || 0) + 1;
    });

    const totalSystemNet = Object.values(docMap).reduce((acc, d) => acc + d.totalNet, 0);

    doctorPerformance = Object.values(docMap).map(d => {
      let topDept = 'General';
      let maxDCount = -1;
      Object.entries(d.deptCounts).forEach(([dep, cnt]) => {
        if (cnt > maxDCount) {
          maxDCount = cnt;
          topDept = dep;
        }
      });

      return {
        doctor: d.doctor,
        department: topDept,
        patientVolume: d.patientVolume,
        totalNet: Math.round(d.totalNet * 100) / 100,
        totalBilled: Math.round(d.totalBilled * 100) / 100,
        totalDiscount: Math.round(d.totalDiscount * 100) / 100,
        totalPaid: Math.round(d.totalPaid * 100) / 100,
        totalPending: Math.round(d.totalPending * 100) / 100,
        avgTicket: d.patientVolume > 0 ? Math.round((d.totalNet / d.patientVolume) * 100) / 100 : 0,
        sharePct: totalSystemNet > 0 ? Math.round((d.totalNet / totalSystemNet) * 1000) / 10 : 0
      };
    }).sort((a, b) => b.totalNet - a.totalNet);

    doctorPerformance.forEach((d, idx) => { d.rank = idx + 1; });

    const topRevDoc = doctorPerformance[0] || null;
    const topVolDoc = [...doctorPerformance].sort((a, b) => b.patientVolume - a.patientVolume)[0] || null;
    const topAvgDoc = [...doctorPerformance].filter(d => d.patientVolume >= 5).sort((a, b) => b.avgTicket - a.avgTicket)[0] || topRevDoc;

    doctorKPIs = {
      doctorColumnName: doctorColProfile.name,
      totalDoctorsCount: doctorPerformance.length,
      topRevenueDoctor: topRevDoc ? { name: topRevDoc.doctor, net: topRevDoc.totalNet, department: topRevDoc.department, patients: topRevDoc.patientVolume } : null,
      topVolumeDoctor: topVolDoc ? { name: topVolDoc.doctor, patients: topVolDoc.patientVolume, net: topVolDoc.totalNet, department: topVolDoc.department } : null,
      topAvgTicketDoctor: topAvgDoc ? { name: topAvgDoc.doctor, avgTicket: topAvgDoc.avgTicket, department: topAvgDoc.department, patients: topAvgDoc.patientVolume } : null
    };

    if (topRevDoc) {
      insights.unshift(`🏆 Top Doctor by Revenue: **${topRevDoc.doctor}** (${topRevDoc.department}) generated **${formatCurrency(topRevDoc.totalNet)}** across **${topRevDoc.patientVolume.toLocaleString()} patients** (${topRevDoc.sharePct}% share).`);
    }
  }

  // 13. Dynamic Available Dimensions Catalog
  const availableDimensions = columnProfiles
    .filter(c => c.type === 'category' || c.type === 'text')
    .filter(c => !isIdentifierColumn(c.name))
    .map(c => {
      let icon = '📊';
      const nm = c.name.toLowerCase();
      if (nm.includes('doctor') || nm.includes('physician') || nm.includes('surgeon')) icon = '👨‍⚕️';
      else if (nm.includes('dept') || nm.includes('department')) icon = '🏥';
      else if (nm.includes('ward') || nm.includes('room') || nm.includes('bed')) icon = '🏨';
      else if (nm.includes('panel') || nm.includes('insurance') || nm.includes('payer')) icon = '🛡️';
      else if (nm.includes('class') || nm.includes('purpose') || nm.includes('category')) icon = '📋';
      else if (nm.includes('pay') || nm.includes('status') || nm.includes('bill')) icon = '💳';
      else if (nm.includes('refer') || nm.includes('agent')) icon = '🤝';
      else if (nm.includes('city') || nm.includes('state') || nm.includes('dist')) icon = '📍';

      return {
        col: c.name,
        label: c.name.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()),
        uniqueCount: c.stats?.uniqueCount || 0,
        icon
      };
    })
    .sort((a, b) => {
      const aIsDoc = a.col.toLowerCase().includes('doctor') ? 0 : 1;
      const bIsDoc = b.col.toLowerCase().includes('doctor') ? 0 : 1;
      return aIsDoc - bIsDoc;
    });

  return {
    rowCount: rows.length,
    colCount: columns.length,
    detectedDomain,
    columns: columnProfiles,
    summaryKPIs,
    categoryBreakdown: categoryBreakdown.slice(0, 150),
    paretoData: paretoData.slice(0, 30),
    paymentBreakdown,
    trendData: trendData.slice(-60),
    anomalies: anomalies.slice(0, 20),
    doctorPerformance,
    doctorKPIs,
    availableDimensions,
    insights
  };
}

function formatCurrency(val) {
  if (val === null || val === undefined) return '0.00';
  return Number(val).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Dynamic Pivot Engine
function generatePivot(rows, rowDim, colDim, valCol, aggMethod = 'sum') {
  if (!rows || rows.length === 0 || !rowDim || !valCol) {
    return { rowKeys: [], colKeys: [], matrix: {}, totals: {} };
  }

  const rowKeySet = new Set();
  const colKeySet = new Set();
  const matrix = {};

  rows.forEach(r => {
    const rKey = String(r[rowDim] || '(Blank)').trim();
    const cKey = colDim ? String(r[colDim] || '(Blank)').trim() : 'Total';
    const val = parseNumber(r[valCol]);

    rowKeySet.add(rKey);
    colKeySet.add(cKey);

    if (!matrix[rKey]) matrix[rKey] = {};
    if (!matrix[rKey][cKey]) {
      matrix[rKey][cKey] = { sum: 0, count: 0, values: [] };
    }
    matrix[rKey][cKey].sum += val;
    matrix[rKey][cKey].count++;
    matrix[rKey][cKey].values.push(val);
  });

  const rowKeys = Array.from(rowKeySet).sort();
  const colKeys = Array.from(colKeySet).sort();

  const formattedMatrix = {};
  const rowTotals = {};
  const colTotals = {};
  let grandTotal = 0;

  rowKeys.forEach(rk => {
    formattedMatrix[rk] = {};
    rowTotals[rk] = 0;

    colKeys.forEach(ck => {
      if (!colTotals[ck]) colTotals[ck] = 0;
      const cell = matrix[rk] && matrix[rk][ck];
      let finalVal = 0;

      if (cell) {
        if (aggMethod === 'sum') finalVal = Math.round(cell.sum * 100) / 100;
        else if (aggMethod === 'count') finalVal = cell.count;
        else if (aggMethod === 'avg') finalVal = Math.round((cell.sum / cell.count) * 100) / 100;
        else if (aggMethod === 'max') finalVal = Math.max(...cell.values);
        else if (aggMethod === 'min') finalVal = Math.min(...cell.values);
      }

      formattedMatrix[rk][ck] = finalVal;
      rowTotals[rk] += finalVal;
      colTotals[ck] += finalVal;
      grandTotal += finalVal;
    });

    rowTotals[rk] = Math.round(rowTotals[rk] * 100) / 100;
  });

  colKeys.forEach(ck => {
    colTotals[ck] = Math.round(colTotals[ck] * 100) / 100;
  });

  return {
    rowKeys,
    colKeys,
    matrix: formattedMatrix,
    rowTotals,
    colTotals,
    grandTotal: Math.round(grandTotal * 100) / 100,
    aggMethod,
    rowDim,
    colDim,
    valCol
  };
}

/**
 * Power BI Custom Visual Engine
 * Supports multi-dimensional grouping, secondary legend breakdown, and metric aggregations.
 */
function buildCustomVisual(rows, { xAxis, yAxis, agg = 'sum', groupBy, topN = 0, sort = 'desc' }) {
  if (!rows || rows.length === 0 || !xAxis || !yAxis) {
    return { labels: [], datasets: [], totals: {} };
  }

  // If no groupBy (single series)
  if (!groupBy || groupBy === xAxis) {
    const map = {};
    rows.forEach(r => {
      const rawX = r[xAxis];
      const xKey = (rawX !== null && rawX !== undefined && String(rawX).trim() !== '') ? String(rawX).trim() : '(Blank)';
      const val = parseNumber(r[yAxis]);

      if (!map[xKey]) {
        map[xKey] = { sum: 0, count: 0, values: [], distinctSet: new Set() };
      }
      map[xKey].sum += val;
      map[xKey].count++;
      map[xKey].values.push(val);
      map[xKey].distinctSet.add(r[yAxis]);
    });

    let items = Object.entries(map).map(([k, stats]) => {
      let finalVal = 0;
      if (agg === 'sum') finalVal = Math.round(stats.sum * 100) / 100;
      else if (agg === 'count') finalVal = stats.count;
      else if (agg === 'avg') finalVal = Math.round((stats.sum / stats.count) * 100) / 100;
      else if (agg === 'max') finalVal = Math.max(...stats.values);
      else if (agg === 'min') finalVal = Math.min(...stats.values);
      else if (agg === 'distinct') finalVal = stats.distinctSet.size;

      return { key: k, value: finalVal, count: stats.count };
    });

    if (sort === 'desc') items.sort((a, b) => b.value - a.value);
    else if (sort === 'asc') items.sort((a, b) => a.value - b.value);
    else if (sort === 'alpha') items.sort((a, b) => a.key.localeCompare(b.key));

    if (topN > 0) items = items.slice(0, topN);

    return {
      labels: items.map(i => i.key),
      datasets: [{
        label: `${agg.toUpperCase()} of ${yAxis}`,
        data: items.map(i => i.value)
      }],
      items
    };
  }

  // Multi-series with Legend Grouping (e.g. X = Department, Legend = Payment Mode)
  const xKeySet = new Set();
  const groupKeySet = new Set();
  const matrix = {};

  rows.forEach(r => {
    const rawX = r[xAxis];
    const xKey = (rawX !== null && rawX !== undefined && String(rawX).trim() !== '') ? String(rawX).trim() : '(Blank)';
    const rawG = r[groupBy];
    const gKey = (rawG !== null && rawG !== undefined && String(rawG).trim() !== '') ? String(rawG).trim() : '(Blank)';
    const val = parseNumber(r[yAxis]);

    xKeySet.add(xKey);
    groupKeySet.add(gKey);

    if (!matrix[xKey]) matrix[xKey] = {};
    if (!matrix[xKey][gKey]) {
      matrix[xKey][gKey] = { sum: 0, count: 0, values: [] };
    }
    matrix[xKey][gKey].sum += val;
    matrix[xKey][gKey].count++;
    matrix[xKey][gKey].values.push(val);
  });

  const labels = Array.from(xKeySet).slice(0, topN > 0 ? topN : 30);
  const groupKeys = Array.from(groupKeySet).slice(0, 10);

  const datasets = groupKeys.map(gk => {
    const data = labels.map(lk => {
      const cell = matrix[lk] && matrix[lk][gk];
      if (!cell) return 0;
      if (agg === 'sum') return Math.round(cell.sum * 100) / 100;
      if (agg === 'count') return cell.count;
      if (agg === 'avg') return Math.round((cell.sum / cell.count) * 100) / 100;
      if (agg === 'max') return Math.max(...cell.values);
      if (agg === 'min') return Math.min(...cell.values);
      return Math.round(cell.sum * 100) / 100;
    });

    return {
      label: gk,
      data
    };
  });

  return { labels, datasets, groupKeys };
}

/**
 * Natural Language "Ask BI" Copilot
 * Understands queries like:
 * - "top 5 doctors by net amount"
 * - "department wise billed amount"
 * - "payment mode distribution"
 * - "average days admitted by department"
 */
function askBI(rows, columns, query) {
  if (!rows || rows.length === 0 || !query) {
    return { success: false, answer: "Please enter a question about your data." };
  }

  const cleanQ = query.toLowerCase().trim();

  // 1. Detect Top N
  let topN = 10;
  const topMatch = cleanQ.match(/top\s*(\d+)/i) || cleanQ.match(/first\s*(\d+)/i);
  if (topMatch) {
    topN = parseInt(topMatch[1], 10);
  }

  // 2. Detect Aggregation
  let agg = 'sum';
  if (cleanQ.includes('average') || cleanQ.includes('avg') || cleanQ.includes('mean')) agg = 'avg';
  else if (cleanQ.includes('count') || cleanQ.includes('how many') || cleanQ.includes('patients') || cleanQ.includes('visits')) agg = 'count';
  else if (cleanQ.includes('highest') || cleanQ.includes('max') || cleanQ.includes('peak')) agg = 'max';
  else if (cleanQ.includes('lowest') || cleanQ.includes('min')) agg = 'min';

  // 3. Separate Numeric (Metric) and Categorical (Dimension) columns from row data
  const sampleRows = rows.slice(0, 50);
  const detectedNumeric = [];
  const detectedDimensions = [];

  columns.forEach(col => {
    if (isIdentifierColumn(col)) {
      detectedDimensions.push(col);
      return;
    }
    let numCount = 0;
    let nonBlank = 0;
    for (const r of sampleRows) {
      const v = r[col];
      if (v !== null && v !== undefined && String(v).trim() !== '') {
        nonBlank++;
        const s = String(v).replace(/[$₹€£,\s]/g, '');
        if (!isNaN(parseFloat(s)) && isFinite(s)) {
          numCount++;
        }
      }
    }
    if (nonBlank > 0 && numCount / nonBlank >= 0.7) {
      detectedNumeric.push(col);
    } else {
      detectedDimensions.push(col);
    }
  });

  // 4. Match Metric Column (Y-Axis)
  let yAxis = null;
  // Specific keywords
  if (cleanQ.includes('billed') || cleanQ.includes('gross')) {
    yAxis = detectedNumeric.find(c => c.toLowerCase().includes('bill') || c.toLowerCase().includes('gross'));
  } else if (cleanQ.includes('concession') || cleanQ.includes('discount')) {
    yAxis = detectedNumeric.find(c => c.toLowerCase().includes('conc') || c.toLowerCase().includes('disc'));
  } else if (cleanQ.includes('net') || cleanQ.includes('collection')) {
    yAxis = detectedNumeric.find(c => c.toLowerCase().includes('net'));
  }

  if (!yAxis) {
    for (const c of detectedNumeric) {
      const colWords = c.toLowerCase().split(/[\s_]+/);
      if (colWords.some(w => w.length > 2 && cleanQ.includes(w))) {
        yAxis = c;
        break;
      }
    }
  }

  if (!yAxis) {
    const defaultMetrics = ['NET AMOUNT', 'net amount', 'BILLED AMOUNT', 'billed amount', 'Gross Amount', 'total_amount', 'Amount'];
    yAxis = detectedNumeric.find(c => defaultMetrics.includes(c)) || detectedNumeric[0] || columns[0];
  }

  // 5. Match Dimension Column (X-Axis)
  let xAxis = null;
  // Specific semantic keywords
  if (cleanQ.includes('doctor') || cleanQ.includes('dr') || cleanQ.includes('physician')) {
    xAxis = detectedDimensions.find(c => {
      const l = c.toLowerCase();
      return l.includes('doctor') || l.includes('physician') || l.includes('consult');
    });
  } else if (cleanQ.includes('department') || cleanQ.includes('dept') || cleanQ.includes('speciality')) {
    xAxis = detectedDimensions.find(c => {
      const l = c.toLowerCase();
      return l.includes('department') || l.includes('dept') || l.includes('speciality');
    });
  } else if (cleanQ.includes('payment') || cleanQ.includes('mode') || cleanQ.includes('method') || cleanQ.includes('channel')) {
    xAxis = detectedDimensions.find(c => {
      const l = c.toLowerCase();
      return l.includes('payment') || l.includes('mode') || l.includes('pay');
    });
  } else if (cleanQ.includes('status')) {
    xAxis = detectedDimensions.find(c => c.toLowerCase().includes('status'));
  } else if (cleanQ.includes('patient')) {
    xAxis = detectedDimensions.find(c => c.toLowerCase().includes('patient'));
  } else if (cleanQ.includes('gender') || cleanQ.includes('sex')) {
    xAxis = detectedDimensions.find(c => c.toLowerCase().includes('gender') || c.toLowerCase().includes('sex'));
  } else if (cleanQ.includes('city') || cleanQ.includes('state') || cleanQ.includes('region')) {
    xAxis = detectedDimensions.find(c => c.toLowerCase().includes('city') || c.toLowerCase().includes('state') || c.toLowerCase().includes('region'));
  }

  if (!xAxis) {
    const dimCandidates = detectedDimensions.filter(c => c !== yAxis);
    for (const c of dimCandidates) {
      const colWords = c.toLowerCase().split(/[\s_]+/);
      if (colWords.some(w => w.length > 2 && cleanQ.includes(w) && !['amount', 'total', 'net', 'gross', 'top'].includes(w))) {
        xAxis = c;
        break;
      }
    }
  }

  if (!xAxis) {
    const candidates = ['PRIMARY DOCTOR', 'DEPARTMENT', 'Consulting Doctor', 'DOCTOR', 'Department', 'PANEL GROUP', 'BILL STATUS', 'GENDER', 'CITY'];
    xAxis = detectedDimensions.find(c => candidates.includes(c)) || detectedDimensions[0] || columns[0];
  }

  // 6. Generate Visual Data
  const visual = buildCustomVisual(rows, { xAxis, yAxis, agg, topN, sort: 'desc' });

  // 6. Formulate Executive Natural Language Summary
  let summary = `Showing ${agg.toUpperCase()} of **${yAxis}** grouped by **${xAxis}** (Top ${topN}).`;
  if (visual.items && visual.items.length > 0) {
    const topItem = visual.items[0];
    summary += ` Leader is **${topItem.key}** with **${formatCurrency(topItem.value)}** (${topItem.count} entries).`;
  }

  return {
    success: true,
    query,
    xAxis,
    yAxis,
    agg,
    topN,
    chartType: visual.labels.length <= 6 ? 'doughnut' : 'bar',
    visual,
    answer: summary
  };
}

module.exports = {
  smartParseSheet,
  analyzeDataset,
  detectColumnType,
  parseNumber,
  parseDate,
  generatePivot,
  buildCustomVisual,
  askBI
};
