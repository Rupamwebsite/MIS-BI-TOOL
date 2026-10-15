const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
require('dotenv').config();

const { initDB, getPool, dbConfig } = require('./db');
const { smartParseSheet, analyzeDataset, generatePivot, parseNumber, parseDate, buildCustomVisual, askBI } = require('./analyzer');
const { generateHospitalBillingSample, generateSalesERPSample } = require('./seedSamples');

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '100mb' }));
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

// Root route: redirect to overview.html
app.get('/', (req, res) => {
  res.redirect('/overview.html');
});
app.get('/index.html', (req, res) => {
  res.redirect('/overview.html');
});

// Static frontend
app.use(express.static(path.join(__dirname, 'public')));

// Upload configuration
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const storage = multer.memoryStorage();
const upload = multer({
  storage,
  limits: { fileSize: 100 * 1024 * 1024 } // 100MB
});

// Helper: Fast persist dataset to MySQL
async function persistDataset(pool, { name, filename, fileSize, primarySheet, sheetsInfo, detectedType, rows, columns, metricsSummary }) {
  const conn = await pool.getConnection();
  try {
    await conn.beginTransaction();

    // 1. Insert dataset record
    const [datasetResult] = await conn.query(
      `INSERT INTO datasets (name, filename, file_size, total_rows, total_columns, primary_sheet, sheets_info, detected_type, metrics_summary)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        name,
        filename,
        fileSize || 0,
        rows.length,
        columns.length,
        primarySheet,
        JSON.stringify(sheetsInfo || [primarySheet]),
        detectedType,
        JSON.stringify(metricsSummary)
      ]
    );

    const datasetId = datasetResult.insertId;

    // 2. Insert columns
    if (metricsSummary && metricsSummary.columns) {
      for (let i = 0; i < metricsSummary.columns.length; i++) {
        const col = metricsSummary.columns[i];
        await conn.query(
          `INSERT INTO dataset_columns (dataset_id, sheet_name, col_index, name, detected_type, sample_values, stats)
           VALUES (?, ?, ?, ?, ?, ?, ?)`,
          [
            datasetId,
            primarySheet,
            i,
            col.name,
            col.type,
            JSON.stringify(col.sampleValues || []),
            JSON.stringify(col.stats || {})
          ]
        );
      }
    }

    // 3. High-Speed Bulk Insert (Chunk size 1500 for lightning speed)
    const chunkSize = 1500;
    for (let i = 0; i < rows.length; i += chunkSize) {
      const chunk = rows.slice(i, i + chunkSize);
      const values = chunk.map((r, idx) => [datasetId, primarySheet, i + idx, JSON.stringify(r)]);
      await conn.query(
        `INSERT INTO dataset_rows (dataset_id, sheet_name, row_index, data) VALUES ?`,
        [values]
      );
    }

    await conn.commit();
    return datasetId;
  } catch (err) {
    await conn.rollback();
    throw err;
  } finally {
    conn.release();
  }
}

// ==========================================
// API ROUTES
// ==========================================

// 1. System Health & Database Stats
app.get('/api/system/status', async (req, res) => {
  try {
    const pool = getPool();
    const [dbRows] = await pool.query('SELECT DATABASE() as db, VERSION() as version');
    const [countRows] = await pool.query('SELECT COUNT(*) as totalDatasets FROM datasets');
    const [rowCounts] = await pool.query('SELECT COUNT(*) as totalRows FROM dataset_rows');

    res.json({
      success: true,
      database: dbRows[0].db,
      mysqlVersion: dbRows[0].version,
      totalDatasets: countRows[0].totalDatasets,
      totalRowsStored: rowCounts[0].totalRows,
      status: 'Online & Connected to MIS_DB'
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 2. High Speed Excel Upload with Smart Header Detection
app.post('/api/upload', upload.single('file'), async (req, res) => {
  const startTime = Date.now();
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'No file uploaded.' });
    }

    const pool = getPool();
    const originalName = req.file.originalname;
    const fileBuffer = req.file.buffer;

    // Fast read workbook with xlsx dense mode
    const workbook = xlsx.read(fileBuffer, { type: 'buffer', cellDates: true, dense: true });
    const sheetNames = workbook.SheetNames;
    if (sheetNames.length === 0) {
      return res.status(400).json({ success: false, error: 'The uploaded Excel file contains no sheets.' });
    }

    // Target sheet
    const targetSheetName = req.body.sheetName || sheetNames[0];
    const worksheet = workbook.Sheets[targetSheetName];
    if (!worksheet) {
      return res.status(400).json({ success: false, error: `Sheet "${targetSheetName}" not found.` });
    }

    // Smart Parse (auto-detect table headers, skip banners, exclude grand total rows)
    const { rows: rawRows, columns, headerRowIndex } = smartParseSheet(worksheet);
    if (rawRows.length === 0) {
      return res.status(400).json({ success: false, error: 'No data rows found in the sheet.' });
    }

    // Run Advanced ERP Analysis
    const analytics = analyzeDataset(rawRows, columns);
    const datasetName = req.body.datasetName || originalName.replace(/\.[^/.]+$/, '');

    // Persist to MySQL MIS_DB in high-speed batches
    const datasetId = await persistDataset(pool, {
      name: datasetName,
      filename: originalName,
      fileSize: req.file.size,
      primarySheet: targetSheetName,
      sheetsInfo: sheetNames,
      detectedType: analytics.detectedDomain,
      rows: rawRows,
      columns,
      metricsSummary: analytics
    });

    const duration = ((Date.now() - startTime) / 1000).toFixed(2);
    console.log(`[Upload Success] "${datasetName}" (${rawRows.length} rows, ${columns.length} cols) processed in ${duration}s`);

    res.json({
      success: true,
      datasetId,
      datasetName,
      sheetNames,
      primarySheet: targetSheetName,
      headerRowIndex,
      processingTimeSec: duration,
      analytics
    });
  } catch (err) {
    console.error('[Upload Error]', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// 3. List all datasets
app.get('/api/datasets', async (req, res) => {
  try {
    const pool = getPool();
    const [rows] = await pool.query(`
      SELECT id, name, filename, file_size, total_rows, total_columns, primary_sheet, 
             detected_type, uploaded_at, metrics_summary
      FROM datasets
      ORDER BY id DESC
    `);

    const formatted = rows.map(r => ({
      ...r,
      metrics_summary: typeof r.metrics_summary === 'string' ? JSON.parse(r.metrics_summary) : r.metrics_summary
    }));

    res.json({ success: true, datasets: formatted });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 4. Get specific dataset by ID
app.get('/api/datasets/:id', async (req, res) => {
  try {
    const pool = getPool();
    const [rows] = await pool.query('SELECT * FROM datasets WHERE id = ?', [req.params.id]);
    if (rows.length === 0) {
      return res.status(404).json({ success: false, error: 'Dataset not found.' });
    }

    const dataset = rows[0];
    const metricsSummary = typeof dataset.metrics_summary === 'string' ? JSON.parse(dataset.metrics_summary) : dataset.metrics_summary;
    const sheetsInfo = typeof dataset.sheets_info === 'string' ? JSON.parse(dataset.sheets_info) : dataset.sheets_info;

    res.json({
      success: true,
      dataset: {
        ...dataset,
        metrics_summary: metricsSummary,
        sheets_info: sheetsInfo
      }
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 5. Paginated, filterable Raw Data records
app.get('/api/datasets/:id/data', async (req, res) => {
  try {
    const pool = getPool();
    const datasetId = req.params.id;
    const page = parseInt(req.query.page) || 1;
    const limit = parseInt(req.query.limit) || 25;
    const offset = (page - 1) * limit;
    const search = req.query.search ? String(req.query.search).toLowerCase().trim() : '';

    const [rowResults] = await pool.query(
      'SELECT data FROM dataset_rows WHERE dataset_id = ? ORDER BY row_index ASC',
      [datasetId]
    );

    let allRows = rowResults.map(r => typeof r.data === 'string' ? JSON.parse(r.data) : r.data);

    // Apply global search filter
    if (search) {
      allRows = allRows.filter(row => {
        return Object.values(row).some(val => 
          val !== null && val !== undefined && String(val).toLowerCase().includes(search)
        );
      });
    }

    const totalFiltered = allRows.length;
    const paginatedRows = allRows.slice(offset, offset + limit);

    res.json({
      success: true,
      page,
      limit,
      totalRows: totalFiltered,
      totalPages: Math.ceil(totalFiltered / limit),
      rows: paginatedRows
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 6. Dynamic Pivot Aggregation
app.post('/api/datasets/:id/pivot', async (req, res) => {
  try {
    const pool = getPool();
    const datasetId = req.params.id;
    const { rowDim, colDim, valCol, aggMethod } = req.body;

    const [rowResults] = await pool.query(
      'SELECT data FROM dataset_rows WHERE dataset_id = ?',
      [datasetId]
    );

    const rows = rowResults.map(r => typeof r.data === 'string' ? JSON.parse(r.data) : r.data);
    const pivotResult = generatePivot(rows, rowDim, colDim, valCol, aggMethod || 'sum');

    res.json({ success: true, pivot: pivotResult });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Helper for case-insensitive and trimmed column value extraction
function getRowValue(row, targetCol) {
  if (!row || !targetCol) return null;
  if (row[targetCol] !== undefined) return row[targetCol];
  const targetKey = String(targetCol).toLowerCase().replace(/[^a-z0-9]/g, '');
  for (const k of Object.keys(row)) {
    if (k.toLowerCase().replace(/[^a-z0-9]/g, '') === targetKey) {
      return row[k];
    }
  }
  return null;
}

// 7. Advanced Multi-Column & Multi-Operator Live Filter Engine
app.post('/api/datasets/:id/filter', async (req, res) => {
  try {
    const pool = getPool();
    const datasetId = req.params.id;
    const { filters, multiFilters, dateRange, metricRange, metricCol, dimCol } = req.body;

    const [rowResults] = await pool.query(
      'SELECT data FROM dataset_rows WHERE dataset_id = ? ORDER BY row_index ASC',
      [datasetId]
    );

    let rows = rowResults.map(r => typeof r.data === 'string' ? JSON.parse(r.data) : r.data);
    const columns = rows.length > 0 ? Object.keys(rows[0]) : [];

    // 1. Basic Dimension / Category Filters
    if (filters && typeof filters === 'object') {
      Object.entries(filters).forEach(([col, val]) => {
        if (val && val !== 'All' && val !== '') {
          const targetVal = String(val).toLowerCase().trim();
          rows = rows.filter(r => {
            const cellVal = getRowValue(r, col);
            return cellVal !== null && cellVal !== undefined && String(cellVal).toLowerCase().trim() === targetVal;
          });
        }
      });
    }

    // 2. Advanced Multi-Filter Rules (e.g. [{ col, op, val }])
    if (Array.isArray(multiFilters) && multiFilters.length > 0) {
      multiFilters.forEach(f => {
        if (!f.col || f.val === undefined || String(f.val).trim() === '') return;
        const targetVal = String(f.val).toLowerCase().trim();

        rows = rows.filter(r => {
          const cellVal = getRowValue(r, f.col);
          if (cellVal === null || cellVal === undefined) return false;

          const numCell = parseNumber(cellVal);
          const numTarget = parseNumber(f.val);

          if (f.op === '=') {
            return String(cellVal).toLowerCase().trim() === targetVal;
          } else if (f.op === '!=') {
            return String(cellVal).toLowerCase().trim() !== targetVal;
          } else if (f.op === 'contains') {
            return String(cellVal).toLowerCase().includes(targetVal);
          } else if (f.op === '>') {
            return numCell > numTarget;
          } else if (f.op === '<') {
            return numCell < numTarget;
          } else if (f.op === '>=') {
            return numCell >= numTarget;
          } else if (f.op === '<=') {
            return numCell <= numTarget;
          }
          return true;
        });
      });
    }

    // 3. Date Range Filter
    let dateColName = dateRange?.col;
    if (!dateColName && (dateRange?.start || dateRange?.end)) {
      // Auto-detect first date column
      dateColName = columns.find(c => {
        const lower = c.toLowerCase();
        return lower.includes('date') || lower.includes('time') || lower.includes('day');
      });
    }

    if (dateColName && (dateRange?.start || dateRange?.end)) {
      rows = rows.filter(r => {
        const rawDate = getRowValue(r, dateColName);
        if (!rawDate) return false;
        const d = parseDate(rawDate);
        if (!d) return true;
        const dStr = d.toISOString().split('T')[0];
        if (dateRange.start && dStr < dateRange.start) return false;
        if (dateRange.end && dStr > dateRange.end) return false;
        return true;
      });
    }

    // 4. Numeric Metric Range (Min / Max)
    if (metricRange && metricRange.col) {
      if (metricRange.min !== undefined && metricRange.min !== '') {
        const minVal = parseNumber(metricRange.min);
        rows = rows.filter(r => parseNumber(getRowValue(r, metricRange.col)) >= minVal);
      }
      if (metricRange.max !== undefined && metricRange.max !== '') {
        const maxVal = parseNumber(metricRange.max);
        rows = rows.filter(r => parseNumber(getRowValue(r, metricRange.col)) <= maxVal);
      }
    }

    // Re-analyze filtered subset with metric and dimension overrides
    const filteredAnalytics = analyzeDataset(rows, columns, { metricCol, dimCol });

    res.json({
      success: true,
      filteredCount: rows.length,
      analytics: filteredAnalytics,
      filteredRows: rows.slice(0, 500)
    });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 8. Delete dataset
app.delete('/api/datasets/:id', async (req, res) => {
  try {
    const pool = getPool();
    await pool.query('DELETE FROM datasets WHERE id = ?', [req.params.id]);
    res.json({ success: true, message: 'Dataset and associated records deleted.' });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 9. Power BI Custom Visual Studio Builder
app.post('/api/datasets/:id/visual', async (req, res) => {
  try {
    const pool = getPool();
    const datasetId = req.params.id;
    const { xAxis, yAxis, agg, groupBy, topN, sort } = req.body;

    const [rowResults] = await pool.query(
      'SELECT data FROM dataset_rows WHERE dataset_id = ? ORDER BY row_index ASC',
      [datasetId]
    );

    const rows = rowResults.map(r => typeof r.data === 'string' ? JSON.parse(r.data) : r.data);
    const visual = buildCustomVisual(rows, { xAxis, yAxis, agg: agg || 'sum', groupBy, topN: parseInt(topN) || 0, sort: sort || 'desc' });

    res.json({ success: true, visual });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 10. Natural Language "Ask BI" Copilot
app.post('/api/datasets/:id/ask', async (req, res) => {
  try {
    const pool = getPool();
    const datasetId = req.params.id;
    const { query } = req.body;

    const [rowResults] = await pool.query(
      'SELECT data FROM dataset_rows WHERE dataset_id = ? ORDER BY row_index ASC',
      [datasetId]
    );

    const rows = rowResults.map(r => typeof r.data === 'string' ? JSON.parse(r.data) : r.data);
    const columns = rows.length > 0 ? Object.keys(rows[0]) : [];

    const answer = askBI(rows, columns, query);
    res.json({ success: true, ...answer });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// 9. Load Sample Datasets
app.post('/api/load-sample', async (req, res) => {
  try {
    const pool = getPool();
    const type = req.body.type || 'hospital';

    let sample;
    if (type === 'sales') {
      sample = generateSalesERPSample();
    } else {
      sample = generateHospitalBillingSample();
    }

    const analytics = analyzeDataset(sample.rows, sample.columns);

    const datasetId = await persistDataset(pool, {
      name: sample.name,
      filename: sample.filename,
      fileSize: 1024 * 128,
      primarySheet: 'Sheet1',
      sheetsInfo: ['Sheet1'],
      detectedType: sample.detectedType,
      rows: sample.rows,
      columns: sample.columns,
      metricsSummary: analytics
    });

    res.json({
      success: true,
      datasetId,
      datasetName: sample.name,
      analytics
    });
  } catch (err) {
    console.error('[Sample Load Error]', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// Start Server
async function startServer() {
  try {
    await initDB();
    app.listen(PORT, () => {
      console.log(`=======================================================`);
      console.log(`🚀 ERP MIS ANALYTICS PLATFORM v2.0 RUNNING`);
      console.log(`📡 URL: http://localhost:${PORT}`);
      console.log(`🗄️ Database: ${dbConfig.database} on MySQL (User: ${dbConfig.user})`);
      console.log(`=======================================================`);
    });
  } catch (err) {
    console.error('Failed to start server:', err.message);
    process.exit(1);
  }
}

startServer();
