const mysql = require('mysql2/promise');
require('dotenv').config();

const dbConfig = {
  host: process.env.DB_HOST || 'localhost',
  user: process.env.DB_USER || 'root',
  password: process.env.DB_PASSWORD || 'Rupam@123',
  database: process.env.DB_NAME || 'MIS_DB',
  waitForConnections: true,
  connectionLimit: 15,
  queueLimit: 0,
  decimalNumbers: true
};

let pool;

async function initDB() {
  try {
    // 1. Connect without database first to ensure MIS_DB exists
    const rootConn = await mysql.createConnection({
      host: dbConfig.host,
      user: dbConfig.user,
      password: dbConfig.password
    });

    await rootConn.query(`CREATE DATABASE IF NOT EXISTS \`${dbConfig.database}\` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci`);
    await rootConn.end();

    // 2. Initialize connection pool
    pool = mysql.createPool(dbConfig);
    console.log(`[Database] Connected to MySQL database "${dbConfig.database}" successfully.`);

    // 3. Create required schema tables
    await createTables();
    return pool;
  } catch (err) {
    console.error('[Database] Connection Error:', err.message);
    throw err;
  }
}

async function createTables() {
  const conn = await pool.getConnection();
  try {
    // Datasets table
    await conn.query(`
      CREATE TABLE IF NOT EXISTS datasets (
        id INT AUTO_INCREMENT PRIMARY KEY,
        name VARCHAR(255) NOT NULL,
        filename VARCHAR(255),
        file_size INT DEFAULT 0,
        total_rows INT DEFAULT 0,
        total_columns INT DEFAULT 0,
        primary_sheet VARCHAR(255) DEFAULT 'Sheet1',
        sheets_info JSON,
        detected_type VARCHAR(100) DEFAULT 'General ERP',
        metrics_summary JSON,
        uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
    `);

    // Column definitions table
    await conn.query(`
      CREATE TABLE IF NOT EXISTS dataset_columns (
        id INT AUTO_INCREMENT PRIMARY KEY,
        dataset_id INT NOT NULL,
        sheet_name VARCHAR(255) DEFAULT 'Sheet1',
        col_index INT NOT NULL,
        name VARCHAR(255) NOT NULL,
        detected_type VARCHAR(50) DEFAULT 'text',
        sample_values JSON,
        stats JSON,
        FOREIGN KEY (dataset_id) REFERENCES datasets(id) ON DELETE CASCADE
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
    `);

    // Rows table
    await conn.query(`
      CREATE TABLE IF NOT EXISTS dataset_rows (
        id INT AUTO_INCREMENT PRIMARY KEY,
        dataset_id INT NOT NULL,
        sheet_name VARCHAR(255) DEFAULT 'Sheet1',
        row_index INT NOT NULL,
        data JSON NOT NULL,
        FOREIGN KEY (dataset_id) REFERENCES datasets(id) ON DELETE CASCADE,
        INDEX idx_dataset_sheet (dataset_id, sheet_name)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
    `);

    // Saved Reports table
    await conn.query(`
      CREATE TABLE IF NOT EXISTS saved_reports (
        id INT AUTO_INCREMENT PRIMARY KEY,
        dataset_id INT NOT NULL,
        title VARCHAR(255) NOT NULL,
        description TEXT,
        filters JSON,
        config JSON,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (dataset_id) REFERENCES datasets(id) ON DELETE CASCADE
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
    `);

    // System activity / audit log
    await conn.query(`
      CREATE TABLE IF NOT EXISTS activity_logs (
        id INT AUTO_INCREMENT PRIMARY KEY,
        action VARCHAR(100) NOT NULL,
        details JSON,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
    `);

    console.log('[Database] Schema tables verified and ready.');
  } finally {
    conn.release();
  }
}

function getPool() {
  if (!pool) {
    throw new Error('Database pool has not been initialized. Call initDB() first.');
  }
  return pool;
}

module.exports = {
  initDB,
  getPool,
  dbConfig
};
