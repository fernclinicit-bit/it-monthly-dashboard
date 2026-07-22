import express from 'express';
import cors from 'cors';
import pg from 'pg';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';
import * as XLSX from 'xlsx/xlsx.mjs';
import * as fs from 'fs';
XLSX.set_fs(fs);

dotenv.config();

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json({ limit: '50mb' }));

const pool = new pg.Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

async function initDb() {
  try {
    const client = await pool.connect();
    try {
      console.log('Connected to PostgreSQL. Initializing tables...');
      
      await client.query(`
        CREATE TABLE IF NOT EXISTS monthly_data (
          month_key VARCHAR(50) PRIMARY KEY,
          month_name VARCHAR(255) NOT NULL,
          total_assets INTEGER NOT NULL,
          asset_value NUMERIC(15, 2) NOT NULL,
          assets_expiring INTEGER NOT NULL,
          assets_broken INTEGER NOT NULL,
          assets_lost INTEGER NOT NULL,
          assets_vacant INTEGER NOT NULL,
          tickets_count INTEGER NOT NULL,
          sla_percent NUMERIC(5, 2) NOT NULL,
          response_time INTEGER NOT NULL,
          resolution_time NUMERIC(5, 2) NOT NULL,
          csat NUMERIC(3, 1) NOT NULL,
          total_software INTEGER NOT NULL,
          licenses_in_use NUMERIC(10, 2) NOT NULL,
          licenses_vacant NUMERIC(10, 2) NOT NULL,
          software_cost NUMERIC(12, 2) NOT NULL,
          software_expiring INTEGER NOT NULL,
          backup_success NUMERIC(5, 2) NOT NULL,
          security_incidents INTEGER NOT NULL,
          antivirus_coverage NUMERIC(5, 2) NOT NULL,
          mfa_coverage NUMERIC(5, 2) NOT NULL,
          repair_count INTEGER NOT NULL,
          repair_cost NUMERIC(12, 2) NOT NULL,
          top_broken_devices JSONB NOT NULL,
          dept_costs JSONB NOT NULL,
          software_expiring_details JSONB NOT NULL,
          assets_expiring_details JSONB NOT NULL,
          ongoing_projects JSONB NOT NULL,
          recommendations JSONB NOT NULL
        )
      `);

      await client.query(`
        CREATE TABLE IF NOT EXISTS assets (
          sn INTEGER PRIMARY KEY,
          date VARCHAR(50) NOT NULL,
          user_name VARCHAR(255) NOT NULL,
          position VARCHAR(255) NOT NULL,
          item_type VARCHAR(255) NOT NULL,
          device_serial VARCHAR(255) NOT NULL,
          status VARCHAR(50) NOT NULL,
          notes TEXT
        )
      `);

      await client.query(`
        CREATE TABLE IF NOT EXISTS tickets (
          sn INTEGER PRIMARY KEY,
          month_key VARCHAR(50) NOT NULL,
          date VARCHAR(50) NOT NULL,
          complainant VARCHAR(255) NOT NULL,
          email VARCHAR(255) NOT NULL,
          anydesk VARCHAR(255) NOT NULL,
          issue TEXT NOT NULL,
          cause TEXT NOT NULL,
          duration VARCHAR(50) NOT NULL,
          responder VARCHAR(255) NOT NULL,
          status VARCHAR(50) NOT NULL,
          cost NUMERIC(12, 2) NOT NULL
        )
      `);

      console.log('Tables initialized successfully.');
    } finally {
      client.release();
    }
  } catch (err) {
    console.warn('WARNING: Failed to connect to PostgreSQL database. Server will start in offline API fallback mode.', err.message);
  }
}

app.get('/api/db-state', async (req, res) => {
  try {
    const monthsResult = await pool.query('SELECT * FROM monthly_data');
    const assetsResult = await pool.query('SELECT * FROM assets ORDER BY sn ASC');
    const ticketsResult = await pool.query('SELECT * FROM tickets ORDER BY sn ASC');

    const dataObj = {};
    monthsResult.rows.forEach(row => {
      dataObj[row.month_key] = {
        monthName: row.month_name,
        totalAssets: Number(row.total_assets),
        assetValue: Number(row.asset_value),
        assetsExpiring: Number(row.assets_expiring),
        assetsBroken: Number(row.assets_broken),
        assetsLost: Number(row.assets_lost),
        assetsVacant: Number(row.assets_vacant),
        ticketsCount: Number(row.tickets_count),
        slaPercent: Number(row.sla_percent),
        responseTime: Number(row.response_time),
        resolutionTime: Number(row.resolution_time),
        csat: Number(row.csat),
        totalSoftware: Number(row.total_software),
        licensesInUse: Number(row.licenses_in_use),
        licensesVacant: Number(row.licenses_vacant),
        softwareCost: Number(row.software_cost),
        softwareExpiring: Number(row.software_expiring),
        backupSuccess: Number(row.backup_success),
        securityIncidents: Number(row.security_incidents),
        antivirusCoverage: Number(row.antivirus_coverage),
        mfaCoverage: Number(row.mfa_coverage),
        repairCount: Number(row.repair_count),
        repairCost: Number(row.repair_cost),
        topBrokenDevices: row.top_broken_devices || [],
        deptCosts: row.dept_costs || {},
        softwareExpiringDetails: row.software_expiring_details || [],
        assetsExpiringDetails: row.assets_expiring_details || [],
        ongoingProjects: row.ongoing_projects || [],
        recommendations: row.recommendations || [],
        ticketsList: []
      };
    });

    ticketsResult.rows.forEach(row => {
      if (dataObj[row.month_key]) {
        dataObj[row.month_key].ticketsList.push({
          sn: Number(row.sn),
          date: row.date,
          complainant: row.complainant,
          email: row.email,
          anydesk: row.anydesk,
          issue: row.issue,
          cause: row.cause,
          duration: row.duration,
          responder: row.responder,
          status: row.status,
          cost: Number(row.cost)
        });
      }
    });

    const assetsList = assetsResult.rows.map(row => ({
      sn: Number(row.sn),
      date: row.date,
      user: row.user_name,
      position: row.position,
      itemType: row.item_type,
      deviceSerial: row.device_serial,
      status: row.status,
      notes: row.notes || ''
    }));

    res.json({ data: dataObj, assetsList });
  } catch (err) {
    console.error('Error fetching DB state:', err);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.get('/api/clinic-data', (req, res) => {
  try {
    const filePath = path.join(process.cwd(), 'Update', 'fern-clinic-customer-report.xls');
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: 'File not found' });
    }
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);
    res.json(data);
  } catch (err) {
    console.error('Error reading clinic data:', err);
    res.status(500).json({ error: 'Failed to read clinic data' });
  }
});

app.post('/api/sync-all', async (req, res) => {
  const { data, assetsList } = req.body;
  if (!data || !assetsList) {
    return res.status(400).json({ error: 'Missing data or assetsList' });
  }

  let client;
  try {
    client = await pool.connect();
    await client.query('BEGIN');

    await client.query('DELETE FROM monthly_data');
    await client.query('DELETE FROM assets');
    await client.query('DELETE FROM tickets');

    for (const asset of assetsList) {
      await client.query(`
        INSERT INTO assets (sn, date, user_name, position, item_type, device_serial, status, notes)
        VALUES ($1, $2, $3, $4, $5, $6, $7, $8)
      `, [
        Number(asset.sn),
        asset.date || '',
        asset.user || 'ส่วนกลาง',
        asset.position || '-',
        asset.itemType || '',
        asset.deviceSerial || '-',
        asset.status || 'ใช้งาน',
        asset.notes || ''
      ]);
    }

    for (const [monthKey, monthData] of Object.entries(data)) {
      await client.query(`
        INSERT INTO monthly_data (
          month_key, month_name, total_assets, asset_value, assets_expiring, assets_broken, assets_lost, assets_vacant,
          tickets_count, sla_percent, response_time, resolution_time, csat, total_software, licenses_in_use, licenses_vacant,
          software_cost, software_expiring, backup_success, security_incidents, antivirus_coverage, mfa_coverage,
          repair_count, repair_cost, top_broken_devices, dept_costs, software_expiring_details, assets_expiring_details,
          ongoing_projects, recommendations
        ) VALUES (
          $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21, $22, $23, $24, $25, $26, $27, $28, $29, $30
        )
      `, [
        monthKey,
        monthData.monthName,
        Number(monthData.totalAssets || 0),
        Number(monthData.assetValue || 0),
        Number(monthData.assetsExpiring || 0),
        Number(monthData.assetsBroken || 0),
        Number(monthData.assetsLost || 0),
        Number(monthData.assetsVacant || 0),
        Number(monthData.ticketsCount || 0),
        Number(monthData.slaPercent || 0),
        Number(monthData.responseTime || 0),
        Number(monthData.resolutionTime || 0),
        Number(monthData.csat || 0),
        Number(monthData.totalSoftware || 0),
        Number(monthData.licensesInUse || 0),
        Number(monthData.licensesVacant || 0),
        Number(monthData.softwareCost || 0),
        Number(monthData.softwareExpiring || 0),
        Number(monthData.backupSuccess || 0),
        Number(monthData.securityIncidents || 0),
        Number(monthData.antivirusCoverage || 0),
        Number(monthData.mfaCoverage || 0),
        Number(monthData.repairCount || 0),
        Number(monthData.repairCost || 0),
        JSON.stringify(monthData.topBrokenDevices || []),
        JSON.stringify(monthData.deptCosts || {}),
        JSON.stringify(monthData.softwareExpiringDetails || []),
        JSON.stringify(monthData.assetsExpiringDetails || []),
        JSON.stringify(monthData.ongoingProjects || []),
        JSON.stringify(monthData.recommendations || [])
      ]);

      const ticketsList = monthData.ticketsList || [];
      for (const ticket of ticketsList) {
        await client.query(`
          INSERT INTO tickets (sn, month_key, date, complainant, email, anydesk, issue, cause, duration, responder, status, cost)
          VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12)
        `, [
          Number(ticket.sn),
          monthKey,
          ticket.date || '',
          ticket.complainant || 'ไม่ระบุชื่อ',
          ticket.email || '-',
          ticket.anydesk || '-',
          ticket.issue || '',
          ticket.cause || '-',
          ticket.duration || '-',
          ticket.responder || '-',
          ticket.status || 'กำลังดำเนินการ',
          Number(ticket.cost || 0)
        ]);
      }
    }

    await client.query('COMMIT');
    res.json({ success: true, message: 'All state synchronized successfully' });
  } catch (err) {
    if (client) {
      await client.query('ROLLBACK');
    }
    console.error('Error syncing all state:', err);
    res.status(500).json({ error: 'Internal Server Error' });
  } finally {
    if (client) {
      client.release();
    }
  }
});

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

app.use(express.static(path.join(__dirname, '../dist')));

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, '../dist/index.html'));
});

initDb().then(() => {
  app.listen(PORT, () => {
    console.log(`Backend Server is running on port ${PORT}`);
  });
});
