import express from 'express';
import cors from 'cors';
import pg from 'pg';
import dotenv from 'dotenv';
import path from 'path';
import crypto from 'crypto';
import { fileURLToPath } from 'url';
import * as XLSX from 'xlsx/xlsx.mjs';
import * as fs from 'fs';
XLSX.set_fs(fs);

dotenv.config();

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(express.json({ limit: '50mb' }));

app.get('/health', (_req, res) => {
  res.status(200).json({ status: 'ok' });
});

const ADMIN_PASSWORD_HASH = process.env.ADMIN_PASSWORD_HASH ||
  '1e630fe2c4c6fecd9f5181b3bd43242407c8efa7e6e7db16204dc447257224db';

app.post('/api/admin/verify', (req, res) => {
  const password = String(req.body?.password || '');
  const suppliedHash = crypto.createHash('sha256').update(password).digest('hex');
  const suppliedBuffer = Buffer.from(suppliedHash, 'hex');
  const expectedBuffer = Buffer.from(ADMIN_PASSWORD_HASH, 'hex');
  const valid = suppliedBuffer.length === expectedBuffer.length &&
    crypto.timingSafeEqual(suppliedBuffer, expectedBuffer);
  res.status(valid ? 200 : 401).json({ valid });
});

const pool = new pg.Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

async function refreshOperationalCounters(client) {
  await client.query(`
    UPDATE monthly_data m
    SET total_assets = (SELECT COUNT(*) FROM assets),
        assets_broken = (SELECT COUNT(*) FROM assets WHERE status = 'รอซ่อม'),
        assets_lost = (SELECT COUNT(*) FROM assets WHERE status = 'สูญหาย'),
        assets_vacant = (SELECT COUNT(*) FROM assets WHERE status = 'ว่าง'),
        tickets_count = (SELECT COUNT(*) FROM tickets t WHERE t.month_key = m.month_key),
        repair_count = (SELECT COUNT(*) FROM tickets t WHERE t.month_key = m.month_key)
  `);
}

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
          notes TEXT,
          details JSONB NOT NULL DEFAULT '{}'::jsonb
        )
      `);
      await client.query(`ALTER TABLE assets ADD COLUMN IF NOT EXISTS details JSONB NOT NULL DEFAULT '{}'::jsonb`);

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
          cost NUMERIC(12, 2) NOT NULL,
          source VARCHAR(50) NOT NULL DEFAULT 'dashboard'
        )
      `);
      await client.query(`ALTER TABLE tickets ADD COLUMN IF NOT EXISTS source VARCHAR(50) NOT NULL DEFAULT 'dashboard'`);

      await client.query(`
        CREATE TABLE IF NOT EXISTS asset_requests (
          id SERIAL PRIMARY KEY,
          requester VARCHAR(255) NOT NULL,
          department VARCHAR(255) NOT NULL,
          item_type VARCHAR(255) NOT NULL,
          purpose TEXT NOT NULL,
          requested_date DATE NOT NULL DEFAULT CURRENT_DATE,
          due_date DATE,
          status VARCHAR(50) NOT NULL DEFAULT 'pending',
          assigned_asset_sn INTEGER,
          reviewer VARCHAR(255),
          issue_date TIMESTAMPTZ,
          return_date TIMESTAMPTZ,
          return_condition VARCHAR(50),
          notes TEXT,
          history JSONB NOT NULL DEFAULT '[]'::jsonb,
          created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
          updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
        )
      `);

      // Workflow is authoritative for assets that are reserved or currently issued.
      // Reconcile on every deploy so stale dashboard snapshots cannot leave them behind.
      await client.query(`
        UPDATE assets a
        SET status = CASE WHEN r.status = 'approved' THEN 'จอง' ELSE 'ใช้งาน' END,
            user_name = r.requester,
            details = a.details || jsonb_build_object(
              'status', CASE WHEN r.status = 'approved' THEN 'จอง' ELSE 'ใช้งาน' END,
              'user', r.requester
            )
        FROM asset_requests r
        WHERE r.assigned_asset_sn = a.sn
          AND r.status IN ('approved', 'issued', 'overdue')
      `);

      // An active request must never point to a deleted asset. Return a dangling
      // request to the approval queue instead of showing conflicting statuses.
      await client.query(`
        UPDATE asset_requests r
        SET status = 'need_info',
            assigned_asset_sn = NULL,
            history = r.history || jsonb_build_array(jsonb_build_object(
              'status', 'need_info',
              'at', NOW(),
              'by', 'System',
              'note', 'ยกเลิกการจัดสรรอัตโนมัติ เนื่องจากไม่พบอุปกรณ์ในทะเบียน'
            )),
            updated_at = NOW()
        WHERE r.status IN ('approved', 'issued', 'overdue')
          AND NOT EXISTS (SELECT 1 FROM assets a WHERE a.sn = r.assigned_asset_sn)
      `);

      await refreshOperationalCounters(client);
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
          cost: Number(row.cost),
          source: row.source
        });
      }
    });

    const assetsList = assetsResult.rows.map(row => ({
      ...(row.details || {}),
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

app.post('/api/tickets', async (req, res) => {
  const { name, department, date, deviceType, issue, priority, assetSerial, email, anydesk } = req.body || {};
  if (!name || !department || !date || !deviceType || !issue) {
    return res.status(400).json({ error: 'กรุณากรอกข้อมูลแจ้งปัญหาให้ครบถ้วน' });
  }

  const monthKey = String(date).slice(0, 7);
  let client;
  try {
    client = await pool.connect();
    await client.query('BEGIN');

    const monthResult = await client.query(
      'SELECT month_key FROM monthly_data WHERE month_key = $1 FOR UPDATE',
      [monthKey]
    );
    if (monthResult.rowCount === 0) {
      throw new Error(`ยังไม่มีข้อมูล Dashboard สำหรับเดือน ${monthKey}`);
    }

    await client.query('SELECT pg_advisory_xact_lock(42001)');
    const snResult = await client.query('SELECT COALESCE(MAX(sn), 0) + 1 AS next_sn FROM tickets');
    const nextSn = Number(snResult.rows[0].next_sn);
    const complainant = `${String(name).trim()} (${String(department).trim()})`;
    const ticketIssue = `[${String(deviceType).trim()}] [${priority || 'medium'}] ${String(issue).trim()}`;

    await client.query(`
      INSERT INTO tickets (
        sn, month_key, date, complainant, email, anydesk, issue, cause,
        duration, responder, status, cost, source
      ) VALUES ($1, $2, $3, $4, $5, $6, $7, '-', '-', '-', 'กำลังดำเนินการ', 0, 'request_form')
    `, [nextSn, monthKey, date, complainant, email || '-', anydesk || '-', ticketIssue]);

    let linkedAssetSn = null;
    if (String(assetSerial || '').trim()) {
      const assetResult = await client.query(
        'SELECT sn, details FROM assets WHERE LOWER(device_serial) = LOWER($1) FOR UPDATE',
        [String(assetSerial).trim()]
      );
      if (assetResult.rowCount === 0) {
        throw new Error(`ไม่พบหมายเลขเครื่อง ${String(assetSerial).trim()} ในทะเบียนทรัพย์สิน`);
      }
      linkedAssetSn = Number(assetResult.rows[0].sn);
      await client.query(`
        UPDATE assets
        SET status = 'รอซ่อม',
            details = details || jsonb_build_object('status', 'รอซ่อม')
        WHERE sn = $1
      `, [linkedAssetSn]);

      const activeRequests = await client.query(`
        SELECT id, history FROM asset_requests
        WHERE assigned_asset_sn = $1 AND status IN ('approved', 'issued', 'overdue')
        FOR UPDATE
      `, [linkedAssetSn]);
      for (const request of activeRequests.rows) {
        const event = {
          status: 'returned',
          at: new Date().toISOString(),
          by: String(name).trim(),
          note: `ส่งเครื่องเข้าซ่อมจากใบแจ้งปัญหา #${nextSn}`
        };
        await client.query(`
          UPDATE asset_requests
          SET status = 'returned',
              return_date = NOW(),
              return_condition = 'ชำรุด',
              history = $1,
              updated_at = NOW()
          WHERE id = $2
        `, [JSON.stringify([...(request.history || []), event]), request.id]);
      }
    }

    await refreshOperationalCounters(client);

    await client.query('COMMIT');
    res.status(201).json({ success: true, sn: nextSn, monthKey, linkedAssetSn });
  } catch (err) {
    if (client) await client.query('ROLLBACK');
    console.error('Error creating IT request ticket:', err);
    const isValidationError = err.message.startsWith('ยังไม่มีข้อมูล') || err.message.startsWith('ไม่พบหมายเลขเครื่อง');
    res.status(isValidationError ? 400 : 500).json({ error: err.message || 'บันทึกคำร้องไม่สำเร็จ' });
  } finally {
    if (client) client.release();
  }
});

app.patch('/api/assets/:sn', async (req, res) => {
  const sn = Number(req.params.sn);
  const { user, position, itemType, additionalEquipment, deviceSerial, status, notes } = req.body || {};
  if (!Number.isInteger(sn) || sn <= 0 || !itemType || !status) {
    return res.status(400).json({ error: 'ข้อมูลทรัพย์สินไม่ถูกต้อง' });
  }

  let client;
  try {
    client = await pool.connect();
    await client.query('BEGIN');

    const assetResult = await client.query('SELECT * FROM assets WHERE sn = $1 FOR UPDATE', [sn]);
    if (assetResult.rowCount === 0) throw new Error('ไม่พบทรัพย์สินที่ต้องการแก้ไข');

    const details = {
      ...(assetResult.rows[0].details || {}),
      user: user || 'ส่วนกลาง',
      position: position || '-',
      itemType,
      additionalEquipment: additionalEquipment || '',
      deviceSerial: deviceSerial || '-',
      status,
      notes: notes || ''
    };
    await client.query(`
      UPDATE assets
      SET user_name = $1, position = $2, item_type = $3, device_serial = $4,
          status = $5, notes = $6, details = $7
      WHERE sn = $8
    `, [
      user || 'ส่วนกลาง',
      position || '-',
      itemType,
      deviceSerial || '-',
      status,
      notes || '',
      JSON.stringify(details),
      sn
    ]);

    if (['ว่าง', 'รอซ่อม', 'สูญหาย'].includes(status)) {
      const activeRequests = await client.query(`
        SELECT id, history FROM asset_requests
        WHERE assigned_asset_sn = $1 AND status IN ('approved', 'issued', 'overdue')
        FOR UPDATE
      `, [sn]);
      for (const request of activeRequests.rows) {
        const event = {
          status: 'returned',
          at: new Date().toISOString(),
          by: 'IT Asset Registry',
          note: `เปลี่ยนสถานะทรัพย์สินเป็น ${status} จากหน้าทะเบียน`
        };
        await client.query(`
          UPDATE asset_requests
          SET status = 'returned',
              return_date = NOW(),
              return_condition = $1,
              history = $2,
              updated_at = NOW()
          WHERE id = $3
        `, [
          status === 'รอซ่อม' ? 'ชำรุด' : status === 'สูญหาย' ? 'สูญหาย' : 'ปกติ',
          JSON.stringify([...(request.history || []), event]),
          request.id
        ]);
      }
    }

    await client.query(`
      UPDATE monthly_data
      SET total_assets = (SELECT COUNT(*) FROM assets),
          assets_broken = (SELECT COUNT(*) FROM assets WHERE status = 'รอซ่อม'),
          assets_lost = (SELECT COUNT(*) FROM assets WHERE status = 'สูญหาย'),
          assets_vacant = (SELECT COUNT(*) FROM assets WHERE status = 'ว่าง')
    `);

    await client.query('COMMIT');
    res.json({ success: true, sn, status });
  } catch (err) {
    if (client) await client.query('ROLLBACK');
    console.error('Error updating asset directly:', err);
    res.status(err.message.startsWith('ไม่พบ') ? 404 : 500).json({ error: err.message || 'แก้ไขทรัพย์สินไม่สำเร็จ' });
  } finally {
    if (client) client.release();
  }
});

app.get('/api/asset-requests', async (_req, res) => {
  try {
    const result = await pool.query(`
      SELECT r.*, a.device_serial, a.item_type AS assigned_item_type
      FROM asset_requests r
      LEFT JOIN assets a ON a.sn = r.assigned_asset_sn
      ORDER BY r.created_at DESC, r.id DESC
    `);
    res.json(result.rows.map(row => ({
      ...row,
      id: Number(row.id),
      assigned_asset_sn: row.assigned_asset_sn === null ? null : Number(row.assigned_asset_sn)
    })));
  } catch (err) {
    console.error('Error fetching asset requests:', err);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.post('/api/asset-requests', async (req, res) => {
  const { requester, department, itemType, purpose, dueDate, notes } = req.body;
  if (!requester || !department || !itemType || !purpose) {
    return res.status(400).json({ error: 'กรุณากรอกข้อมูลคำขอให้ครบ' });
  }
  try {
    const event = { status: 'pending', at: new Date().toISOString(), by: requester, note: 'ส่งคำขอเบิกอุปกรณ์' };
    const result = await pool.query(`
      INSERT INTO asset_requests (requester, department, item_type, purpose, due_date, notes, history)
      VALUES ($1, $2, $3, $4, $5, $6, $7)
      RETURNING *
    `, [requester, department, itemType, purpose, dueDate || null, notes || '', JSON.stringify([event])]);
    res.status(201).json(result.rows[0]);
  } catch (err) {
    console.error('Error creating asset request:', err);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.patch('/api/asset-requests/:id/action', async (req, res) => {
  const id = Number(req.params.id);
  const { action, reviewer, assetSn, condition, note } = req.body;
  const allowedActions = ['approve', 'reject', 'issue', 'return'];
  if (!Number.isInteger(id) || !allowedActions.includes(action)) {
    return res.status(400).json({ error: 'คำสั่งไม่ถูกต้อง' });
  }

  let client;
  try {
    client = await pool.connect();
    await client.query('BEGIN');
    const requestResult = await client.query('SELECT * FROM asset_requests WHERE id = $1 FOR UPDATE', [id]);
    if (requestResult.rowCount === 0) {
      await client.query('ROLLBACK');
      return res.status(404).json({ error: 'ไม่พบคำขอ' });
    }
    const request = requestResult.rows[0];
    let nextStatus;
    let assignedSn = request.assigned_asset_sn;
    let assetStatus = null;

    if (action === 'approve') {
      if (request.status !== 'pending' && request.status !== 'need_info') throw new Error('คำขอนี้ไม่อยู่ในสถานะรออนุมัติ');
      assignedSn = Number(assetSn);
      const assetResult = await client.query('SELECT * FROM assets WHERE sn = $1 FOR UPDATE', [assignedSn]);
      if (assetResult.rowCount === 0) throw new Error('ไม่พบอุปกรณ์ที่เลือก');
      if (assetResult.rows[0].status !== 'ว่าง') throw new Error('อุปกรณ์นี้ไม่ว่างหรือถูกจองแล้ว');
      nextStatus = 'approved';
      assetStatus = 'จอง';
    } else if (action === 'reject') {
      if (request.status !== 'pending' && request.status !== 'need_info') throw new Error('คำขอนี้ไม่สามารถปฏิเสธได้');
      nextStatus = 'rejected';
    } else if (action === 'issue') {
      if (request.status !== 'approved') throw new Error('ต้องอนุมัติคำขอก่อนส่งมอบ');
      nextStatus = 'issued';
      assetStatus = 'ใช้งาน';
    } else {
      if (request.status !== 'issued' && request.status !== 'overdue') throw new Error('รายการนี้ยังไม่ได้ส่งมอบ');
      nextStatus = 'returned';
      assetStatus = condition === 'ชำรุด' ? 'รอซ่อม' : condition === 'สูญหาย' ? 'สูญหาย' : 'ว่าง';
    }

    if (assetStatus && assignedSn) {
      await client.query(`
        UPDATE assets
        SET status = $1::text,
            user_name = CASE WHEN $1::text = 'ว่าง' THEN 'ส่วนกลาง' ELSE $2::text END,
            details = details || jsonb_build_object('status', $1::text, 'user', CASE WHEN $1::text = 'ว่าง' THEN 'ส่วนกลาง' ELSE $2::text END)
        WHERE sn = $3
      `, [assetStatus, request.requester, assignedSn]);
    }

    if (action === 'return' && condition === 'ชำรุด' && assignedSn) {
      const monthResult = await client.query(`
        SELECT month_key
        FROM monthly_data
        ORDER BY (month_key = TO_CHAR(CURRENT_DATE, 'YYYY-MM')) DESC, month_key DESC
        LIMIT 1
      `);
      if (monthResult.rowCount > 0) {
        const repairMonthKey = monthResult.rows[0].month_key;
        await client.query('SELECT pg_advisory_xact_lock(42001)');
        const snResult = await client.query('SELECT COALESCE(MAX(sn), 0) + 1 AS next_sn FROM tickets');
        const nextTicketSn = Number(snResult.rows[0].next_sn);
        const assetResult = await client.query('SELECT device_serial, item_type FROM assets WHERE sn = $1', [assignedSn]);
        const asset = assetResult.rows[0] || {};
        await client.query(`
          INSERT INTO tickets (
            sn, month_key, date, complainant, email, anydesk, issue, cause,
            duration, responder, status, cost, source
          ) VALUES ($1, $2, TO_CHAR(NOW(), 'YYYY-MM-DD HH24:MI'), $3, '-', '-', $4, '-',
                    '-', $5, 'กำลังดำเนินการ', 0, 'asset_workflow')
        `, [
          nextTicketSn,
          repairMonthKey,
          `${request.requester} (${request.department})`,
          `[${asset.item_type || request.item_type}] เครื่อง ${asset.device_serial || assignedSn} ชำรุดจากการรับคืน`,
          reviewer || 'IT'
        ]);
      }
    }

    const event = { status: nextStatus, at: new Date().toISOString(), by: reviewer || request.requester, note: note || '' };
    const updated = await client.query(`
      UPDATE asset_requests
      SET status = $1::text,
          assigned_asset_sn = $2,
          reviewer = COALESCE($3, reviewer),
          issue_date = CASE WHEN $1::text = 'issued' THEN NOW() ELSE issue_date END,
          return_date = CASE WHEN $1::text = 'returned' THEN NOW() ELSE return_date END,
          return_condition = CASE WHEN $1::text = 'returned' THEN $4::text ELSE return_condition END,
          notes = CASE WHEN $5 <> '' THEN CONCAT_WS(E'\n', NULLIF(notes, ''), $5) ELSE notes END,
          history = history || $6::jsonb,
          updated_at = NOW()
      WHERE id = $7
      RETURNING *
    `, [nextStatus, assignedSn, reviewer || null, condition || null, note || '', JSON.stringify([event]), id]);
    await refreshOperationalCounters(client);
    await client.query('COMMIT');
    res.json(updated.rows[0]);
  } catch (err) {
    if (client) await client.query('ROLLBACK');
    console.error('Error updating asset request:', err);
    res.status(400).json({ error: err.message || 'ไม่สามารถดำเนินการได้' });
  } finally {
    if (client) client.release();
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

app.get('/api/inventory-data', (req, res) => {
  try {
    const filePath = path.join(process.cwd(), 'Update', 'Inventory_Inventory_Results.xlsx');
    if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'Inventory file not found' });
    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const formatDate = (value) => {
      if (!value) return '';
      const d = value instanceof Date ? value : new Date(value);
      if (Number.isNaN(d.getTime())) return String(value);
      return new Intl.DateTimeFormat('th-TH', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(d);
    };
    res.json(rows.map((row, i) => ({
      sn: row['Nember'] || i + 1,
      submittedOn: formatDate(row['Submitted on']),
      respondent: row['Respondents'],
      date: formatDate(row['วันที่เบิกใช้งาน']),
      user: row['บุคคลเบิกใช้อุปกรณ์'] || 'ส่วนกลาง/ไม่ระบุ',
      position: row['ตำแหน่ง'] || '-',
      itemType: row['รายการอุปกรณ์หลัก'] || 'อุปกรณ์เสริม/อื่นๆ',
      additionalEquipment: row['อุปกรณ์เพิ่มเติมที่ต้องการเบิก'],
      softwareApp: row['ซอต์ฟแวร์/ App'],
      registeredEmail: row['เมลที่ลงทะเบียน'],
      deviceSerial: row['หมายเลขอุปกรณ์ (เช่น  Ipad 016)'] || '-',
      additionalSerial: row['หมายเลขอุปกรณ์ เพิ่มเติม  (เช่น  สาย อะเเดปเตอร์ ipad-011))'],
      returnDueDate: formatDate(row['กำหนดคืนอุปกรณ์']),
      status: row['สถานะ'] || 'ใช้งาน',
      notes: row['หมายเหตุ'],
      inspectionDate: formatDate(row['วันที่ตรวจสอบ']),
      purchaseDate: formatDate(row['วันที่ซื้อ']),
      warrantyEndDate: formatDate(row['วันหมดประกัน']),
      expense: Number(row['ค่าใช้จ่าย']) || 0
    })));
  } catch (err) {
    console.error('Error reading inventory data:', err);
    res.status(500).json({ error: 'Failed to read inventory data' });
  }
});

app.post('/api/sync-all', async (req, res) => {
  const { data, assetsList } = req.body;
  if (!data || !assetsList) {
    return res.status(400).json({ error: 'Missing data or assetsList' });
  }

  let client;
  try {
    const finiteNumber = (value, fallback = 0) => {
      const parsed = Number(value);
      return Number.isFinite(parsed) ? parsed : fallback;
    };
    const integerNumber = (value, fallback = 0) => Math.trunc(finiteNumber(value, fallback));

    client = await pool.connect();
    await client.query('BEGIN');

    await client.query('DELETE FROM monthly_data');
    // Tickets created by API workflows are authoritative and must survive older
    // dashboard snapshots sent by automatic synchronization.
    await client.query(`DELETE FROM tickets WHERE source = 'dashboard'`);

    // Keep assets assigned to active requests even if an older browser snapshot
    // does not contain them. This prevents dangling issue/return workflows.
    await client.query(`
      CREATE TEMP TABLE preserved_workflow_assets ON COMMIT DROP AS
      SELECT a.*
      FROM assets a
      JOIN asset_requests r ON r.assigned_asset_sn = a.sn
      WHERE r.status IN ('approved', 'issued', 'overdue')
    `);

    // Keep workflow-controlled status/user authoritative when a browser sends an
    // older full asset snapshot during automatic synchronization.
    await client.query('DELETE FROM assets');
    await client.query(`
      INSERT INTO assets (sn, date, user_name, position, item_type, device_serial, status, notes, details)
      SELECT
        COALESCE(NULLIF(asset_row.item->>'sn', '')::integer, asset_row.ordinality::integer),
        COALESCE(asset_row.item->>'date', ''),
        COALESCE(assignment.requester, NULLIF(asset_row.item->>'user', ''), 'ส่วนกลาง'),
        COALESCE(NULLIF(asset_row.item->>'position', ''), '-'),
        COALESCE(asset_row.item->>'itemType', ''),
        COALESCE(NULLIF(asset_row.item->>'deviceSerial', ''), '-'),
        COALESCE(
          CASE WHEN assignment.status = 'approved' THEN 'จอง'
               WHEN assignment.status IN ('issued', 'overdue') THEN 'ใช้งาน'
          END,
          NULLIF(asset_row.item->>'status', ''),
          'ใช้งาน'
        ),
        COALESCE(asset_row.item->>'notes', ''),
        asset_row.item || jsonb_build_object(
          'user', COALESCE(assignment.requester, NULLIF(asset_row.item->>'user', ''), 'ส่วนกลาง'),
          'status', COALESCE(
            CASE WHEN assignment.status = 'approved' THEN 'จอง'
                 WHEN assignment.status IN ('issued', 'overdue') THEN 'ใช้งาน'
            END,
            NULLIF(asset_row.item->>'status', ''),
            'ใช้งาน'
          )
        )
      FROM jsonb_array_elements($1::jsonb) WITH ORDINALITY AS asset_row(item, ordinality)
      LEFT JOIN asset_requests assignment
        ON assignment.assigned_asset_sn = COALESCE(NULLIF(asset_row.item->>'sn', '')::integer, asset_row.ordinality::integer)
       AND assignment.status IN ('approved', 'issued', 'overdue')
    `, [JSON.stringify(assetsList)]);

    await client.query(`
      INSERT INTO assets (sn, date, user_name, position, item_type, device_serial, status, notes, details)
      SELECT sn, date, user_name, position, item_type, device_serial, status, notes, details
      FROM preserved_workflow_assets preserved
      WHERE NOT EXISTS (SELECT 1 FROM assets current_asset WHERE current_asset.sn = preserved.sn)
    `);

    const ticketPayload = [];
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
        integerNumber(monthData.totalAssets),
        finiteNumber(monthData.assetValue),
        integerNumber(monthData.assetsExpiring),
        integerNumber(monthData.assetsBroken),
        integerNumber(monthData.assetsLost),
        integerNumber(monthData.assetsVacant),
        integerNumber(monthData.ticketsCount),
        finiteNumber(monthData.slaPercent),
        integerNumber(monthData.responseTime),
        finiteNumber(monthData.resolutionTime),
        finiteNumber(monthData.csat),
        integerNumber(monthData.totalSoftware),
        finiteNumber(monthData.licensesInUse),
        finiteNumber(monthData.licensesVacant),
        finiteNumber(monthData.softwareCost),
        integerNumber(monthData.softwareExpiring),
        finiteNumber(monthData.backupSuccess),
        integerNumber(monthData.securityIncidents),
        finiteNumber(monthData.antivirusCoverage),
        finiteNumber(monthData.mfaCoverage),
        integerNumber(monthData.repairCount),
        finiteNumber(monthData.repairCost),
        JSON.stringify(monthData.topBrokenDevices || []),
        JSON.stringify(monthData.deptCosts || {}),
        JSON.stringify(monthData.softwareExpiringDetails || []),
        JSON.stringify(monthData.assetsExpiringDetails || []),
        JSON.stringify(monthData.ongoingProjects || []),
        JSON.stringify(monthData.recommendations || [])
      ]);

      const ticketsList = monthData.ticketsList || [];
      for (const ticket of ticketsList) {
        const ticketSn = Number(ticket.sn);
        if (!Number.isInteger(ticketSn) || ticketSn <= 0) continue;
        ticketPayload.push({
          sn: ticketSn,
          monthKey,
          date: ticket.date || '',
          complainant: ticket.complainant || 'ไม่ระบุชื่อ',
          email: ticket.email || '-',
          anydesk: ticket.anydesk || '-',
          issue: ticket.issue || '',
          cause: ticket.cause || '-',
          duration: ticket.duration || '-',
          responder: ticket.responder || '-',
          status: ticket.status || 'กำลังดำเนินการ',
          cost: finiteNumber(ticket.cost),
          source: ticket.source === 'request_form' ? 'request_form' : 'dashboard'
        });
      }
    }

    if (ticketPayload.length > 0) {
      await client.query(`
        INSERT INTO tickets (sn, month_key, date, complainant, email, anydesk, issue, cause, duration, responder, status, cost, source)
        SELECT
          (ticket.item->>'sn')::integer,
          ticket.item->>'monthKey',
          COALESCE(ticket.item->>'date', ''),
          COALESCE(NULLIF(ticket.item->>'complainant', ''), 'ไม่ระบุชื่อ'),
          COALESCE(NULLIF(ticket.item->>'email', ''), '-'),
          COALESCE(NULLIF(ticket.item->>'anydesk', ''), '-'),
          COALESCE(ticket.item->>'issue', ''),
          COALESCE(NULLIF(ticket.item->>'cause', ''), '-'),
          COALESCE(NULLIF(ticket.item->>'duration', ''), '-'),
          COALESCE(NULLIF(ticket.item->>'responder', ''), '-'),
          COALESCE(NULLIF(ticket.item->>'status', ''), 'กำลังดำเนินการ'),
          COALESCE((ticket.item->>'cost')::numeric, 0),
          COALESCE(NULLIF(ticket.item->>'source', ''), 'dashboard')
        FROM jsonb_array_elements($1::jsonb) AS ticket(item)
        ON CONFLICT (sn) DO UPDATE SET
          month_key = EXCLUDED.month_key,
          date = EXCLUDED.date,
          complainant = EXCLUDED.complainant,
          email = EXCLUDED.email,
          anydesk = EXCLUDED.anydesk,
          issue = EXCLUDED.issue,
          cause = EXCLUDED.cause,
          duration = EXCLUDED.duration,
          responder = EXCLUDED.responder,
          status = EXCLUDED.status,
          cost = EXCLUDED.cost,
          source = CASE WHEN tickets.source = 'request_form' THEN tickets.source ELSE EXCLUDED.source END
      `, [JSON.stringify(ticketPayload)]);
    }

    await refreshOperationalCounters(client);

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
