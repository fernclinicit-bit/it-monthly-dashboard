import React, { useState, useEffect, useRef } from 'react';
import Chart from 'chart.js/auto';
import * as XLSX from 'xlsx';
import { 
  Laptop, 
  Ticket, 
  FileCode, 
  ShieldCheck, 
  Wrench, 
  Lightbulb, 
  Edit3, 
  Printer, 
  Database,
  CheckCircle,
  AlertTriangle,
  Clock,
  ThumbsUp,
  X,
  Upload,
  Download,
  FileSpreadsheet
} from 'lucide-react';

// Initial blank data - use Excel import to load real data
const initialDashboardData = {
  "blank": {
    monthName: "รอนำเข้าข้อมูล",
    // Asset
    totalAssets: 0,
    assetValue: 0,
    assetsExpiring: 0,
    assetsBroken: 0,
    assetsLost: 0,
    // Support
    ticketsCount: 0,
    slaPercent: 0,
    responseTime: 0,
    resolutionTime: 0,
    csat: 0,
    // Software
    totalSoftware: 0,
    licensesInUse: 0,
    licensesVacant: 0,
    softwareCost: 0,
    softwareExpiring: 0,
    // Security
    backupSuccess: 0,
    securityIncidents: 0,
    antivirusCoverage: 0,
    mfaCoverage: 0,
    // Repair
    repairCount: 0,
    repairCost: 0,
    topBrokenDevices: [],
    deptCosts: {},
    softwareExpiringDetails: [],
    assetsExpiringDetails: [],
    // Improvement
    automationsDone: 0,
    aiApps: 0,
    hoursSaved: 0,
    ongoingProjects: [],
    recommendations: []
  }
};

export default function App() {
  const [data, setData] = useState(initialDashboardData);
  const [currentMonth, setCurrentMonth] = useState("blank");
  const [activeModal, setActiveModal] = useState(null); // 'edit', 'expiringAssets', 'expiringSoftware', 'topBrokenDevices'
  const [importStatus, setImportStatus] = useState(null); // { type: 'success' | 'error', message: string }
  
  // Form input states
  const [formInputs, setFormInputs] = useState({});

  // Hidden file input ref for xlsx import
  const fileInputRef = useRef(null);

  // Chart canvas refs
  const assetCanvasRef = useRef(null);
  const softwareCanvasRef = useRef(null);
  const repairCanvasRef = useRef(null);

  // Chart instances trackers
  const assetChartInst = useRef(null);
  const softwareChartInst = useRef(null);
  const repairChartInst = useRef(null);

  const activeData = data[currentMonth];

  // Helper currency formatter
  const formatThaiBaht = (value) => {
    return new Intl.NumberFormat('th-TH', { style: 'currency', currency: 'THB', maximumFractionDigits: 0 }).format(value);
  };

  // Re-render Chart.js graphs whenever activeData updates
  useEffect(() => {
    if (!activeData) return;

    // --- CHART 1: ASSETS (Doughnut) ---
    if (assetCanvasRef.current) {
      if (assetChartInst.current) assetChartInst.current.destroy();

      const normalAssets = activeData.totalAssets - activeData.assetsBroken - activeData.assetsLost - activeData.assetsExpiring;
      const ctx = assetCanvasRef.current.getContext('2d');
      assetChartInst.current = new Chart(ctx, {
        type: 'doughnut',
        data: {
          labels: ['ปกติ', 'ใกล้หมดอายุ', 'ชำรุด', 'สูญหาย'],
          datasets: [{
            data: [normalAssets, activeData.assetsExpiring, activeData.assetsBroken, activeData.assetsLost],
            backgroundColor: [
              'rgba(16, 185, 129, 0.7)',
              'rgba(245, 158, 11, 0.7)',
              'rgba(239, 68, 68, 0.7)',
              'rgba(156, 163, 175, 0.7)'
            ],
            borderColor: ['#10b981', '#f59e0b', '#ef4444', '#9ca3af'],
            borderWidth: 1
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: {
              position: 'right',
              labels: {
                color: '#e5e7eb',
                font: { size: 9, family: 'Sarabun' }
              }
            }
          },
          cutout: '70%'
        }
      });
    }

    // --- CHART 2: SOFTWARE LICENSES (Stacked horizontal bar) ---
    if (softwareCanvasRef.current) {
      if (softwareChartInst.current) softwareChartInst.current.destroy();

      const ctx = softwareCanvasRef.current.getContext('2d');
      softwareChartInst.current = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: ['สิทธิ์ใช้งาน (Licenses)'],
          datasets: [
            {
              label: 'ใช้งานอยู่ (In Use)',
              data: [activeData.licensesInUse],
              backgroundColor: 'rgba(59, 130, 246, 0.75)',
              borderColor: '#3b82f6',
              borderWidth: 1
            },
            {
              label: 'ว่าง (Vacant)',
              data: [activeData.licensesVacant],
              backgroundColor: 'rgba(6, 182, 212, 0.75)',
              borderColor: '#06b6d4',
              borderWidth: 1
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          indexAxis: 'y',
          scales: {
            x: {
              stacked: true,
              grid: { color: 'rgba(255,255,255,0.05)' },
              ticks: { color: '#9ca3af', font: { size: 8 } }
            },
            y: {
              stacked: true,
              grid: { display: false },
              ticks: { color: '#9ca3af', font: { size: 9, family: 'Sarabun' } }
            }
          },
          plugins: {
            legend: {
              position: 'bottom',
              labels: {
                color: '#e5e7eb',
                boxWidth: 12,
                font: { size: 9, family: 'Sarabun' }
              }
            }
          }
        }
      });
    }

    // --- CHART 3: REPAIRS BY DEPT (Horizontal bar) ---
    if (repairCanvasRef.current) {
      if (repairChartInst.current) repairChartInst.current.destroy();

      const ctx = repairCanvasRef.current.getContext('2d');
      const depts = Object.keys(activeData.deptCosts);
      const values = Object.values(activeData.deptCosts);

      repairChartInst.current = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: depts,
          datasets: [{
            label: 'ค่าใช้จ่ายซ่อม (บาท)',
            data: values,
            backgroundColor: 'rgba(245, 158, 11, 0.7)',
            borderColor: '#f59e0b',
            borderWidth: 1
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          indexAxis: 'y',
          scales: {
            x: {
              grid: { color: 'rgba(255,255,255,0.05)' },
              ticks: { color: '#9ca3af', font: { size: 8 } }
            },
            y: {
              grid: { display: false },
              ticks: { color: '#e5e7eb', font: { size: 9, family: 'Sarabun' } }
            }
          },
          plugins: {
            legend: { display: false }
          }
        }
      });
    }

    // Cleanup chart instances on unmount/re-render
    return () => {
      if (assetChartInst.current) assetChartInst.current.destroy();
      if (softwareChartInst.current) softwareChartInst.current.destroy();
      if (repairChartInst.current) repairChartInst.current.destroy();
    };
  }, [currentMonth, data]);

  // ========================================
  // XLSX IMPORT / EXPORT / TEMPLATE FUNCTIONS
  // ========================================

  // Define the field mapping for the main "Dashboard" sheet
  const FIELD_MAP = [
    { key: 'monthName', header: 'เดือน (Month Name)', example: 'กรกฎาคม 2026' },
    { key: 'monthKey', header: 'รหัสเดือน (Month Key, เช่น 2026-07)', example: '2026-07' },
    // Asset
    { key: 'totalAssets', header: 'จำนวนอุปกรณ์ทั้งหมด', example: 1450 },
    { key: 'assetValue', header: 'มูลค่าทรัพย์สิน IT (บาท)', example: 85200000 },
    { key: 'assetsExpiring', header: 'อุปกรณ์ใกล้หมดอายุ', example: 38 },
    { key: 'assetsBroken', header: 'อุปกรณ์ชำรุด', example: 8 },
    { key: 'assetsLost', header: 'อุปกรณ์สูญหาย', example: 1 },
    // Support
    { key: 'ticketsCount', header: 'จำนวน Ticket', example: 280 },
    { key: 'slaPercent', header: 'SLA Compliance (%)', example: 98.8 },
    { key: 'responseTime', header: 'Response Time เฉลี่ย (นาที)', example: 8 },
    { key: 'resolutionTime', header: 'Resolution Time เฉลี่ย (ชม.)', example: 1.8 },
    { key: 'csat', header: 'CSAT คะแนนความพึงพอใจ (จาก 5)', example: 4.9 },
    // Software
    { key: 'totalSoftware', header: 'โปรแกรมทั้งหมด', example: 45 },
    { key: 'licensesInUse', header: 'License ใช้งาน', example: 2450 },
    { key: 'licensesVacant', header: 'License ว่าง', example: 350 },
    { key: 'softwareCost', header: 'ค่าใช้จ่าย Software (บาท/เดือน)', example: 1280000 },
    { key: 'softwareExpiring', header: 'โปรแกรมใกล้หมดอายุ', example: 3 },
    // Security
    { key: 'backupSuccess', header: 'Backup สำเร็จ (%)', example: 99.98 },
    { key: 'securityIncidents', header: 'Security Incident (ครั้ง)', example: 0 },
    { key: 'antivirusCoverage', header: 'Antivirus Coverage (%)', example: 100 },
    { key: 'mfaCoverage', header: 'MFA Coverage (%)', example: 100 },
    // Repair
    { key: 'repairCount', header: 'จำนวนการซ่อม', example: 12 },
    { key: 'repairCost', header: 'ค่าใช้จ่ายการซ่อม (บาท)', example: 145000 },
    // Improvement
    { key: 'automationsDone', header: 'Automation ที่ทำเสร็จ', example: 5 },
    { key: 'aiApps', header: 'AI ที่นำมาใช้', example: 4 },
    { key: 'hoursSaved', header: 'ชั่วโมงที่ลดลงจาก Automation', example: 320 },
  ];

  // Download a blank .xlsx template
  const downloadTemplate = () => {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Dashboard Data (main metrics)
    const dashHeaders = FIELD_MAP.map(f => f.header);
    const dashExamples = FIELD_MAP.map(f => f.example);
    const dashWs = XLSX.utils.aoa_to_sheet([dashHeaders, dashExamples]);
    // Set column widths
    dashWs['!cols'] = dashHeaders.map(h => ({ wch: Math.max(h.length + 4, 18) }));
    XLSX.utils.book_append_sheet(wb, dashWs, 'Dashboard');

    // Sheet 2: Top 10 อุปกรณ์เสียบ่อย
    const repairHeaders = ['รหัสเดือน (Month Key)', 'ชื่ออุปกรณ์', 'จำนวนครั้งที่เสีย', 'ค่าใช้จ่ายซ่อม (บาท)'];
    const repairExample = ['2026-07', 'Google TPU v5e Node', 4, 80000];
    const repairWs = XLSX.utils.aoa_to_sheet([repairHeaders, repairExample]);
    repairWs['!cols'] = repairHeaders.map(h => ({ wch: Math.max(h.length + 4, 18) }));
    XLSX.utils.book_append_sheet(wb, repairWs, 'Top10 อุปกรณ์เสียบ่อย');

    // Sheet 3: ค่าใช้จ่ายต่อแผนก
    const deptHeaders = ['รหัสเดือน (Month Key)', 'ชื่อแผนก', 'ค่าใช้จ่ายซ่อม (บาท)'];
    const deptExample = ['2026-07', 'AI Research', 85000];
    const deptWs = XLSX.utils.aoa_to_sheet([deptHeaders, deptExample]);
    deptWs['!cols'] = deptHeaders.map(h => ({ wch: Math.max(h.length + 4, 18) }));
    XLSX.utils.book_append_sheet(wb, deptWs, 'ค่าใช้จ่ายต่อแผนก');

    // Sheet 4: โครงการที่กำลังดำเนินการ
    const projHeaders = ['รหัสเดือน (Month Key)', 'ชื่อโครงการ', 'รายละเอียดความคืบหน้า'];
    const projExample = ['2026-07', 'Gemini Auto-IT Agent', 'นำโมเดล Gemini มาช่วยตอบและแก้ปัญหาไอที คืบหน้า 85%'];
    const projWs = XLSX.utils.aoa_to_sheet([projHeaders, projExample]);
    projWs['!cols'] = projHeaders.map(h => ({ wch: Math.max(h.length + 4, 25) }));
    XLSX.utils.book_append_sheet(wb, projWs, 'โครงการดำเนินการ');

    // Sheet 5: Recommendation
    const recHeaders = ['รหัสเดือน (Month Key)', 'ข้อเสนอแนะ'];
    const recExample = ['2026-07', 'แนะนำจัดทำแผนงบประมาณเพื่อเปลี่ยนผ่านจาก TPU v4 Nodes'];
    const recWs = XLSX.utils.aoa_to_sheet([recHeaders, recExample]);
    recWs['!cols'] = recHeaders.map(h => ({ wch: Math.max(h.length + 4, 40) }));
    XLSX.utils.book_append_sheet(wb, recWs, 'Recommendation');

    // Sheet 6: อุปกรณ์ใกล้หมดอายุ (Expiring Assets Details)
    const expAssetHeaders = ['รหัสเดือน (Month Key)', 'รหัสทรัพย์สิน', 'ประเภท', 'รุ่น/โมเดล', 'แผนก', 'วันที่หมดอายุ'];
    const expAssetExample = ['2026-07', 'AST-TPU-042', 'Server Node', 'Google TPU v4 Node', 'AI Research', '10 ส.ค. 2026'];
    const expAssetWs = XLSX.utils.aoa_to_sheet([expAssetHeaders, expAssetExample]);
    expAssetWs['!cols'] = expAssetHeaders.map(h => ({ wch: Math.max(h.length + 4, 20) }));
    XLSX.utils.book_append_sheet(wb, expAssetWs, 'อุปกรณ์ใกล้หมดอายุ');

    // Sheet 7: โปรแกรมใกล้หมดอายุ (Expiring Software Details)
    const expSwHeaders = ['รหัสเดือน (Month Key)', 'ชื่อซอฟต์แวร์', 'จำนวน Licenses', 'วันหมดสัญญา', 'สถานะ'];
    const expSwExample = ['2026-07', 'Google Cloud Platform', 500, '15 ส.ค. 2026', 'ใกล้หมดอายุ'];
    const expSwWs = XLSX.utils.aoa_to_sheet([expSwHeaders, expSwExample]);
    expSwWs['!cols'] = expSwHeaders.map(h => ({ wch: Math.max(h.length + 4, 22) }));
    XLSX.utils.book_append_sheet(wb, expSwWs, 'โปรแกรมใกล้หมดอายุ');

    XLSX.writeFile(wb, 'IT_Dashboard_Template.xlsx');
  };

  // Export current data to .xlsx
  const exportToXlsx = () => {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Dashboard metrics (one row per month)
    const dashHeaders = FIELD_MAP.map(f => f.header);
    const rows = Object.entries(data).map(([monthKey, d]) => {
      return FIELD_MAP.map(f => {
        if (f.key === 'monthKey') return monthKey;
        return d[f.key] ?? '';
      });
    });
    const dashWs = XLSX.utils.aoa_to_sheet([dashHeaders, ...rows]);
    dashWs['!cols'] = dashHeaders.map(h => ({ wch: Math.max(h.length + 4, 18) }));
    XLSX.utils.book_append_sheet(wb, dashWs, 'Dashboard');

    // Sheet 2: Top 10 broken devices (all months)
    const repairHeaders = ['รหัสเดือน', 'ชื่ออุปกรณ์', 'จำนวนครั้งที่เสีย', 'ค่าใช้จ่ายซ่อม (บาท)'];
    const repairRows = [];
    Object.entries(data).forEach(([monthKey, d]) => {
      (d.topBrokenDevices || []).forEach(dev => {
        repairRows.push([monthKey, dev.name, dev.count, dev.cost]);
      });
    });
    const repairWs = XLSX.utils.aoa_to_sheet([repairHeaders, ...repairRows]);
    repairWs['!cols'] = repairHeaders.map(h => ({ wch: Math.max(h.length + 4, 18) }));
    XLSX.utils.book_append_sheet(wb, repairWs, 'Top10 อุปกรณ์เสียบ่อย');

    // Sheet 3: Department costs (all months)
    const deptHeaders = ['รหัสเดือน', 'ชื่อแผนก', 'ค่าใช้จ่ายซ่อม (บาท)'];
    const deptRows = [];
    Object.entries(data).forEach(([monthKey, d]) => {
      Object.entries(d.deptCosts || {}).forEach(([dept, cost]) => {
        deptRows.push([monthKey, dept, cost]);
      });
    });
    const deptWs = XLSX.utils.aoa_to_sheet([deptHeaders, ...deptRows]);
    deptWs['!cols'] = deptHeaders.map(h => ({ wch: Math.max(h.length + 4, 18) }));
    XLSX.utils.book_append_sheet(wb, deptWs, 'ค่าใช้จ่ายต่อแผนก');

    // Sheet 4: Ongoing projects (all months)
    const projHeaders = ['รหัสเดือน', 'ชื่อโครงการ', 'รายละเอียดความคืบหน้า'];
    const projRows = [];
    Object.entries(data).forEach(([monthKey, d]) => {
      (d.ongoingProjects || []).forEach(proj => {
        projRows.push([monthKey, proj.title, proj.desc]);
      });
    });
    const projWs = XLSX.utils.aoa_to_sheet([projHeaders, ...projRows]);
    projWs['!cols'] = projHeaders.map(h => ({ wch: Math.max(h.length + 4, 25) }));
    XLSX.utils.book_append_sheet(wb, projWs, 'โครงการดำเนินการ');

    // Sheet 5: Recommendations (all months)
    const recHeaders = ['รหัสเดือน', 'ข้อเสนอแนะ'];
    const recRows = [];
    Object.entries(data).forEach(([monthKey, d]) => {
      (d.recommendations || []).forEach(rec => {
        recRows.push([monthKey, rec]);
      });
    });
    const recWs = XLSX.utils.aoa_to_sheet([recHeaders, ...recRows]);
    recWs['!cols'] = recHeaders.map(h => ({ wch: Math.max(h.length + 4, 40) }));
    XLSX.utils.book_append_sheet(wb, recWs, 'Recommendation');

    // Sheet 6: Expiring assets details (all months)
    const expAssetHeaders = ['รหัสเดือน', 'รหัสทรัพย์สิน', 'ประเภท', 'รุ่น/โมเดล', 'แผนก', 'วันที่หมดอายุ'];
    const expAssetRows = [];
    Object.entries(data).forEach(([monthKey, d]) => {
      (d.assetsExpiringDetails || []).forEach(a => {
        expAssetRows.push([monthKey, a.id, a.type, a.model, a.dept, a.expDate]);
      });
    });
    const expAssetWs = XLSX.utils.aoa_to_sheet([expAssetHeaders, ...expAssetRows]);
    expAssetWs['!cols'] = expAssetHeaders.map(h => ({ wch: Math.max(h.length + 4, 20) }));
    XLSX.utils.book_append_sheet(wb, expAssetWs, 'อุปกรณ์ใกล้หมดอายุ');

    // Sheet 7: Expiring software details (all months)
    const expSwHeaders = ['รหัสเดือน', 'ชื่อซอฟต์แวร์', 'จำนวน Licenses', 'วันหมดสัญญา', 'สถานะ'];
    const expSwRows = [];
    Object.entries(data).forEach(([monthKey, d]) => {
      (d.softwareExpiringDetails || []).forEach(s => {
        expSwRows.push([monthKey, s.name, s.licenses, s.expiringDate, s.status]);
      });
    });
    const expSwWs = XLSX.utils.aoa_to_sheet([expSwHeaders, ...expSwRows]);
    expSwWs['!cols'] = expSwHeaders.map(h => ({ wch: Math.max(h.length + 4, 22) }));
    XLSX.utils.book_append_sheet(wb, expSwWs, 'โปรแกรมใกล้หมดอายุ');

    XLSX.writeFile(wb, `IT_Dashboard_Export_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  // Import data from uploaded .xlsx file
  const handleImportXlsx = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const binaryStr = evt.target.result;
        const wb = XLSX.read(binaryStr, { type: 'binary' });

        // --- Parse Sheet 1: Dashboard (by column index) ---
        const dashSheet = wb.Sheets[wb.SheetNames[0]];
        if (!dashSheet) throw new Error('ไม่พบ Sheet แรก (Dashboard)');
        const allRows = XLSX.utils.sheet_to_json(dashSheet, { header: 1 });
        // Skip header row (index 0), data starts from row 1
        const dataRows = allRows.slice(1).filter(r => r && r.length > 1);
        if (dataRows.length === 0) throw new Error('ไม่พบข้อมูลใน Sheet Dashboard');

        const newData = {};

        // Column order matches FIELD_MAP: 
        // 0=monthName, 1=monthKey, 2=totalAssets, 3=assetValue, ...
        dataRows.forEach(cols => {
          const monthName = String(cols[0] || '').trim();
          const monthKey = String(cols[1] || '').trim();

          if (!monthKey || !monthName) return;

          const monthData = {
            monthName,
            topBrokenDevices: [],
            deptCosts: {},
            ongoingProjects: [],
            recommendations: [],
            softwareExpiringDetails: [],
            assetsExpiringDetails: []
          };

          // Map remaining columns by index (starting from index 2)
          FIELD_MAP.forEach((field, idx) => {
            if (field.key === 'monthKey' || field.key === 'monthName') return;
            const rawVal = cols[idx];
            monthData[field.key] = rawVal !== undefined && rawVal !== '' ? Number(rawVal) : 0;
          });

          newData[monthKey] = monthData;
        });

        // --- Parse Sheet 2: Top 10 Broken Devices (optional) ---
        if (wb.SheetNames.length >= 2) {
          const repairSheet = wb.Sheets[wb.SheetNames[1]];
          if (repairSheet) {
            const rows = XLSX.utils.sheet_to_json(repairSheet, { header: 1 }).slice(1);
            rows.forEach(cols => {
              const mk = String(cols[0] || '').trim();
              if (mk && newData[mk]) {
                newData[mk].topBrokenDevices.push({
                  name: String(cols[1] || ''),
                  count: Number(cols[2]) || 0,
                  cost: Number(cols[3]) || 0
                });
              }
            });
          }
        }

        // --- Parse Sheet 3: Department Costs (optional) ---
        if (wb.SheetNames.length >= 3) {
          const deptSheet = wb.Sheets[wb.SheetNames[2]];
          if (deptSheet) {
            const rows = XLSX.utils.sheet_to_json(deptSheet, { header: 1 }).slice(1);
            rows.forEach(cols => {
              const mk = String(cols[0] || '').trim();
              if (mk && newData[mk]) {
                const deptName = String(cols[1] || '');
                const deptCost = Number(cols[2]) || 0;
                if (deptName) newData[mk].deptCosts[deptName] = deptCost;
              }
            });
          }
        }

        // --- Parse Sheet 4: Ongoing Projects (optional) ---
        if (wb.SheetNames.length >= 4) {
          const projSheet = wb.Sheets[wb.SheetNames[3]];
          if (projSheet) {
            const rows = XLSX.utils.sheet_to_json(projSheet, { header: 1 }).slice(1);
            rows.forEach(cols => {
              const mk = String(cols[0] || '').trim();
              if (mk && newData[mk]) {
                newData[mk].ongoingProjects.push({
                  title: String(cols[1] || ''),
                  desc: String(cols[2] || '')
                });
              }
            });
          }
        }

        // --- Parse Sheet 5: Recommendations (optional) ---
        if (wb.SheetNames.length >= 5) {
          const recSheet = wb.Sheets[wb.SheetNames[4]];
          if (recSheet) {
            const rows = XLSX.utils.sheet_to_json(recSheet, { header: 1 }).slice(1);
            rows.forEach(cols => {
              const mk = String(cols[0] || '').trim();
              if (mk && newData[mk]) {
                const recText = String(cols[1] || '');
                if (recText) newData[mk].recommendations.push(recText);
              }
            });
          }
        }

        // --- Parse Sheet 6: Expiring Assets Details (optional) ---
        if (wb.SheetNames.length >= 6) {
          const expAssetSheet = wb.Sheets[wb.SheetNames[5]];
          if (expAssetSheet) {
            const rows = XLSX.utils.sheet_to_json(expAssetSheet, { header: 1 }).slice(1);
            rows.forEach(cols => {
              const mk = String(cols[0] || '').trim();
              if (mk && newData[mk]) {
                newData[mk].assetsExpiringDetails.push({
                  id: String(cols[1] || ''),
                  type: String(cols[2] || ''),
                  model: String(cols[3] || ''),
                  dept: String(cols[4] || ''),
                  expDate: String(cols[5] || '')
                });
              }
            });
          }
        }

        // --- Parse Sheet 7: Expiring Software Details (optional) ---
        if (wb.SheetNames.length >= 7) {
          const expSwSheet = wb.Sheets[wb.SheetNames[6]];
          if (expSwSheet) {
            const rows = XLSX.utils.sheet_to_json(expSwSheet, { header: 1 }).slice(1);
            rows.forEach(cols => {
              const mk = String(cols[0] || '').trim();
              if (mk && newData[mk]) {
                newData[mk].softwareExpiringDetails.push({
                  name: String(cols[1] || ''),
                  licenses: Number(cols[2]) || 0,
                  expiringDate: String(cols[3] || ''),
                  status: String(cols[4] || '')
                });
              }
            });
          }
        }

        // Update state
        const monthKeys = Object.keys(newData);
        if (monthKeys.length === 0) throw new Error('ไม่พบข้อมูลเดือนที่สามารถนำเข้าได้');

        setData(newData);
        setCurrentMonth(monthKeys[0]);
        setImportStatus({ type: 'success', message: `นำเข้าสำเร็จ! พบข้อมูล ${monthKeys.length} เดือน` });
        setTimeout(() => setImportStatus(null), 4000);
      } catch (err) {
        setImportStatus({ type: 'error', message: `นำเข้าล้มเหลว: ${err.message}` });
        setTimeout(() => setImportStatus(null), 5000);
      }
    };
    reader.readAsBinaryString(file);
    // Reset the file input so re-uploading the same file triggers onChange
    e.target.value = '';
  };

  // Sync state data with form inputs when opening the edit modal
  const openEditModal = () => {
    setFormInputs({
      // Assets
      totalAssets: activeData.totalAssets,
      assetValue: activeData.assetValue,
      assetsExpiring: activeData.assetsExpiring,
      assetsBroken: activeData.assetsBroken,
      assetsLost: activeData.assetsLost,
      // Support
      ticketsCount: activeData.ticketsCount,
      slaPercent: activeData.slaPercent,
      responseTime: activeData.responseTime,
      resolutionTime: activeData.resolutionTime,
      csat: activeData.csat,
      // Software
      totalSoftware: activeData.totalSoftware,
      licensesInUse: activeData.licensesInUse,
      licensesVacant: activeData.licensesVacant,
      softwareCost: activeData.softwareCost,
      softwareExpiring: activeData.softwareExpiring,
      // Security
      backupSuccess: activeData.backupSuccess,
      securityIncidents: activeData.securityIncidents,
      antivirusCoverage: activeData.antivirusCoverage,
      mfaCoverage: activeData.mfaCoverage,
      // Repair
      repairCount: activeData.repairCount,
      repairCost: activeData.repairCost,
      // Recommendations text
      recommendations: activeData.recommendations.join('\n')
    });
    setActiveModal('edit');
  };

  // Save edit form modifications back to monthly state data
  const handleEditSubmit = (e) => {
    e.preventDefault();
    setData(prev => ({
      ...prev,
      [currentMonth]: {
        ...prev[currentMonth],
        totalAssets: Number(formInputs.totalAssets),
        assetValue: Number(formInputs.assetValue),
        assetsExpiring: Number(formInputs.assetsExpiring),
        assetsBroken: Number(formInputs.assetsBroken),
        assetsLost: Number(formInputs.assetsLost),

        ticketsCount: Number(formInputs.ticketsCount),
        slaPercent: Number(formInputs.slaPercent),
        responseTime: Number(formInputs.responseTime),
        resolutionTime: Number(formInputs.resolutionTime),
        csat: Number(formInputs.csat),

        totalSoftware: Number(formInputs.totalSoftware),
        licensesInUse: Number(formInputs.licensesInUse),
        licensesVacant: Number(formInputs.licensesVacant),
        softwareCost: Number(formInputs.softwareCost),
        softwareExpiring: Number(formInputs.softwareExpiring),

        backupSuccess: Number(formInputs.backupSuccess),
        securityIncidents: Number(formInputs.securityIncidents),
        antivirusCoverage: Number(formInputs.antivirusCoverage),
        mfaCoverage: Number(formInputs.mfaCoverage),

        repairCount: Number(formInputs.repairCount),
        repairCost: Number(formInputs.repairCost),
        
        recommendations: formInputs.recommendations.trim() ? formInputs.recommendations.split('\n') : []
      }
    }));
    setActiveModal(null);
  };

  // Helper component to render circular gauges for security block
  const CircularProgress = ({ value, label }) => {
    const radius = 27;
    const maxOffset = 170; // 2 * PI * radius
    const strokeOffset = maxOffset - (value / 100) * maxOffset;

    let strokeColor = 'var(--danger)';
    if (value >= 99) strokeColor = 'var(--success)';
    else if (value >= 95) strokeColor = 'var(--warning)';

    return (
      <div className="gauge-item">
        <div className="gauge-circle">
          <svg className="gauge-svg">
            <circle className="gauge-bg" cx="30" cy="30" r={radius}></circle>
            <circle 
              className="gauge-bar" 
              cx="30" 
              cy="30" 
              r={radius}
              style={{ strokeDashoffset: strokeOffset, stroke: strokeColor }}
            ></circle>
          </svg>
          <div className="gauge-text">{Math.round(value)}%</div>
        </div>
        <div className="metric-label">{label}</div>
      </div>
    );
  };

  return (
    <>
      {/* SIDEBAR NAVIGATION CONTROL PANEL */}
      <aside className="sidebar no-print">
        <div className="logo-container">
          <div className="logo-icon">MD</div>
          <div className="logo-text">
            <h1>IT Dashboard</h1>
            <p>React Executive Report</p>
          </div>
        </div>

        {/* Month Dropdown Selection */}
        <div className="control-group">
          <label className="control-label">เลือกเดือนที่ต้องการรายงาน</label>
          <select 
            className="month-selector"
            value={currentMonth}
            onChange={(e) => setCurrentMonth(e.target.value)}
          >
            {Object.keys(data).map(key => (
              <option key={key} value={key}>{data[key].monthName}</option>
            ))}
          </select>
        </div>

        {/* Data Customizer Trigger */}
        <div className="control-group">
          <label className="control-label">การจัดการข้อมูล</label>
          <button onClick={openEditModal} className="sidebar-btn">
            <Edit3 size={16} />
            แก้ไขข้อมูลรายงานนี้
          </button>
        </div>

        {/* XLSX Database Operations */}
        <div className="control-group">
          <label className="control-label">ฐานข้อมูล Excel (.xlsx)</label>
          <input
            type="file"
            ref={fileInputRef}
            accept=".xlsx,.xls"
            style={{ display: 'none' }}
            onChange={handleImportXlsx}
          />
          <button onClick={() => fileInputRef.current?.click()} className="sidebar-btn" style={{ backgroundColor: '#059669' }}>
            <Upload size={16} />
            นำเข้าข้อมูลจาก Excel
          </button>
          <button onClick={exportToXlsx} className="sidebar-btn secondary" style={{ marginTop: '6px' }}>
            <Download size={16} />
            ส่งออกข้อมูลเป็น Excel
          </button>
          <button onClick={downloadTemplate} className="sidebar-btn secondary" style={{ marginTop: '6px' }}>
            <FileSpreadsheet size={16} />
            ดาวน์โหลดเทมเพลต Excel
          </button>
          {importStatus && (
            <div style={{
              marginTop: '8px',
              padding: '8px 12px',
              borderRadius: '6px',
              fontSize: '0.78rem',
              fontWeight: '500',
              backgroundColor: importStatus.type === 'success' ? 'rgba(16, 185, 129, 0.15)' : 'rgba(239, 68, 68, 0.15)',
              color: importStatus.type === 'success' ? 'var(--success)' : 'var(--danger)',
              border: `1px solid ${importStatus.type === 'success' ? 'rgba(16, 185, 129, 0.3)' : 'rgba(239, 68, 68, 0.3)'}`
            }}>
              {importStatus.type === 'success' ? <CheckCircle size={13} style={{ marginRight: '6px', verticalAlign: 'text-bottom' }} /> : <AlertTriangle size={13} style={{ marginRight: '6px', verticalAlign: 'text-bottom' }} />}
              {importStatus.message}
            </div>
          )}
        </div>

        {/* PDF Printing Trigger */}
        <div className="control-group" style={{ marginTop: '10px' }}>
          <label className="control-label">ส่งออกเอกสาร</label>
          <button onClick={() => window.print()} className="sidebar-btn secondary">
            <Printer size={16} />
            บันทึกเป็น PDF / พิมพ์
          </button>
        </div>

        <div className="org-info">
          <p><strong>หน่วยงาน:</strong> ฝ่ายเทคโนโลยีสารสนเทศ (IT)</p>
          <p><strong>องค์กร:</strong> Google DeepMind Research</p>
          <p style={{ marginTop: '8px', fontSize: '0.7rem', color: 'var(--text-muted)' }}>
            ปรับปรุงข้อมูลล่าสุด: <br />18 กรกฎาคม 2026
          </p>
        </div>
      </aside>

      {/* DASHBOARD GRID CONTENT */}
      <main className="dashboard-container">
        <header className="dashboard-header">
          <div className="dashboard-title">
            <h2>รายงานสรุปการดำเนินงานเทคโนโลยีสารสนเทศ (IT Monthly Dashboard)</h2>
            <p>ประจำเดือน {activeData.monthName}</p>
          </div>
          <div className="header-status status-indicator">
            <span className="status-dot"></span>
            <span>ระบบรายงานพร้อมทำงาน</span>
          </div>
        </header>

        <section className="dashboard-grid">
          {/* CARD 1: ASSETS */}
          <article className="card asset-card">
            <div className="card-header">
              <h3 className="card-title">
                <span className="card-icon"><Laptop size={18} style={{ color: 'var(--primary)' }} /></span>
                ทรัพย์สินและอุปกรณ์ (Asset)
              </h3>
              <button 
                onClick={() => setActiveModal('expiringAssets')} 
                className="btn-details"
              >
                ดูข้อมูลหมดอายุ
              </button>
            </div>
            <div className="metrics-row">
              <div className="metric-item full-width">
                <div className="metric-label">จำนวนอุปกรณ์ทั้งหมด</div>
                <div className="metric-value highlight-primary">{activeData.totalAssets.toLocaleString()} เครื่อง</div>
              </div>
              <div className="metric-item full-width">
                <div className="metric-label">มูลค่าทรัพย์สิน IT รวม</div>
                <div className="metric-value">{formatThaiBaht(activeData.assetValue)}</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">ใกล้หมดอายุ</div>
                <div className="metric-value highlight-warning">{activeData.assetsExpiring} เครื่อง</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">ชำรุด</div>
                <div className="metric-value highlight-danger">{activeData.assetsBroken} เครื่อง</div>
              </div>
              <div className="metric-item full-width" style={{ marginTop: '-6px' }}>
                <div className="metric-label">สูญหาย</div>
                <div className="metric-value highlight-danger">{activeData.assetsLost} เครื่อง</div>
              </div>
            </div>
            <div className="card-chart-container">
              <canvas ref={assetCanvasRef}></canvas>
            </div>
          </article>

          {/* CARD 2: SUPPORT */}
          <article className="card support-card">
            <div className="card-header">
              <h3 className="card-title">
                <span className="card-icon"><Ticket size={18} style={{ color: 'var(--violet)' }} /></span>
                การสนับสนุนผู้ใช้ (Support)
              </h3>
            </div>
            <div className="metrics-row">
              <div className="metric-item">
                <div className="metric-label">จำนวน Ticket ทั้งหมด</div>
                <div className="metric-value highlight-primary">{activeData.ticketsCount.toLocaleString()} ใบ</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">สถิติ SLA Compliance</div>
                <div className="metric-value highlight-success">{activeData.slaPercent}%</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">Response Time เฉลี่ย</div>
                <div className="metric-value">{activeData.responseTime} นาที</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">Resolution Time เฉลี่ย</div>
                <div className="metric-value">{activeData.resolutionTime} ชม.</div>
              </div>
              <div className="metric-item full-width">
                <div className="metric-label">คะแนนความพึงพอใจผู้ใช้ (CSAT)</div>
                <div className="metric-value highlight-warning">{activeData.csat.toFixed(1)} / 5.0</div>
              </div>
            </div>
          </article>

          {/* CARD 3: SOFTWARE */}
          <article className="card software-card">
            <div className="card-header">
              <h3 className="card-title">
                <span className="card-icon"><FileCode size={18} style={{ color: 'var(--secondary)' }} /></span>
                ซอฟต์แวร์และลิขสิทธิ์ (Software)
              </h3>
              <button 
                onClick={() => setActiveModal('expiringSoftware')} 
                className="btn-details"
              >
                ดูโปรแกรมหมดสัญญา
              </button>
            </div>
            <div className="metrics-row">
              <div className="metric-item">
                <div className="metric-label">โปรแกรมทั้งหมด</div>
                <div className="metric-value">{activeData.totalSoftware} โปรแกรม</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">โปรแกรมใกล้หมดสัญญา</div>
                <div className="metric-value highlight-danger">{activeData.softwareExpiring} โปรแกรม</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">License ใช้งาน</div>
                <div className="metric-value highlight-primary">{activeData.licensesInUse.toLocaleString()} Core/User</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">License ว่าง</div>
                <div className="metric-value highlight-secondary">{activeData.licensesVacant.toLocaleString()} Core/User</div>
              </div>
              <div className="metric-item full-width">
                <div className="metric-label">ค่าใช้จ่ายซอฟต์แวร์รวมรายเดือน</div>
                <div className="metric-value">{formatThaiBaht(activeData.softwareCost)}</div>
              </div>
            </div>
            <div className="card-chart-container">
              <canvas ref={softwareCanvasRef}></canvas>
            </div>
          </article>

          {/* CARD 4: CYBER SECURITY */}
          <article className="card security-card">
            <div className="card-header">
              <h3 className="card-title">
                <span className="card-icon"><ShieldCheck size={18} style={{ color: 'var(--success)' }} /></span>
                ความปลอดภัยข้อมูล (Security)
              </h3>
            </div>
            <div className="metrics-row">
              <div className="metric-item full-width">
                <div className="metric-label">Security Incident (เหตุการณ์คุกคาม)</div>
                <div className={`metric-value ${activeData.securityIncidents > 0 ? 'highlight-danger' : 'highlight-success'}`}>
                  {activeData.securityIncidents} ครั้ง
                </div>
              </div>
            </div>
            <div className="gauges-container">
              <CircularProgress value={activeData.backupSuccess} label="Backup สำเร็จ" />
              <CircularProgress value={activeData.antivirusCoverage} label="Antivirus Coverage" />
              <CircularProgress value={activeData.mfaCoverage} label="MFA Coverage" />
            </div>
          </article>

          {/* CARD 5: REPAIR & COST */}
          <article className="card repair-card">
            <div className="card-header">
              <h3 className="card-title">
                <span className="card-icon"><Wrench size={18} style={{ color: 'var(--warning)' }} /></span>
                การซ่อมบำรุงและแผนก (Repair)
              </h3>
              <button 
                onClick={() => setActiveModal('topBrokenDevices')} 
                className="btn-details"
              >
                ดู Top 10 เสียบ่อย
              </button>
            </div>
            <div className="metrics-row">
              <div className="metric-item">
                <div className="metric-label">จำนวนงานซ่อม</div>
                <div className="metric-value highlight-warning">{activeData.repairCount} ครั้ง</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">ค่าซ่อมบำรุงสะสม</div>
                <div className="metric-value">{formatThaiBaht(activeData.repairCost)}</div>
              </div>
            </div>
            <div className="card-chart-container" style={{ height: '100px' }}>
              <canvas ref={repairCanvasRef}></canvas>
            </div>
            <div className="repair-list-summary">
              <div className="metric-label" style={{ marginBottom: '2px' }}>อุปกรณ์แจ้งซ่อมสูงสุด 3 อันดับแรก:</div>
              {activeData.topBrokenDevices.filter(d => d.count > 0).slice(0, 3).map((device, idx) => (
                <div key={idx} className="repair-list-item">
                  <span className="repair-item-name">{device.name}</span>
                  <span className="repair-item-count">{device.count} เครื่อง</span>
                </div>
              ))}
              {activeData.topBrokenDevices.filter(d => d.count > 0).length === 0 && (
                <div className="repair-list-item">
                  <span className="repair-item-name">ไม่มีประวัติอุปกรณ์ชำรุดในเดือนนี้</span>
                </div>
              )}
            </div>
          </article>

          {/* CARD 6: IMPROVEMENT */}
          <article className="card improvement-card">
            <div className="card-header">
              <h3 className="card-title">
                <span className="card-icon"><Lightbulb size={18} style={{ color: 'var(--danger)' }} /></span>
                การปรับปรุงและเทคโนโลยี (Improvement)
              </h3>
            </div>
            <div className="metrics-row">
              <div className="metric-item">
                <div className="metric-label">Automation ที่ทำเสร็จ</div>
                <div className="metric-value highlight-primary">{activeData.automationsDone} รายการ</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">AI ที่นำมาประยุกต์ใช้</div>
                <div className="metric-value highlight-secondary">{activeData.aiApps} โมเดล</div>
              </div>
              <div className="metric-item full-width">
                <div className="metric-label">ประหยัดชั่วโมงการทำงาน</div>
                <div className="metric-value highlight-success">{activeData.hoursSaved} ชั่วโมง/เดือน</div>
              </div>
            </div>
            <div className="project-list">
              <div className="metric-label" style={{ marginBottom: '2px' }}>โครงการหลักที่กำลังดำเนินการ:</div>
              {activeData.ongoingProjects.map((proj, idx) => (
                <div key={idx} className="project-item">
                  <div className="project-item-title">{proj.title}</div>
                  <div className="project-item-desc">{proj.desc}</div>
                </div>
              ))}
            </div>
          </article>
        </section>

        {/* BOTTOM SECTION: RECOMMENDATIONS */}
        <section className="card recommendation-card">
          <h3 className="card-title" style={{ color: 'white', marginBottom: '12px' }}>
            <span className="card-icon"><Lightbulb size={18} style={{ color: '#fff' }} /></span>
            ข้อเสนอแนะและแนวทางปฏิบัติ (Recommendation)
          </h3>
          <div className="recommendation-content">
            {activeData.recommendations.length > 0 ? (
              <ul className="recommendation-list">
                {activeData.recommendations.map((rec, idx) => (
                  <li key={idx}>{rec}</li>
                ))}
              </ul>
            ) : (
              <p>ไม่มีข้อเสนอแนะเพิ่มเติมสำหรับเดือนนี้ ระบบอยู่ในเกณฑ์ปกติ</p>
            )}
          </div>
        </section>
      </main>

      {/* MODAL 1: EDIT FORM CONSOLE */}
      {activeModal === 'edit' && (
        <div className="modal-overlay active">
          <div className="modal large">
            <header className="modal-header">
              <h3>✏️ แก้ไขข้อมูลประจำเดือน <span style={{ color: 'var(--primary)' }}>{activeData.monthName}</span></h3>
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>
            <div className="modal-body">
              <form onSubmit={handleEditSubmit}>
                
                {/* Hardware inputs */}
                <div className="form-section">
                  <h4 className="form-section-title">1. ทรัพย์สินและฮาร์ดแวร์ (Assets)</h4>
                  <div className="form-grid">
                    <div className="form-group">
                      <label>จำนวนอุปกรณ์ทั้งหมด (เครื่อง)</label>
                      <input 
                        type="number" 
                        value={formInputs.totalAssets || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, totalAssets: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>มูลค่าทรัพย์สินไอทีรวม (บาท)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetValue || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetValue: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>อุปกรณ์ใกล้หมดอายุ (เครื่อง)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetsExpiring || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetsExpiring: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>อุปกรณ์ชำรุด (เครื่อง)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetsBroken || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetsBroken: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>อุปกรณ์สูญหาย (เครื่อง)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetsLost || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetsLost: e.target.value }))}
                        required 
                      />
                    </div>
                  </div>
                </div>

                {/* Support inputs */}
                <div className="form-section">
                  <h4 className="form-section-title">2. บริการช่วยเหลือผู้ใช้ (Support)</h4>
                  <div className="form-grid">
                    <div className="form-group">
                      <label>จำนวน Ticket ทั้งหมด (ใบ)</label>
                      <input 
                        type="number" 
                        value={formInputs.ticketsCount || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, ticketsCount: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>สถิติการบรรลุข้อตกลง SLA (%)</label>
                      <input 
                        type="number" 
                        step="0.1" 
                        value={formInputs.slaPercent || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, slaPercent: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>Response Time เฉลี่ย (นาที)</label>
                      <input 
                        type="number" 
                        value={formInputs.responseTime || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, responseTime: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>Resolution Time เฉลี่ย (ชั่วโมง)</label>
                      <input 
                        type="number" 
                        step="0.1" 
                        value={formInputs.resolutionTime || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, resolutionTime: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group full-width">
                      <label>คะแนนความพึงพอใจผู้ใช้ CSAT (คะแนนเต็ม 5)</label>
                      <input 
                        type="number" 
                        step="0.1" 
                        max="5" 
                        value={formInputs.csat || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, csat: e.target.value }))}
                        required 
                      />
                    </div>
                  </div>
                </div>

                {/* Software inputs */}
                <div className="form-section">
                  <h4 className="form-section-title">3. ลิขสิทธิ์ซอฟต์แวร์ (Software)</h4>
                  <div className="form-grid">
                    <div className="form-group">
                      <label>โปรแกรมซอฟต์แวร์ทั้งหมด</label>
                      <input 
                        type="number" 
                        value={formInputs.totalSoftware || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, totalSoftware: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>โปรแกรมใกล้สัญญาหมดสัญญา</label>
                      <input 
                        type="number" 
                        value={formInputs.softwareExpiring || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, softwareExpiring: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>สิทธิ์/บัญชีใช้งานอยู่ (Licenses)</label>
                      <input 
                        type="number" 
                        value={formInputs.licensesInUse || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, licensesInUse: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>สิทธิ์/บัญชีว่าง (Licenses)</label>
                      <input 
                        type="number" 
                        value={formInputs.licensesVacant || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, licensesVacant: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group full-width">
                      <label>ค่าใช้จ่ายซอฟต์แวร์รวม (บาท)</label>
                      <input 
                        type="number" 
                        value={formInputs.softwareCost || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, softwareCost: e.target.value }))}
                        required 
                      />
                    </div>
                  </div>
                </div>

                {/* Security inputs */}
                <div className="form-section">
                  <h4 className="form-section-title">4. ความปลอดภัยข้อมูลและการกู้คืน (Security)</h4>
                  <div className="form-grid">
                    <div className="form-group">
                      <label>การสำรองข้อมูล (Backup) สำเร็จ (%)</label>
                      <input 
                        type="number" 
                        step="0.1" 
                        max="100" 
                        value={formInputs.backupSuccess || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, backupSuccess: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>ภัยคุกคาม Security Incident (ครั้ง)</label>
                      <input 
                        type="number" 
                        value={formInputs.securityIncidents || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, securityIncidents: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>ความคุ้มครองระบบ Antivirus (%)</label>
                      <input 
                        type="number" 
                        step="0.1" 
                        max="100" 
                        value={formInputs.antivirusCoverage || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, antivirusCoverage: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>การเปิดใช้งาน MFA ในระบบหลัก (%)</label>
                      <input 
                        type="number" 
                        step="0.1" 
                        max="100" 
                        value={formInputs.mfaCoverage || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, mfaCoverage: e.target.value }))}
                        required 
                      />
                    </div>
                  </div>
                </div>

                {/* Repair inputs */}
                <div className="form-section">
                  <h4 className="form-section-title">5. การแจ้งส่งซ่อมและค่าบำรุงรักษา (Repair)</h4>
                  <div className="form-grid">
                    <div className="form-group">
                      <label>จำนวนครั้งการส่งซ่อม</label>
                      <input 
                        type="number" 
                        value={formInputs.repairCount || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, repairCount: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>ค่าซ่อมแซมและบำรุงรักษาอุปกรณ์รวม (บาท)</label>
                      <input 
                        type="number" 
                        value={formInputs.repairCost || ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, repairCost: e.target.value }))}
                        required 
                      />
                    </div>
                  </div>
                </div>

                {/* Recommendation inputs */}
                <div className="form-section">
                  <h4 className="form-section-title">6. ข้อเสนอแนะและสรุปเชิงผู้บริหาร (Recommendations)</h4>
                  <div className="form-group full-width">
                    <label>ข้อเสนอแนะ (ใส่หนึ่งข้อเสนอแนะต่อ 1 บรรทัด)</label>
                    <textarea 
                      rows="4" 
                      value={formInputs.recommendations || ''} 
                      onChange={(e) => setFormInputs(p => ({ ...p, recommendations: e.target.value }))}
                      placeholder="แนะนำให้ปรับปรุงการ..."
                    ></textarea>
                  </div>
                </div>

                <footer className="form-footer">
                  <button 
                    type="button" 
                    onClick={() => setActiveModal(null)} 
                    className="sidebar-btn secondary" 
                    style={{ width: 'auto', padding: '10px 20px' }}
                  >
                    ยกเลิก
                  </button>
                  <button 
                    type="submit" 
                    className="sidebar-btn" 
                    style={{ width: 'auto', padding: '10px 24px' }}
                  >
                    บันทึกการแก้ไข
                  </button>
                </footer>
              </form>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: EXPIRING HARDWARE LIST */}
      {activeModal === 'expiringAssets' && (
        <div className="modal-overlay active">
          <div className="modal large">
            <header className="modal-header">
              <h3>รายละเอียดอุปกรณ์ใกล้หมดอายุ ({activeData.monthName})</h3>
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>
            <div className="modal-body">
              <div style={{ overflowX: 'auto' }}>
                <table className="details-table">
                  <thead>
                    <tr>
                      <th>รหัสทรัพย์สิน</th>
                      <th>ประเภท</th>
                      <th>รุ่น</th>
                      <th>แผนก</th>
                      <th>วันที่หมดอายุ</th>
                    </tr>
                  </thead>
                  <tbody>
                    {activeData.assetsExpiringDetails.length > 0 ? (
                      activeData.assetsExpiringDetails.map((asset, idx) => (
                        <tr key={idx}>
                          <td><strong>{asset.id}</strong></td>
                          <td>{asset.type}</td>
                          <td>{asset.model}</td>
                          <td>{asset.dept}</td>
                          <td><span style={{ color: 'var(--warning)' }}>{asset.expDate}</span></td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan="5" style={{ textAlign: 'center' }}>ไม่มีข้อมูลการเตือนหมดอายุของฮาร์ดแวร์</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 3: EXPIRING SOFTWARE AGREEMENTS */}
      {activeModal === 'expiringSoftware' && (
        <div className="modal-overlay active">
          <div className="modal large">
            <header className="modal-header">
              <h3>รายละเอียดโปรแกรม/สิทธิ์ใกล้หมดอายุ ({activeData.monthName})</h3>
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>
            <div className="modal-body">
              <div style={{ overflowX: 'auto' }}>
                <table className="details-table">
                  <thead>
                    <tr>
                      <th>ชื่อซอฟต์แวร์/โปรแกรม</th>
                      <th>จำนวนสิทธิ์ (Licenses)</th>
                      <th>วันหมดสัญญา</th>
                      <th>สถานะ</th>
                    </tr>
                  </thead>
                  <tbody>
                    {activeData.softwareExpiringDetails.length > 0 ? (
                      activeData.softwareExpiringDetails.map((soft, idx) => (
                        <tr key={idx}>
                          <td><strong>{soft.name}</strong></td>
                          <td>{soft.licenses} Licenses</td>
                          <td>{soft.expiringDate}</td>
                          <td>
                            <span 
                              style={{ 
                                color: soft.status.includes('หมดอายุ') ? 'var(--danger)' : 'var(--warning)', 
                                fontWeight: '600' 
                              }}
                            >
                              {soft.status}
                            </span>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan="4" style={{ textAlign: 'center' }}>ไม่มีข้อมูลโปรแกรมใกล้หมดสัญญา</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 4: TOP 10 BROKEN DEVICES */}
      {activeModal === 'topBrokenDevices' && (
        <div className="modal-overlay active">
          <div className="modal large">
            <header className="modal-header">
              <h3>ทำเนียบอุปกรณ์ชำรุด (Top 10 อุปกรณ์เสียบ่อย)</h3>
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>
            <div className="modal-body">
              <div style={{ overflowX: 'auto' }}>
                <table className="details-table">
                  <thead>
                    <tr>
                      <th>อันดับ</th>
                      <th>อุปกรณ์/รุ่น</th>
                      <th>จำนวนครั้งที่เสีย</th>
                      <th>ค่าใช้จ่ายในการซ่อมรวม (โดยประมาณ)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[...activeData.topBrokenDevices].sort((a,b) => b.count - a.count).map((device, idx) => (
                      <tr key={idx}>
                        <td><strong>{idx + 1}</strong></td>
                        <td>{device.name}</td>
                        <td><span className="repair-item-count">{device.count} ครั้ง</span></td>
                        <td>{device.cost > 0 ? formatThaiBaht(device.cost) : '0 บาท'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
