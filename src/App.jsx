import React, { useState, useEffect, useRef } from 'react';
import Chart from 'chart.js/auto';
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
  X
} from 'lucide-react';

// Initial Mock Data (Google DeepMind Research IT Operations)
const initialDashboardData = {
  "2026-07": {
    monthName: "กรกฎาคม 2026",
    // Asset
    totalAssets: 1450,
    assetValue: 85200000,
    assetsExpiring: 38,
    assetsBroken: 8,
    assetsLost: 1,
    // Support
    ticketsCount: 280,
    slaPercent: 98.8,
    responseTime: 8,
    resolutionTime: 1.8,
    csat: 4.9,
    // Software
    totalSoftware: 45,
    licensesInUse: 2450,
    licensesVacant: 350,
    softwareCost: 1280000,
    softwareExpiring: 3,
    // Security
    backupSuccess: 99.98,
    securityIncidents: 0,
    antivirusCoverage: 100.0,
    mfaCoverage: 100.0,
    // Repair
    repairCount: 12,
    repairCost: 145000,
    topBrokenDevices: [
      { name: "Google TPU v5e Node", count: 4, cost: 80000 },
      { name: "NVIDIA H100 Workstation", count: 3, cost: 45000 },
      { name: "MacBook Pro M3 Max 16\"", count: 2, cost: 12000 },
      { name: "Dell Precision 7960", count: 1, cost: 5000 },
      { name: "ASUS ProArt AI Monitor", count: 1, cost: 2000 },
      { name: "Logitech MX Master 3S", count: 1, cost: 1000 },
      { name: "Google Pixel Fold 2", count: 0, cost: 0 },
      { name: "Cisco Catalyst Switch 9300", count: 0, cost: 0 },
      { name: "Supermicro Server H12", count: 0, cost: 0 },
      { name: "HP Color LaserJet Enterprise", count: 0, cost: 0 }
    ],
    deptCosts: {
      "AI Research": 85000,
      "Robotics Lab": 35000,
      "Infrastructure": 15000,
      "Operations": 6000,
      "Finance": 4000
    },
    softwareExpiringDetails: [
      { name: "Google Cloud Platform (GCP Enterprise)", expiringDate: "15 ส.ค. 2026", licenses: 500, status: "ใกล้หมดอายุ" },
      { name: "Slack Enterprise Grid", expiringDate: "28 ส.ค. 2026", licenses: 1200, status: "ใกล้หมดอายุ" },
      { name: "Weights & Biases (W&B Teams)", expiringDate: "12 ก.ย. 2026", licenses: 350, status: "แจ้งเตือนล่วงหน้า" }
    ],
    assetsExpiringDetails: [
      { id: "AST-TPU-042", type: "Server Node", model: "Google TPU v4 Node", dept: "AI Research", expDate: "10 ส.ค. 2026" },
      { id: "AST-GPU-109", type: "Workstation", model: "NVIDIA A100 Workstation", dept: "Robotics Lab", expDate: "15 ส.ค. 2026" },
      { id: "AST-LAP-882", type: "Laptop", model: "MacBook Pro M1 Max 16\"", dept: "Operations", expDate: "20 ส.ค. 2026" }
    ],
    // Improvement
    automationsDone: 5,
    aiApps: 4,
    hoursSaved: 320,
    ongoingProjects: [
      { title: "Gemini Auto-IT Agent", desc: "นำโมเดล Gemini มาช่วยตอบและแก้ปัญหาไอทีอัตโนมัติ คืบหน้า 85%" },
      { title: "TPU Cluster Optimization", desc: "ปรับปรุงการจัดสรรทรัพยากรการคำนวณสำหรับโมเดลวิจัย คืบหน้า 70%" },
      { title: "Zero Trust Security", desc: "ยกระดับระบบตรวจสอบสิทธิ์เข้าถึงข้อมูลโมเดลวิจัย คืบหน้า 90%" }
    ],
    recommendations: [
      "แนะนำจัดทำแผนงบประมาณเพื่อเปลี่ยนผ่านจาก TPU v4 Nodes ที่ใกล้หมดอายุไปยัง TPU v6e เพื่อเพิ่มประสิทธิภาพการคำนวณและลดการใช้พลังงาน",
      "เสนออนุมัติต่ออายุสัญญา Google Cloud Platform (GCP Enterprise) ล่วงหน้าก่อนหมดอายุในวันที่ 15 ส.ค. เพื่อเลี่ยงผลกระทบต่อ Training Pipelines ของทีมนักวิจัย",
      "ผลลัพธ์ของโครงการ Gemini Auto-IT Agent ช่วยลดภาระงาน Support Tickets ได้ถึง 35% ในเดือนนี้ แนะนำขยายผลให้รองรับภาษาท้องถิ่นเพิ่มเติมในสาขาเอเชียแปซิฟิก"
    ]
  },
  "2026-06": {
    monthName: "มิถุนายน 2026",
    // Asset
    totalAssets: 1420,
    assetValue: 84100000,
    assetsExpiring: 22,
    assetsBroken: 10,
    assetsLost: 0,
    // Support
    ticketsCount: 310,
    slaPercent: 98.1,
    responseTime: 9.5,
    resolutionTime: 2.1,
    csat: 4.8,
    // Software
    totalSoftware: 44,
    licensesInUse: 2400,
    licensesVacant: 400,
    softwareCost: 1250000,
    softwareExpiring: 1,
    // Security
    backupSuccess: 99.95,
    securityIncidents: 0,
    antivirusCoverage: 100.0,
    mfaCoverage: 100.0,
    // Repair
    repairCount: 15,
    repairCost: 168000,
    topBrokenDevices: [
      { name: "Google TPU v5e Node", count: 5, cost: 100000 },
      { name: "NVIDIA H100 Workstation", count: 4, cost: 60000 },
      { name: "MacBook Pro M3 Max 16\"", count: 3, cost: 8000 },
      { name: "Dell Precision 7960", count: 2, cost: 10000 },
      { name: "ASUS ProArt AI Monitor", count: 1, cost: 2000 },
      { name: "Logitech MX Master 3S", count: 0, cost: 0 },
      { name: "Google Pixel Fold 2", count: 0, cost: 0 },
      { name: "Cisco Catalyst Switch 9300", count: 0, cost: 0 },
      { name: "Supermicro Server H12", count: 0, cost: 0 },
      { name: "HP Color LaserJet Enterprise", count: 0, cost: 0 }
    ],
    deptCosts: {
      "AI Research": 110000,
      "Robotics Lab": 40000,
      "Infrastructure": 10000,
      "Operations": 8000,
      "Finance": 0
    },
    softwareExpiringDetails: [
      { name: "Weights & Biases (W&B Teams)", expiringDate: "12 ก.ค. 2026", licenses: 350, status: "ต่ออายุแล้ว" }
    ],
    assetsExpiringDetails: [
      { id: "AST-GPU-004", type: "Workstation", model: "NVIDIA A100 Workstation", dept: "AI Research", expDate: "15 มิ.ย. 2026" }
    ],
    // Improvement
    automationsDone: 3,
    aiApps: 3,
    hoursSaved: 280,
    ongoingProjects: [
      { title: "Gemini Auto-IT Agent", desc: "เริ่มทดสอบระบบคัดกรอง Ticket อัตโนมัติ คืบหน้า 50%" },
      { title: "TPU Cluster Optimization", desc: "วิเคราะห์การใช้พลังงานในคลัสเตอร์ขนาดใหญ่ คืบหน้า 45%" }
    ],
    recommendations: [
      "มีการจัดเก็บและคืนอุปกรณ์คอมพิวเตอร์อย่างมีประสิทธิภาพในเดือนนี้ ส่งผลให้อัตราอุปกรณ์สูญหายเป็นศูนย์ (0 เครื่อง)",
      "การเคลมประกันและส่งซ่อมของ TPU Nodes ในแผนก AI Research ส่งผลกระทบต่อเวลาของนักวิจัยบางส่วน แนะนำให้จัดหา Nodes สำรองเพิ่มสำหรับความล้มเหลวแบบฮาร์ดแวร์"
    ]
  }
};

export default function App() {
  const [data, setData] = useState(initialDashboardData);
  const [currentMonth, setCurrentMonth] = useState("2026-07");
  const [activeModal, setActiveModal] = useState(null); // 'edit', 'expiringAssets', 'expiringSoftware', 'topBrokenDevices'
  
  // Form input states
  const [formInputs, setFormInputs] = useState({});

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
