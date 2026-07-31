import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { 
  Users, 
  DollarSign, 
  TrendingUp, 
  PieChart, 
  ArrowLeft,
  Activity
} from 'lucide-react';
import { Pie, Bar, Doughnut } from 'react-chartjs-2';
import 'chart.js/auto';

const ClinicDashboard = () => {
  const navigate = useNavigate();
  const [loading, setLoading] = useState(true);
  const [data, setData] = useState([]);
  const [error, setError] = useState(null);

  useEffect(() => {
    const fetchClinicData = async () => {
      try {
        const API_BASE = import.meta.env.VITE_API_BASE || (window.location.hostname === 'localhost' ? 'http://localhost:5000' : 'https://it-monthly-dashboard-new.onrender.com');
        const response = await fetch(`${API_BASE}/api/clinic-data`);
        if (!response.ok) {
          throw new Error('Failed to fetch clinic data');
        }
        const jsonData = await response.json();
        setData(jsonData);
      } catch (err) {
        console.error(err);
        setError(err.message);
      } finally {
        setLoading(false);
      }
    };

    fetchClinicData();
  }, []);

  // Compute Metrics
  const totalLeads = data.length;
  const totalBookings = data.reduce((sum, row) => sum + (Number(row['ยอดจอง']) || 0), 0);
  const totalTransfers = data.reduce((sum, row) => sum + (Number(row['ยอดสลิปรวม']) || 0), 0);

  // Chart Data preparation
  const channels = {};
  const interests = {};
  const statuses = {};

  data.forEach(row => {
    const channel = row['ช่องทาง'] || 'ไม่ระบุ';
    channels[channel] = (channels[channel] || 0) + 1;

    const interest = row['ความสนใจ'] || 'ไม่ระบุ';
    interests[interest] = (interests[interest] || 0) + 1;

    const status = row['สถานะ'] || 'ไม่ระบุ';
    statuses[status] = (statuses[status] || 0) + 1;
  });

  const channelChartData = {
    labels: Object.keys(channels),
    datasets: [{
      data: Object.values(channels),
      backgroundColor: ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6'],
      borderWidth: 0,
    }]
  };

  const interestChartData = {
    labels: Object.keys(interests),
    datasets: [{
      label: 'จำนวนลูกค้า',
      data: Object.values(interests),
      backgroundColor: '#3b82f6',
      borderRadius: 4,
    }]
  };

  const statusChartData = {
    labels: Object.keys(statuses),
    datasets: [{
      data: Object.values(statuses),
      backgroundColor: ['#10b981', '#f59e0b', '#ef4444', '#6366f1'],
      borderWidth: 0,
    }]
  };

  const chartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { position: 'bottom', labels: { color: '#e5e7eb' } }
    },
    scales: {
      x: { display: false },
      y: { display: false }
    }
  };

  const barChartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false }
    },
    scales: {
      x: { ticks: { color: '#9ca3af' }, grid: { display: false } },
      y: { ticks: { color: '#9ca3af', stepSize: 1 }, grid: { color: '#374151' } }
    }
  };

  if (loading) return <div style={{ color: 'var(--text-primary)', padding: '20px' }}>กำลังโหลดข้อมูล...</div>;
  if (error) return <div style={{ color: 'red', padding: '20px' }}>เกิดข้อผิดพลาด: {error}</div>;

  return (
    <div style={{ backgroundColor: '#111827', minHeight: '100vh', color: 'var(--text-primary)', padding: '20px' }}>
      <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
        <div>
          <button onClick={() => navigate('/')} style={{ background: 'none', border: 'none', color: '#9ca3af', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '5px', marginBottom: '10px' }}>
            <ArrowLeft size={16} /> กลับไปยังหน้าระบบ IT
          </button>
          <h1 style={{ margin: 0, fontSize: '24px', color: '#f3f4f6' }}>รายงานลูกค้า Fern Clinic (Customer Report)</h1>
          <p style={{ margin: '5px 0 0 0', color: '#9ca3af' }}>ข้อมูลดึงจากไฟล์ Update/fern-clinic-customer-report.xls อัตโนมัติ</p>
        </div>
      </header>

      {/* KPI Cards */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))', gap: '20px', marginBottom: '20px' }}>
        <div style={{ backgroundColor: 'var(--bg-card)', padding: '20px', borderRadius: '12px', display: 'flex', alignItems: 'center', gap: '15px' }}>
          <div style={{ padding: '12px', backgroundColor: 'rgba(59, 130, 246, 0.2)', borderRadius: '50%', color: '#3b82f6' }}>
            <Users size={24} />
          </div>
          <div>
            <p style={{ margin: 0, color: '#9ca3af', fontSize: '14px' }}>จำนวนลูกค้าทั้งหมด (Leads)</p>
            <h3 style={{ margin: 0, fontSize: '24px', color: '#f3f4f6' }}>{totalLeads} คน</h3>
          </div>
        </div>
        <div style={{ backgroundColor: 'var(--bg-card)', padding: '20px', borderRadius: '12px', display: 'flex', alignItems: 'center', gap: '15px' }}>
          <div style={{ padding: '12px', backgroundColor: 'rgba(16, 185, 129, 0.2)', borderRadius: '50%', color: '#10b981' }}>
            <DollarSign size={24} />
          </div>
          <div>
            <p style={{ margin: 0, color: '#9ca3af', fontSize: '14px' }}>ยอดจองรวม (Booking)</p>
            <h3 style={{ margin: 0, fontSize: '24px', color: '#f3f4f6' }}>฿{totalBookings.toLocaleString()}</h3>
          </div>
        </div>
        <div style={{ backgroundColor: 'var(--bg-card)', padding: '20px', borderRadius: '12px', display: 'flex', alignItems: 'center', gap: '15px' }}>
          <div style={{ padding: '12px', backgroundColor: 'rgba(245, 158, 11, 0.2)', borderRadius: '50%', color: '#f59e0b' }}>
            <TrendingUp size={24} />
          </div>
          <div>
            <p style={{ margin: 0, color: '#9ca3af', fontSize: '14px' }}>ยอดโอนจริง (Transfer)</p>
            <h3 style={{ margin: 0, fontSize: '24px', color: '#f3f4f6' }}>฿{totalTransfers.toLocaleString()}</h3>
          </div>
        </div>
      </div>

      {/* Charts */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: '20px', marginBottom: '20px' }}>
        <div style={{ backgroundColor: 'var(--bg-card)', padding: '20px', borderRadius: '12px' }}>
          <h3 style={{ margin: '0 0 15px 0', fontSize: '16px', color: '#e5e7eb', display: 'flex', alignItems: 'center', gap: '8px' }}>
            <PieChart size={18} color="#3b82f6" /> สัดส่วนลูกค้าตามช่องทาง
          </h3>
          <div style={{ height: '250px' }}>
            <Pie data={channelChartData} options={chartOptions} />
          </div>
        </div>

        <div style={{ backgroundColor: 'var(--bg-card)', padding: '20px', borderRadius: '12px' }}>
          <h3 style={{ margin: '0 0 15px 0', fontSize: '16px', color: '#e5e7eb', display: 'flex', alignItems: 'center', gap: '8px' }}>
            <Activity size={18} color="#10b981" /> ความสนใจของลูกค้า
          </h3>
          <div style={{ height: '250px' }}>
            <Bar data={interestChartData} options={barChartOptions} />
          </div>
        </div>

        <div style={{ backgroundColor: 'var(--bg-card)', padding: '20px', borderRadius: '12px' }}>
          <h3 style={{ margin: '0 0 15px 0', fontSize: '16px', color: '#e5e7eb', display: 'flex', alignItems: 'center', gap: '8px' }}>
            <TrendingUp size={18} color="#f59e0b" /> สถานะการติดตาม
          </h3>
          <div style={{ height: '250px' }}>
            <Doughnut data={statusChartData} options={chartOptions} />
          </div>
        </div>
      </div>

      {/* Data Table */}
      <div style={{ backgroundColor: 'var(--bg-card)', padding: '20px', borderRadius: '12px', overflowX: 'auto' }}>
        <h3 style={{ margin: '0 0 15px 0', fontSize: '16px', color: '#e5e7eb' }}>ตารางข้อมูลลูกค้าล่าสุด</h3>
        <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
          <thead>
            <tr style={{ borderBottom: '1px solid #374151', color: '#9ca3af' }}>
              <th style={{ padding: '12px 8px' }}>ชื่อลูกค้า</th>
              <th style={{ padding: '12px 8px' }}>ช่องทาง</th>
              <th style={{ padding: '12px 8px' }}>ความสนใจ</th>
              <th style={{ padding: '12px 8px' }}>สถานะ</th>
              <th style={{ padding: '12px 8px' }}>ยอดจอง</th>
              <th style={{ padding: '12px 8px' }}>เจ้าของงาน</th>
            </tr>
          </thead>
          <tbody>
            {data.map((row, index) => (
              <tr key={index} style={{ borderBottom: '1px solid #374151' }}>
                <td style={{ padding: '12px 8px', color: '#f3f4f6' }}>{row['ชื่อลูกค้า']}</td>
                <td style={{ padding: '12px 8px', color: '#9ca3af' }}>{row['ช่องทาง']}</td>
                <td style={{ padding: '12px 8px', color: '#9ca3af' }}>{row['ความสนใจ']}</td>
                <td style={{ padding: '12px 8px' }}>
                  <span style={{ 
                    padding: '4px 8px', 
                    borderRadius: '4px', 
                    fontSize: '12px',
                    backgroundColor: row['สถานะ'] === 'ต้องติดตาม' ? 'rgba(245, 158, 11, 0.2)' : 'rgba(16, 185, 129, 0.2)',
                    color: row['สถานะ'] === 'ต้องติดตาม' ? '#f59e0b' : '#10b981'
                  }}>
                    {row['สถานะ']}
                  </span>
                </td>
                <td style={{ padding: '12px 8px', color: '#f3f4f6' }}>฿{(Number(row['ยอดจอง']) || 0).toLocaleString()}</td>
                <td style={{ padding: '12px 8px', color: '#9ca3af' }}>{row['เจ้าของงาน']}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default ClinicDashboard;
