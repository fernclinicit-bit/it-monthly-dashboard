import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { 
  ArrowLeft, 
  Send, 
  Monitor, 
  AlertCircle, 
  User, 
  Briefcase, 
  Calendar 
} from 'lucide-react';
import './LarkForm.css';

const LarkForm = () => {
  const navigate = useNavigate();
  const [formData, setFormData] = useState({
    name: '',
    department: '',
    date: new Date().toISOString().split('T')[0],
    deviceType: '',
    assetSerial: '',
    anydesk: '',
    issue: '',
    priority: 'medium'
  });
  
  const [submitted, setSubmitted] = useState(false);
  const [submitting, setSubmitting] = useState(false);

  const handleChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => ({
      ...prev,
      [name]: value
    }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setSubmitting(true);
    try {
      const API_BASE = import.meta.env.VITE_API_BASE || (window.location.hostname === 'localhost' ? 'http://localhost:5000' : 'https://it-monthly-dashboard-new.onrender.com');

      const response = await fetch(`${API_BASE}/api/tickets`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(formData)
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'บันทึกคำร้องไม่สำเร็จ');
      
      setSubmitted(true);
      setTimeout(() => {
        setSubmitted(false);
        setFormData({
          name: '',
          department: '',
          date: new Date().toISOString().split('T')[0],
          deviceType: '',
          assetSerial: '',
          anydesk: '',
          issue: '',
          priority: 'medium'
        });
      }, 3000);
    } catch (error) {
      console.error("Error submitting form:", error);
      alert("เกิดข้อผิดพลาดในการส่งข้อมูล: " + error.message);
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="it-request-page">
      <div className="lark-form-header">
        <button className="back-btn" onClick={() => navigate('/')}>
          <ArrowLeft size={20} />
          กลับไปแดชบอร์ด
        </button>
      </div>

      <div className="lark-form-wrapper">
        <div className="form-card">
          <div className="form-title-section">
            <div className="icon-wrapper">
              <AlertCircle size={28} className="text-blue-600" />
            </div>
            <h2>แบบฟอร์มแจ้งปัญหา IT (IT Request Form)</h2>
            <p>กรุณากรอกรายละเอียดปัญหาหรือคำร้องของคุณ เพื่อให้ทีม IT ดำเนินการ</p>
          </div>

          {submitted ? (
            <div className="success-message">
              <div className="success-icon-wrapper">
                <svg className="checkmark" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 52 52">
                  <circle className="checkmark__circle" cx="26" cy="26" r="25" fill="none"/>
                  <path className="checkmark__check" fill="none" d="M14.1 27.2l7.1 7.2 16.7-16.8"/>
                </svg>
              </div>
              <h3>ส่งคำร้องสำเร็จ!</h3>
              <p>ทีม IT ได้รับคำร้องของคุณแล้ว และจะรีบดำเนินการโดยเร็วที่สุด</p>
            </div>
          ) : (
            <form onSubmit={handleSubmit} className="modern-form">
              <div className="form-grid">
                <div className="form-group">
                  <label htmlFor="name">
                    <User size={16} />
                    ชื่อ-นามสกุลผู้แจ้ง <span className="required">*</span>
                  </label>
                  <input 
                    type="text" 
                    id="name" 
                    name="name" 
                    value={formData.name}
                    onChange={handleChange}
                    placeholder="เช่น สมหมาย ใจดี" 
                    required 
                  />
                </div>

                <div className="form-group">
                  <label htmlFor="department">
                    <Briefcase size={16} />
                    แผนก/ฝ่าย <span className="required">*</span>
                  </label>
                  <input 
                    type="text" 
                    id="department" 
                    name="department" 
                    value={formData.department}
                    onChange={handleChange}
                    placeholder="เช่น บัญชี, การตลาด" 
                    required 
                  />
                </div>
              </div>

              <div className="form-grid">
                <div className="form-group">
                  <label htmlFor="date">
                    <Calendar size={16} />
                    วันที่แจ้ง
                  </label>
                  <input 
                    type="date" 
                    id="date" 
                    name="date" 
                    value={formData.date}
                    onChange={handleChange}
                  />
                </div>

                <div className="form-group">
                  <label htmlFor="deviceType">
                    <Monitor size={16} />
                    ประเภทอุปกรณ์ <span className="required">*</span>
                  </label>
                  <select 
                    id="deviceType" 
                    name="deviceType" 
                    value={formData.deviceType}
                    onChange={handleChange}
                    required
                  >
                    <option value="" disabled>-- เลือกประเภทอุปกรณ์ --</option>
                    <option value="PC">คอมพิวเตอร์ตั้งโต๊ะ (PC)</option>
                    <option value="Notebook">โน้ตบุ๊ก (Notebook)</option>
                    <option value="Printer">เครื่องปริ้นเตอร์ (Printer)</option>
                    <option value="Network">ระบบเครือข่าย/อินเทอร์เน็ต (Network)</option>
                    <option value="Software">โปรแกรม/ระบบ (Software)</option>
                    <option value="Other">อื่นๆ</option>
                  </select>
                </div>
              </div>

              <div className="form-group full-width">
                <label htmlFor="assetSerial">
                  <Monitor size={16} />
                  หมายเลขเครื่องจากทะเบียน (ถ้ามี)
                </label>
                <input
                  type="text"
                  id="assetSerial"
                  name="assetSerial"
                  value={formData.assetSerial}
                  onChange={handleChange}
                  placeholder="เช่น ASUS-019, MC-002"
                />
                <small>เมื่อระบุหมายเลขเครื่อง ระบบจะเปลี่ยนสถานะเครื่องเป็น “รอซ่อม” และซิงค์กับทะเบียนอัตโนมัติ</small>
              </div>

              <div className="form-group full-width">
                <label htmlFor="anydesk">
                  <Monitor size={16} />
                  เลขที่ AnyDesk (ถ้ามี)
                  <a
                    className="anydesk-download-link"
                    href="https://anydesk.com/en/downloads/thank-you?dv=win_exe"
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    ดาวน์โหลด AnyDesk
                  </a>
                </label>
                <input
                  type="text"
                  id="anydesk"
                  name="anydesk"
                  value={formData.anydesk}
                  onChange={handleChange}
                  placeholder="เช่น 123 456 789"
                  inputMode="numeric"
                />
              </div>

              <div className="form-group full-width">
                <label htmlFor="issue">รายละเอียดปัญหา <span className="required">*</span></label>
                <textarea 
                  id="issue" 
                  name="issue" 
                  rows="4" 
                  value={formData.issue}
                  onChange={handleChange}
                  placeholder="อธิบายปัญหาที่พบ เช่น เปิดเครื่องไม่ติด, ปริ้นไม่ออก..." 
                  required
                ></textarea>
              </div>

              <div className="form-group full-width">
                <label>ระดับความเร่งด่วน</label>
                <div className="radio-group">
                  <label className={`radio-card ${formData.priority === 'low' ? 'active low' : ''}`}>
                    <input 
                      type="radio" 
                      name="priority" 
                      value="low" 
                      checked={formData.priority === 'low'}
                      onChange={handleChange}
                    />
                    <div className="radio-content">
                      <span className="radio-title">ปกติ</span>
                      <span className="radio-desc">สามารถรอได้ 1-2 วัน</span>
                    </div>
                  </label>
                  <label className={`radio-card ${formData.priority === 'medium' ? 'active medium' : ''}`}>
                    <input 
                      type="radio" 
                      name="priority" 
                      value="medium" 
                      checked={formData.priority === 'medium'}
                      onChange={handleChange}
                    />
                    <div className="radio-content">
                      <span className="radio-title">ด่วน</span>
                      <span className="radio-desc">ต้องการใช้ภายในวันนี้</span>
                    </div>
                  </label>
                  <label className={`radio-card ${formData.priority === 'high' ? 'active high' : ''}`}>
                    <input 
                      type="radio" 
                      name="priority" 
                      value="high" 
                      checked={formData.priority === 'high'}
                      onChange={handleChange}
                    />
                    <div className="radio-content">
                      <span className="radio-title">ด่วนที่สุด</span>
                      <span className="radio-desc">กระทบการทำงานทั้งหมด</span>
                    </div>
                  </label>
                </div>
              </div>

              <div className="form-actions">
                <button type="submit" className="submit-btn" disabled={submitting}>
                  <Send size={18} />
                  {submitting ? 'กำลังส่งข้อมูล...' : 'ส่งคำร้อง'}
                </button>
              </div>
            </form>
          )}
        </div>
      </div>
    </div>
  );
};

export default LarkForm;
