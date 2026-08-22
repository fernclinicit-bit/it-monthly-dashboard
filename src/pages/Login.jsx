import { useState } from 'react';
import { Eye, EyeOff, LockKeyhole, UserRound } from 'lucide-react';
import fernLogo from '../assets/fern-aesthetique-logo.png';
import { storeAuth } from '../auth';
import './Login.css';

const API_BASE = window.location.hostname === 'localhost' ? 'http://localhost:5000' : '';

export default function Login({ onLogin }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const submit = async (event) => {
    event.preventDefault();
    setError('');
    setLoading(true);
    try {
      const response = await fetch(`${API_BASE}/api/auth/login`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: username.trim(), password })
      });
      const result = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(result.message || 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
      const auth = { token: result.token, user: result.user };
      storeAuth(auth);
      onLogin(auth);
    } catch (loginError) {
      setError(loginError.message || 'ไม่สามารถเข้าสู่ระบบได้');
    } finally {
      setLoading(false);
    }
  };

  return (
    <main className="login-page">
      <section className="login-card" aria-labelledby="login-title">
        <img src={fernLogo} alt="FERN AESTHETIQUE" className="login-logo" />
        <div className="login-heading">
          <span className="login-icon"><LockKeyhole size={24} /></span>
          <div>
            <h1 id="login-title">เข้าสู่ระบบ IT Dashboard</h1>
            <p>กรอกบัญชีผู้ใช้เพื่อเข้าสู่ระบบตามสิทธิ์ที่ได้รับ</p>
          </div>
        </div>

        <form onSubmit={submit} className="login-form">
          <label htmlFor="login-user">User</label>
          <div className="login-input-wrap">
            <UserRound size={19} />
            <input id="login-user" autoComplete="username" value={username} onChange={(e) => setUsername(e.target.value)} placeholder="ชื่อผู้ใช้งาน" required autoFocus />
          </div>

          <label htmlFor="login-password">Password</label>
          <div className="login-input-wrap">
            <LockKeyhole size={19} />
            <input id="login-password" type={showPassword ? 'text' : 'password'} autoComplete="current-password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="รหัสผ่าน" required />
            <button type="button" className="password-toggle" onClick={() => setShowPassword((value) => !value)} aria-label={showPassword ? 'ซ่อนรหัสผ่าน' : 'แสดงรหัสผ่าน'}>
              {showPassword ? <EyeOff size={19} /> : <Eye size={19} />}
            </button>
          </div>

          {error && <div className="login-error" role="alert">{error}</div>}
          <button className="login-submit" type="submit" disabled={loading}>
            {loading ? 'กำลังตรวจสอบ...' : 'เข้าสู่ระบบ'}
          </button>
        </form>

        <div className="login-access-note">
          ระบบกำหนดสิทธิ์ 3 ระดับ: ผู้ดูรายงาน • ผู้ใช้งานทั่วไป • ผู้ดูแล IT
        </div>
      </section>
    </main>
  );
}
