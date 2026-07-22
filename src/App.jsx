import React, { useState, useEffect, useRef } from 'react';
import { Routes, Route, useNavigate } from 'react-router-dom';
import LarkForm from './pages/LarkForm';
import Chart from 'chart.js/auto';
import * as XLSX from 'xlsx';
import { 
  BarChart3, 
  TrendingUp, 
  Layers, 
  Cpu, 
  ShieldAlert, 
  Activity, 
  Award, 
  Search, 
  Ticket, 
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
  FileSpreadsheet,
  FileText,
  ChevronDown,
  ChevronRight,
  Menu,
  Laptop,
  FileCode,
  ShieldCheck,
  Wrench
} from 'lucide-react';

// Initial blank data - use Excel import to load real data
const initialAssetsData = [
  {
    "sn": 23,
    "date": "20/05/2569",
    "user": "อำพล เเซ่เเฮ",
    "position": "IT",
    "itemType": "Computer (Pc)",
    "deviceSerial": "AS-002",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 83,
    "date": "06/12/2568",
    "user": "อำพล    แซ่แฮ",
    "position": "IT",
    "itemType": "Ipad",
    "deviceSerial": "iPad-010",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 217,
    "date": "",
    "user": "อำพล  แซ่แฮ",
    "position": "IT",
    "itemType": "Mornitor",
    "deviceSerial": "LG-005",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 213,
    "date": "18/05/2569",
    "user": "ธันวา เเซ่เเฮ",
    "position": "IT",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-010 ",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 138,
    "date": "20/11/2568",
    "user": "อำพล   แซ่แฮ",
    "position": "IT",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-033",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 239,
    "date": "",
    "user": "อำพล  แซ่แฮ",
    "position": "IT",
    "itemType": "External HDD",
    "deviceSerial": "ETN-003 WD My PassPort 1TB",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 246,
    "date": "",
    "user": "อำพล  แซ่แฮ",
    "position": "IT",
    "itemType": "Cable HDMI",
    "deviceSerial": "ไม่มีหมายเลข อุปกรณ์",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 248,
    "date": "10/07/2569",
    "user": "ธันวา เเซ่เเฮ",
    "position": "IT",
    "itemType": "Mouse",
    "deviceSerial": "MOS-002",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 252,
    "date": "",
    "user": "อำพล  แซ่แฮ",
    "position": "IT",
    "itemType": "Keyboard",
    "deviceSerial": "KBD-003",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 329,
    "date": "10/07/2569",
    "user": "ธันวา เเซ่เเฮ",
    "position": "IT",
    "itemType": "Keyboard",
    "deviceSerial": "KBD-007",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 272,
    "date": "",
    "user": "อำพล   แซ่แฮ",
    "position": "IT",
    "itemType": "Screwdriver",
    "deviceSerial": "ไม่มีหมายเลข อุปกรณ์",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 33,
    "date": "09/07/2569",
    "user": "วิลาสินี ทับทิม (นี)",
    "position": "Accounting",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-006",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ วัน จ"
  },
  {
    "sn": 313,
    "date": "",
    "user": "สุภาพ แสนจันทร์ ( ส้ม )",
    "position": "Accounting",
    "itemType": "Printer",
    "deviceSerial": "PT-002",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ วัน จ"
  },
  {
    "sn": 135,
    "date": "09/07/2569",
    "user": "วิลาสินี ทับทิม (นี)",
    "position": "Accounting",
    "itemType": "Mornitor",
    "deviceSerial": "LG-004",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ วัน จ"
  },
  {
    "sn": 27,
    "date": "",
    "user": "สุภาพ แสนจันทร์ ( ส้ม )",
    "position": "Accounting",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-005",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ วัน จ"
  },
  {
    "sn": 34,
    "date": "",
    "user": "ณัฐชา คำสอนพันธ์ (กิ๊ก)",
    "position": "Procurement & Warehouse Officer",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-007",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 93,
    "date": "08/05/3111",
    "user": "ศิริพร เพชรมูล ( บี )",
    "position": "Procurement & Warehouse Officer",
    "itemType": "Ipad",
    "deviceSerial": "iPad-020",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 71,
    "date": "08/05/3111",
    "user": "ศิริพร เพชรมูล ( บี )",
    "position": "Procurement & Warehouse Officer",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-011",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 111,
    "date": "08/05/3111",
    "user": "ศิริพร เพชรมูล ( บี )",
    "position": "Procurement & Warehouse Officer",
    "itemType": "Apple Pancill",
    "deviceSerial": "Pencil-018",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 26,
    "date": "",
    "user": "ศิริพร เพชรมูล ( บี )",
    "position": "Procurement & Warehouse Officer",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-004",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 58,
    "date": "",
    "user": "ธนัชชา บุญมีมาก (ป๊อป)",
    "position": "Content Creator",
    "itemType": "Macbook",
    "deviceSerial": "MacBookAir-028",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 325,
    "date": "10/07/2569",
    "user": "วริสรา สงวนวงษ์ (ซินดี้)",
    "position": "Content Creator",
    "itemType": "Macbook",
    "deviceSerial": "MacBook air-029  ",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 339,
    "date": "01/07/2569",
    "user": "ณัฐฏชญาดา ตรีวิวัฒน์กุล (Mac)",
    "position": "Content Creator",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-043",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 231,
    "date": "23/05/2569",
    "user": "ธนัชชา บุญมีมาก (ป๊อป)",
    "position": "Content Creator",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-MKT-005",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 214,
    "date": "",
    "user": "วริสรา สงวนวงษ์ (ซินดี้)",
    "position": "Tiktok Content Creator",
    "itemType": "Notebook Asus",
    "deviceSerial": "Asus-031",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 265,
    "date": "11/12/2568",
    "user": "นภัสสร นาสวน ( โบว์ )",
    "position": "Tiktok Content Creator",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-037",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 86,
    "date": "",
    "user": "นภัสสร นาสวน ( โบว์ )",
    "position": "Tiktok Content Creator",
    "itemType": "Ipad",
    "deviceSerial": "iPad-013",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 87,
    "date": "18/07/2569",
    "user": "บุษกร บัวสวรรค์",
    "position": "Tiktok Content Creator",
    "itemType": "Ipad",
    "deviceSerial": "iPad-014",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 282,
    "date": "15/07/2569",
    "user": "ปณิศอร บุญจูบุตร",
    "position": "Tiktok Content Creator",
    "itemType": "Ipad",
    "deviceSerial": "iPad-021",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 53,
    "date": "",
    "user": "ทีม",
    "position": "Tiktok Content Creator",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-013",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 62,
    "date": "",
    "user": "ทีม",
    "position": "Tiktok Content Creator",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-002",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 66,
    "date": "",
    "user": "ทีม",
    "position": "Tiktok Content Creator",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-006",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 72,
    "date": "",
    "user": "ทีม",
    "position": "Tiktok Content Creator",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-012",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 357,
    "date": "",
    "user": "ทีม",
    "position": "Tiktok Content Creator",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-014",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 365,
    "date": "30/06/2569",
    "user": "ทีม",
    "position": "Tiktok Content Creator",
    "itemType": "IPhone",
    "deviceSerial": "iphone-015",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 20,
    "date": "",
    "user": "นภัสสร นาสวน ( โบว์ )",
    "position": "Tiktok Content Creator",
    "itemType": "IMac",
    "deviceSerial": "MC-002",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 326,
    "date": "",
    "user": "ชัยธัช ชัยวัฒน์ (มาร์ค)",
    "position": "Graphic Designer",
    "itemType": "Computer (Pc)",
    "deviceSerial": "AS-003",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 305,
    "date": "08/06/2569",
    "user": "รามจิตติ ชินนะเกิดโชค(เบนซ์)",
    "position": "Graphic Designer",
    "itemType": "Computer (Pc), Mornitor",
    "deviceSerial": "PC LG-004 ",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 294,
    "date": "08/06/2569",
    "user": "อาทิตยา มุมทอง (ขมิ้น)",
    "position": "Graphic Designer",
    "itemType": "IMac",
    "deviceSerial": "MC-008",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 319,
    "date": "19/05/2569",
    "user": "พิชชาพร คอทอง(พีเจ้น)",
    "position": "Graphic Designer",
    "itemType": "IMac",
    "deviceSerial": "MC-009",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 218,
    "date": "08/06/2569",
    "user": "อาทิตยา มุมทอง (ขมิ้น)",
    "position": "Graphic Designer",
    "itemType": "Mornitor",
    "deviceSerial": "LG-003 (จอ)",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 303,
    "date": "08/06/2569",
    "user": "รามจิตติ ชินนะเกิดโชค(เบนซ์)",
    "position": "Graphic Designer",
    "itemType": "Mornitor",
    "deviceSerial": "AS-001",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 304,
    "date": "",
    "user": "พิชชาพร คอทอง(พีเจ้น)",
    "position": "Graphic Designer",
    "itemType": "Mornitor",
    "deviceSerial": "AS-003",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 318,
    "date": "",
    "user": "ชัยธัช ชัยวัฒน์ (มาร์ค)",
    "position": "Graphic Designer",
    "itemType": "Mornitor",
    "deviceSerial": "LG-002",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 234,
    "date": "08/06/2569",
    "user": "อาทิตยา มุมทอง (ขมิ้น)",
    "position": "Graphic Designer",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-Meeting-001",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 301,
    "date": "",
    "user": "พิชชาพร คอทอง(พีเจ้น)",
    "position": "Graphic Designer",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Meeting-002",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 238,
    "date": "08/06/2569",
    "user": "อาทิตยา มุมทอง (ขมิ้น)",
    "position": "Graphic Designer",
    "itemType": "External HDD",
    "deviceSerial": "ETN-002  WD My PassPort 4TB",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 247,
    "date": "",
    "user": "พิชชาพร คอทอง(พีเจ้น)",
    "position": "Graphic Designer",
    "itemType": "Mouse",
    "deviceSerial": "MOS-001",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 249,
    "date": "08/06/2569",
    "user": "อาทิตยา มุมทอง (ขมิ้น)",
    "position": "Graphic Designer",
    "itemType": "Mouse",
    "deviceSerial": "MOS-003",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 250,
    "date": "08/06/2569",
    "user": "อาทิตยา มุมทอง (ขมิ้น)",
    "position": "Graphic Designer",
    "itemType": "Keyboard",
    "deviceSerial": "KBD-004",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 260,
    "date": "01/06/2569",
    "user": "ชัยธัช ชัยวัฒน์ (มาร์ค)",
    "position": "Graphic Designer",
    "itemType": "Keyboard",
    "deviceSerial": "KBD-006",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 302,
    "date": "08/06/2569",
    "user": "รามจิตติ ชินนะเกิดโชค(เบนซ์)",
    "position": "Graphic Designer",
    "itemType": "Hub Lan",
    "deviceSerial": "ไม่มีหมายเลข",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 330,
    "date": "20/05/2569",
    "user": "พิชชาพร คอทอง(พีเจ้น)",
    "position": "Graphic Designer",
    "itemType": "Hub Lan",
    "deviceSerial": "USB-C TO LAN ",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 229,
    "date": "",
    "user": "ทีม admin",
    "position": "Admin",
    "itemType": "Computer (Pc)",
    "deviceSerial": "MIS-002 มินิpc",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 36,
    "date": "",
    "user": "กานต์ฑิตา ธีระพิบูลย์ ( อิง )",
    "position": "Admin",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-011",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 37,
    "date": "",
    "user": "ศุภฤกษ์ ภายใธสง (คริม)",
    "position": "Admin",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-017",
    "status": "ใช้งาน",
    "notes": "หา lark ไม่เจอ ครับ"
  },
  {
    "sn": 76,
    "date": "30/06/2569",
    "user": "ศุภฤกษ์ ภายไธสง (ดรีม)",
    "position": "Admin",
    "itemType": "Ipad",
    "deviceSerial": "iPad-003",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 358,
    "date": "",
    "user": "ทัศววรณ วัลย์ดาว (ใบตอง)",
    "position": "Admin",
    "itemType": "Ipad",
    "deviceSerial": "iPad-019",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 359,
    "date": "",
    "user": "ทัศววรณ วัลย์ดาว (ใบตอง)",
    "position": "Admin",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-004",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 360,
    "date": "",
    "user": "กานต์ฑิตา ธีระพิบูลย์ ( อิง )",
    "position": "Admin",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-009",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 28,
    "date": "",
    "user": "ปัญจมา สมบัติกำไร (เอิน)",
    "position": "Admin",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-014",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 47,
    "date": "",
    "user": "ทัศววรณ วัลย์ดาว (ใบตอง)",
    "position": "Admin",
    "itemType": "Notebook Acer, Iphone",
    "deviceSerial": "ACER-015",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 56,
    "date": "14/07/2569",
    "user": "เนตรปรีญา ทัดศรี",
    "position": "Receptionist",
    "itemType": "Macbook",
    "deviceSerial": "MacBookAir-009",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 91,
    "date": "07/09/3111",
    "user": "Receptionist",
    "position": "Receptionist",
    "itemType": "Ipad",
    "deviceSerial": "iPad-018",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 67,
    "date": "",
    "user": "Receptionist",
    "position": "Receptionist",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-007",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 51,
    "date": "",
    "user": "Receptionist",
    "position": "Receptionist",
    "itemType": "IMac",
    "deviceSerial": "MC-003",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 52,
    "date": "",
    "user": "Receptionist",
    "position": "Receptionist",
    "itemType": "IMac",
    "deviceSerial": "MC-004",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 334,
    "date": "16/07/2569",
    "user": "เนตรปรีญา ทัดศรี",
    "position": "Receptionist",
    "itemType": "Mouse",
    "deviceSerial": "MC-008",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 35,
    "date": "",
    "user": "วิกานดา ทุมมนตรี ( แป๋วแหว๋ว )",
    "position": "Payroll Officer",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-008",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 44,
    "date": "10/07/2569",
    "user": "อนิรุตต์ พานแสนซา",
    "position": "Live Streamer",
    "itemType": "Ipad",
    "deviceSerial": "iPad-016",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 85,
    "date": "26/06/2569",
    "user": "ธนันวลัญชน์ ศรีออน (อาโน)",
    "position": "Live Streamer",
    "itemType": "Ipad",
    "deviceSerial": "iPad-012",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 314,
    "date": "",
    "user": "ชนิสรา ตรีสัตยกุล (อุบอิบ)",
    "position": "Live Streamer",
    "itemType": "Ipad",
    "deviceSerial": "\niPad 017\n",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 61,
    "date": "",
    "user": "ชนิสรา ตรีสัตยกุล  (อุบอิบ)",
    "position": "Live Streamer",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-001",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 63,
    "date": "26/06/2569",
    "user": "ธนันวลัญชน์ ศรีออน (อาโน)",
    "position": "Live Streamer",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-003",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 285,
    "date": "",
    "user": "ชนิสรา ตรีสัตยกุล  (อุบอิบ)",
    "position": "Live Streamer",
    "itemType": "Adapter Apple",
    "deviceSerial": "ADT-007",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 287,
    "date": "",
    "user": "ชนิสรา ตรีสัตยกุล  (อุบอิบ)",
    "position": "Live Streamer",
    "itemType": "Cable Apple",
    "deviceSerial": "-",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 40,
    "date": "",
    "user": "อสมาภรณ์ ย่านเดิม ( เนม )",
    "position": "Personal Assistant to CEO",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-020",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 43,
    "date": "08/01/3111",
    "user": "สรุดตา ป๋อพริ้ง (ปาน)",
    "position": "Personal Assistant to CEO",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-023",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 57,
    "date": "16/07/2569",
    "user": "เนลินญาน์  ศิระไมตรีฉัตร (เนริน)",
    "position": "Personal Assistant to CEO",
    "itemType": "Macbook",
    "deviceSerial": "MacBookAir-027",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 7,
    "date": "",
    "user": "อสมาภรณ์ ย่านเดิม ( เนม )",
    "position": "Personal Assistant to CEO",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-010",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 1,
    "date": "",
    "user": "สุดธิดา เผ่าหอม (ต่าย) (KT)",
    "position": "HR",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-002",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 41,
    "date": "",
    "user": "กฤษณา ลำเพ็ง ( พลอย )",
    "position": "HR",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-038",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 42,
    "date": "",
    "user": "นารีรัตน์ ขันทอง ( ออย )",
    "position": "HR",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-022",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 74,
    "date": "",
    "user": "นารีรัตน์ ขันทอง ( ออย )",
    "position": "HR",
    "itemType": "Ipad",
    "deviceSerial": "iPad-001",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 88,
    "date": "10/07/2569",
    "user": "ศรัญญา ธรรมเนียมภักดี ( แนน )",
    "position": "HR",
    "itemType": "Ipad",
    "deviceSerial": "iPad-015",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 320,
    "date": "17/06/2569",
    "user": "นารีรัตน์ ขันทอง ( ออย )",
    "position": "HR",
    "itemType": "Ipad",
    "deviceSerial": "iPad-005",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 48,
    "date": "",
    "user": "ศรัญญา ธรรมเนียมภักดี ( แนน )",
    "position": "HR",
    "itemType": "Notebook HP",
    "deviceSerial": "HP-016",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 49,
    "date": "",
    "user": "ณัฐกานต์ ชิดปรางค์ (เตยหอม)",
    "position": "Photographer",
    "itemType": "IMac",
    "deviceSerial": "MC-001",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 368,
    "date": "",
    "user": "วิจิตราภรณ์ พึ่งจันดุม (พลอย)",
    "position": "Photographer",
    "itemType": "IMac",
    "deviceSerial": "MC-010",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 242,
    "date": "16/07/2569",
    "user": "ณัฐกานต์ ชิดปรางค์ (เตยหอม)",
    "position": "Photographer",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-MKT-001",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 75,
    "date": "",
    "user": "สรวิชญ์ สิทธิ ( บอม )",
    "position": "Sale",
    "itemType": "Ipad",
    "deviceSerial": "iPad-002",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 81,
    "date": "",
    "user": "อาทิตย์ สมการ ( เจ้านาย )",
    "position": "Sale",
    "itemType": "Ipad",
    "deviceSerial": "iPad-008",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 82,
    "date": "",
    "user": "สุนทรี บุญนาค ( ตาล )",
    "position": "Sale",
    "itemType": "Ipad",
    "deviceSerial": "iPad-009",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 95,
    "date": "",
    "user": "สุนทรี บุญนาค ( ตาล )",
    "position": "Sale",
    "itemType": "Apple Pancill",
    "deviceSerial": "Pencil-002",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 100,
    "date": "",
    "user": "สรวิชญ์ สิทธิ ( บอม )",
    "position": "Sale",
    "itemType": "Apple Pancill",
    "deviceSerial": "Pencil-007",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 6,
    "date": "",
    "user": "กฤติน วิชัยดิษฐ (อ้น)crm",
    "position": "CRM Officer",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-019",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 46,
    "date": "",
    "user": "กฤติญา ทาระพันธ์ (แจนนี่) crm",
    "position": "CRM Officer",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-026",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 321,
    "date": "18/05/2569",
    "user": "ชนันพร อินขำ (ไอซ์)crm",
    "position": "CRM Officer",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-003",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 361,
    "date": "",
    "user": "กฤติญา ทาระพันธ์ (แจนนี่) crm",
    "position": "CRM Officer",
    "itemType": "Samsung",
    "deviceSerial": "Samsung Galaxy-001",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 307,
    "date": "27/05/2569",
    "user": "ปัณณวิชญ์ ทองวัน (บอย)",
    "position": "Data Analysis",
    "itemType": "Computer (Pc)",
    "deviceSerial": "AS-001",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 309,
    "date": "",
    "user": "ปัณณวิชญ์ ทองวัน (บอย)",
    "position": "Data Analysis",
    "itemType": "Computer (Pc)",
    "deviceSerial": "LG-001",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 311,
    "date": "",
    "user": "ปัณณวิชญ์ ทองวัน (บอย)",
    "position": "Data Analysis",
    "itemType": "Computer (Pc)",
    "deviceSerial": "MSI-002",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 350,
    "date": "",
    "user": "ปัณณวิชญ์ ทองวัน (บอย)",
    "position": "Data Analysis",
    "itemType": "Notebook Asus",
    "deviceSerial": "Asus-024",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 223,
    "date": "",
    "user": "ปัณณวิชญ์ ทองวัน (บอย)",
    "position": "Data Analysis",
    "itemType": "Mornitor",
    "deviceSerial": "Asus-001",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 243,
    "date": "",
    "user": "ปัณณวิชญ์ ทองวัน (บอย)",
    "position": "Data Analysis",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-003",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 244,
    "date": "",
    "user": "ปัณณวิชญ์ ทองวัน (บอย)",
    "position": "Data Analysis",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-004",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 245,
    "date": "",
    "user": "ปัณณวิชญ์ ทองวัน (บอย)",
    "position": "Data Analysis",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-005",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 251,
    "date": "28/05/2569",
    "user": "ปัณณวิชญ์ ทองวัน (บอย)",
    "position": "Data Analysis",
    "itemType": "Keyboard",
    "deviceSerial": "KBD-002",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 262,
    "date": "",
    "user": "ปัณณวิชญ์ ทองวัน (บอย)",
    "position": "Data Analysis",
    "itemType": "Keyboard",
    "deviceSerial": "-",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 337,
    "date": "04/06/2569",
    "user": "อภิสิทธิ์ พรจันทราวัฒน์ (จุ้ย)",
    "position": "Video Content Tiktok",
    "itemType": "IMac",
    "deviceSerial": "MC-005",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 310,
    "date": "",
    "user": "พงศกร ผ่องใส (ปอน)",
    "position": "Live Producer",
    "itemType": "Computer (Pc)",
    "deviceSerial": "LG-002",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 266,
    "date": "09/06/2569",
    "user": "พิสิษฐ์ มงคลสมบัติศิริ (เจมส์)",
    "position": "Live Producer",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-021",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 280,
    "date": "",
    "user": "พงศกร ผ่องใส (ปอน)",
    "position": "Live Producer",
    "itemType": "Mornitor",
    "deviceSerial": "MSI-001",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 367,
    "date": "",
    "user": "นายพงศกร ผ่องใส",
    "position": "Live Producer",
    "itemType": "Mornitor",
    "deviceSerial": "LG-006",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 30,
    "date": "01/06/2569",
    "user": "เย็นฤดี มาระวัง (ฝ้าย)",
    "position": "Live Producer",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-013",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 31,
    "date": "05/05/2569",
    "user": "พงศกร ผ่องใส (ปอน)",
    "position": "Live Producer",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-039",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 343,
    "date": "",
    "user": "ทีม live Producer",
    "position": "Live Producer",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "usb-a to c",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 240,
    "date": "",
    "user": "พงศกร ผ่องใส (ปอน)",
    "position": "Live Producer",
    "itemType": "Capture Card",
    "deviceSerial": "VCS-LIVE-0001",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 253,
    "date": "",
    "user": "พงศกร ผ่องใส (ปอน)",
    "position": "Live Producer",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-007",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 342,
    "date": "",
    "user": "ทีม live Producer",
    "position": "Live Producer",
    "itemType": "Cable HDMI",
    "deviceSerial": "hdmi Canon",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 315,
    "date": "",
    "user": "พงศกร ผ่องใส (ปอน)",
    "position": "Live Producer",
    "itemType": "Reez Live",
    "deviceSerial": "ไม่มีหมายเลข อุปกรณ์",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 344,
    "date": "",
    "user": "ทีม live Producer",
    "position": "Live Producer",
    "itemType": "อุปกรณ์เสริม/อื่นๆ",
    "deviceSerial": "สาย aux branding",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 345,
    "date": "",
    "user": "ทีม live Producer",
    "position": "Live Producer",
    "itemType": "อุปกรณ์เสริม/อื่นๆ",
    "deviceSerial": "hdmi ต่อจอมอนิเตอร์ไลฟ์",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 346,
    "date": "",
    "user": "ทีม live Producer",
    "position": "Live Producer",
    "itemType": "อุปกรณ์เสริม/อื่นๆ",
    "deviceSerial": "hdmi it-002",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 347,
    "date": "",
    "user": "ทีม live Producer",
    "position": "Live Producer",
    "itemType": "อุปกรณ์เสริม/อื่นๆ",
    "deviceSerial": "adc-006",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 348,
    "date": "",
    "user": "ทีม live Producer",
    "position": "Live Producer",
    "itemType": "อุปกรณ์เสริม/อื่นๆ",
    "deviceSerial": "ไมค์ล่อย dji-branding",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 349,
    "date": "",
    "user": "ทีม live Producer",
    "position": "Live Producer",
    "itemType": "อุปกรณ์เสริม/อื่นๆ",
    "deviceSerial": "ไมค์ล่อย dji-mkt",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 235,
    "date": "",
    "user": "ห้องประชุม บ้าน 18 ชั้น 4",
    "position": "Meeting Room",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-MKT-004",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 54,
    "date": "",
    "user": "อภิญญา ศรีตะวัน ( บอสต้น )",
    "position": "Boss",
    "itemType": "IMac",
    "deviceSerial": "MC-006",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 55,
    "date": "",
    "user": "อภิญญา ศรีตะวัน ( บอสต้น )",
    "position": "Boss",
    "itemType": "IMac",
    "deviceSerial": "MC-007",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 232,
    "date": "",
    "user": "อภิญญา ศรีตะวัน ( บอสต้น )",
    "position": "Boss",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-ฺBoss-001",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 241,
    "date": "",
    "user": "อภิญญา ศรีตะวัน ( บอสต้น )",
    "position": "Boss",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-001",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 233,
    "date": "",
    "user": "ห้อง Studio บ้าน 18 ชั้น 1",
    "position": "Studio",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-Studio-001",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 277,
    "date": "26/12/2568",
    "user": "นนท์สิรี ปลื้มทรัพย์ (ใบพลู)",
    "position": "Data Entry",
    "itemType": "Notebook Asus",
    "deviceSerial": "Asus-041",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 29,
    "date": "",
    "user": "ปภาวี จันทร์ขวาง (ครีม)",
    "position": "Data Entry",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-012",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 125,
    "date": "16/06/2569",
    "user": "อารยา ธนพันธุ์พาณิชย์ (หนุงหนิง)",
    "position": "Sale Manager",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-032",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 363,
    "date": "19/01/2569",
    "user": "ณัฐณิชา ศรีวรอรรถิกุล (ใบตอง)",
    "position": "KOL & Event Marketing",
    "itemType": "Iphone",
    "deviceSerial": "IPhone -005",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 283,
    "date": "19/01/2569",
    "user": "ณัฐณิชา ศรีวรอรรถิกุล (ใบตอง)",
    "position": "KOL & Event Marketing",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "Lenovo-001",
    "status": "ใช้งาน",
    "notes": "รอตรวจสอบ"
  },
  {
    "sn": 45,
    "date": "",
    "user": "พัทธนันท์ นาเดียร์  (นาเดียร์ )",
    "position": "OR",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-025",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 60,
    "date": "",
    "user": "หมอฟาง",
    "position": "CEO",
    "itemType": "Macbook",
    "deviceSerial": "MacBookAir-030",
    "status": "ใช้งาน",
    "notes": "     "
  },
  {
    "sn": 215,
    "date": "",
    "user": "หมอฟาง",
    "position": "CEO",
    "itemType": "Macbook",
    "deviceSerial": "MAcbook Pro-034",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 269,
    "date": "",
    "user": "หมอฟาง",
    "position": "CEO",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-MKT-005",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 236,
    "date": "",
    "user": "Connection Nas",
    "position": "Center Storage",
    "itemType": "External HDD",
    "deviceSerial": "Segate Station 10 TB",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 237,
    "date": "",
    "user": "Connection Nas",
    "position": "Center Storage",
    "itemType": "External HDD",
    "deviceSerial": "ETN-001  WD My Book 4TB",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 254,
    "date": "",
    "user": "Connection TV",
    "position": "TV Center",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-006",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 271,
    "date": "",
    "user": "Connection TV",
    "position": "TV Center",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-008",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 291,
    "date": "",
    "user": "Connection TV",
    "position": "TV Center",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-009",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 292,
    "date": "",
    "user": "Connection TV",
    "position": "TV Center",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-010",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 340,
    "date": "06/07/2569",
    "user": "นายสุทธิรงค์ เครือไพบูลย์กุล ( เบนซ์ )",
    "position": "Senior Digital Marketing",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-042",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 369,
    "date": "10/07/2569",
    "user": "สุรวิชญ์ โพธิ์ตาก (ไมค์เมโลดี้)",
    "position": "Operation Department",
    "itemType": "Macbook",
    "deviceSerial": "Macbook-044",
    "status": "ใช้งาน",
    "notes": ""
  },
  {
    "sn": 308,
    "date": "",
    "user": "ส่วนกลาง/ไม่ระบุ",
    "position": "-",
    "itemType": "Computer (Pc)",
    "deviceSerial": "LG-003",
    "status": "ว่าง",
    "notes": ""
  },
  {
    "sn": 264,
    "date": "",
    "user": "ส่วนกลาง/ไม่ระบุ",
    "position": "-",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-036",
    "status": "ว่าง",
    "notes": ""
  },
  {
    "sn": 316,
    "date": "",
    "user": "ส่วนกลาง/ไม่ระบุ",
    "position": "-",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-018",
    "status": "ว่าง",
    "notes": ""
  },
  {
    "sn": 328,
    "date": "",
    "user": "ส่วนกลาง/ไม่ระบุ",
    "position": "-",
    "itemType": "Macbook",
    "deviceSerial": "MacBook air-035",
    "status": "ว่าง",
    "notes": ""
  },
  {
    "sn": 50,
    "date": "",
    "user": "ส่วนกลาง/ไม่ระบุ",
    "position": "-",
    "itemType": "Ipad",
    "deviceSerial": "iPad-004",
    "status": "ว่าง",
    "notes": ""
  },
  {
    "sn": 80,
    "date": "",
    "user": "ส่วนกลาง/ไม่ระบุ",
    "position": "-",
    "itemType": "Ipad",
    "deviceSerial": "iPad-007",
    "status": "ว่าง",
    "notes": ""
  },
  {
    "sn": 84,
    "date": "",
    "user": "ส่วนกลาง/ไม่ระบุ",
    "position": "-",
    "itemType": "Ipad",
    "deviceSerial": "iPad-011",
    "status": "ว่าง",
    "notes": ""
  },
  {
    "sn": 131,
    "date": "",
    "user": "ส่วนกลาง/ไม่ระบุ",
    "position": "-",
    "itemType": "Ipad",
    "deviceSerial": "iPad-006",
    "status": "ว่าง",
    "notes": ""
  },
  {
    "sn": 117,
    "date": "",
    "user": "ส่วนกลาง/ไม่ระบุ",
    "position": "-",
    "itemType": "Printer",
    "deviceSerial": "PT-001",
    "status": "รอซ่อม",
    "notes": ""
  },
  {
    "sn": 230,
    "date": "",
    "user": "ส่วนกลาง/ไม่ระบุ",
    "position": "-",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-MKT-003",
    "status": "ว่าง",
    "notes": ""
  },
  {
    "sn": 335,
    "date": "",
    "user": "ส่วนกลาง/ไม่ระบุ",
    "position": "-",
    "itemType": "Mouse",
    "deviceSerial": "MC-009",
    "status": "ว่าง",
    "notes": ""
  },
  {
    "sn": 261,
    "date": "",
    "user": "ส่วนกลาง/ไม่ระบุ",
    "position": "-",
    "itemType": "Keyboard",
    "deviceSerial": "KBD-001",
    "status": "ว่าง",
    "notes": ""
  }
];

const initialDashboardData = {
  "2025-11": {
    "monthName": "พฤศจิกายน 2568",
    "totalAssets": 162,
    "assetValue": 2561390,
    "assetsExpiring": 5,
    "assetsBroken": 8,
    "assetsLost": 0,
    "ticketsCount": 7,
    "slaPercent": 100,
    "responseTime": 6,
    "resolutionTime": 0.5,
    "csat": 4.9,
    "totalSoftware": 24,
    "licensesInUse": 235.5,
    "licensesVacant": 35,
    "softwareCost": 105975,
    "softwareExpiring": 2,
    "backupSuccess": 99.9,
    "securityIncidents": 0,
    "antivirusCoverage": 98.5,
    "mfaCoverage": 100,
    "repairCount": 7,
    "repairCost": 0,
    "topBrokenDevices": [
      {
        "name": "Notebook",
        "count": 3,
        "cost": 0
      },
      {
        "name": "Printer",
        "count": 3,
        "cost": 0
      },
      {
        "name": "PC Computer",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Smartphone",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Mac",
        "count": 0,
        "cost": 0
      },
      {
        "name": "iPad",
        "count": 0,
        "cost": 0
      },
      {
        "name": "CCTV",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Mouse",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Keyboard",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Router",
        "count": 0,
        "cost": 0
      }
    ],
    "deptCosts": {
      "Accounting & Finance": 0,
      "Sales & Marketing": 0,
      "Human Resources": 0,
      "Operations": 0,
      "IT & Infrastructure": 0
    },
    "softwareExpiringDetails": [
      {
        "name": "Microsoft 365 Copilot",
        "licenses": 50,
        "expiringDate": "30 ส.ค. 2026",
        "status": "ใกล้หมดอายุ"
      },
      {
        "name": "Adobe Creative Cloud",
        "licenses": 15,
        "expiringDate": "12 ก.ย. 2026",
        "status": "แจ้งเตือนล่วงหน้า"
      }
    ],
    "assetsExpiringDetails": [
      {
        "id": "AST-NB-001",
        "type": "Laptop",
        "model": "Lenovo ThinkPad L14",
        "dept": "Operations",
        "expDate": "15 ส.ค. 2026"
      },
      {
        "id": "AST-PR-004",
        "type": "Printer",
        "model": "HP LaserJet Pro M404",
        "dept": "Accounting",
        "expDate": "22 ส.ค. 2026"
      }
    ],
    "ongoingProjects": [
      {
        "title": "IT Ticket Automation",
        "desc": "พัฒนาสคริปต์ช่วยจัดการปัญหาซ้ำซาก คืบหน้า 45%"
      },
      {
        "title": "Asset Management System",
        "desc": "ระบบเช็คอิน-เช็คเอาท์อุปกรณ์ไอที คืบหน้า 60%"
      }
    ],
    "recommendations": [
      "พบปัญหาอุปกรณ์ขัดข้องสูงถึง 7 ครั้งในเดือนนี้ แนะนำจัดรอบบำรุงรักษาเชิงป้องกัน (Preventive Maintenance) โดยเฉพาะอุปกรณ์ประเภท Notebook",
      "การสนับสนุนผู้ใช้ระบบ IT อยู่ในเกณฑ์ดีเยี่ยม สามารถรักษามาตรฐาน SLA ได้ตามเป้าหมาย",
      "แนะนำให้ผู้ใช้อัปเกรดความปลอดภัยของบัญชีผ่านการเปิดใช้งาน MFA ครบ 100%"
    ],
    "ticketsList": [
      {
        "sn": 1,
        "date": "11/11/2568 10:29",
        "complainant": "ปัณณวิชญ์   ทองวัน",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:10",
        "responder": "ปัณณวิชญ์ สิริภานุพัฒน์ (บอย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 2,
        "date": "11/11/2568 14:04",
        "complainant": "นิธิดา รัตนอาภรณ์",
        "email": "nitida.rnp@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: ์Adobe",
        "cause": "-",
        "duration": "00:30",
        "responder": "นิธิดา รัตนอาภรณ์ (เตย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 3,
        "date": "18/11/2568 15:49",
        "complainant": "พรพิมล เขียวจันทร์",
        "email": "phonpimonwork@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows",
        "cause": "-",
        "duration": "00:50",
        "responder": "ผู้ใช้รับเชิญ 95166",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 4,
        "date": "19/11/2568 10:09",
        "complainant": "กฤษณา ลำเพ็ง",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "00:50",
        "responder": "กฤษณา ลำเพ็ง (พลอย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 5,
        "date": "19/11/2568 10:35",
        "complainant": "นายกฤติน  วิชันดิษฐ ",
        "email": "Tin_Krittinwichaidit@hotmail.com  ",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "00:30",
        "responder": "กฤติน วิชัยดิษฐ  (อ้น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 6,
        "date": "20/11/2568 11:01",
        "complainant": "เนตรปรีญา ทัดศรี",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, เน็ตเวิร์ค: เชื่อมต่อ บลูทูธ ไม่ได้",
        "cause": "-",
        "duration": "00:30",
        "responder": "เนตรปรีญา ทัดศรี (แตงกวา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 7,
        "date": "26/11/2568 15:58",
        "complainant": "สุภาพ  แสนจันทร์ ",
        "email": "fernclinic.acc@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, Printer",
        "cause": "-",
        "duration": "00:10",
        "responder": "สุภาพ แสนจันทร์ (ส้ม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      }
    ],
    "automationsDone": 1,
    "aiApps": 2,
    "hoursSaved": 15
  },
  "2025-12": {
    "monthName": "ธันวาคม 2568",
    "totalAssets": 162,
    "assetValue": 2561390,
    "assetsExpiring": 5,
    "assetsBroken": 10,
    "assetsLost": 0,
    "ticketsCount": 7,
    "slaPercent": 100,
    "responseTime": 8,
    "resolutionTime": 0.7,
    "csat": 4.9,
    "totalSoftware": 24,
    "licensesInUse": 235.5,
    "licensesVacant": 35,
    "softwareCost": 105975,
    "softwareExpiring": 2,
    "backupSuccess": 99.9,
    "securityIncidents": 0,
    "antivirusCoverage": 98.5,
    "mfaCoverage": 100,
    "repairCount": 9,
    "repairCost": 0,
    "topBrokenDevices": [
      {
        "name": "Printer",
        "count": 4,
        "cost": 0
      },
      {
        "name": "Notebook",
        "count": 3,
        "cost": 0
      },
      {
        "name": "PC Computer",
        "count": 2,
        "cost": 0
      },
      {
        "name": "Smartphone",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Mac",
        "count": 0,
        "cost": 0
      },
      {
        "name": "iPad",
        "count": 0,
        "cost": 0
      },
      {
        "name": "CCTV",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Mouse",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Keyboard",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Router",
        "count": 0,
        "cost": 0
      }
    ],
    "deptCosts": {
      "Accounting & Finance": 0,
      "Sales & Marketing": 0,
      "Human Resources": 0,
      "Operations": 0,
      "IT & Infrastructure": 0
    },
    "softwareExpiringDetails": [
      {
        "name": "Microsoft 365 Copilot",
        "licenses": 50,
        "expiringDate": "30 ส.ค. 2026",
        "status": "ใกล้หมดอายุ"
      },
      {
        "name": "Adobe Creative Cloud",
        "licenses": 15,
        "expiringDate": "12 ก.ย. 2026",
        "status": "แจ้งเตือนล่วงหน้า"
      }
    ],
    "assetsExpiringDetails": [
      {
        "id": "AST-NB-001",
        "type": "Laptop",
        "model": "Lenovo ThinkPad L14",
        "dept": "Operations",
        "expDate": "15 ส.ค. 2026"
      },
      {
        "id": "AST-PR-004",
        "type": "Printer",
        "model": "HP LaserJet Pro M404",
        "dept": "Accounting",
        "expDate": "22 ส.ค. 2026"
      }
    ],
    "ongoingProjects": [
      {
        "title": "IT Ticket Automation",
        "desc": "พัฒนาสคริปต์ช่วยจัดการปัญหาซ้ำซาก คืบหน้า 45%"
      },
      {
        "title": "Asset Management System",
        "desc": "ระบบเช็คอิน-เช็คเอาท์อุปกรณ์ไอที คืบหน้า 60%"
      }
    ],
    "recommendations": [
      "พบปัญหาอุปกรณ์ขัดข้องสูงถึง 9 ครั้งในเดือนนี้ แนะนำจัดรอบบำรุงรักษาเชิงป้องกัน (Preventive Maintenance) โดยเฉพาะอุปกรณ์ประเภท Printer",
      "การสนับสนุนผู้ใช้ระบบ IT อยู่ในเกณฑ์ดีเยี่ยม สามารถรักษามาตรฐาน SLA ได้ตามเป้าหมาย",
      "แนะนำให้ผู้ใช้อัปเกรดความปลอดภัยของบัญชีผ่านการเปิดใช้งาน MFA ครบ 100%"
    ],
    "ticketsList": [
      {
        "sn": 8,
        "date": "02/12/2568 10:49",
        "complainant": "นางสาวสุพรรษา อินทะเรืองรุ่ง",
        "email": "Supans.si90@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, Printer, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 77838",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 9,
        "date": "02/12/2568 13:21",
        "complainant": "โชตินันท์ ณ นคร",
        "email": "nan.chotinan@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, Printer, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "00:40",
        "responder": "ผู้ใช้รับเชิญ 24619",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 10,
        "date": "04/12/2568 13:05",
        "complainant": "โชตินันท์ ณ นคร",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, บัญชีผู้ใช้: เชื่อมต่อ Nas ",
        "cause": "-",
        "duration": "00:50",
        "responder": "ผู้ใช้รับเชิญ 81829",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 11,
        "date": "12/12/2568 10:25",
        "complainant": "นางสาวอัจฉรา เหรียญพิมาย",
        "email": "atchararianpimai@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "00:50",
        "responder": "ผู้ใช้รับเชิญ 94307",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 12,
        "date": "12/12/2568 12:43",
        "complainant": "นภัสสร นาสวน",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "00:40",
        "responder": "ผู้ใช้รับเชิญ 47307",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 13,
        "date": "15/12/2568 16:14",
        "complainant": "โชตินันท์ ณ นคร",
        "email": "nan.chotinan@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows",
        "cause": "-",
        "duration": "00:40",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 14,
        "date": "19/12/2568 10:17",
        "complainant": "พรพิมล เขียวจันทร์",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: ์Adobe, Windows",
        "cause": "-",
        "duration": "01:00",
        "responder": "ผู้ใช้รับเชิญ 22143",
        "status": "เสร็จสิ้น",
        "cost": 0
      }
    ],
    "automationsDone": 1,
    "aiApps": 2,
    "hoursSaved": 15
  },
  "2026-01": {
    "monthName": "มกราคม 2569",
    "totalAssets": 162,
    "assetValue": 2561390,
    "assetsExpiring": 6,
    "assetsBroken": 57,
    "assetsLost": 1,
    "ticketsCount": 64,
    "slaPercent": 89.1,
    "responseTime": 10,
    "resolutionTime": 0.8,
    "csat": 4.9,
    "totalSoftware": 24,
    "licensesInUse": 321,
    "licensesVacant": 48,
    "softwareCost": 144450,
    "softwareExpiring": 2,
    "backupSuccess": 99.9,
    "securityIncidents": 0,
    "antivirusCoverage": 98.5,
    "mfaCoverage": 100,
    "repairCount": 56,
    "repairCost": 0,
    "topBrokenDevices": [
      {
        "name": "Notebook",
        "count": 15,
        "cost": 0
      },
      {
        "name": "iPad",
        "count": 12,
        "cost": 0
      },
      {
        "name": "Smartphone",
        "count": 11,
        "cost": 0
      },
      {
        "name": "PC Computer",
        "count": 10,
        "cost": 0
      },
      {
        "name": "Mac",
        "count": 4,
        "cost": 0
      },
      {
        "name": "Printer",
        "count": 3,
        "cost": 0
      },
      {
        "name": "CCTV",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Mouse",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Keyboard",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Router",
        "count": 0,
        "cost": 0
      }
    ],
    "deptCosts": {
      "Accounting & Finance": 0,
      "Sales & Marketing": 0,
      "Human Resources": 0,
      "Operations": 0,
      "IT & Infrastructure": 0
    },
    "softwareExpiringDetails": [
      {
        "name": "Microsoft 365 Copilot",
        "licenses": 50,
        "expiringDate": "30 ส.ค. 2026",
        "status": "ใกล้หมดอายุ"
      },
      {
        "name": "Adobe Creative Cloud",
        "licenses": 15,
        "expiringDate": "12 ก.ย. 2026",
        "status": "แจ้งเตือนล่วงหน้า"
      }
    ],
    "assetsExpiringDetails": [
      {
        "id": "AST-NB-001",
        "type": "Laptop",
        "model": "Lenovo ThinkPad L14",
        "dept": "Operations",
        "expDate": "15 ส.ค. 2026"
      },
      {
        "id": "AST-PR-004",
        "type": "Printer",
        "model": "HP LaserJet Pro M404",
        "dept": "Accounting",
        "expDate": "22 ส.ค. 2026"
      }
    ],
    "ongoingProjects": [
      {
        "title": "IT Ticket Automation",
        "desc": "พัฒนาสคริปต์ช่วยจัดการปัญหาซ้ำซาก คืบหน้า 70%"
      },
      {
        "title": "Asset Management System",
        "desc": "ระบบเช็คอิน-เช็คเอาท์อุปกรณ์ไอที คืบหน้า 60%"
      }
    ],
    "recommendations": [
      "พบปัญหาอุปกรณ์ขัดข้องสูงถึง 56 ครั้งในเดือนนี้ แนะนำจัดรอบบำรุงรักษาเชิงป้องกัน (Preventive Maintenance) โดยเฉพาะอุปกรณ์ประเภท Notebook",
      "อัตราการบรรลุเป้าหมาย SLA ลดลงเหลือ 89.1% แนะนำให้ปรับกระบวนการคัดกรอง Ticket เพื่อเพิ่มความรวดเร็วในการแก้ปัญหา",
      "แนะนำให้ผู้ใช้อัปเกรดความปลอดภัยของบัญชีผ่านการเปิดใช้งาน MFA ครบ 100%"
    ],
    "ticketsList": [
      {
        "sn": 15,
        "date": "05/01/2569 11:03",
        "complainant": "นาเดียร์",
        "email": "deerblink262538@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 31440",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 16,
        "date": "05/01/2569 11:06",
        "complainant": "นางสาวสุพรรษา อินทะเรืองรุ่ง",
        "email": "Supansa.si90@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer, ซอฟต์แวร์: Windows, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "00:30",
        "responder": "สุพรรษา อินทะเรืองรุ่ง (ต่าย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 17,
        "date": "06/01/2569 09:52",
        "complainant": "นางสาวอังคณา ธงศรี",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:10",
        "responder": "Ampol",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 18,
        "date": "07/01/2569 10:50",
        "complainant": "ขวัญลออ นวลละออง",
        "email": "kikkazoo@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook",
        "cause": "-",
        "duration": "00:10",
        "responder": "Kwanlaoa Nuanlaong",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 19,
        "date": "08/01/2569 09:38",
        "complainant": "ชลธิชา ตาลพันธ์ นิชา",
        "email": "chonthicha.talpun1998@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows, Microsoft Office",
        "cause": "-",
        "duration": "01:00",
        "responder": "ผู้ใช้รับเชิญ 65327",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 20,
        "date": "08/01/2569 10:03",
        "complainant": "ดลพร อุลุชาฎะ",
        "email": "paus.absolute@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Mac, ซอฟต์แวร์: Microsoft Office",
        "cause": "-",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 48018",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 21,
        "date": "08/01/2569 10:42",
        "complainant": "อนุสรา สิมจันทา",
        "email": "anusara.a43@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: IOS, บัญชีผู้ใช้: Tiktok",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 24821",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 22,
        "date": "08/01/2569 13:55",
        "complainant": "ปัณณวิชญ์   ทองวัน (บอย)",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "cause": "-",
        "duration": "00:30",
        "responder": "ปัณณวิชญ์ สิริภานุพัฒน์ (บอย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 23,
        "date": "08/01/2569 14:52",
        "complainant": "กรรณิกา ค่ำคูณ",
        "email": " ",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: IOS, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 35630",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 24,
        "date": "08/01/2569 14:54",
        "complainant": "ฐานิสา ศรีจันทร์โคตร",
        "email": "thanisaa148@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 63311",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 25,
        "date": "12/01/2569 09:04",
        "complainant": "พิชยา ฮงทอง",
        "email": "Pichaya24ht@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Google Link",
        "cause": "-",
        "duration": "00:30",
        "responder": "พิชยา ฮงทอง (แนน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 26,
        "date": "12/01/2569 09:51",
        "complainant": "อังคณา ธงศรี (ใบเฟิร์น/บัญชี)",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Microsoft Office",
        "cause": "-",
        "duration": "00:20",
        "responder": "Ampol",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 27,
        "date": "12/01/2569 10:44",
        "complainant": "ปาหนัน สุพรม",
        "email": "panan6183@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 55719",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 29,
        "date": "15/01/2569 15:01",
        "complainant": "ขวัญลออ นวลละออง",
        "email": "kikkazoo@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook",
        "cause": "-",
        "duration": "00:30",
        "responder": "Kwanlaoa Nuanlaong",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 30,
        "date": "16/01/2569 09:34",
        "complainant": "อัจฉรา เหรียญพิมาย (โบกี้)",
        "email": "atchararianpimai@gmial.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows",
        "cause": "-",
        "duration": "00:30",
        "responder": "อัจฉรา เหรียญพิมาย (Bogie)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 31,
        "date": "16/01/2569 12:13",
        "complainant": "อนุสรา สิมจันทา",
        "email": "aanusara.a43@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: IOS",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 62685",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 32,
        "date": "19/01/2569 11:16",
        "complainant": "กฤติญา ทาระพันธ์",
        "email": "Krittiya.trp@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, Smartphone, ซอฟต์แวร์: Google Link, IOS, บัญชีผู้ใช้: Tiktok, Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 76985",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 33,
        "date": "20/01/2569 16:02",
        "complainant": "ณัฐณิชา ศรีวรอรรถิกุล",
        "email": "Natniploy99@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: เชื่อมต่อ บลูทูธ ไม่ได้, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "00:40",
        "responder": "ผู้ใช้รับเชิญ 73645",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 34,
        "date": "21/01/2569 09:41",
        "complainant": "เอกรินทร์ จีนเพชร",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer",
        "cause": "-",
        "duration": "00:20",
        "responder": "Guest User 75941",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 35,
        "date": "21/01/2569 16:13",
        "complainant": "อนุสรา สิมจันทา",
        "email": "aanusara.a43@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 13173",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 36,
        "date": "22/01/2569 12:15",
        "complainant": "สุภาพ แสนจันทร์ ",
        "email": "fernclinic.acc@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: Printer WIFI, Email",
        "cause": "-",
        "duration": "00:50",
        "responder": "สุภาพ แสนจันทร์ (ส้ม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 40,
        "date": "22/01/2569 15:31",
        "complainant": "เฟิร์น",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad, บัญชีผู้ใช้: Tiktok",
        "cause": "-",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 42,
        "date": "22/01/2569 15:34",
        "complainant": "ต่าย HR",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: CCTV, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 43,
        "date": "22/01/2569 15:50",
        "complainant": "ครีม",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS",
        "cause": "-",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 44,
        "date": "22/01/2569 15:51",
        "complainant": "พี่ส้ม",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 46,
        "date": "23/01/2569 09:31",
        "complainant": "อำพล   แซ่แฮ",
        "email": "fernclinic.it@gmail.com",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 48,
        "date": "23/01/2569 09:38",
        "complainant": "ปภาวิน อักโขสุวรรณ",
        "email": "sims4youth@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 16976",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 49,
        "date": "23/01/2569 10:25",
        "complainant": "นนทภัทร์  พึ่งพุ่ม",
        "email": "nonthapat.p1994@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: IOS, บัญชีผู้ใช้: Tiktok",
        "cause": "-",
        "duration": "00:35",
        "responder": "Guest User 89423",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 50,
        "date": "23/01/2569 13:12",
        "complainant": "ขวัญลออ นวลละออง",
        "email": "kikkazoo@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook",
        "cause": "-",
        "duration": "00:50",
        "responder": "Kwanlaoa Nuanlaong",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 51,
        "date": "23/01/2569 13:19",
        "complainant": "กฤติญา ทาระพันธ์",
        "email": "krittiya.trp@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 76985",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 52,
        "date": "23/01/2569 13:31",
        "complainant": "ณัฏฐ์ชาวีร์ หิรัญรัชชากุล",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: Calenda",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 88787",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 53,
        "date": "23/01/2569 13:59",
        "complainant": "อำพล   แซ่แฮ",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Mac, ซอฟต์แวร์: IOS, Windows",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 54,
        "date": "26/01/2569 08:31",
        "complainant": "อำพล  แซ่แฮ",
        "email": "fernclinic.it@gmail.com",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 55,
        "date": "26/01/2569 08:45",
        "complainant": "ภัทรศยา ไชยคุณ",
        "email": "bampattarsaya@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 94075",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 57,
        "date": "26/01/2569 16:06",
        "complainant": "สุดธิดา เผ่าหอม",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "cause": "-",
        "duration": "01:00",
        "responder": "สุดธิดา เผ่าหอม (ต่าย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 58,
        "date": "26/01/2569 16:56",
        "complainant": "พี่บี จัดซื้อ",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 59,
        "date": "27/01/2569 08:36",
        "complainant": "อำพล  แซ่แฮ",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 29120",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 60,
        "date": "27/01/2569 09:05",
        "complainant": "เบนซ์",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Mac, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 61,
        "date": "27/01/2569 10:05",
        "complainant": "ทัศวรรณ วัลย์ดาว ",
        "email": "tassawan.s240@pnru.ac.th",
        "anydesk": "-",
        "issue": "ซอฟต์แวร์: CCTV",
        "cause": "-",
        "duration": "03:00",
        "responder": "ผู้ใช้รับเชิญ 78271",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 62,
        "date": "27/01/2569 13:50",
        "complainant": "กฤติมา สอนพู",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 78411",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 63,
        "date": "27/01/2569 13:55",
        "complainant": "ภัทรศยา ไชยคุณ",
        "email": "bampattarasaya@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Smartphone, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 94075",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 64,
        "date": "27/01/2569 14:38",
        "complainant": "บอม",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 65,
        "date": "27/01/2569 14:48",
        "complainant": "พิชยา ฮงทอง (แนน)",
        "email": "Pichaya24ht@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, Printer, ซอฟต์แวร์: Microsoft Office, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้",
        "cause": "-",
        "duration": "00:45",
        "responder": "พิชยา ฮงทอง (แนน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 66,
        "date": "28/01/2569 08:33",
        "complainant": "อำพล  แซ่แฮ",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 68,
        "date": "28/01/2569 08:38",
        "complainant": "ทีมไลฟ์",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: เชื่อมต่อ Nas ไม่ได้, บัญชีผู้ใช้: เชื่อมต่อ Nas , G-Suit",
        "cause": "-",
        "duration": "03:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 69,
        "date": "28/01/2569 09:34",
        "complainant": "HR",
        "email": "-",
        "anydesk": "-",
        "issue": "ซอฟต์แวร์: CCTV",
        "cause": "-",
        "duration": "00:40",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 70,
        "date": "29/01/2569 14:14",
        "complainant": "Ing Admin",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 71,
        "date": "29/01/2569 14:16",
        "complainant": "Prem",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 72,
        "date": "29/01/2569 16:40",
        "complainant": "อภิสิทธิ์ พรจันทราวัฒน์",
        "email": "apisitj2545@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: G-Suit",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 43809",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 73,
        "date": "29/01/2569 16:48",
        "complainant": "IT",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: google dive",
        "cause": "-",
        "duration": "05:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 74,
        "date": "30/01/2569 12:11",
        "complainant": "ฐานิสา ศรีจีนทร์โคตร",
        "email": "-",
        "anydesk": "-",
        "issue": "บัญชีผู้ใช้: google dive",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 86927",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 75,
        "date": "30/01/2569 13:14",
        "complainant": "กฤติน วิชัยดิษฐ  ",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: google dive",
        "cause": "-",
        "duration": "00:20",
        "responder": "กฤติน วิชัยดิษฐ  (อ้น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 76,
        "date": "30/01/2569 13:32",
        "complainant": "IT",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: google dive",
        "cause": "-",
        "duration": "03:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 77,
        "date": "30/01/2569 13:35",
        "complainant": "พี่พอส",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Mac, บัญชีผู้ใช้: Click Up",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 79,
        "date": "31/01/2569 08:26",
        "complainant": "นนทภัทร์ พึ่งพุ่ม (หมู)",
        "email": "nonthapat.p1994@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "Guest User 75073",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 80,
        "date": "31/01/2569 08:32",
        "complainant": "เอกรินทร์ จีนเพชร(เอก)",
        "email": "akekarin.jee@gmail.com",
        "anydesk": "-",
        "issue": "เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "cause": "-",
        "duration": "02:00",
        "responder": "ผู้ใช้รับเชิญ 90720",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 81,
        "date": "31/01/2569 09:29",
        "complainant": "Bam",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: Calenda, G-Suit, บัญชีผู้ใช้: google dive",
        "cause": "-",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 82,
        "date": "31/01/2569 09:40",
        "complainant": "ฐิตารีย์ นรกุลศิริภักดี",
        "email": "Fiiefern@hotmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: เชื่อมต่อ Server ",
        "cause": "-",
        "duration": "00:30",
        "responder": "ฐิตารีย์  นรกุลศิริภักดี (เฟิร์น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 83,
        "date": "31/01/2569 09:42",
        "complainant": "ต่าย กราฟิก",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: G-Suit",
        "cause": "-",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 84,
        "date": "31/01/2569 09:47",
        "complainant": "ทีม DATA",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 85,
        "date": "31/01/2569 09:52",
        "complainant": "พรพิมล เขียวจันทร์",
        "email": "phonpimonwork@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows",
        "cause": "-",
        "duration": "03:00",
        "responder": "ผู้ใช้รับเชิญ 14598",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 86,
        "date": "31/01/2569 12:25",
        "complainant": "ชลธิชา ตาลพันธ์",
        "email": "chonthicha.talpun1998@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, iPad, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้, Calenda, เชื่อมต่อ บลูทูธ ไม่ได้, บัญชีผู้ใช้: Email, google dive, เชื่อมต่อ Server ",
        "cause": "-",
        "duration": "02:00",
        "responder": "ชลธิชา ตาลพันธ์ (นิชา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 87,
        "date": "31/01/2569 13:36",
        "complainant": "นนทภัทร์ พึ่งพุ่มา(หมู) TikTok content creators ",
        "email": "nonthapat.p1994@gmail.com",
        "anydesk": "-",
        "issue": "บัญชีผู้ใช้: Capcut",
        "cause": "-",
        "duration": "00:30",
        "responder": "Guest User 75513",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 88,
        "date": "31/01/2569 16:00",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      }
    ],
    "automationsDone": 6,
    "aiApps": 2,
    "hoursSaved": 90
  },
  "2026-02": {
    "monthName": "กุมภาพันธ์ 2569",
    "totalAssets": 162,
    "assetValue": 2561390,
    "assetsExpiring": 6,
    "assetsBroken": 20,
    "assetsLost": 0,
    "ticketsCount": 35,
    "slaPercent": 80,
    "responseTime": 67,
    "resolutionTime": 5.6,
    "csat": 4.8,
    "totalSoftware": 24,
    "licensesInUse": 277.5,
    "licensesVacant": 42,
    "softwareCost": 124875,
    "softwareExpiring": 2,
    "backupSuccess": 99.9,
    "securityIncidents": 0,
    "antivirusCoverage": 98.5,
    "mfaCoverage": 100,
    "repairCount": 19,
    "repairCost": 0,
    "topBrokenDevices": [
      {
        "name": "Notebook",
        "count": 6,
        "cost": 0
      },
      {
        "name": "Printer",
        "count": 4,
        "cost": 0
      },
      {
        "name": "CCTV",
        "count": 4,
        "cost": 0
      },
      {
        "name": "PC Computer",
        "count": 2,
        "cost": 0
      },
      {
        "name": "Mac",
        "count": 1,
        "cost": 0
      },
      {
        "name": "iPad",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Mouse",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Smartphone",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Keyboard",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Router",
        "count": 0,
        "cost": 0
      }
    ],
    "deptCosts": {
      "Accounting & Finance": 0,
      "Sales & Marketing": 0,
      "Human Resources": 0,
      "Operations": 0,
      "IT & Infrastructure": 0
    },
    "softwareExpiringDetails": [
      {
        "name": "Microsoft 365 Copilot",
        "licenses": 50,
        "expiringDate": "30 ส.ค. 2026",
        "status": "ใกล้หมดอายุ"
      },
      {
        "name": "Adobe Creative Cloud",
        "licenses": 15,
        "expiringDate": "12 ก.ย. 2026",
        "status": "แจ้งเตือนล่วงหน้า"
      }
    ],
    "assetsExpiringDetails": [
      {
        "id": "AST-NB-001",
        "type": "Laptop",
        "model": "Lenovo ThinkPad L14",
        "dept": "Operations",
        "expDate": "15 ส.ค. 2026"
      },
      {
        "id": "AST-PR-004",
        "type": "Printer",
        "model": "HP LaserJet Pro M404",
        "dept": "Accounting",
        "expDate": "22 ส.ค. 2026"
      }
    ],
    "ongoingProjects": [
      {
        "title": "IT Ticket Automation",
        "desc": "พัฒนาสคริปต์ช่วยจัดการปัญหาซ้ำซาก คืบหน้า 60%"
      },
      {
        "title": "Asset Management System",
        "desc": "ระบบเช็คอิน-เช็คเอาท์อุปกรณ์ไอที คืบหน้า 60%"
      }
    ],
    "recommendations": [
      "พบปัญหาอุปกรณ์ขัดข้องสูงถึง 19 ครั้งในเดือนนี้ แนะนำจัดรอบบำรุงรักษาเชิงป้องกัน (Preventive Maintenance) โดยเฉพาะอุปกรณ์ประเภท Notebook",
      "อัตราการบรรลุเป้าหมาย SLA ลดลงเหลือ 80% แนะนำให้ปรับกระบวนการคัดกรอง Ticket เพื่อเพิ่มความรวดเร็วในการแก้ปัญหา",
      "แนะนำให้ผู้ใช้อัปเกรดความปลอดภัยของบัญชีผ่านการเปิดใช้งาน MFA ครบ 100%"
    ],
    "ticketsList": [
      {
        "sn": 89,
        "date": "03/02/2569 07:51",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "01:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 90,
        "date": "03/02/2569 09:24",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: เชื่อมต่อ Nas ไม่ได้, บัญชีผู้ใช้: เชื่อมต่อ Server ",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 91,
        "date": "03/02/2569 09:28",
        "complainant": "ปาหนัน สุพรม",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "03:00",
        "responder": "ปาหนัน สุพรม (พู่กัน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 92,
        "date": "03/02/2569 12:29",
        "complainant": "สุภาพ  แสนจันทร์ (พี่ส่ม)ี",
        "email": "khanoomjeen@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "01:00",
        "responder": "ผู้ใช้รับเชิญ 81286",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 93,
        "date": "04/02/2569 07:52",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 94,
        "date": "04/02/2569 09:10",
        "complainant": "สุภาพ แสนจันทร์ ",
        "email": "khanoomjeen@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: CCTV, ซอฟต์แวร์: CCTV, บัญชีผู้ใช้: CCTV",
        "cause": "-",
        "duration": "01:30",
        "responder": "ผู้ใช้รับเชิญ 81286",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 95,
        "date": "04/02/2569 10:44",
        "complainant": "พิชยา ฮงทอง",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "cause": "-",
        "duration": "01:00",
        "responder": "Guest User 16356",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 96,
        "date": "04/02/2569 10:45",
        "complainant": "พิชยา ฮงทอง",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "00:30",
        "responder": "Guest User 16356",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 98,
        "date": "05/02/2569 09:51",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 99,
        "date": "05/02/2569 09:52",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: CCTV, ซอฟต์แวร์: CCTV",
        "cause": "-",
        "duration": "48:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 101,
        "date": "05/02/2569 09:54",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: CCTV, ซอฟต์แวร์: CCTV",
        "cause": "-",
        "duration": "48:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 102,
        "date": "05/02/2569 15:11",
        "complainant": "ชลธิชา ตาลพันธ์",
        "email": "chonthicha.talpun1998@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: Calenda, G-Suit, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:30",
        "responder": "ชลธิชา ตาลพันธ์ (นิชา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 103,
        "date": "06/02/2569 10:56",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 105,
        "date": "06/02/2569 14:01",
        "complainant": "สุรวิชญ์ โพธิ์ตาก",
        "email": "p.surwit@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Mac, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: เชื่อมต่อ Nas ไม่ได้, บัญชีผู้ใช้: เชื่อมต่อ Nas ",
        "cause": "-",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 44149",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 106,
        "date": "09/02/2569 08:42",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 107,
        "date": "10/02/2569 09:06",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 108,
        "date": "10/02/2569 09:46",
        "complainant": "อัจฉรา เหรียญพิมาย (โบกี้)",
        "email": "atchararianpimai@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook",
        "cause": "-",
        "duration": "00:10",
        "responder": "อัจฉรา เหรียญพิมาย (Bogie)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 109,
        "date": "10/02/2569 10:12",
        "complainant": "พิชยา ฮงทอง",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, เน็ตเวิร์ค: Calenda",
        "cause": "-",
        "duration": "00:30",
        "responder": "Guest User 22874",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 110,
        "date": "13/02/2569 10:07",
        "complainant": "นภัสสร นาสวน",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer",
        "cause": "-",
        "duration": "00:30",
        "responder": "Guest User 75198",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 111,
        "date": "16/02/2569 09:22",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 113,
        "date": "16/02/2569 10:58",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "04:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 114,
        "date": "17/02/2569 08:40",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 115,
        "date": "17/02/2569 08:56",
        "complainant": "กฤษณา ลำเพ็ง",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows",
        "cause": "-",
        "duration": "00:30",
        "responder": "Guest User 67573",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 116,
        "date": "17/02/2569 09:13",
        "complainant": "อนุสรา สิมจันทา",
        "email": "aanusara.a43@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: IOS",
        "cause": "-",
        "duration": "00:10",
        "responder": "ผู้ใช้รับเชิญ 52165",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 117,
        "date": "17/02/2569 13:57",
        "complainant": "กฤติมา สอนพูน",
        "email": "kittima1712@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Microsoft Office, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้, บัญชีผู้ใช้: google dive",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 14230",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 118,
        "date": "18/02/2569 08:28",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "โปรดให้เนื้อหาต้นฉบับที่ต้องการสรุปมาเพื่อให้ฉันสามารถสรุปเป็นภาษาไทยให้คุณได้ครับ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 119,
        "date": "23/02/2569 09:38",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "โปรดให้เนื้อหาหลักที่ต้องการสรุปมา ฉันจะสรุปให้เป็นภาษาไทยตามที่ขอค่ะ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 120,
        "date": "23/02/2569 09:39",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: CCTV, ซอฟต์แวร์: CCTV",
        "cause": "-",
        "duration": "72:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 121,
        "date": "23/02/2569 11:00",
        "complainant": "นางสาวสุพรรษา อินทะเรืองรุ่ง",
        "email": "Supansa.si90@gmail.com",
        "anydesk": "-",
        "issue": "บัญชีผู้ใช้: เชื่อมต่อ Server , Email",
        "cause": "-",
        "duration": "00:30",
        "responder": "สุพรรษา อินทะเรืองรุ่ง (ต่าย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 123,
        "date": "24/02/2569 08:47",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "กรุณาให้เนื้อหาหลัก (正文) ที่ต้องการสรุปเป็นภาษาไทยมาให้ฉันนะครับ/คะ ฉันจะสรุปเนื้อหานั้นให้ถูกต้องและกระชับตามที่คุณต้องการโดยทันที",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 132,
        "date": "25/02/2569 15:39",
        "complainant": "นิชา",
        "email": "Chonthich.Talpun1998@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook",
        "cause": "-",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 53591",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 136,
        "date": "26/02/2569 09:00",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 140,
        "date": "26/02/2569 11:12",
        "complainant": "ปาหนัน สุพรม พู่กัน",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Mouse",
        "cause": "-",
        "duration": "00:20",
        "responder": "ปาหนัน สุพรม (พู่กัน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 143,
        "date": "27/02/2569 09:04",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 144,
        "date": "28/02/2569 08:47",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      }
    ],
    "automationsDone": 4,
    "aiApps": 2,
    "hoursSaved": 60
  },
  "2026-03": {
    "monthName": "มีนาคม 2569",
    "totalAssets": 162,
    "assetValue": 2561390,
    "assetsExpiring": 5,
    "assetsBroken": 9,
    "assetsLost": 0,
    "ticketsCount": 17,
    "slaPercent": 100,
    "responseTime": 6,
    "resolutionTime": 0.5,
    "csat": 4.9,
    "totalSoftware": 24,
    "licensesInUse": 250.5,
    "licensesVacant": 38,
    "softwareCost": 112725,
    "softwareExpiring": 2,
    "backupSuccess": 99.9,
    "securityIncidents": 0,
    "antivirusCoverage": 98.5,
    "mfaCoverage": 100,
    "repairCount": 8,
    "repairCost": 0,
    "topBrokenDevices": [
      {
        "name": "Notebook",
        "count": 7,
        "cost": 0
      },
      {
        "name": "Mac",
        "count": 1,
        "cost": 0
      },
      {
        "name": "PC Computer",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Printer",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Smartphone",
        "count": 0,
        "cost": 0
      },
      {
        "name": "iPad",
        "count": 0,
        "cost": 0
      },
      {
        "name": "CCTV",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Mouse",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Keyboard",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Router",
        "count": 0,
        "cost": 0
      }
    ],
    "deptCosts": {
      "Accounting & Finance": 0,
      "Sales & Marketing": 0,
      "Human Resources": 0,
      "Operations": 0,
      "IT & Infrastructure": 0
    },
    "softwareExpiringDetails": [
      {
        "name": "Microsoft 365 Copilot",
        "licenses": 50,
        "expiringDate": "30 ส.ค. 2026",
        "status": "ใกล้หมดอายุ"
      },
      {
        "name": "Adobe Creative Cloud",
        "licenses": 15,
        "expiringDate": "12 ก.ย. 2026",
        "status": "แจ้งเตือนล่วงหน้า"
      }
    ],
    "assetsExpiringDetails": [
      {
        "id": "AST-NB-001",
        "type": "Laptop",
        "model": "Lenovo ThinkPad L14",
        "dept": "Operations",
        "expDate": "15 ส.ค. 2026"
      },
      {
        "id": "AST-PR-004",
        "type": "Printer",
        "model": "HP LaserJet Pro M404",
        "dept": "Accounting",
        "expDate": "22 ส.ค. 2026"
      }
    ],
    "ongoingProjects": [
      {
        "title": "IT Ticket Automation",
        "desc": "พัฒนาสคริปต์ช่วยจัดการปัญหาซ้ำซาก คืบหน้า 50%"
      },
      {
        "title": "Asset Management System",
        "desc": "ระบบเช็คอิน-เช็คเอาท์อุปกรณ์ไอที คืบหน้า 60%"
      }
    ],
    "recommendations": [
      "พบปัญหาอุปกรณ์ขัดข้องสูงถึง 8 ครั้งในเดือนนี้ แนะนำจัดรอบบำรุงรักษาเชิงป้องกัน (Preventive Maintenance) โดยเฉพาะอุปกรณ์ประเภท Notebook",
      "การสนับสนุนผู้ใช้ระบบ IT อยู่ในเกณฑ์ดีเยี่ยม สามารถรักษามาตรฐาน SLA ได้ตามเป้าหมาย",
      "แนะนำให้ผู้ใช้อัปเกรดความปลอดภัยของบัญชีผ่านการเปิดใช้งาน MFA ครบ 100%"
    ],
    "ticketsList": [
      {
        "sn": 145,
        "date": "03/03/2569 08:34",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 146,
        "date": "04/03/2569 08:30",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 147,
        "date": "10/03/2569 08:36",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 148,
        "date": "10/03/2569 09:51",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "เทรนนิ่ง Lark",
        "cause": "-",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 150,
        "date": "11/03/2569 08:28",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 151,
        "date": "12/03/2569 08:20",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 167,
        "date": "18/03/2569 09:06",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 169,
        "date": "18/03/2569 09:06",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook",
        "cause": "-",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 171,
        "date": "19/03/2569 10:09",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 172,
        "date": "19/03/2569 10:09",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 173,
        "date": "19/03/2569 10:10",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows",
        "cause": "-",
        "duration": "00:30",
        "responder": "Base Assistant",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 175,
        "date": "19/03/2569 16:46",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: CCTV",
        "cause": "-",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 177,
        "date": "19/03/2569 16:48",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "cause": "-",
        "duration": "00:30",
        "responder": "Base Assistant",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 178,
        "date": "20/03/2569 12:54",
        "complainant": "พิชยา ฮงทอง (แนน)",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Microsoft Office",
        "cause": "-",
        "duration": "00:20",
        "responder": "พิชยา ฮงทอง (แนน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 181,
        "date": "30/03/2569 08:47",
        "complainant": "ต่าย HR",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Microsoft Office",
        "cause": "-",
        "duration": "0030",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 184,
        "date": "31/03/2569 13:06",
        "complainant": "ช่างภาพ ไมค์",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Mac, ซอฟต์แวร์: IOS",
        "cause": "-",
        "duration": "00:40",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 186,
        "date": "31/03/2569 15:42",
        "complainant": "กิ๊กจัดซื้อ",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows",
        "cause": "-",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      }
    ],
    "automationsDone": 2,
    "aiApps": 2,
    "hoursSaved": 30
  },
  "2026-04": {
    "monthName": "เมษายน 2569",
    "totalAssets": 162,
    "assetValue": 2561390,
    "assetsExpiring": 5,
    "assetsBroken": 7,
    "assetsLost": 0,
    "ticketsCount": 8,
    "slaPercent": 50,
    "responseTime": 19,
    "resolutionTime": 1.6,
    "csat": 4.7,
    "totalSoftware": 24,
    "licensesInUse": 237,
    "licensesVacant": 36,
    "softwareCost": 106650,
    "softwareExpiring": 2,
    "backupSuccess": 99.9,
    "securityIncidents": 0,
    "antivirusCoverage": 98.5,
    "mfaCoverage": 100,
    "repairCount": 6,
    "repairCost": 0,
    "topBrokenDevices": [
      {
        "name": "Printer",
        "count": 2,
        "cost": 0
      },
      {
        "name": "Notebook",
        "count": 2,
        "cost": 0
      },
      {
        "name": "Mac",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Router",
        "count": 1,
        "cost": 0
      },
      {
        "name": "PC Computer",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Smartphone",
        "count": 0,
        "cost": 0
      },
      {
        "name": "iPad",
        "count": 0,
        "cost": 0
      },
      {
        "name": "CCTV",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Mouse",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Keyboard",
        "count": 0,
        "cost": 0
      }
    ],
    "deptCosts": {
      "Accounting & Finance": 0,
      "Sales & Marketing": 0,
      "Human Resources": 0,
      "Operations": 0,
      "IT & Infrastructure": 0
    },
    "softwareExpiringDetails": [
      {
        "name": "Microsoft 365 Copilot",
        "licenses": 50,
        "expiringDate": "30 ส.ค. 2026",
        "status": "ใกล้หมดอายุ"
      },
      {
        "name": "Adobe Creative Cloud",
        "licenses": 15,
        "expiringDate": "12 ก.ย. 2026",
        "status": "แจ้งเตือนล่วงหน้า"
      }
    ],
    "assetsExpiringDetails": [
      {
        "id": "AST-NB-001",
        "type": "Laptop",
        "model": "Lenovo ThinkPad L14",
        "dept": "Operations",
        "expDate": "15 ส.ค. 2026"
      },
      {
        "id": "AST-PR-004",
        "type": "Printer",
        "model": "HP LaserJet Pro M404",
        "dept": "Accounting",
        "expDate": "22 ส.ค. 2026"
      }
    ],
    "ongoingProjects": [
      {
        "title": "IT Ticket Automation",
        "desc": "พัฒนาสคริปต์ช่วยจัดการปัญหาซ้ำซาก คืบหน้า 45%"
      },
      {
        "title": "Asset Management System",
        "desc": "ระบบเช็คอิน-เช็คเอาท์อุปกรณ์ไอที คืบหน้า 60%"
      }
    ],
    "recommendations": [
      "พบปัญหาอุปกรณ์ขัดข้องสูงถึง 6 ครั้งในเดือนนี้ แนะนำจัดรอบบำรุงรักษาเชิงป้องกัน (Preventive Maintenance) โดยเฉพาะอุปกรณ์ประเภท Printer",
      "อัตราการบรรลุเป้าหมาย SLA ลดลงเหลือ 50% แนะนำให้ปรับกระบวนการคัดกรอง Ticket เพื่อเพิ่มความรวดเร็วในการแก้ปัญหา",
      "แนะนำให้ผู้ใช้อัปเกรดความปลอดภัยของบัญชีผ่านการเปิดใช้งาน MFA ครบ 100%"
    ],
    "ticketsList": [
      {
        "sn": 188,
        "date": "03/04/2569 09:48",
        "complainant": "กฤษณา ลำเพ็ง",
        "email": "-",
        "anydesk": "-",
        "issue": "ติดตั้งเครื่องปริ้นบ้าน 18 ในโน้ตบุ้ค",
        "cause": "-",
        "duration": "00:20",
        "responder": "กฤษณา ลำเพ็ง (พลอย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 189,
        "date": "07/04/2569 08:31",
        "complainant": "สุภาพ  แสนจันทร์ ",
        "email": "fernclinic.acc@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: เชื่อมต่อ Server , Printer WIFI",
        "cause": "-",
        "duration": "02:00",
        "responder": "สุภาพ แสนจันทร์ (ส้ม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 190,
        "date": "07/04/2569 12:08",
        "complainant": "ชลธิชา สุวาส",
        "email": "chonticha.suw@northbkk.ac.th",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:30",
        "responder": "ชลธิชา สุวาส (จิน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 191,
        "date": "09/04/2569 09:12",
        "complainant": "เอกรินทร์ จีนเพชร",
        "email": "-",
        "anydesk": "-",
        "issue": "บัญชีผู้ใช้: Kumoo, Email",
        "cause": "-",
        "duration": "02:00",
        "responder": "ผู้ใช้รับเชิญ 55420",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 192,
        "date": "22/04/2569 10:56",
        "complainant": "เนตรปรีญา ทัดศรี",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer",
        "cause": "-",
        "duration": "03:00",
        "responder": "เนตรปรีญา ทัดศรี(แตงกวา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 193,
        "date": "23/04/2569 15:21",
        "complainant": "ชลธิชา สุวาส",
        "email": "chonticha.suw@northbkk.ac.th",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Microsoft Office, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้",
        "cause": "-",
        "duration": "01:00",
        "responder": "ผู้ใช้รับเชิญ 98623",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 194,
        "date": "24/04/2569 13:22",
        "complainant": "เนตรปรีญา ทัดศรี",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Mac, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้",
        "cause": "-",
        "duration": "01:00",
        "responder": "เนตรปรีญา ทัดศรี(แตงกวา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 195,
        "date": "30/04/2569 08:25",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "ซอฟต์แวร์: Config System, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "cause": "-",
        "duration": "03:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      }
    ],
    "automationsDone": 1,
    "aiApps": 2,
    "hoursSaved": 15
  },
  "2026-05": {
    "monthName": "พฤษภาคม 2569",
    "totalAssets": 162,
    "assetValue": 2561390,
    "assetsExpiring": 5,
    "assetsBroken": 6,
    "assetsLost": 0,
    "ticketsCount": 8,
    "slaPercent": 100,
    "responseTime": 5,
    "resolutionTime": 0.4,
    "csat": 4.9,
    "totalSoftware": 24,
    "licensesInUse": 237,
    "licensesVacant": 36,
    "softwareCost": 106650,
    "softwareExpiring": 2,
    "backupSuccess": 99.9,
    "securityIncidents": 0,
    "antivirusCoverage": 98.5,
    "mfaCoverage": 100,
    "repairCount": 5,
    "repairCost": 0,
    "topBrokenDevices": [
      {
        "name": "Printer",
        "count": 1,
        "cost": 0
      },
      {
        "name": "iPad",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Notebook",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Mouse",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Smartphone",
        "count": 1,
        "cost": 0
      },
      {
        "name": "PC Computer",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Mac",
        "count": 0,
        "cost": 0
      },
      {
        "name": "CCTV",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Keyboard",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Router",
        "count": 0,
        "cost": 0
      }
    ],
    "deptCosts": {
      "Accounting & Finance": 0,
      "Sales & Marketing": 0,
      "Human Resources": 0,
      "Operations": 0,
      "IT & Infrastructure": 0
    },
    "softwareExpiringDetails": [
      {
        "name": "Microsoft 365 Copilot",
        "licenses": 50,
        "expiringDate": "30 ส.ค. 2026",
        "status": "ใกล้หมดอายุ"
      },
      {
        "name": "Adobe Creative Cloud",
        "licenses": 15,
        "expiringDate": "12 ก.ย. 2026",
        "status": "แจ้งเตือนล่วงหน้า"
      }
    ],
    "assetsExpiringDetails": [
      {
        "id": "AST-NB-001",
        "type": "Laptop",
        "model": "Lenovo ThinkPad L14",
        "dept": "Operations",
        "expDate": "15 ส.ค. 2026"
      },
      {
        "id": "AST-PR-004",
        "type": "Printer",
        "model": "HP LaserJet Pro M404",
        "dept": "Accounting",
        "expDate": "22 ส.ค. 2026"
      }
    ],
    "ongoingProjects": [
      {
        "title": "IT Ticket Automation",
        "desc": "พัฒนาสคริปต์ช่วยจัดการปัญหาซ้ำซาก คืบหน้า 45%"
      },
      {
        "title": "Asset Management System",
        "desc": "ระบบเช็คอิน-เช็คเอาท์อุปกรณ์ไอที คืบหน้า 60%"
      }
    ],
    "recommendations": [
      "การสนับสนุนผู้ใช้ระบบ IT อยู่ในเกณฑ์ดีเยี่ยม สามารถรักษามาตรฐาน SLA ได้ตามเป้าหมาย",
      "แนะนำให้ผู้ใช้อัปเกรดความปลอดภัยของบัญชีผ่านการเปิดใช้งาน MFA ครบ 100%"
    ],
    "ticketsList": [
      {
        "sn": 196,
        "date": "06/05/2569 10:44",
        "complainant": "เบญจภรณ์ เอี่ยมต้นเค้า",
        "email": "benjaporn.atk@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Printer, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "cause": "-",
        "duration": "01:00",
        "responder": "เนตรปรีญา ทัดศรี(แตงกวา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 197,
        "date": "18/05/2569 14:04",
        "complainant": "กฤติญา ทาระพันธ์",
        "email": "Krittiya.trp@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad, Notebook, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 41070",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 198,
        "date": "19/05/2569 13:30",
        "complainant": "ปาหนัน สุพรม",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Mouse",
        "cause": "-",
        "duration": "00:20",
        "responder": "ปาหนัน สุพรม (พู่กัน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 200,
        "date": "20/05/2569 13:22",
        "complainant": "ธนัชชา บุญมีมาก",
        "email": "poppopss31886@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Smartphone, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:10",
        "responder": "ธนัชชา บุญมีมาก (ป๊อป)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 201,
        "date": "20/05/2569 14:49",
        "complainant": "ชนันพร อินขำ",
        "email": "cnppcy156cm@gmail.com",
        "anydesk": "-",
        "issue": "เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: google dive",
        "cause": "-",
        "duration": "00:40",
        "responder": "ผู้ใช้รับเชิญ 40358",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 202,
        "date": "22/05/2569 09:02",
        "complainant": "ณัฏชณินภา กำจร",
        "email": "water.work1308@gmail.com",
        "anydesk": "1805513405",
        "issue": "บัญชีผู้ใช้: Lark, Chat GPT",
        "cause": "-",
        "duration": "00:05",
        "responder": "ผู้ใช้รับเชิญ 57602",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 203,
        "date": "22/05/2569 10:16",
        "complainant": "ชนันพร อินขำ",
        "email": "cnppcy156cm@gmail.com",
        "anydesk": "-",
        "issue": "บัญชีผู้ใช้: Tiktok",
        "cause": "-",
        "duration": "00:05",
        "responder": "ชนันพร อินขำ (ไอซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 204,
        "date": "22/05/2569 10:37",
        "complainant": "ปัญจมา สมบัติกำไร",
        "email": "-",
        "anydesk": "1267304100",
        "issue": "บัญชีผู้ใช้: Tiktok",
        "cause": "-",
        "duration": "00:05",
        "responder": "Guest User 61903",
        "status": "เสร็จสิ้น",
        "cost": 0
      }
    ],
    "automationsDone": 1,
    "aiApps": 2,
    "hoursSaved": 15
  },
  "2026-06": {
    "monthName": "มิถุนายน 2569",
    "totalAssets": 162,
    "assetValue": 2561390,
    "assetsExpiring": 5,
    "assetsBroken": 10,
    "assetsLost": 0,
    "ticketsCount": 25,
    "slaPercent": 96,
    "responseTime": 5,
    "resolutionTime": 0.3,
    "csat": 4.9,
    "totalSoftware": 24,
    "licensesInUse": 262.5,
    "licensesVacant": 39,
    "softwareCost": 118125,
    "softwareExpiring": 2,
    "backupSuccess": 99.9,
    "securityIncidents": 0,
    "antivirusCoverage": 98.5,
    "mfaCoverage": 100,
    "repairCount": 9,
    "repairCost": 0,
    "topBrokenDevices": [
      {
        "name": "Notebook",
        "count": 3,
        "cost": 0
      },
      {
        "name": "Mac",
        "count": 3,
        "cost": 0
      },
      {
        "name": "Mouse",
        "count": 1,
        "cost": 0
      },
      {
        "name": "iPad",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Printer",
        "count": 1,
        "cost": 0
      },
      {
        "name": "PC Computer",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Smartphone",
        "count": 0,
        "cost": 0
      },
      {
        "name": "CCTV",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Keyboard",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Router",
        "count": 0,
        "cost": 0
      }
    ],
    "deptCosts": {
      "Accounting & Finance": 0,
      "Sales & Marketing": 0,
      "Human Resources": 0,
      "Operations": 0,
      "IT & Infrastructure": 0
    },
    "softwareExpiringDetails": [
      {
        "name": "Microsoft 365 Copilot",
        "licenses": 50,
        "expiringDate": "30 ส.ค. 2026",
        "status": "ใกล้หมดอายุ"
      },
      {
        "name": "Adobe Creative Cloud",
        "licenses": 15,
        "expiringDate": "12 ก.ย. 2026",
        "status": "แจ้งเตือนล่วงหน้า"
      }
    ],
    "assetsExpiringDetails": [
      {
        "id": "AST-NB-001",
        "type": "Laptop",
        "model": "Lenovo ThinkPad L14",
        "dept": "Operations",
        "expDate": "15 ส.ค. 2026"
      },
      {
        "id": "AST-PR-004",
        "type": "Printer",
        "model": "HP LaserJet Pro M404",
        "dept": "Accounting",
        "expDate": "22 ส.ค. 2026"
      }
    ],
    "ongoingProjects": [
      {
        "title": "IT Ticket Automation",
        "desc": "พัฒนาสคริปต์ช่วยจัดการปัญหาซ้ำซาก คืบหน้า 55%"
      },
      {
        "title": "Asset Management System",
        "desc": "ระบบเช็คอิน-เช็คเอาท์อุปกรณ์ไอที คืบหน้า 60%"
      }
    ],
    "recommendations": [
      "พบปัญหาอุปกรณ์ขัดข้องสูงถึง 9 ครั้งในเดือนนี้ แนะนำจัดรอบบำรุงรักษาเชิงป้องกัน (Preventive Maintenance) โดยเฉพาะอุปกรณ์ประเภท Notebook",
      "การสนับสนุนผู้ใช้ระบบ IT อยู่ในเกณฑ์ดีเยี่ยม สามารถรักษามาตรฐาน SLA ได้ตามเป้าหมาย",
      "แนะนำให้ผู้ใช้อัปเกรดความปลอดภัยของบัญชีผ่านการเปิดใช้งาน MFA ครบ 100%"
    ],
    "ticketsList": [
      {
        "sn": 205,
        "date": "01/06/2569 08:44",
        "complainant": "เอก",
        "email": "-",
        "anydesk": "-",
        "issue": "บัญชีผู้ใช้: Lark",
        "cause": "-",
        "duration": "00.05",
        "responder": "ธันวา เเซ่เเฮ (ไนซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 206,
        "date": "01/06/2569 11:02",
        "complainant": "รามจิตติ ชินนะเกิด",
        "email": "-",
        "anydesk": "1207995468",
        "issue": "บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00.05",
        "responder": "รามจิตติ ชินนะเกิดโชค (เบนซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 208,
        "date": "02/06/2569 08:46",
        "complainant": "กานต์ฑิตา ธีระพิบูลย์",
        "email": "kantita313@gmail.com",
        "anydesk": "1540706084",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Google Link, Windows, Lark, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "กานต์ฑิตา ธีระพิบูลย์ (อิง)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 209,
        "date": "02/06/2569 11:27",
        "complainant": "รามจิตติ",
        "email": "-",
        "anydesk": "1207995468",
        "issue": "บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "รามจิตติ ชินนะเกิดโชค (เบนซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 210,
        "date": "02/06/2569 11:41",
        "complainant": "ชนันพร อินขำ",
        "email": "cnppcy156cm@gmail.com",
        "anydesk": "-",
        "issue": "บัญชีผู้ใช้: Chat GPT",
        "cause": "-",
        "duration": "00.02",
        "responder": "Guest User 94269",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 211,
        "date": "02/06/2569 12:09",
        "complainant": "พิชชาพร คอทอง",
        "email": "Pidchaporn@fernnasthetic.com",
        "anydesk": "1483295639",
        "issue": "บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:30",
        "responder": "พิชชาพร คอทอง (พีเจ้น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 212,
        "date": "04/06/2569 09:24",
        "complainant": "เย็นฤดี มาระวัง",
        "email": "yenrudee42830@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Notebook",
        "cause": "-",
        "duration": "00.05",
        "responder": "เย็นฤดี มาระวัง (ฝ้าย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 213,
        "date": "05/06/2569 10:14",
        "complainant": "อาทิตยา มุมทอง",
        "email": "graphicfernclinic@gmail.com",
        "anydesk": "1 955 477 996",
        "issue": "บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:10",
        "responder": "อาทิตยา มุมทอง (ขมิ้น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 219,
        "date": "08/06/2569 11:01",
        "complainant": "เย็นฤดี มาระวัง",
        "email": "yenrudee42830@gmail.com",
        "anydesk": "Lenovo-013",
        "issue": "ฮาร์ดแวร์: Notebook, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "00.10",
        "responder": "เย็นฤดี มาระวัง (ฝ้าย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 220,
        "date": "09/06/2569 09:34",
        "complainant": "เย็นฤดี มาระวัง",
        "email": "yenrudee42830@gmail.com",
        "anydesk": "asus-032",
        "issue": "ฮาร์ดแวร์: Mouse",
        "cause": "-",
        "duration": "00.03",
        "responder": "ผู้ใช้รับเชิญ 66959",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 221,
        "date": "09/06/2569 09:51",
        "complainant": "พิสิษฐ์ มงคลสมบัติศิริ",
        "email": "-",
        "anydesk": "1677760326",
        "issue": "บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00.03",
        "responder": "พิสิษฐ์ มงคลสมบัติศิริ (เจมส์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 223,
        "date": "09/06/2569 10:34",
        "complainant": "เอมปวีภร์ วัชระตระการพงศ์",
        "email": "Aimpavee@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Mac",
        "cause": "-",
        "duration": "00.23",
        "responder": "เอมปวีภร์ วัชระตระการพงศ์ (กิ๊ฟ)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 224,
        "date": "10/06/2569 15:57",
        "complainant": "เนลินญาน์  ศิระไมตรีฉัตร (เนลิน)",
        "email": "-",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Mac",
        "cause": "-",
        "duration": "00.30",
        "responder": "เนลินญาน์ ศิระไมตรีฉัตร (เนลิน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 225,
        "date": "11/06/2569 08:21",
        "complainant": "อาทิตยา มุมทอง",
        "email": "Athittaya9597@gmail.com",
        "anydesk": "478 845 465",
        "issue": "ฮาร์ดแวร์: Mac",
        "cause": "-",
        "duration": "00:20",
        "responder": "อาทิตยา มุมทอง (ขมิ้น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 226,
        "date": "15/06/2569 09:08",
        "complainant": "มนัสนันท์ เทพแก้ว ",
        "email": "mantsanantk@gmail.com",
        "anydesk": "1 375 615 606",
        "issue": "บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00.01",
        "responder": "Guest User 33774",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 227,
        "date": "15/06/2569 10:30",
        "complainant": "",
        "email": "-",
        "anydesk": "-",
        "issue": "เช็คอุปกรณ์ Admin CRM บัญชี Producer Live",
        "cause": "-",
        "duration": "02:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 228,
        "date": "15/06/2569 12:50",
        "complainant": "เนลินญาน์  ศิระไมตรีฉัตร ",
        "email": "-",
        "anydesk": "https://meet.google.com/dmr-jcgs-nfz",
        "issue": "บัญชีผู้ใช้: google meet",
        "cause": "-",
        "duration": "00:10",
        "responder": "เนลินญาน์ ศิระไมตรีฉัตร (เนลิน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 257,
        "date": "16/06/2569 14:18",
        "complainant": "นภัสสร นาสวน",
        "email": "-",
        "anydesk": "- ",
        "issue": "ฮาร์ดแวร์: iPad",
        "cause": "-",
        "duration": "00.02",
        "responder": "นภัสสร นาสวน (โบว์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 258,
        "date": "16/06/2569 15:40",
        "complainant": "อิศราภรณ์ ปิ่นงาม",
        "email": "isaraphornxb@gmail.com",
        "anydesk": "1901650433",
        "issue": "เข้า IG Fern Clinic ในคอมกับโทรศัพท์เลขาให้หน่อยค่ะ เพื่อให้ฝั่ง content ทำการปรับแก้ไอจีได้",
        "cause": "-",
        "duration": "00:20",
        "responder": "อิศราภรณ์ ปิ่นงาม (เบนซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 259,
        "date": "16/06/2569 16:08",
        "complainant": "ชนันพร อินขำ",
        "email": "cnppcy156cm@gmail.com",
        "anydesk": "-",
        "issue": "บัญชีผู้ใช้: Tiktok",
        "cause": "-",
        "duration": "00.15",
        "responder": "ชนันพร อินขำ (ไอซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 260,
        "date": "18/06/2569 08:15",
        "complainant": "อาทิตยา มุมทอง",
        "email": "athittaya9597@gmail.com",
        "anydesk": "478 845 465",
        "issue": "บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00:20",
        "responder": "อาทิตยา มุมทอง (ขมิ้น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 261,
        "date": "19/06/2569 10:34",
        "complainant": "อารยา ธนพันธุ์พาณิชย์",
        "email": "Noungning282@gmail.com",
        "anydesk": "-",
        "issue": "ขอบัญชี",
        "cause": "-",
        "duration": "00.10",
        "responder": "Guest User 72644",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 262,
        "date": "23/06/2569 10:02",
        "complainant": "ชัยธัช ชัยวัฒน์",
        "email": "-",
        "anydesk": "1 207 995 468",
        "issue": "บัญชีผู้ใช้: Email, Lark",
        "cause": "-",
        "duration": "00:30",
        "responder": "รามจิตติ ชินนะเกิดโชค (เบนซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 263,
        "date": "29/06/2569 09:41",
        "complainant": "ณัฐกานต์ ชิดปรางค์",
        "email": "Bitoey.nat@gmail.com",
        "anydesk": "1 380 377 906",
        "issue": "ฮาร์ดแวร์: Printer",
        "cause": "-",
        "duration": "00:30",
        "responder": "ณัฐกานต์ ชิดปรางค์ (เตยหอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 264,
        "date": "29/06/2569 13:09",
        "complainant": "บุษกร บัวสวรรค์",
        "email": "จำเมลไม่ได้แต่ลดล็อกอินที่ชื่อว่า video team",
        "anydesk": "iMac dr.fern 2",
        "issue": "กดล็อกอิน CapCut เข้าเมลที่ชื่อวิดีโอทีม แล้วมันให้กรอกวันเดือนปีเกิด กดออกจากหน้านี้ไม่ได้เลย น่าจะเป็นวันเดือนปีเกิดที่ตรงกับเมล โลกบังคับออกแล้วก็รีสตาร์ตเครื่องแล้วก็เป็นเหมือนเดิม",
        "cause": "-",
        "duration": "00:20",
        "responder": "บุษกร บัวสวรรค์ (เรนนี่)",
        "status": "เสร็จสิ้น",
        "cost": 0
      }
    ],
    "automationsDone": 3,
    "aiApps": 2,
    "hoursSaved": 45
  },
  "2026-07": {
    "monthName": "กรกฎาคม 2569",
    "totalAssets": 162,
    "assetValue": 2561390,
    "assetsExpiring": 5,
    "assetsBroken": 8,
    "assetsLost": 0,
    "ticketsCount": 12,
    "slaPercent": 100,
    "responseTime": 5,
    "resolutionTime": 0.1,
    "csat": 4.9,
    "totalSoftware": 24,
    "licensesInUse": 243,
    "licensesVacant": 36,
    "softwareCost": 109350,
    "softwareExpiring": 2,
    "backupSuccess": 99.9,
    "securityIncidents": 0,
    "antivirusCoverage": 98.5,
    "mfaCoverage": 100,
    "repairCount": 8,
    "repairCost": 27800,
    "topBrokenDevices": [
      {
        "name": "PC Computer",
        "count": 3,
        "cost": 13900
      },
      {
        "name": "iPad",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Notebook",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Keyboard",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Printer",
        "count": 1,
        "cost": 0
      },
      {
        "name": "Smartphone",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Mac",
        "count": 0,
        "cost": 0
      },
      {
        "name": "CCTV",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Mouse",
        "count": 0,
        "cost": 0
      },
      {
        "name": "Router",
        "count": 0,
        "cost": 0
      }
    ],
    "deptCosts": {
      "Accounting & Finance": 6950,
      "Sales & Marketing": 5560,
      "Human Resources": 4170,
      "Operations": 8340,
      "IT & Infrastructure": 2780
    },
    "softwareExpiringDetails": [
      {
        "name": "Microsoft 365 Copilot",
        "licenses": 50,
        "expiringDate": "30 ส.ค. 2026",
        "status": "ใกล้หมดอายุ"
      },
      {
        "name": "Adobe Creative Cloud",
        "licenses": 15,
        "expiringDate": "12 ก.ย. 2026",
        "status": "แจ้งเตือนล่วงหน้า"
      }
    ],
    "assetsExpiringDetails": [
      {
        "id": "AST-NB-001",
        "type": "Laptop",
        "model": "Lenovo ThinkPad L14",
        "dept": "Operations",
        "expDate": "15 ส.ค. 2026"
      },
      {
        "id": "AST-PR-004",
        "type": "Printer",
        "model": "HP LaserJet Pro M404",
        "dept": "Accounting",
        "expDate": "22 ส.ค. 2026"
      }
    ],
    "ongoingProjects": [
      {
        "title": "IT Ticket Automation",
        "desc": "พัฒนาสคริปต์ช่วยจัดการปัญหาซ้ำซาก คืบหน้า 45%"
      },
      {
        "title": "Asset Management System",
        "desc": "ระบบเช็คอิน-เช็คเอาท์อุปกรณ์ไอที คืบหน้า 60%"
      }
    ],
    "recommendations": [
      "พบปัญหาอุปกรณ์ขัดข้องสูงถึง 7 ครั้งในเดือนนี้ แนะนำจัดรอบบำรุงรักษาเชิงป้องกัน (Preventive Maintenance) โดยเฉพาะอุปกรณ์ประเภท PC Computer",
      "การสนับสนุนผู้ใช้ระบบ IT อยู่ในเกณฑ์ดีเยี่ยม สามารถรักษามาตรฐาน SLA ได้ตามเป้าหมาย",
      "แนะนำให้ผู้ใช้อัปเกรดความปลอดภัยของบัญชีผ่านการเปิดใช้งาน MFA ครบ 100%"
    ],
    "ticketsList": [
      {
        "sn": 301,
        "date": "22/07/2569 10:00",
        "complainant": "คุณมินตรา (Facebook)",
        "email": "089-234-8891",
        "anydesk": "-",
        "issue": "ความสนใจ: คอร์สลดน้ำหนัก | AI ตรวจสลิป: รอสลิป (ขาด 2,900 บาท)",
        "cause": "-",
        "duration": "00:00",
        "responder": "Sale A",
        "status": "ต้องติดตาม",
        "cost": 2900
      },
      {
        "sn": 302,
        "date": "22/07/2569 10:30",
        "complainant": "คุณต้น (LINE)",
        "email": "082-775-4410",
        "anydesk": "-",
        "issue": "ความสนใจ: ปรึกษาผิวหน้า | AI ตรวจสลิป: ยอดตรงกัน (AI ตรวจสอบผ่าน)",
        "cause": "-",
        "duration": "00:00",
        "responder": "Sale B",
        "status": "เสร็จสิ้น",
        "cost": 500
      },
      {
        "sn": 303,
        "date": "22/07/2569 11:00",
        "complainant": "Nana Beauty (TikTok)",
        "email": "-",
        "anydesk": "-",
        "issue": "ความสนใจ: รีวิวสินค้า | AI ตรวจสลิป: รอยอดจอง (ยังไม่มียอดให้ AI ตรวจ)",
        "cause": "-",
        "duration": "00:00",
        "responder": "Unassigned",
        "status": "รอดำเนินการ",
        "cost": 0
      },
      {
        "sn": 265,
        "date": "01/07/2569 09:11",
        "complainant": "ชนันพร อินขำ",
        "email": "cnppcy156cm@gmail.com",
        "anydesk": "-",
        "issue": "บัญชีผู้ใช้: Chat GPT",
        "cause": "-",
        "duration": "07.00",
        "responder": "ชนันพร อินขำ (ไอซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 266,
        "date": "01/07/2569 13:47",
        "complainant": "ศุภฤกษ์ ภายไธสง",
        "email": "Supharoekphaithaisong@gmail.com",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: IOS, บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00.10",
        "responder": "ศุภฤกษ์ ภายไธสง (ดรีม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 267,
        "date": "01/07/2569 16:14",
        "complainant": "สุรวิชญ์ โพธิ์ตาก (ไมค์)",
        "email": "อยากได้อีเมล Gmail สำหรับ ตำแหน่ง Operation Coordinator พร้อมรหัส",
        "anydesk": "ไม่มี ยังไม่โหลดจ้า ",
        "issue": "ท้องเสียบางวันเวลาเช้า ๆ ",
        "cause": "-",
        "duration": "00:15",
        "responder": "สุรวิชญ์ โพธิ์ตาก (ไมค์เมโลดี้)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 268,
        "date": "03/07/2569 11:09",
        "complainant": "พิสิษฐ์ มงคลสมบัติศิริ",
        "email": "-",
        "anydesk": "1677760326",
        "issue": "ฮาร์ดแวร์: Notebook, บัญชีผู้ใช้: Printer WIFI",
        "cause": "-",
        "duration": "00:20",
        "responder": "พิสิษฐ์ มงคลสมบัติศิริ (เจมส์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 269,
        "date": "08/07/2569 15:57",
        "complainant": "วิลาสินี ทับทิม",
        "email": "Vila.neena@gmail.com",
        "anydesk": "ดาวโหลดแล้ว",
        "issue": "ไม่มีเครื่องปริ้นท์ในโปรแกรมพีคค่ะ/พี่นีปริ้นท์งานไม่ได้เลยค่ะ/SET ให้ด้วยค่ะ",
        "cause": "-",
        "duration": "00:30",
        "responder": "Guest User 35635",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 271,
        "date": "13/07/2569 14:04",
        "complainant": "รามจิตติ ชินนะเกิดโชค",
        "email": "-",
        "anydesk": "1801560835",
        "issue": "ฮาร์ดแวร์: PC Computer, บัญชีผู้ใช้: Email, Kumoo, google dive",
        "cause": "เสื่อมตามสภาพ",
        "duration": "00.25",
        "responder": "รามจิตติ ชินนะเกิดโชค (เบนซ์)",
        "status": "เสร็จสิ้น",
        "cost": 13900
      },
      {
        "sn": 272,
        "date": "14/07/2569 13:37",
        "complainant": "วิลาสินี ทับทิม",
        "email": "VILA.NEENA@GMAIL.COM",
        "anydesk": "-",
        "issue": "ฮาร์ดแวร์: Keyboard , บัญชีผู้ใช้: Lark, Chat GPT",
        "cause": "-",
        "duration": "00.30",
        "responder": "Guest User 35635",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 274,
        "date": "16/07/2569 08:45",
        "complainant": "เนลินญาน์  ศิระไมตรีฉัตร",
        "email": "-",
        "anydesk": "-",
        "issue": "บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00.05",
        "responder": "Guest User 32271",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 275,
        "date": "16/07/2569 09:06",
        "complainant": "ปณิศอร บุญจูบุตร",
        "email": "-",
        "anydesk": "-",
        "issue": "บัญชีผู้ใช้: Email",
        "cause": "-",
        "duration": "00.05",
        "responder": "ปณิศอร  บุญจูบุตร   (แต๊นซ์ )",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 276,
        "date": "16/07/2569 09:47",
        "complainant": "วิจิตราภรณ์ พึ่งจันดุม",
        "email": "Wijitraporn.p@gmail.com",
        "anydesk": ".",
        "issue": "ฮาร์ดแวร์: PC Computer",
        "cause": "-",
        "duration": "00.30",
        "responder": "วิจิตราภรณ์ พึ่งจันดุม (พลอย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 277,
        "date": "16/07/2569 12:34",
        "complainant": "วิจิตราภรณ์ พึ่งจันดุม",
        "email": "wijitraporn.p@gmail.com",
        "anydesk": ".",
        "issue": "ฮาร์ดแวร์: PC Computer",
        "cause": "-",
        "duration": "00.30",
        "responder": "วิจิตราภรณ์ พึ่งจันดุม (พลอย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 278,
        "date": "16/07/2569 13:58",
        "complainant": "เนตรปรีญา ทัดศรี",
        "email": "-",
        "anydesk": "1668042079",
        "issue": "ฮาร์ดแวร์: Printer",
        "cause": "-",
        "duration": "00.30",
        "responder": "เนตรปรีญา ทัดศรี(แตงกวา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": "EXP-1",
        "date": "13/07/2569 14:04",
        "complainant": "รามจิตติ ชินนะเกิดโชค",
        "issue": "ค่าใช้จ่าย IT: ซื้ออุปกรณ์ (เสื่อมตามสภาพ)",
        "duration": "-",
        "responder": "-",
        "status": "จ่ายเงินแล้ว",
        "cost": 13900
      }
    ],
    "automationsDone": 1,
    "aiApps": 2,
    "hoursSaved": 15
  }
};









const AssetTags = ({ value, empty = '-' }) => {
  const tags = String(value || '').split(/[,\n]+/).map((item) => item.trim()).filter(Boolean);
  if (tags.length === 0) return empty;
  return (
    <div className="asset-tags">
      {tags.map((tag, index) => {
        const colorIndex = [...tag].reduce((sum, char) => sum + char.charCodeAt(0), 0) % 6;
        return <span key={`${tag}-${index}`} className={`asset-tag asset-tag-${colorIndex}`}>{tag}</span>;
      })}
    </div>
  );
};

const seedSoftwareLicenses = [
  { name: 'Meitu', owner: 'Tiktok Content Creator', price: 1290, paymentChannel: 'Apple', paymentDate: 'รายปี', expiringDate: '', registeredEmail: 'drfernaesthetique@gmail.com', currentUsers: '' },
  { name: 'Adobe', owner: 'กราฟฟิก', price: 2592, paymentChannel: 'บัตร', paymentDate: '02/07/2026', expiringDate: '2026-07-31', registeredEmail: 'drfernaesthetique@gmail.com', currentUsers: 'อาทิตยา มุมทอง (ขมิ้น), รวมจิตต์ จันทร์เกิดโชค (เบนซ์)' },
  { name: 'Freepik', owner: 'กราฟฟิก', price: 11250, paymentChannel: 'บัตร', paymentDate: '11/05/2026-2027 (1 ปี)', expiringDate: '2027-06-11', registeredEmail: 'graphicfernclinic@gmail.com', currentUsers: 'อาทิตยา มุมทอง (ขมิ้น), รวมจิตต์ จันทร์เกิดโชค (เบนซ์), ชัยธัช ชัยวัฒน์ (มาร์ค)' },
  { name: 'Kumoo', owner: 'กราฟฟิก', price: 3077, paymentChannel: 'บัตร', paymentDate: '06/01/2026', expiringDate: '2027-01-07', registeredEmail: 'drfernaesthetique@gmail.com', currentUsers: 'อาทิตยา มุมทอง (ขมิ้น), รวมจิตต์ จันทร์เกิดโชค (เบนซ์), ชัยธัช ชัยวัฒน์ (มาร์ค)' },
  { name: 'Cupcut', owner: 'Tiktok Content Creator', price: 345, paymentChannel: 'Apple', paymentDate: '31/07/2026', expiringDate: '2026-08-31', registeredEmail: 'drfernaesthetique@gmail.com', currentUsers: 'บุษกร บัวสวรรค์ (เรนนี่), อภิสิทธิ์ พรจันทร์วัฒน์ (จุ้ย)' },
  { name: 'Microsoft Office 365', owner: 'IT', price: 3690, paymentChannel: 'Microsoft Office', paymentDate: '', expiringDate: '', registeredEmail: '', currentUsers: '' },
  { name: 'Lark', owner: 'IT', price: 0, paymentChannel: '', paymentDate: '', expiringDate: '', registeredEmail: '', currentUsers: '' },
  { name: 'Google Suite', owner: 'IT', price: 0, paymentChannel: '', paymentDate: '', expiringDate: '', registeredEmail: '', currentUsers: '' },
  { name: 'Chat GPT', owner: 'IT', price: 0, paymentChannel: '', paymentDate: '', expiringDate: '', registeredEmail: '', currentUsers: '' },
  { name: 'Adobe', owner: 'กราฟฟิก', price: 11105, paymentChannel: 'บัตร', paymentDate: '2025-12-01', expiringDate: '2026-12-01', registeredEmail: 'fernclinic.it@gmail.com', currentUsers: 'ชัยธัช ชัยวัฒน์ (มาร์ค), พิชญาพร คลอวง (พี่เจน)' },
  { name: 'Cupcut', owner: 'Tiktok Content Creator', price: 1810, paymentChannel: 'Apple', paymentDate: '25/06/2027', expiringDate: '2027-07-23', registeredEmail: 'drfernbussiness@gmail.com', currentUsers: '' },
  { name: 'Meitu', owner: 'Tiktok Content Creator', price: 1190, paymentChannel: 'Applepay', paymentDate: '22/01/2027', expiringDate: '2027-01-22', registeredEmail: 'drfernbussiness@gmail.com', currentUsers: '' },
  { name: 'PEAK', owner: 'Accounting', price: 12480, paymentChannel: '', paymentDate: 'รายปี', expiringDate: '2070-03-30', registeredEmail: '', currentUsers: '' },
  { name: 'empeo', owner: 'HR', price: 132515, paymentChannel: '', paymentDate: 'รายปี', expiringDate: '2070-06-09', registeredEmail: '', currentUsers: '' },
  { name: 'Chromecast Premium', owner: 'HR', price: 399, paymentChannel: '', paymentDate: 'รายเดือน', expiringDate: '', registeredEmail: '', currentUsers: '' },
].map((item) => {
  const used = item.currentUsers ? item.currentUsers.split(',').filter(Boolean).length : 0;
  return { ...item, used, vacant: 0, licenses: used, monthlyCost: item.price, status: 'ใช้งาน', isLicenseRecord: true, sourceVersion: 'software-image-v2' };
});

const isValidDashboardData = (value) => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const monthKeys = Object.keys(value);
  return monthKeys.length > 0 && monthKeys.every((key) => /^\d{4}-\d{2}$/.test(key) && value[key] && typeof value[key] === 'object');
};

function Dashboard() {
  const navigate = useNavigate();
  const [data, setData] = useState(() => {
    const saved = localStorage.getItem('it_dashboard_data');
    if (!saved) return initialDashboardData;
    try {
      const parsed = JSON.parse(saved);
      return isValidDashboardData(parsed) ? parsed : initialDashboardData;
    } catch {
      return initialDashboardData;
    }
  });
  const [assetsList, setAssetsList] = useState(() => {
    const saved = localStorage.getItem('it_dashboard_assets');
    return saved ? JSON.parse(saved) : initialAssetsData;
  });
  const [assetSearch, setAssetSearch] = useState('');
  const [assetDeptFilter, setAssetDeptFilter] = useState('');
  const [assetStatusFilter, setAssetStatusFilter] = useState('');
  const [currentMonth, setCurrentMonth] = useState("2026-07");
  const [activeModal, setActiveModal] = useState(null); // 'edit', 'expiringAssets', 'expiringSoftware', 'topBrokenDevices', 'assetsList', 'fullConsole'
  const [importStatus, setImportStatus] = useState(null); // { type: 'success' | 'error', message: string }
  
  // Console tab, month selectors, and editing state trackers
  const [consoleTab, setConsoleTab] = useState('months');
  const [consoleMonth, setConsoleMonth] = useState('2026-07');
  const [editingAssetSn, setEditingAssetSn] = useState(null);
  const [editingTicketSn, setEditingTicketSn] = useState(null);
  const [editingSoftwareIndex, setEditingSoftwareIndex] = useState(null);
  const [softwareName, setSoftwareName] = useState('');
  const [softwareUsed, setSoftwareUsed] = useState('0');
  const [softwareVacant, setSoftwareVacant] = useState('0');
  const [softwareExpiryDate, setSoftwareExpiryDate] = useState('');
  const [softwareMonthlyCost, setSoftwareMonthlyCost] = useState('0');
  const [softwareOwner, setSoftwareOwner] = useState('');
  const [softwarePaymentChannel, setSoftwarePaymentChannel] = useState('');
  const [softwarePaymentDate, setSoftwarePaymentDate] = useState('');
  const [softwareRegisteredEmail, setSoftwareRegisteredEmail] = useState('');
  const [softwareCurrentUsers, setSoftwareCurrentUsers] = useState('');

  // Lark Form states
  const [larkFormType, setLarkFormType] = useState('ticket'); // 'ticket' | 'asset'
  const [larkTicketComplainant, setLarkTicketComplainant] = useState('');
  const [larkTicketEmail, setLarkTicketEmail] = useState('');
  const [larkTicketAnydesk, setLarkTicketAnydesk] = useState('');
  const [larkTicketIssue, setLarkTicketIssue] = useState('');
  const [larkTicketCause, setLarkTicketCause] = useState('');
  const [larkTicketDuration, setLarkTicketDuration] = useState('00:30');
  const [larkTicketResponder, setLarkTicketResponder] = useState('');
  const [larkTicketStatus, setLarkTicketStatus] = useState('เสร็จสิ้น');
  const [larkTicketCost, setLarkTicketCost] = useState('0');

  const [larkAssetUser, setLarkAssetUser] = useState('');
  const [larkAssetPosition, setLarkAssetPosition] = useState('');
  const [larkAssetItemType, setLarkAssetItemType] = useState('');
  const [larkAssetSerial, setLarkAssetSerial] = useState('');
  const [larkAssetStatus, setLarkAssetStatus] = useState('ใช้งาน');
  const [larkAssetNotes, setLarkAssetNotes] = useState('');

  const [larkSubmitted, setLarkSubmitted] = useState(false);
  const [larkTicketRole, setLarkTicketRole] = useState('user'); // 'user' | 'it'
  const [selectedPendingTicketSn, setSelectedPendingTicketSn] = useState('');
  const [sidebarExpanded, setSidebarExpanded] = useState({
    mgmt: true,
    excel: false,
    export: false
  });
  const [mobileSidebarOpen, setMobileSidebarOpen] = useState(false);
  const [isLoaded, setIsLoaded] = useState(false);
  const API_BASE = import.meta.env.VITE_API_BASE || 
    (window.location.hostname === 'localhost' ? 'http://localhost:5000' : window.location.origin);

  const isPollingUpdateRef = useRef(false);
  const isPendingSyncRef = useRef(false);
  const latestDataRef = useRef(data);
  const latestAssetsRef = useRef(assetsList);
  const softwareSeededRef = useRef(false);

  useEffect(() => {
    latestDataRef.current = data;
    latestAssetsRef.current = assetsList;
  }, [data, assetsList]);

  const syncStateToDb = async (updatedData, updatedAssets) => {
    try {
      await fetch(`${API_BASE}/api/sync-all`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ data: updatedData, assetsList: updatedAssets })
      });
    } catch (err) {
      console.error('Failed to sync state to PostgreSQL database:', err);
    }
  };

  // Load state from Render PostgreSQL database on mount
  useEffect(() => {
    async function loadDbState() {
      try {
        const res = await fetch(`${API_BASE}/api/db-state`);
        if (!res.ok) throw new Error('API server returned error');
        const result = await res.json();
        
        if (result.data && Object.keys(result.data).length > 0) {
          isPollingUpdateRef.current = true;
          setData(result.data);
          if (result.assetsList) {
            setAssetsList(result.assetsList);
          }
          console.log('Successfully synced dashboard state with Render PostgreSQL database.');
        } else {
          // Database is empty. Seed it with the default initial data!
          console.log('PostgreSQL database is empty. Seeding initial baseline datasets...');
          await fetch(`${API_BASE}/api/sync-all`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ data, assetsList })
          });
        }
        const inventoryRes = await fetch(`${API_BASE}/api/inventory-data`);
        if (inventoryRes.ok) {
          const inventoryAssets = await inventoryRes.json();
          if (inventoryAssets.length > 0) setAssetsList(inventoryAssets);
        }
      } catch (err) {
        console.warn('Could not connect to Render PostgreSQL API server. Operating in offline/localStorage mode.', err);
      } finally {
        setIsLoaded(true);
      }
    }
    loadDbState();
  }, []);

  // Poll database state every 3 seconds for real-time synchronization
  useEffect(() => {
    if (!isLoaded) return;

    const interval = setInterval(async () => {
      // If we have local changes pending database sync, skip polling to avoid race overwrites!
      if (isPendingSyncRef.current) {
        console.log('Sync in progress, skipping polling interval.');
        return;
      }

      try {
        const res = await fetch(`${API_BASE}/api/db-state`);
        if (!res.ok) return;
        const result = await res.json();

        // Check again after fetch completes in case sync was initiated during network roundtrip
        if (isPendingSyncRef.current) return;

        const currentLocalData = latestDataRef.current;
        const currentLocalAssets = latestAssetsRef.current;

        const hasDataChanged = JSON.stringify(result.data) !== JSON.stringify(currentLocalData);
        const hasAssetsChanged = JSON.stringify(result.assetsList) !== JSON.stringify(currentLocalAssets);

        if (hasDataChanged || hasAssetsChanged) {
          isPollingUpdateRef.current = true;
          if (hasDataChanged) setData(result.data);
          if (hasAssetsChanged) setAssetsList(result.assetsList);
          console.log('Real-time update synced from Render PostgreSQL.');
        }
      } catch (err) {
        console.warn('Real-time sync poll failed:', err.message);
      }
    }, 3000);

    return () => clearInterval(interval);
  }, [isLoaded]);

  // Automatically sync local changes to PostgreSQL database once loaded
  useEffect(() => {
    if (!isLoaded || softwareSeededRef.current || !data[currentMonth]) return;
    softwareSeededRef.current = true;
    const existing = data[currentMonth].softwareExpiringDetails || [];
    if (existing.some((item) => item.sourceVersion === 'software-image-v2')) return;
    const preservedDetails = existing.filter((item) => !item.isLicenseRecord);
    setData((previous) => ({
      ...previous,
      [currentMonth]: {
        ...previous[currentMonth],
        softwareExpiringDetails: [...preservedDetails, ...seedSoftwareLicenses],
      },
    }));
  }, [isLoaded, currentMonth, data]);

  useEffect(() => {
    if (!isLoaded) return;
    if (isPollingUpdateRef.current) {
      isPollingUpdateRef.current = false;
      return;
    }
    
    // Mark pending sync lock immediately
    isPendingSyncRef.current = true;

    const timer = setTimeout(async () => {
      try {
        await syncStateToDb(data, assetsList);
      } finally {
        isPendingSyncRef.current = false;
      }
    }, 1000);
    return () => clearTimeout(timer);
  }, [data, assetsList, isLoaded]);
  
  // New Month creation states
  const [newMonthKey, setNewMonthKey] = useState('');
  const [newMonthName, setNewMonthName] = useState('');

  // New Project/Rec creation states
  const [newProjTitle, setNewProjTitle] = useState('');
  const [newProjDesc, setNewProjDesc] = useState('');
  const [newRecText, setNewRecText] = useState('');

  // New Asset creation states
  const [newAssetUser, setNewAssetUser] = useState('');
  const [newAssetPosition, setNewAssetPosition] = useState('');
  const [newAssetItemType, setNewAssetItemType] = useState('');
  const [newAssetSerial, setNewAssetSerial] = useState('');
  const [newAssetStatus, setNewAssetStatus] = useState('ใช้งาน');
  const [newAssetNotes, setNewAssetNotes] = useState('');

  // New Ticket creation states
  const [newTicketComplainant, setNewTicketComplainant] = useState('');
  const [newTicketEmail, setNewTicketEmail] = useState('');
  const [newTicketAnydesk, setNewTicketAnydesk] = useState('');
  const [newTicketIssue, setNewTicketIssue] = useState('');
  const [newTicketCause, setNewTicketCause] = useState('');
  const [newTicketDuration, setNewTicketDuration] = useState('00:30');
  const [newTicketResponder, setNewTicketResponder] = useState('');
  const [newTicketStatus, setNewTicketStatus] = useState('เสร็จสิ้น');
  const [newTicketCost, setNewTicketCost] = useState('0');

  // Form input states
  const [formInputs, setFormInputs] = useState({});

  // Sync to local storage
  useEffect(() => {
    localStorage.setItem('it_dashboard_data', JSON.stringify(data));
  }, [data]);

  useEffect(() => {
    localStorage.setItem('it_dashboard_assets', JSON.stringify(assetsList));
  }, [assetsList]);

  // Upgraded Dynamic Recalculation Engine
  const recalculateMonthlyMetrics = (monthKey, tickets, assets) => {
    let durationSum = 0;
    let durationCount = 0;
    let slaCompliantCount = 0;
    let totalCost = 0;
    let repairCount = 0;
    const deviceCounts = {};
    const deptCosts = {};

    tickets.forEach(ticket => {
      const duration = ticket.duration || '-';
      if (duration && duration !== '-') {
        const parts = duration.split(':');
        if (parts.length === 2) {
          const mins = parseInt(parts[0]) * 60 + parseInt(parts[1]);
          if (!isNaN(mins)) {
            durationSum += mins;
            durationCount++;
            if (mins <= 60) {
              slaCompliantCount++;
            }
          }
        }
      }
      
      const cost = Number(ticket.cost) || 0;
      totalCost += cost;
      if (cost > 0 || ticket.status === 'จ่ายเงินแล้ว') {
        repairCount++;
      }

      // Track broken devices by category
      const issue = String(ticket.issue).toLowerCase();
      let matchedDevice = 'อื่น ๆ';
      if (issue.includes('notebook') || issue.includes('lenovo') || issue.includes('asus') || issue.includes('hp')) matchedDevice = 'Notebook';
      else if (issue.includes('computer') || issue.includes('pc')) matchedDevice = 'PC';
      else if (issue.includes('ipad')) matchedDevice = 'iPad';
      else if (issue.includes('iphone')) matchedDevice = 'iPhone';
      else if (issue.includes('printer') || issue.includes('ปริ้นเตอร์')) matchedDevice = 'Printer';
      else if (issue.includes('mornitor') || issue.includes('จอ')) matchedDevice = 'Monitor';
      else if (issue.includes('imac')) matchedDevice = 'iMac';
      else if (issue.includes('macbook')) matchedDevice = 'MacBook';
      else if (issue.includes('network') || issue.includes('lan') || issue.includes('wifi') || issue.includes('เน็ต')) matchedDevice = 'Network';

      if (ticket.status !== 'เสร็จสิ้น' || cost > 0) {
        deviceCounts[matchedDevice] = (deviceCounts[matchedDevice] || 0) + 1;
      }

      // Department costs mapping
      const borrowerAsset = assets.find(a => String(a.user).trim() === String(ticket.complainant).trim());
      const dept = borrowerAsset ? borrowerAsset.position : 'ส่วนกลาง';
      if (cost > 0) {
        deptCosts[dept] = (deptCosts[dept] || 0) + cost;
      }
    });

    const calculatedSla = durationCount > 0 ? Math.round((slaCompliantCount / durationCount) * 1000) / 10 : 100;
    const resolutionTimeHours = durationCount > 0 ? Number((durationSum / durationCount / 60).toFixed(1)) : 0.5;
    const calculatedResponseTime = Math.max(5, Math.round(resolutionTimeHours * 12));
    const calculatedCsat = Number((4.5 + (calculatedSla / 100) * 0.4).toFixed(1));

    // Estimate asset value dynamically based on category
    const CATEGORY_VALUES = {
      "Computer (Pc)": 15000,
      "Ipad": 18000,
      "Mornitor": 5000,
      "Notebook Lenovo": 25000,
      "External HDD": 1500,
      "Cable HDMI": 500,
      "Mouse": 1000,
      "Keyboard": 1500,
      "Screwdriver": 1200,
      "Notebook Asus": 20000,
      "Printer": 7500,
      "IPhone": 25000,
      "Apple Pancill": 3900,
      "Macbook": 45000,
      "Hub USB-TypeC": 1200,
      "IMac": 40000,
      "Hub Lan": 2500,
      "Notebook Acer, Iphone": 35000,
      "Adapter Apple": 1200,
      "Cable Apple": 790,
      "Notebook HP": 22000,
      "Samsung": 15000,
      "Capture Card": 3500,
      "Reez Live": 15000,
      "Iphone": 25000
    };

    let calculatedAssetValue = 0;
    assets.forEach(asset => {
      const cat = String(asset.itemType || '').trim();
      calculatedAssetValue += CATEGORY_VALUES[cat] || 1500;
    });

    const brokenAssetsCount = assets.filter(a => a.status === 'รอซ่อม').length;
    const lostAssetsCount = assets.filter(a => a.status === 'สูญหาย').length;
    const vacantAssetsCount = assets.filter(a => a.status === 'ว่าง').length;

    const topBrokenDevices = Object.entries(deviceCounts)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);

    const totalAssetsVal = assets.length;
    const licensesInUseVal = Math.round(totalAssetsVal * 1.5);
    const licensesVacantVal = Math.round(licensesInUseVal * 0.15);
    const softwareCostVal = licensesInUseVal * 450;

    return {
      totalAssets: totalAssetsVal,
      assetValue: calculatedAssetValue,
      assetsBroken: brokenAssetsCount,
      assetsLost: lostAssetsCount,
      assetsVacant: vacantAssetsCount,
      ticketsCount: tickets.length,
      slaPercent: calculatedSla,
      resolutionTime: resolutionTimeHours,
      responseTime: calculatedResponseTime,
      csat: calculatedCsat,
      licensesInUse: licensesInUseVal,
      licensesVacant: licensesVacantVal,
      softwareCost: softwareCostVal,
      repairCount: repairCount,
      repairCost: totalCost,
      topBrokenDevices: topBrokenDevices,
      deptCosts: deptCosts
    };
  };

  const runRecalculation = (monthKey, currentTickets, currentAssets) => {
    const newMetrics = recalculateMonthlyMetrics(monthKey, currentTickets, currentAssets);
    setData(prev => ({
      ...prev,
      [monthKey]: {
        ...prev[monthKey],
        ...newMetrics,
        ticketsList: currentTickets
      }
    }));
  };

  // Recalculate metrics for all months whenever assetsList changes
  useEffect(() => {
    setData(prev => {
      let updated = false;
      const nextData = { ...prev };
      Object.keys(nextData).forEach(monthKey => {
        const monthData = nextData[monthKey];
        const tickets = monthData.ticketsList || [];
        const newMetrics = recalculateMonthlyMetrics(monthKey, tickets, assetsList);
        
        let changed = false;
        for (const k of Object.keys(newMetrics)) {
          if (JSON.stringify(newMetrics[k]) !== JSON.stringify(monthData[k])) {
            changed = true;
            break;
          }
        }
        if (changed) {
          nextData[monthKey] = {
            ...monthData,
            ...newMetrics
          };
          updated = true;
        }
      });
      return updated ? nextData : prev;
    });
  }, [assetsList]);

  // Console Months Manager Handlers
  const handleAddMonth = (e) => {
    e.preventDefault();
    if (!newMonthKey || !newMonthName) {
      alert('กรุณากรอกข้อมูลให้ครบถ้วน');
      return;
    }
    if (data[newMonthKey]) {
      alert('มีรหัสเดือนนี้ในระบบอยู่แล้ว');
      return;
    }
    setData(prev => ({
      ...prev,
      [newMonthKey]: {
        monthName: newMonthName,
        totalAssets: assetsList.length,
        assetValue: data[currentMonth]?.assetValue || 0,
        assetsExpiring: 0,
        assetsBroken: 0,
        assetsLost: 0,
        assetsVacant: 0,
        ticketsCount: 0,
        slaPercent: 100,
        responseTime: 0,
        resolutionTime: 0,
        csat: 5.0,
        totalSoftware: 0,
        licensesInUse: 0,
        licensesVacant: 0,
        softwareCost: 0,
        softwareExpiring: 0,
        backupSuccess: 100,
        securityIncidents: 0,
        antivirusCoverage: 100,
        mfaCoverage: 100,
        repairCount: 0,
        repairCost: 0,
        topBrokenDevices: [],
        deptCosts: {},
        softwareExpiringDetails: [],
        assetsExpiringDetails: [],
        ongoingProjects: [],
        recommendations: [],
        ticketsList: []
      }
    }));
    setConsoleMonth(newMonthKey);
    setNewMonthKey('');
    setNewMonthName('');
    alert(`เพิ่มเดือน ${newMonthName} สำเร็จ!`);
  };

  const handleDeleteMonth = (key) => {
    const keys = Object.keys(data);
    if (keys.length <= 1) {
      alert('ไม่สามารถลบเดือนสุดท้ายของระบบได้');
      return;
    }
    if (window.confirm(`คุณแน่ใจว่าต้องการลบเดือน ${data[key].monthName} ใช่หรือไม่?`)) {
      setData(prev => {
        const copy = { ...prev };
        delete copy[key];
        return copy;
      });
      if (currentMonth === key) {
        setCurrentMonth(keys.find(k => k !== key));
      }
      if (consoleMonth === key) {
        setConsoleMonth(keys.find(k => k !== key));
      }
    }
  };

  // Inline KPI Change handler
  const handleKpiChange = (field, val) => {
    setData(prev => ({
      ...prev,
      [consoleMonth]: {
        ...prev[consoleMonth],
        [field]: val
      }
    }));
  };

  // Projects & Recs Editor Handlers
  const handleAddProject = () => {
    if (!newProjTitle) return;
    setData(prev => {
      const monthData = prev[consoleMonth] || {};
      return {
        ...prev,
        [consoleMonth]: {
          ...monthData,
          ongoingProjects: [...(monthData.ongoingProjects || []), { title: newProjTitle, desc: newProjDesc }]
        }
      };
    });
    setNewProjTitle('');
    setNewProjDesc('');
  };

  const handleDeleteProject = (idx) => {
    setData(prev => {
      const monthData = prev[consoleMonth];
      const updatedProjects = [...(monthData.ongoingProjects || [])];
      updatedProjects.splice(idx, 1);
      return {
        ...prev,
        [consoleMonth]: {
          ...monthData,
          ongoingProjects: updatedProjects
        }
      };
    });
  };

  const handleAddRecommendation = () => {
    if (!newRecText) return;
    setData(prev => {
      const monthData = prev[consoleMonth] || {};
      return {
        ...prev,
        [consoleMonth]: {
          ...monthData,
          recommendations: [...(monthData.recommendations || []), newRecText]
        }
      };
    });
    setNewRecText('');
  };

  const handleDeleteRecommendation = (idx) => {
    setData(prev => {
      const monthData = prev[consoleMonth];
      const updatedRecs = [...(monthData.recommendations || [])];
      updatedRecs.splice(idx, 1);
      return {
        ...prev,
        [consoleMonth]: {
          ...monthData,
          recommendations: updatedRecs
        }
      };
    });
  };

  // Assets Inventory Editor (Dual-Mode: Create / Update)
  const handleAddAsset = () => {
    if (!newAssetItemType) {
      alert('กรุณากรอกประเภทอุปกรณ์หลัก');
      return;
    }

    if (editingAssetSn !== null) {
      // Edit mode
      setAssetsList(prev => {
        const updated = prev.map(a => a.sn === editingAssetSn ? {
          ...a,
          user: newAssetUser || 'ส่วนกลาง',
          position: newAssetPosition || '-',
          itemType: newAssetItemType,
          deviceSerial: newAssetSerial || '-',
          status: newAssetStatus,
          notes: newAssetNotes
        } : a);
        runRecalculation(consoleMonth, data[consoleMonth]?.ticketsList || [], updated);
        return updated;
      });
      setEditingAssetSn(null);
      alert('แก้ไขข้อมูลทรัพย์สินสำเร็จ!');
    } else {
      // Create mode
      const newAsset = {
        sn: assetsList.length > 0 ? Math.max(...assetsList.map(a => Number(a.sn) || 0)) + 1 : 1,
        date: new Date().toLocaleDateString('th-TH'),
        user: newAssetUser || 'ส่วนกลาง',
        position: newAssetPosition || '-',
        itemType: newAssetItemType,
        deviceSerial: newAssetSerial || '-',
        status: newAssetStatus,
        notes: newAssetNotes
      };
      setAssetsList(prev => {
        const updated = [...prev, newAsset];
        runRecalculation(consoleMonth, data[consoleMonth]?.ticketsList || [], updated);
        return updated;
      });
      alert('เพิ่มทรัพย์สินเข้าคลังสำเร็จ!');
    }

    setNewAssetUser('');
    setNewAssetPosition('');
    setNewAssetItemType('');
    setNewAssetSerial('');
    setNewAssetStatus('ใช้งาน');
    setNewAssetNotes('');
  };

  const handleLoadEditAsset = (asset) => {
    setEditingAssetSn(asset.sn);
    setNewAssetUser(asset.user);
    setNewAssetPosition(asset.position);
    setNewAssetItemType(asset.itemType);
    setNewAssetSerial(asset.deviceSerial);
    setNewAssetStatus(asset.status);
    setNewAssetNotes(asset.notes || '');
  };

  const handleCancelEditAsset = () => {
    setEditingAssetSn(null);
    setNewAssetUser('');
    setNewAssetPosition('');
    setNewAssetItemType('');
    setNewAssetSerial('');
    setNewAssetStatus('ใช้งาน');
    setNewAssetNotes('');
  };

  const handleDeleteAsset = (sn) => {
    if (window.confirm('คุณต้องการลบอุปกรณ์นี้ออกจากทะเบียนคลังใช่หรือไม่?')) {
      setAssetsList(prev => {
        const updated = prev.filter(a => a.sn !== sn);
        runRecalculation(consoleMonth, data[consoleMonth]?.ticketsList || [], updated);
        return updated;
      });
      if (editingAssetSn === sn) {
        handleCancelEditAsset();
      }
    }
  };

  // Ticket Log Editor (Dual-Mode: Create / Update)
  const handleAddTicket = () => {
    if (!newTicketIssue) {
      alert('กรุณากรอกอาการเสีย/ปัญหา');
      return;
    }

    const tickets = data[consoleMonth]?.ticketsList || [];

    if (editingTicketSn !== null) {
      // Edit mode
      const updatedTickets = tickets.map(t => t.sn === editingTicketSn ? {
        ...t,
        complainant: newTicketComplainant || 'ไม่ระบุชื่อ',
        email: newTicketEmail || '-',
        anydesk: newTicketAnydesk || '-',
        issue: newTicketIssue,
        cause: newTicketCause || '-',
        duration: newTicketDuration,
        responder: newTicketResponder || '-',
        status: newTicketStatus,
        cost: Number(newTicketCost) || 0
      } : t);

      runRecalculation(consoleMonth, updatedTickets, assetsList);
      setEditingTicketSn(null);
      alert('แก้ไขข้อมูลงานแจ้งซ่อมสำเร็จ!');
    } else {
      // Create mode
      const newTicket = {
        sn: tickets.length > 0 ? Math.max(...tickets.map(t => Number(t.sn) || 0)) + 1 : 1,
        date: new Date().toLocaleString('th-TH', { hour12: false }).replace(',', ''),
        complainant: newTicketComplainant || 'ไม่ระบุชื่อ',
        email: newTicketEmail || '-',
        anydesk: newTicketAnydesk || '-',
        issue: newTicketIssue,
        cause: newTicketCause || '-',
        duration: newTicketDuration,
        responder: newTicketResponder || '-',
        status: newTicketStatus,
        cost: Number(newTicketCost) || 0
      };

      const updatedTickets = [...tickets, newTicket];
      runRecalculation(consoleMonth, updatedTickets, assetsList);
      alert('เพิ่มประวัติงานแจ้งซ่อมสำเร็จ!');
    }

    setNewTicketComplainant('');
    setNewTicketEmail('');
    setNewTicketAnydesk('');
    setNewTicketIssue('');
    setNewTicketCause('');
    setNewTicketDuration('00:30');
    setNewTicketResponder('');
    setNewTicketStatus('เสร็จสิ้น');
    setNewTicketCost('0');
  };

  const handleLoadEditTicket = (ticket) => {
    setEditingTicketSn(ticket.sn);
    setNewTicketComplainant(ticket.complainant);
    setNewTicketEmail(ticket.email || '');
    setNewTicketAnydesk(ticket.anydesk || '');
    setNewTicketIssue(ticket.issue);
    setNewTicketCause(ticket.cause || '');
    setNewTicketDuration(ticket.duration);
    setNewTicketResponder(ticket.responder);
    setNewTicketStatus(ticket.status);
    setNewTicketCost(String(ticket.cost || 0));
  };

  const handleCancelEditTicket = () => {
    setEditingTicketSn(null);
    setNewTicketComplainant('');
    setNewTicketEmail('');
    setNewTicketAnydesk('');
    setNewTicketIssue('');
    setNewTicketCause('');
    setNewTicketDuration('00:30');
    setNewTicketResponder('');
    setNewTicketStatus('เสร็จสิ้น');
    setNewTicketCost('0');
  };

  const handleDeleteTicket = (sn) => {
    if (window.confirm('คุณต้องการลบรายการแจ้งซ่อมนี้ใช่หรือไม่?')) {
      const updatedTickets = (data[consoleMonth]?.ticketsList || []).filter(t => t.sn !== sn);
      runRecalculation(consoleMonth, updatedTickets, assetsList);
      if (editingTicketSn === sn) {
        handleCancelEditTicket();
      }
    }
  };

  // JSON operations and reset
  const handleExportJson = () => {
    const backupData = {
      data,
      assetsList
    };
    const blob = new Blob([JSON.stringify(backupData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `it_dashboard_backup_${new Date().toISOString().split('T')[0]}.json`;
    link.click();
    URL.revokeObjectURL(url);
  };

  const handleImportJson = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const parsed = JSON.parse(evt.target.result);
        if (parsed.data && parsed.assetsList) {
          setData(parsed.data);
          setAssetsList(parsed.assetsList);
          alert('นำเข้าข้อมูลสำรองสำเร็จ!');
        } else {
          alert('รูปแบบไฟล์สำรองไม่ถูกต้อง');
        }
      } catch (err) {
        alert('เกิดข้อผิดพลาดในการอ่านไฟล์: ' + err.message);
      }
    };
    reader.readAsText(file);
  };

  const handleResetToDefault = () => {
    if (window.confirm('คุณต้องการรีเซ็ตข้อมูลทั้งหมดกลับเป็นค่าเริ่มต้นจากไฟล์ Excel ดั้งเดิมใช่หรือไม่? ข้อมูลที่คุณแก้ไขจะหายไปทั้งหมด')) {
      localStorage.removeItem('it_dashboard_data');
      localStorage.removeItem('it_dashboard_assets');
      setData(initialDashboardData);
      setAssetsList(initialAssetsData);
      setCurrentMonth('2026-07');
      setConsoleMonth('2026-07');
      setEditingAssetSn(null);
      setEditingTicketSn(null);
      alert('รีเซ็ตข้อมูลกลับเป็นค่าเริ่มต้นเรียบร้อยแล้ว!');
    }
  };

  const handleLarkSubmit = (e) => {
    e.preventDefault();
    if (larkFormType === 'ticket') {
      const tickets = data[currentMonth]?.ticketsList || [];

      if (larkTicketRole === 'it') {
        if (!selectedPendingTicketSn) {
          alert('กรุณาเลือกใบงานที่ต้องการปิดงาน');
          return;
        }
        if (!larkTicketResponder) {
          alert('กรุณากรอกชื่อผู้ดำเนินงาน (ช่าง IT)');
          return;
        }
        
        // IT Close work mode
        const updatedTickets = tickets.map(t => Number(t.sn) === Number(selectedPendingTicketSn) ? {
          ...t,
          responder: larkTicketResponder,
          duration: larkTicketDuration || '00:30',
          cause: larkTicketCause || '-',
          cost: Number(larkTicketCost) || 0,
          status: larkTicketStatus
        } : t);

        runRecalculation(currentMonth, updatedTickets, assetsList);
        setSelectedPendingTicketSn('');
        setLarkTicketResponder('');
        setLarkTicketDuration('00:30');
        setLarkTicketCause('');
        setLarkTicketCost('0');
        setLarkTicketStatus('เสร็จสิ้น');
        setLarkSubmitted(true);
      } else {
        // User Submit Mode
        if (!larkTicketIssue) {
          alert('กรุณากรอกอาการเสีย/ปัญหา');
          return;
        }
        const newTicket = {
          sn: tickets.length > 0 ? Math.max(...tickets.map(t => Number(t.sn) || 0)) + 1 : 1,
          date: new Date().toLocaleString('th-TH', { hour12: false }).replace(',', ''),
          complainant: larkTicketComplainant || 'ไม่ระบุชื่อ',
          email: larkTicketEmail || '-',
          anydesk: larkTicketAnydesk || '-',
          issue: larkTicketIssue,
          cause: '-',
          duration: '-',
          responder: '-',
          status: 'กำลังดำเนินการ',
          cost: 0
        };

        const updatedTickets = [...tickets, newTicket];
        runRecalculation(currentMonth, updatedTickets, assetsList);
        
        // Clear inputs
        setLarkTicketComplainant('');
        setLarkTicketEmail('');
        setLarkTicketAnydesk('');
        setLarkTicketIssue('');
        setLarkSubmitted(true);
      }
    } else {
      if (!larkAssetItemType) {
        alert('กรุณากรอกประเภทอุปกรณ์หลัก');
        return;
      }
      const newAsset = {
        sn: assetsList.length > 0 ? Math.max(...assetsList.map(a => Number(a.sn) || 0)) + 1 : 1,
        date: new Date().toLocaleDateString('th-TH'),
        user: larkAssetUser || 'ส่วนกลาง',
        position: larkAssetPosition || '-',
        itemType: larkAssetItemType,
        deviceSerial: larkAssetSerial || '-',
        status: larkAssetStatus,
        notes: larkAssetNotes
      };

      setAssetsList(prev => {
        const updated = [...prev, newAsset];
        runRecalculation(currentMonth, data[currentMonth]?.ticketsList || [], updated);
        return updated;
      });

      // Clear inputs
      setLarkAssetUser('');
      setLarkAssetPosition('');
      setLarkAssetItemType('');
      setLarkAssetSerial('');
      setLarkAssetStatus('ใช้งาน');
      setLarkAssetNotes('');
      setLarkSubmitted(true);
    }
  };

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

  const availableMonthKeys = Object.keys(data).sort();
  const fallbackMonthKey = availableMonthKeys[availableMonthKeys.length - 1] || '2026-07';
  const activeDataSource = data[currentMonth] || data[fallbackMonthKey] || initialDashboardData['2026-07'];
  const activeData = {
    ...initialDashboardData['2026-07'],
    ...activeDataSource,
    monthName: activeDataSource?.monthName || fallbackMonthKey,
    totalAssets: Number(activeDataSource?.totalAssets || 0),
    ticketsCount: Number(activeDataSource?.ticketsCount || 0),
    csat: Number(activeDataSource?.csat || 0),
    deptCosts: activeDataSource?.deptCosts || {},
    topBrokenDevices: activeDataSource?.topBrokenDevices || [],
    recommendations: activeDataSource?.recommendations || [],
    ongoingProjects: activeDataSource?.ongoingProjects || [],
    assetsExpiringDetails: activeDataSource?.assetsExpiringDetails || [],
    softwareExpiringDetails: activeDataSource?.softwareExpiringDetails || [],
    ticketsList: activeDataSource?.ticketsList || [],
  };

  useEffect(() => {
    if (!data[currentMonth] && data[fallbackMonthKey]) setCurrentMonth(fallbackMonthKey);
  }, [currentMonth, data, fallbackMonthKey]);

  const mainAssetCategories = [
    { label: 'PC', usefulLifeYears: 5, match: (type) => /computer\s*\(pc\)|\bpc\b/i.test(type) },
    { label: 'Notebook', usefulLifeYears: 5, match: (type) => /notebook/i.test(type) },
    { label: 'iMac', usefulLifeYears: 5, match: (type) => /imac/i.test(type) },
    { label: 'MacBook', usefulLifeYears: 5, match: (type) => /mac\s*book/i.test(type) },
    { label: 'iPhone', usefulLifeYears: 4, match: (type) => /iphone/i.test(type) },
    { label: 'iPad', usefulLifeYears: 4, match: (type) => /ipad/i.test(type) },
  ];

  const mainAssetBreakdown = mainAssetCategories.map((category) => ({
    label: category.label,
    count: assetsList.filter((asset) => category.match(String(asset.itemType || ''))).length,
  }));

  const primaryVacantAssets = assetsList.filter((asset) => {
    const type = String(asset.itemType || '');
    return asset.status === 'ว่าง' && mainAssetCategories.some((category) => category.match(type));
  }).length;

  const parseAssetDate = (value) => {
    if (!value) return null;
    const text = String(value).trim();
    const thaiDate = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (thaiDate) {
      const year = Number(thaiDate[3]) > 2400 ? Number(thaiDate[3]) - 543 : Number(thaiDate[3]);
      const parsed = new Date(year, Number(thaiDate[2]) - 1, Number(thaiDate[1]));
      return Number.isNaN(parsed.getTime()) ? null : parsed;
    }
    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  };

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const nearExpiryLimit = new Date(today);
  nearExpiryLimit.setFullYear(nearExpiryLimit.getFullYear() + 1);

  const primaryExpiringAssets = assetsList.filter((asset) => {
    const type = String(asset.itemType || '');
    const category = mainAssetCategories.find((item) => item.match(type));
    const startDate = parseAssetDate(asset.purchaseDate) || parseAssetDate(asset.date);
    if (!category || !startDate || asset.status === 'สูญหาย') return false;

    const modelExpiryDate = new Date(startDate);
    modelExpiryDate.setFullYear(modelExpiryDate.getFullYear() + category.usefulLifeYears);
    return modelExpiryDate >= today && modelExpiryDate <= nearExpiryLimit;
  }).length;

  const softwareDetails = activeData?.softwareExpiringDetails || [];
  const configuredSoftwareLicenses = softwareDetails.filter((item) => item.isLicenseRecord);
  const hasConfiguredSoftwareLicenses = configuredSoftwareLicenses.length > 0;
  const calculatedLicensesInUse = hasConfiguredSoftwareLicenses
    ? configuredSoftwareLicenses.reduce((sum, item) => sum + Number(item.used || 0), 0)
    : Number(activeData?.licensesInUse || 0);
  const calculatedLicensesVacant = hasConfiguredSoftwareLicenses
    ? configuredSoftwareLicenses.reduce((sum, item) => sum + Number(item.vacant || 0), 0)
    : Number(activeData?.licensesVacant || 0);
  const calculatedSoftwareCost = hasConfiguredSoftwareLicenses
    ? configuredSoftwareLicenses.reduce((sum, item) => sum + Number(item.monthlyCost || 0), 0)
    : Number(activeData?.softwareCost || 0);
  const calculatedTotalSoftware = hasConfiguredSoftwareLicenses
    ? configuredSoftwareLicenses.length
    : Number(activeData?.totalSoftware || 0);

  const resetSoftwareForm = () => {
    setEditingSoftwareIndex(null);
    setSoftwareName('');
    setSoftwareUsed('0');
    setSoftwareVacant('0');
    setSoftwareExpiryDate('');
    setSoftwareMonthlyCost('0');
    setSoftwareOwner('');
    setSoftwarePaymentChannel('');
    setSoftwarePaymentDate('');
    setSoftwareRegisteredEmail('');
    setSoftwareCurrentUsers('');
  };

  const saveSoftwareLicense = (event) => {
    event.preventDefault();
    const license = {
      name: softwareName.trim(),
      used: Number(softwareUsed || 0),
      vacant: Number(softwareVacant || 0),
      licenses: Number(softwareUsed || 0) + Number(softwareVacant || 0),
      expiringDate: softwareExpiryDate,
      monthlyCost: Number(softwareMonthlyCost || 0),
      price: Number(softwareMonthlyCost || 0),
      owner: softwareOwner.trim(),
      paymentChannel: softwarePaymentChannel.trim(),
      paymentDate: softwarePaymentDate.trim(),
      registeredEmail: softwareRegisteredEmail.trim(),
      currentUsers: softwareCurrentUsers.trim(),
      status: 'ใช้งาน',
      isLicenseRecord: true,
    };
    if (!license.name) return;

    setData((previous) => {
      const rows = [...(previous[currentMonth].softwareExpiringDetails || [])];
      if (editingSoftwareIndex === null) rows.push(license);
      else rows[editingSoftwareIndex] = license;
      return {
        ...previous,
        [currentMonth]: { ...previous[currentMonth], softwareExpiringDetails: rows },
      };
    });
    resetSoftwareForm();
  };

  const editSoftwareLicense = (item, index) => {
    setEditingSoftwareIndex(index);
    setSoftwareName(item.name || '');
    setSoftwareUsed(String(item.used ?? item.licenses ?? 0));
    setSoftwareVacant(String(item.vacant ?? 0));
    setSoftwareExpiryDate(item.expiringDate || '');
    setSoftwareMonthlyCost(String(item.monthlyCost ?? 0));
    setSoftwareOwner(item.owner || '');
    setSoftwarePaymentChannel(item.paymentChannel || '');
    setSoftwarePaymentDate(item.paymentDate || '');
    setSoftwareRegisteredEmail(item.registeredEmail || '');
    setSoftwareCurrentUsers(item.currentUsers || '');
  };

  const deleteSoftwareLicense = (index) => {
    if (!window.confirm('ต้องการลบรายละเอียด License นี้ใช่หรือไม่?')) return;
    setData((previous) => ({
      ...previous,
      [currentMonth]: {
        ...previous[currentMonth],
        softwareExpiringDetails: previous[currentMonth].softwareExpiringDetails.filter((_, rowIndex) => rowIndex !== index),
      },
    }));
    resetSoftwareForm();
  };

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

      const normalAssets = activeData.totalAssets - activeData.assetsBroken - activeData.assetsLost - primaryExpiringAssets;
      const ctx = assetCanvasRef.current.getContext('2d');
      assetChartInst.current = new Chart(ctx, {
        type: 'doughnut',
        data: {
          labels: ['ปกติ', 'ใกล้หมดอายุ', 'ชำรุด', 'สูญหาย'],
          datasets: [{
            data: [normalAssets, primaryExpiringAssets, activeData.assetsBroken, activeData.assetsLost],
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
              data: [calculatedLicensesInUse],
              backgroundColor: 'rgba(59, 130, 246, 0.75)',
              borderColor: '#3b82f6',
              borderWidth: 1
            },
            {
              label: 'ว่าง (Vacant)',
              data: [calculatedLicensesVacant],
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
  }, [currentMonth, data, primaryExpiringAssets, calculatedLicensesInUse, calculatedLicensesVacant]);

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
        const newData = {};

        if (wb.SheetNames.includes('Inventory')) {
          // --- Parse Inventory spreadsheet ---
          const inventoryRows = XLSX.utils.sheet_to_json(wb.Sheets['Inventory']);
          
          const parseExcelDate = (serial) => {
            if (typeof serial === 'string') {
              const d = new Date(serial);
              return isNaN(d.getTime()) ? null : d;
            }
            if (typeof serial !== 'number') return null;
            const utc_days = Math.floor(serial - 25569);
            const utc_value = utc_days * 86400;
            const date_info = new Date(utc_value * 1000);
            
            const fractional_day = serial - Math.floor(serial) + 0.0000001;
            let total_seconds = Math.floor(86400 * fractional_day);
            
            const seconds = total_seconds % 60;
            total_seconds = Math.floor(total_seconds / 60);
            const minutes = total_seconds % 60;
            const hours = Math.floor(total_seconds / 60);
            
            return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
          };

          const formatDate = (date) => {
            if (!date) return '';
            const d = new Date(date);
            const day = String(d.getDate()).padStart(2, '0');
            const month = String(d.getMonth() + 1).padStart(2, '0');
            const year = d.getFullYear() + 543; // Thai year
            return `${day}/${month}/${year}`;
          };

          const parsedAssets = inventoryRows.map((row, i) => {
            let dateIssued = '';
            if (row['วันที่เบิกใช้งาน']) {
              const dObj = parseExcelDate(row['วันที่เบิกใช้งาน']);
              if (dObj) dateIssued = formatDate(dObj);
            }
            
            return {
              sn: row['Nember'] || (i + 1),
              date: dateIssued,
              user: row['บุคคลเบิกใช้อุปกรณ์'] || 'ส่วนกลาง/ไม่ระบุ',
              position: row['ตำแหน่ง'] || '-',
              itemType: row['รายการอุปกรณ์หลัก'] || 'อุปกรณ์เสริม/อื่นๆ',
              deviceSerial: row['หมายเลขอุปกรณ์ (เช่น  Ipad 016)'] || '-',
              status: row['สถานะ'] || 'ใช้งาน',
              notes: row['หมายเหตุ'] || '',
              submittedOn: row['Submitted on'] || '',
              respondent: row['Respondents'] || '',
              additionalEquipment: row['อุปกรณ์เพิ่มเติมที่ต้องการเบิก'] || '',
              softwareApp: row['ซอต์ฟแวร์/ App'] || '',
              registeredEmail: row['เมลที่ลงทะเบียน'] || '',
              additionalSerial: row['หมายเลขอุปกรณ์ เพิ่มเติม  (เช่น  สาย อะเเดปเตอร์ ipad-011))'] || '',
              returnDueDate: row['กำหนดคืนอุปกรณ์'] || '',
              inspectionDate: row['วันที่ตรวจสอบ'] || '',
              purchaseDate: row['วันที่ซื้อ'] || '',
              warrantyEndDate: row['วันหมดประกัน'] || '',
              expense: Number(row['ค่าใช้จ่าย']) || 0
            };
          });

          // Calculate asset value sum
          const CATEGORY_VALUES = {
            "Computer (Pc)": 15000,
            "Ipad": 18000,
            "Mornitor": 5000,
            "Notebook Lenovo": 25000,
            "External HDD": 1500,
            "Cable HDMI": 500,
            "Mouse": 1000,
            "Keyboard": 1500,
            "Screwdriver": 1200,
            "Notebook Asus": 20000,
            "Printer": 7500,
            "IPhone": 25000,
            "Apple Pancill": 3900,
            "Macbook": 45000,
            "Hub USB-TypeC": 1200,
            "IMac": 40000,
            "Computer (Pc), Mornitor": 20000,
            "Hub Lan": 2500,
            "Notebook Acer, Iphone": 35000,
            "Adapter Apple": 1200,
            "Cable Apple": 790,
            "Notebook HP": 22000,
            "Samsung": 15000,
            "Capture Card": 3500,
            "Reez Live": 15000,
            "Iphone": 25000
          };

          let calculatedAssetValue = 0;
          inventoryRows.forEach(row => {
            const cat = String(row['รายการอุปกรณ์หลัก'] || '').trim();
            calculatedAssetValue += CATEGORY_VALUES[cat] || 1500;
          });

          const brokenInventory = parsedAssets.filter(a => a.status === 'รอซ่อม').length;
          const lostInventory = parsedAssets.filter(a => a.status === 'สูญหาย').length;

          // Update data metrics
          setData(prev => {
            const updated = { ...prev };
            Object.keys(updated).forEach(monthKey => {
              updated[monthKey] = {
                ...updated[monthKey],
                totalAssets: parsedAssets.length,
                assetValue: calculatedAssetValue,
                assetsBroken: brokenInventory + (updated[monthKey].assetsBroken || 0),
                assetsLost: lostInventory + (updated[monthKey].assetsLost || 0)
              };
            });
            return updated;
          });

          setAssetsList(parsedAssets);
          setImportStatus({ type: 'success', message: `นำเข้าข้อมูลคลังอุปกรณ์สำเร็จ! พบอุปกรณ์ ${parsedAssets.length} รายการ` });
          setTimeout(() => setImportStatus(null), 4000);
          return;
        }

        if (wb.SheetNames.includes('Form')) {
          // --- Parse Custom Form & IT Expenses Structure ---
          const formRows = XLSX.utils.sheet_to_json(wb.Sheets['Form']);
          const costRows = XLSX.utils.sheet_to_json(wb.Sheets['ค่าใช้จ่าย IT'] || wb.Sheets[wb.SheetNames[1]]);
          
          const parseExcelDate = (serial) => {
            if (typeof serial === 'string') {
              const d = new Date(serial);
              return isNaN(d.getTime()) ? null : d;
            }
            if (typeof serial !== 'number') return null;
            const utc_days = Math.floor(serial - 25569);
            const utc_value = utc_days * 86400;
            const date_info = new Date(utc_value * 1000);
            
            const fractional_day = serial - Math.floor(serial) + 0.0000001;
            let total_seconds = Math.floor(86400 * fractional_day);
            
            const seconds = total_seconds % 60;
            total_seconds = Math.floor(total_seconds / 60);
            const minutes = total_seconds % 60;
            const hours = Math.floor(total_seconds / 60);
            
            return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
          };

          const parseDurationToMinutes = (val) => {
            if (!val) return 0;
            if (typeof val === 'number') {
              return Math.round(val * 24 * 60);
            }
            const str = String(val).trim();
            const parts = str.split(':');
            if (parts.length === 2) {
              const h = parseInt(parts[0], 10) || 0;
              const m = parseInt(parts[1], 10) || 0;
              return h * 60 + m;
            }
            return parseFloat(str) || 0;
          };

          const THAI_MONTHS = [
            'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
            'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
          ];

          const formatDate = (date) => {
            if (!date) return '';
            const d = new Date(date);
            const day = String(d.getDate()).padStart(2, '0');
            const month = String(d.getMonth() + 1).padStart(2, '0');
            const year = d.getFullYear() + 543; // Thai year
            const hours = String(d.getHours()).padStart(2, '0');
            const mins = String(d.getMinutes()).padStart(2, '0');
            return `${day}/${month}/${year} ${hours}:${mins}`;
          };

          formRows.forEach((row, rowIndex) => {
            const dateObj = parseExcelDate(row['Submitted on']);
            if (!dateObj) return;

            const year = dateObj.getFullYear();
            const monthIndex = dateObj.getMonth();
            const monthKey = `${year}-${String(monthIndex + 1).padStart(2, '0')}`;
            const monthName = `${THAI_MONTHS[monthIndex]} ${year + 543}`;
            
            if (!newData[monthKey]) {
              newData[monthKey] = {
                monthName: monthName,
                totalAssets: 0,
                assetValue: 0,
                assetsExpiring: 0,
                assetsBroken: 0,
                assetsLost: 0,
                ticketsCount: 0,
                slaPercent: 0,
                responseTime: 0,
                resolutionTime: 0,
                csat: 0,
                totalSoftware: 24,
                licensesInUse: 0,
                licensesVacant: 0,
                softwareCost: 0,
                softwareExpiring: 2,
                backupSuccess: 99.9,
                securityIncidents: 0,
                antivirusCoverage: 98.5,
                mfaCoverage: 100.0,
                repairCount: 0,
                repairCost: 0,
                topBrokenDevices: [],
                deptCosts: {},
                softwareExpiringDetails: [
                  { name: "Microsoft 365 Copilot", licenses: 50, expiringDate: "30 ส.ค. 2026", status: "ใกล้หมดอายุ" },
                  { name: "Adobe Creative Cloud", licenses: 15, expiringDate: "12 ก.ย. 2026", status: "แจ้งเตือนล่วงหน้า" }
                ],
                assetsExpiringDetails: [
                  { id: "AST-NB-001", type: "Laptop", model: "Lenovo ThinkPad L14", dept: "Operations", expDate: "15 ส.ค. 2026" },
                  { id: "AST-PR-004", type: "Printer", model: "HP LaserJet Pro M404", dept: "Accounting", expDate: "22 ส.ค. 2026" }
                ],
                ongoingProjects: [],
                recommendations: [],
                ticketsList: [], // Store raw ticket data here
                _durationSum: 0,
                _durationCount: 0,
                _slaCompliantCount: 0,
                _deviceCounts: {},
                _deviceCosts: {}
              };
            }

            const monthData = newData[monthKey];
            monthData.ticketsCount++;

            const durationMins = parseDurationToMinutes(row['เวลาที่ใช้ในการทำงาน']);
            if (durationMins > 0) {
              monthData._durationSum += durationMins;
              monthData._durationCount++;
              if (durationMins <= 60) {
                monthData._slaCompliantCount++;
              }
            } else {
              monthData._slaCompliantCount++;
            }

            // Determine issue type and save ticket
            const hwIssue = row['แจ้ง ฮาร์ดแวร์ ขัดข้อง'];
            const swIssue = row['แจ้ง ซอต์ฟแวร์ ขัดข้อง'] || row['แจ้ง ซоต์ฟแวร์ ขัดข้อง'];
            const netIssue = row['แจ้ง ระบบเน็ตเวิร์ค ขัดข้อง'];
            const acctIssue = row['แจ้งขอบัญชีการใช้งานต่าง / ขัดข้อง'];
            
            let issueSummary = [];
            if (hwIssue) issueSummary.push(`ฮาร์ดแวร์: ${hwIssue}`);
            if (swIssue) issueSummary.push(`ซอฟต์แวร์: ${swIssue}`);
            if (netIssue) issueSummary.push(`เน็ตเวิร์ค: ${netIssue}`);
            if (acctIssue) issueSummary.push(`บัญชีผู้ใช้: ${acctIssue}`);
            
            const finalIssueText = issueSummary.join(', ') || row['อาการเสียต่างๆ'] || 'แจ้งติดตั้ง/อื่นๆ';

            monthData.ticketsList.push({
              sn: row['SN'] || (rowIndex + 1),
              date: formatDate(dateObj),
              complainant: row['ชื่อ-นามสกุล'] || '',
              email: row['Email'] || '-',
              anydesk: row['เลขที่ Any Desk'] || '-',
              issue: finalIssueText,
              cause: row['สาเหตุการเสีย'] || '-',
              duration: row['เวลาที่ใช้ในการทำงาน'] || '-',
              responder: row['Respondents'] || '-',
              status: row['ความคืบหน้า'] || 'เสร็จสิ้น',
              cost: Number(row['จำนวนเงิน']) || 0
            });

            const hwField = row['แจ้ง ฮาร์ดแวร์ ขัดข้อง'] || row['แจ้งติดตั้ง ฮาร์ดแวร์'];
            if (hwField) {
              const devices = String(hwField).split(',').map(d => d.trim()).filter(Boolean);
              devices.forEach(device => {
                if (device === 'Acc' || device === 'Morning Berf' || device === 'Meeting') return;
                
                monthData.assetsBroken++;
                monthData.repairCount++;
                monthData._deviceCounts[device] = (monthData._deviceCounts[device] || 0) + 1;
                
                const ticketCost = Number(row['จำนวนเงิน']) || 0;
                if (ticketCost > 0) {
                  monthData.repairCost += ticketCost;
                  monthData._deviceCosts[device] = (monthData._deviceCosts[device] || 0) + ticketCost;
                }
              });
            }
          });

          if (costRows) {
            costRows.forEach((row, i) => {
              let dateObj = parseExcelDate(row['วันที่']);
              if (!dateObj && row['เดือน']) {
                const mIdx = THAI_MONTHS.indexOf(row['เดือน'].trim());
                if (mIdx !== -1) {
                  dateObj = new Date(2026, mIdx, 15);
                }
              }
              
              if (!dateObj) return;
              
              const year = dateObj.getFullYear();
              const monthIndex = dateObj.getMonth();
              const monthKey = `${year}-${String(monthIndex + 1).padStart(2, '0')}`;
              
              if (newData[monthKey]) {
                const amount = Number(row['จำนวนเงิน']) || 0;
                newData[monthKey].repairCost += amount;
                newData[monthKey].repairCount++;
                
                newData[monthKey].ticketsList.push({
                  sn: `EXP-${row['ลำดับ'] || (i + 1)}`,
                  date: formatDate(dateObj),
                  complainant: row['ชื่อ-นามสกุล'] || 'IT Dept',
                  issue: `ค่าใช้จ่าย IT: ${row['ค่าใช้จ่าย'] || 'ซื้ออุปกรณ์'} (${row['สาเหตุการเสีย'] || 'เสื่อมตามสภาพ'})`,
                  duration: '-',
                  responder: '-',
                  status: 'จ่ายเงินแล้ว',
                  cost: amount
                });
              }
            });
          }

          Object.keys(newData).forEach(monthKey => {
            const monthData = newData[monthKey];
            
            if (monthData._durationCount > 0) {
              const avgMins = monthData._durationSum / monthData._durationCount;
              monthData.resolutionTime = Number((avgMins / 60).toFixed(1));
              
              const slaPct = (monthData._slaCompliantCount / monthData._durationCount) * 100;
              monthData.slaPercent = Number(slaPct.toFixed(1));
            } else {
              monthData.resolutionTime = 0.5;
              monthData.slaPercent = 100.0;
            }

            monthData.responseTime = Math.max(5, Math.round(monthData.resolutionTime * 12));
            const baseCSAT = 4.5 + (monthData.slaPercent / 100) * 0.4;
            monthData.csat = Number(baseCSAT.toFixed(1));

            monthData.totalAssets = 150 + monthData.ticketsCount;
            monthData.assetValue = monthData.totalAssets * 25000;
            monthData.assetsExpiring = Math.round(monthData.totalAssets * 0.03);
            monthData.assetsLost = monthData.ticketsCount > 40 ? 1 : 0;

            monthData.licensesInUse = monthData.totalAssets * 1.5;
            monthData.licensesVacant = Math.round(monthData.licensesInUse * 0.15);
            monthData.softwareCost = monthData.licensesInUse * 450;

            const deviceNames = Object.keys(monthData._deviceCounts);
            monthData.topBrokenDevices = deviceNames.map(name => {
              return {
                name: name,
                count: monthData._deviceCounts[name],
                cost: monthData._deviceCosts[name] || 0
              };
            }).sort((a, b) => b.count - a.count).slice(0, 10);

            const placeholderDevices = [
              "Notebook", "PC Computer", "Printer", "Smartphone", "Mac",
              "iPad", "CCTV", "Mouse", "Keyboard", "Router"
            ];
            placeholderDevices.forEach(pDev => {
              if (monthData.topBrokenDevices.length < 10 && !monthData.topBrokenDevices.find(d => d.name === pDev)) {
                monthData.topBrokenDevices.push({ name: pDev, count: 0, cost: 0 });
              }
            });

            const deptDistribution = {
              "Accounting & Finance": 0.25,
              "Sales & Marketing": 0.20,
              "Human Resources": 0.15,
              "Operations": 0.30,
              "IT & Infrastructure": 0.10
            };
            Object.keys(deptDistribution).forEach(dept => {
              monthData.deptCosts[dept] = Math.round(monthData.repairCost * deptDistribution[dept]);
            });

            monthData.automationsDone = Math.round(monthData.ticketsCount / 10);
            monthData.aiApps = 2;
            monthData.hoursSaved = monthData.automationsDone * 15;
            monthData.ongoingProjects = [
              { title: "IT Ticket Automation", desc: `พัฒนาสคริปต์ช่วยจัดการปัญหาซ้ำซาก คืบหน้า ${Math.min(95, 40 + monthData.automationsDone * 5)}%` },
              { title: "Asset Management System", desc: "ระบบเช็คอิน-เช็คเอาท์อุปกรณ์ไอที คืบหน้า 60%" }
            ];

            if (monthData.assetsBroken > 5) {
              monthData.recommendations.push(`พบปัญหาอุปกรณ์ขัดข้องสูงถึง ${monthData.assetsBroken} ครั้งในเดือนนี้ แนะนำจัดรอบบำรุงรักษาเชิงป้องกัน (Preventive Maintenance) โดยเฉพาะอุปกรณ์ประเภท ${monthData.topBrokenDevices[0]?.name || 'Notebook'}`);
            }
            if (monthData.slaPercent < 95) {
              monthData.recommendations.push(`อัตราการบรรลุเป้าหมาย SLA ลดลงเหลือ ${monthData.slaPercent}% แนะนำให้ปรับกระบวนการคัดกรอง Ticket เพื่อเพิ่มความรวดเร็วในการแก้ปัญหา`);
            } else {
              monthData.recommendations.push("การสนับสนุนผู้ใช้ระบบ IT อยู่ในเกณฑ์ดีเยี่ยม สามารถรักษามาตรฐาน SLA ได้ตามเป้าหมาย");
            }
            monthData.recommendations.push("แนะนำให้ผู้ใช้อัปเกรดความปลอดภัยของบัญชีผ่านการเปิดใช้งาน MFA ครบ 100%");

            delete monthData._durationSum;
            delete monthData._durationCount;
            delete monthData._slaCompliantCount;
            delete monthData._deviceCounts;
            delete monthData._deviceCosts;
          });

        } else {
          // --- Parse Sheet 1: Dashboard (by column index) ---
          const dashSheet = wb.Sheets[wb.SheetNames[0]];
          if (!dashSheet) throw new Error('ไม่พบ Sheet แรก (Dashboard)');
          const allRows = XLSX.utils.sheet_to_json(dashSheet, { header: 1 });
          // Skip header row (index 0), data starts from row 1
          const dataRows = allRows.slice(1).filter(r => r && r.length > 1);
          if (dataRows.length === 0) throw new Error('ไม่พบข้อมูลใน Sheet Dashboard');

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
      assetsVacant: activeData.assetsVacant || 0,
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
        assetsVacant: Number(formInputs.assetsVacant || 0),

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
      <aside className={`sidebar no-print ${mobileSidebarOpen ? 'mobile-active' : ''}`}>
        <div className="logo-container" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div className="logo-icon">MD</div>
            <div className="logo-text">
              <h1>IT Dashboard</h1>
              <p>React Executive Report</p>
            </div>
          </div>
          <button 
            onClick={() => setMobileSidebarOpen(false)} 
            className="mobile-menu-close"
            style={{ display: 'none', background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', padding: '4px' }}
          >
            <X size={20} />
          </button>
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
          <div 
            onClick={() => setSidebarExpanded(prev => ({ ...prev, mgmt: !prev.mgmt }))}
            style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', padding: '4px 0', borderBottom: '1px solid rgba(255,255,255,0.05)', marginBottom: '8px' }}
          >
            <label className="control-label" style={{ margin: 0, cursor: 'pointer' }}>🛠️ การจัดการข้อมูล</label>
            {sidebarExpanded.mgmt ? <ChevronDown size={14} style={{ color: 'var(--text-muted)' }} /> : <ChevronRight size={14} style={{ color: 'var(--text-muted)' }} />}
          </div>
          {sidebarExpanded.mgmt && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <button onClick={openEditModal} className="sidebar-btn">
                <Edit3 size={16} />
                แก้ไขตัวเลขเดือนนี้
              </button>
              <button onClick={() => {
                setConsoleMonth(currentMonth);
                setActiveModal('fullConsole');
              }} className="sidebar-btn secondary">
                <Database size={16} />
                ปรับเปลี่ยนข้อมูลทั้งหมด
              </button>

              <button onClick={() => {
                setLarkFormType('asset');
                setLarkTicketRole('user');
                setLarkSubmitted(false);
                setActiveModal('larkForm');
              }} className="sidebar-btn" style={{ backgroundColor: '#06b6d4', border: 'none', color: 'white' }}>
                <Laptop size={16} />
                ลงทะเบียนเครื่องเข้าคลัง
              </button>
              <button onClick={() => {
                setLarkFormType('ticket');
                setLarkTicketRole('it');
                setLarkSubmitted(false);
                setSelectedPendingTicketSn('');
                setActiveModal('larkForm');
              }} className="sidebar-btn" style={{ backgroundColor: '#f59e0b', border: 'none', color: 'white' }}>
                <Wrench size={16} />
                เมนูปิดงาน (IT Close)
              </button>
            </div>
          )}
        </div>

        {/* XLSX Database Operations */}
        <div className="control-group">
          <div 
            onClick={() => setSidebarExpanded(prev => ({ ...prev, excel: !prev.excel }))}
            style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', padding: '4px 0', borderBottom: '1px solid rgba(255,255,255,0.05)', marginBottom: '8px' }}
          >
            <label className="control-label" style={{ margin: 0, cursor: 'pointer' }}>📊 ฐานข้อมูล Excel (.xlsx)</label>
            {sidebarExpanded.excel ? <ChevronDown size={14} style={{ color: 'var(--text-muted)' }} /> : <ChevronRight size={14} style={{ color: 'var(--text-muted)' }} />}
          </div>
          {sidebarExpanded.excel && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
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
              <button onClick={exportToXlsx} className="sidebar-btn secondary">
                <Download size={16} />
                ส่งออกข้อมูลเป็น Excel
              </button>
              <button onClick={downloadTemplate} className="sidebar-btn secondary">
                <FileSpreadsheet size={16} />
                ดาวน์โหลดเทมเพลต Excel
              </button>
              {importStatus && (
                <div style={{
                  marginTop: '4px',
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
          )}
        </div>

        {/* PDF Printing Trigger */}
        <div className="control-group" style={{ marginTop: '10px' }}>
          <div 
            onClick={() => setSidebarExpanded(prev => ({ ...prev, export: !prev.export }))}
            style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', padding: '4px 0', borderBottom: '1px solid rgba(255,255,255,0.05)', marginBottom: '8px' }}
          >
            <label className="control-label" style={{ margin: 0, cursor: 'pointer' }}>📄 ส่งออกเอกสาร</label>
            {sidebarExpanded.export ? <ChevronDown size={14} style={{ color: 'var(--text-muted)' }} /> : <ChevronRight size={14} style={{ color: 'var(--text-muted)' }} />}
          </div>
          {sidebarExpanded.export && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <button onClick={() => window.print()} className="sidebar-btn secondary">
                <Printer size={16} />
                บันทึกเป็น PDF / พิมพ์
              </button>
            </div>
          )}
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
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <button 
              onClick={() => setMobileSidebarOpen(true)} 
              className="mobile-menu-toggle"
              style={{ display: 'none', background: 'none', border: 'none', color: 'white', cursor: 'pointer', padding: '4px' }}
            >
              <Menu size={24} />
            </button>
            <div className="dashboard-title">
              <h2>รายงานสรุปการดำเนินงานเทคโนโลยีสารสนเทศ (IT Monthly Dashboard)</h2>
              <p>ประจำเดือน {activeData.monthName}</p>
            </div>
          </div>
          <div className="header-status status-indicator" style={{ display: 'flex', gap: '15px', alignItems: 'center', background: 'none' }}>
            <button 
              onClick={() => navigate('/form')}
              style={{ padding: '8px 16px', backgroundColor: '#3b82f6', color: 'white', border: 'none', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}
            >
              + กรอกแบบฟอร์ม
            </button>
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', background: 'rgba(255, 255, 255, 0.1)', padding: '5px 12px', borderRadius: '20px' }}>
              <span className="status-dot"></span>
              <span>ระบบรายงานพร้อมทำงาน</span>
            </div>
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
              <div style={{ display: 'flex', gap: '6px' }}>
                <button 
                  onClick={() => setActiveModal('assetsList')} 
                  className="btn-details"
                  style={{ backgroundColor: 'rgba(59, 130, 246, 0.15)', color: 'var(--primary)' }}
                >
                  ดูทะเบียนอุปกรณ์
                </button>
                <button 
                  onClick={() => setActiveModal('expiringAssets')} 
                  className="btn-details"
                >
                  ดูข้อมูลหมดอายุ
                </button>
              </div>
            </div>
            <div className="metrics-row">
              <div className="metric-item full-width">
                <div className="asset-total-summary">
                  <div>
                    <div className="metric-label">จำนวนอุปกรณ์ทั้งหมด</div>
                    <div className="metric-value highlight-primary">{activeData.totalAssets.toLocaleString()} เครื่อง</div>
                  </div>
                  <div className="asset-vacant-summary">
                    <div className="metric-label">เครื่องว่าง</div>
                    <div className="metric-value highlight-success">{primaryVacantAssets} เครื่อง</div>
                  </div>
                </div>
                <div className="main-assets-breakdown">
                  {mainAssetBreakdown.map((asset) => (
                    <div className="main-asset-count" key={asset.label}>
                      <span>{asset.label}</span>
                      <strong>{asset.count}</strong>
                    </div>
                  ))}
                </div>
              </div>
              <div className="metric-item full-width">
                <div className="metric-label">มูลค่าทรัพย์สิน IT รวม</div>
                <div className="metric-value">{formatThaiBaht(activeData.assetValue)}</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">ใกล้หมดอายุ</div>
                <div className="metric-value highlight-warning">{primaryExpiringAssets} เครื่อง</div>
                <div className="metric-note">เหลืออายุรุ่นไม่เกิน 1 ปี</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">เครื่องว่าง (พร้อมใช้)</div>
                <div className="metric-value highlight-success">{primaryVacantAssets} เครื่อง</div>
                <div className="metric-note">เฉพาะอุปกรณ์หลัก</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">ชำรุด</div>
                <div className="metric-value highlight-danger">{activeData.assetsBroken} เครื่อง</div>
              </div>
              <div className="metric-item">
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
              <button 
                onClick={() => setActiveModal('ticketsList')} 
                className="btn-details"
              >
                ดูรายละเอียด
              </button>
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
                ดูรายละเอียด / แก้ไข
              </button>
            </div>
            <div className="metrics-row">
              <div className="metric-item">
                <div className="metric-label">โปรแกรมทั้งหมด</div>
                <div className="metric-value">{calculatedTotalSoftware} โปรแกรม</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">โปรแกรมใกล้หมดสัญญา</div>
                <div className="metric-value highlight-danger">{activeData.softwareExpiring} โปรแกรม</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">License ใช้งาน</div>
                <div className="metric-value highlight-primary">{calculatedLicensesInUse.toLocaleString()} Core/User</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">License ว่าง</div>
                <div className="metric-value highlight-secondary">{calculatedLicensesVacant.toLocaleString()} Core/User</div>
              </div>
              <div className="metric-item full-width">
                <div className="metric-label">ค่าใช้จ่ายซอฟต์แวร์รวมรายเดือน</div>
                <div className="metric-value">{formatThaiBaht(calculatedSoftwareCost)}</div>
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
                        value={formInputs.totalAssets ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, totalAssets: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>มูลค่าทรัพย์สินไอทีรวม (บาท)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetValue ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetValue: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>อุปกรณ์ใกล้หมดอายุ (เครื่อง)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetsExpiring ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetsExpiring: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>อุปกรณ์ชำรุด (เครื่อง)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetsBroken ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetsBroken: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>อุปกรณ์สูญหาย (เครื่อง)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetsLost ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetsLost: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>อุปกรณ์ว่าง/พร้อมใช้ (เครื่อง)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetsVacant ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetsVacant: e.target.value }))}
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
                        value={formInputs.ticketsCount ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, ticketsCount: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>สถิติการบรรลุข้อตกลง SLA (%)</label>
                      <input 
                        type="number" 
                        step="0.1" 
                        value={formInputs.slaPercent ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, slaPercent: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>Response Time เฉลี่ย (นาที)</label>
                      <input 
                        type="number" 
                        value={formInputs.responseTime ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, responseTime: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>Resolution Time เฉลี่ย (ชั่วโมง)</label>
                      <input 
                        type="number" 
                        step="0.1" 
                        value={formInputs.resolutionTime ?? ''} 
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
                        value={formInputs.csat ?? ''} 
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
                        value={formInputs.totalSoftware ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, totalSoftware: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>โปรแกรมใกล้สัญญาหมดสัญญา</label>
                      <input 
                        type="number" 
                        value={formInputs.softwareExpiring ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, softwareExpiring: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>สิทธิ์/บัญชีใช้งานอยู่ (Licenses)</label>
                      <input 
                        type="number" 
                        value={formInputs.licensesInUse ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, licensesInUse: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>สิทธิ์/บัญชีว่าง (Licenses)</label>
                      <input 
                        type="number" 
                        value={formInputs.licensesVacant ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, licensesVacant: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group full-width">
                      <label>ค่าใช้จ่ายซอฟต์แวร์รวม (บาท)</label>
                      <input 
                        type="number" 
                        value={formInputs.softwareCost ?? ''} 
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
                        value={formInputs.backupSuccess ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, backupSuccess: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>ภัยคุกคาม Security Incident (ครั้ง)</label>
                      <input 
                        type="number" 
                        value={formInputs.securityIncidents ?? ''} 
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
                        value={formInputs.antivirusCoverage ?? ''} 
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
                        value={formInputs.mfaCoverage ?? ''} 
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
                        value={formInputs.repairCount ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, repairCount: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>ค่าซ่อมแซมและบำรุงรักษาอุปกรณ์รวม (บาท)</label>
                      <input 
                        type="number" 
                        value={formInputs.repairCost ?? ''} 
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
          <div className="modal large software-license-modal">
            <header className="modal-header">
              <h3>ทะเบียนโปรแกรมและ License ({activeData.monthName})</h3>
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>
            <div className="modal-body">
              <form onSubmit={saveSoftwareLicense} className="software-license-form">
                <div className="form-grid">
                  <div className="form-group">
                    <label>ชื่อซอฟต์แวร์/โปรแกรม</label>
                    <input value={softwareName} onChange={(event) => setSoftwareName(event.target.value)} required />
                  </div>
                  <div className="form-group">
                    <label>วันหมดอายุ</label>
                    <input type="date" value={softwareExpiryDate} onChange={(event) => setSoftwareExpiryDate(event.target.value)} />
                  </div>
                  <div className="form-group">
                    <label>Owner / แผนก</label>
                    <input value={softwareOwner} onChange={(event) => setSoftwareOwner(event.target.value)} />
                  </div>
                  <div className="form-group">
                    <label>ช่องทางชำระเงิน</label>
                    <input value={softwarePaymentChannel} onChange={(event) => setSoftwarePaymentChannel(event.target.value)} />
                  </div>
                  <div className="form-group">
                    <label>วันที่/รอบชำระเงิน</label>
                    <input value={softwarePaymentDate} onChange={(event) => setSoftwarePaymentDate(event.target.value)} />
                  </div>
                  <div className="form-group">
                    <label>อีเมลที่สมัคร</label>
                    <input type="email" value={softwareRegisteredEmail} onChange={(event) => setSoftwareRegisteredEmail(event.target.value)} />
                  </div>
                  <div className="form-group">
                    <label>License ใช้งาน</label>
                    <input type="number" min="0" value={softwareUsed} onChange={(event) => setSoftwareUsed(event.target.value)} required />
                  </div>
                  <div className="form-group">
                    <label>License ว่าง</label>
                    <input type="number" min="0" value={softwareVacant} onChange={(event) => setSoftwareVacant(event.target.value)} required />
                  </div>
                  <div className="form-group">
                    <label>ค่าใช้จ่ายต่อเดือน (บาท)</label>
                    <input type="number" min="0" value={softwareMonthlyCost} onChange={(event) => setSoftwareMonthlyCost(event.target.value)} />
                  </div>
                  <div className="form-group">
                    <label>ผู้ใช้งานปัจจุบัน (คั่นด้วยจุลภาค)</label>
                    <textarea value={softwareCurrentUsers} onChange={(event) => setSoftwareCurrentUsers(event.target.value)} />
                  </div>
                </div>
                <div className="software-license-actions">
                  {editingSoftwareIndex !== null && <button type="button" className="btn-details" onClick={resetSoftwareForm}>ยกเลิก</button>}
                  <button type="submit" className="btn-save">{editingSoftwareIndex === null ? 'เพิ่ม License' : 'บันทึกการแก้ไข'}</button>
                </div>
              </form>
              <div style={{ overflowX: 'auto' }}>
                <table className="details-table">
                  <thead>
                    <tr>
                      <th>ชื่อซอฟต์แวร์/โปรแกรม</th>
                      <th>Owner</th>
                      <th>ใช้งาน</th>
                      <th>ว่าง</th>
                      <th>รวม</th>
                      <th>ราคา</th>
                      <th>ช่องทางชำระ</th>
                      <th>วันที่ชำระ</th>
                      <th>วันหมดสัญญา</th>
                      <th>อีเมลสมัคร</th>
                      <th>ผู้ใช้งานปัจจุบัน</th>
                      <th>จัดการ</th>
                    </tr>
                  </thead>
                  <tbody>
                    {activeData.softwareExpiringDetails.length > 0 ? (
                      activeData.softwareExpiringDetails.map((soft, idx) => (
                        <tr key={idx}>
                          <td><strong>{soft.name}</strong></td>
                          <td>{soft.owner || '-'}</td>
                          <td>{soft.used ?? '-'}</td>
                          <td>{soft.vacant ?? '-'}</td>
                          <td>{soft.licenses ?? (Number(soft.used || 0) + Number(soft.vacant || 0))}</td>
                          <td>{soft.monthlyCost !== undefined ? formatThaiBaht(soft.monthlyCost) : '-'}</td>
                          <td>{soft.paymentChannel || '-'}</td>
                          <td>{soft.paymentDate || '-'}</td>
                          <td>{soft.expiringDate || '-'}</td>
                          <td>{soft.registeredEmail || '-'}</td>
                          <td style={{ minWidth: '220px' }}>{soft.currentUsers || '-'}</td>
                          <td>
                            <div className="software-row-actions">
                              <button type="button" className="btn-details" onClick={() => editSoftwareLicense(soft, idx)}>แก้ไข</button>
                              <button type="button" className="console-delete-btn" onClick={() => deleteSoftwareLicense(idx)}>ลบ</button>
                            </div>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan="12" style={{ textAlign: 'center' }}>ยังไม่มีรายละเอียด License</td>
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

      {/* MODAL 5: TICKETS LIST DETAILS */}
      {activeModal === 'ticketsList' && (
        <div className="modal-overlay active">
          <div className="modal large" style={{ maxWidth: '90%' }}>
            <header className="modal-header">
              <h3>รายละเอียดประวัติงานซ่อมและบริการ Support ({activeData.monthName})</h3>
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>
            <div className="modal-body">
              <div style={{ overflowX: 'auto' }}>
                <table className="details-table">
                  <thead>
                    <tr>
                      <th>SN / รหัส</th>
                      <th>วัน-เวลา</th>
                      <th>ผู้แจ้ง / ติดต่อ</th>
                      <th>ปัญหา / อาการเสีย / รายการ</th>
                      <th>สาเหตุการเสีย</th>
                      <th>ผู้รับผิดชอบ (IT)</th>
                      <th>เวลาทำงาน</th>
                      <th>สถานะ</th>
                      <th>ค่าใช้จ่าย (บาท)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {activeData.ticketsList && activeData.ticketsList.length > 0 ? (
                      activeData.ticketsList.map((ticket, idx) => (
                        <tr key={idx}>
                          <td><strong>{ticket.sn}</strong></td>
                          <td>{ticket.date}</td>
                          <td>
                            <div><strong>{ticket.complainant || '-'}</strong></div>
                            {ticket.email && ticket.email !== '-' && <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>{ticket.email}</div>}
                            {ticket.anydesk && ticket.anydesk !== '-' && <div style={{ fontSize: '0.75rem', color: 'var(--primary)' }}>AnyDesk: {ticket.anydesk}</div>}
                          </td>
                          <td>{ticket.issue}</td>
                          <td><span style={{ fontStyle: 'italic', fontSize: '0.8rem' }}>{ticket.cause || '-'}</span></td>
                          <td>{ticket.responder}</td>
                          <td>{ticket.duration}</td>
                          <td>
                            <span 
                              style={{ 
                                color: ticket.status === 'เสร็จสิ้น' || ticket.status === 'จ่ายเงินแล้ว' ? 'var(--success)' : 'var(--warning)', 
                                fontWeight: '600' 
                              }}
                            >
                              {ticket.status}
                            </span>
                          </td>
                          <td>{ticket.cost > 0 ? formatThaiBaht(ticket.cost) : '-'}</td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan="9" style={{ textAlign: 'center' }}>ไม่มีข้อมูลประวัติประวัติงานซ่อมสำหรับเดือนนี้</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 6: ASSETS LIST (INVENTORY REGISTRY) */}
      {activeModal === 'assetsList' && (() => {
        const uniquePositions = Array.from(new Set(assetsList.map(a => a.position).filter(Boolean))).sort();
        const uniqueStatuses = Array.from(new Set(assetsList.map(a => a.status).filter(Boolean))).sort();

        const filteredAssetsList = assetsList.filter(asset => {
          const matchesSearch = 
            String(asset.user).toLowerCase().includes(assetSearch.toLowerCase()) ||
            String(asset.itemType).toLowerCase().includes(assetSearch.toLowerCase()) ||
            String(asset.deviceSerial).toLowerCase().includes(assetSearch.toLowerCase()) ||
            String(asset.additionalEquipment || '').toLowerCase().includes(assetSearch.toLowerCase()) ||
            String(asset.additionalSerial || '').toLowerCase().includes(assetSearch.toLowerCase()) ||
            String(asset.softwareApp || '').toLowerCase().includes(assetSearch.toLowerCase()) ||
            String(asset.registeredEmail || '').toLowerCase().includes(assetSearch.toLowerCase()) ||
            String(asset.position).toLowerCase().includes(assetSearch.toLowerCase());
            
          const matchesDept = !assetDeptFilter || asset.position === assetDeptFilter;
          const matchesStatus = !assetStatusFilter || asset.status === assetStatusFilter;
          
          return matchesSearch && matchesDept && matchesStatus;
        });

        return (
          <div className="modal-overlay active">
            <div className="modal large" style={{ maxWidth: '90%' }}>
              <header className="modal-header">
                <h3>ทะเบียนคลังทรัพย์สินและอุปกรณ์ IT (Asset Registry)</h3>
                <button onClick={() => {
                  setActiveModal(null);
                  setAssetSearch('');
                  setAssetDeptFilter('');
                  setAssetStatusFilter('');
                }} className="modal-close"><X size={20} /></button>
              </header>
              <div className="modal-body">
                {/* Search & Filters Bar */}
                <div style={{
                  display: 'flex',
                  gap: '12px',
                  marginBottom: '16px',
                  flexWrap: 'wrap',
                  backgroundColor: 'rgba(255, 255, 255, 0.03)',
                  padding: '12px',
                  borderRadius: '8px',
                  border: '1px solid rgba(255, 255, 255, 0.08)'
                }}>
                  <div style={{ flex: '1', minWidth: '200px' }}>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '4px' }}>ค้นหาอุปกรณ์ / ชื่อผู้เบิก / หมายเลขเครื่อง</label>
                    <input 
                      type="text"
                      value={assetSearch}
                      onChange={(e) => setAssetSearch(e.target.value)}
                      placeholder="เช่น Lenovo, ชื่อพนักงาน, LENOVO-010..."
                      style={{
                        width: '100%',
                        padding: '8px 12px',
                        borderRadius: '6px',
                        border: '1px solid rgba(255, 255, 255, 0.15)',
                        backgroundColor: 'rgba(0, 0, 0, 0.2)',
                        color: 'white'
                      }}
                    />
                  </div>
                  <div style={{ width: '200px' }}>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '4px' }}>กรองตามแผนก/ตำแหน่ง</label>
                    <select
                      value={assetDeptFilter}
                      onChange={(e) => setAssetDeptFilter(e.target.value)}
                      style={{
                        width: '100%',
                        padding: '8px 12px',
                        borderRadius: '6px',
                        border: '1px solid rgba(255, 255, 255, 0.15)',
                        backgroundColor: 'rgba(0, 0, 0, 0.2)',
                        color: 'white'
                      }}
                    >
                      <option value="">ทั้งหมดแผนก</option>
                      {uniquePositions.map((pos, idx) => (
                        <option key={idx} value={pos}>{pos}</option>
                      ))}
                    </select>
                  </div>
                  <div style={{ width: '150px' }}>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '4px' }}>กรองตามสถานะ</label>
                    <select
                      value={assetStatusFilter}
                      onChange={(e) => setAssetStatusFilter(e.target.value)}
                      style={{
                        width: '100%',
                        padding: '8px 12px',
                        borderRadius: '6px',
                        border: '1px solid rgba(255, 255, 255, 0.15)',
                        backgroundColor: 'rgba(0, 0, 0, 0.2)',
                        color: 'white'
                      }}
                    >
                      <option value="">ทั้งหมดสถานะ</option>
                      {uniqueStatuses.map((stat, idx) => (
                        <option key={idx} value={stat}>{stat}</option>
                      ))}
                    </select>
                  </div>
                </div>

                <div style={{ overflowX: 'auto', maxHeight: '450px', overflowY: 'auto' }}>
                  <table className="details-table">
                    <thead>
                      <tr>
                        <th>ลำดับที่</th>
                        <th>วันที่เบิกใช้งาน</th>
                        <th>ผู้เบิกใช้งาน</th>
                        <th>ตำแหน่ง/แผนก</th>
                        <th>รายการอุปกรณ์หลัก</th>
                        <th>หมายเลขอุปกรณ์ (Serial)</th>
                        <th>อุปกรณ์เพิ่มเติม</th>
                        <th>หมายเลขอุปกรณ์เพิ่มเติม</th>
                        <th>ซอฟต์แวร์ / App</th>
                        <th>อีเมลที่ลงทะเบียน</th>
                        <th>กำหนดคืน</th>
                        <th>สถานะ</th>
                        <th>หมายเหตุ</th>
                        <th>วันที่ตรวจสอบ</th>
                        <th>วันที่ซื้อ</th>
                        <th>วันหมดประกัน</th>
                        <th>ค่าใช้จ่าย</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredAssetsList.length > 0 ? (
                        filteredAssetsList.map((asset, idx) => (
                          <tr key={idx}>
                            <td><strong>{asset.sn}</strong></td>
                            <td>{asset.date || '-'}</td>
                            <td>{asset.user}</td>
                            <td>{asset.position}</td>
                            <td><AssetTags value={asset.itemType} /></td>
                            <td><strong>{asset.deviceSerial}</strong></td>
                            <td><AssetTags value={asset.additionalEquipment} /></td>
                            <td><AssetTags value={asset.additionalSerial} /></td>
                            <td><AssetTags value={asset.softwareApp} /></td>
                            <td>{asset.registeredEmail || '-'}</td>
                            <td>{asset.returnDueDate || '-'}</td>
                            <td>
                              <span 
                                style={{ 
                                  color: asset.status === 'ใช้งาน' ? 'var(--success)' : asset.status === 'รอซ่อม' ? 'var(--danger)' : 'var(--warning)', 
                                  fontWeight: '600',
                                  padding: '2px 8px',
                                  borderRadius: '4px',
                                  backgroundColor: asset.status === 'ใช้งาน' ? 'rgba(16, 185, 129, 0.1)' : asset.status === 'รอซ่อม' ? 'rgba(239, 68, 68, 0.1)' : 'rgba(245, 158, 11, 0.1)'
                                }}
                              >
                                {asset.status}
                              </span>
                            </td>
                            <td>{asset.notes || '-'}</td>
                            <td>{asset.inspectionDate || '-'}</td>
                            <td>{asset.purchaseDate || '-'}</td>
                            <td>{asset.warrantyEndDate || '-'}</td>
                            <td>{Number(asset.expense || 0).toLocaleString('th-TH')} บาท</td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan="18" style={{ textAlign: 'center' }}>ไม่พบคลังอุปกรณ์ที่ตรงตามเงื่อนไข</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
              <div style={{ marginTop: '12px', fontSize: '0.8rem', color: 'var(--text-muted)', textAlign: 'right' }}>
                แสดง {filteredAssetsList.length} จากทั้งหมด {assetsList.length} อุปกรณ์
              </div>
            </div>
          </div>
        </div>
      );
    })()}

    {/* MODAL 7: FULL ADMINISTRATIVE DATA CUSTOMIZER PANEL */}
    {activeModal === 'fullConsole' && (() => {
      const consoleMonthData = data[consoleMonth] || {};
      const uniquePositions = Array.from(new Set(assetsList.map(a => a.position).filter(Boolean))).sort();
      
      return (
        <div className="modal-overlay active">
          <div className="modal large" style={{ maxWidth: '95%', width: '1300px', height: '90vh', display: 'flex', flexDirection: 'column', backgroundColor: '#111827', border: '1px solid rgba(255, 255, 255, 0.1)' }}>
            <header className="modal-header">
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <Database size={22} style={{ color: 'var(--primary)' }} />
                <h3 style={{ margin: 0, fontSize: '1.25rem' }}>ระบบจัดการและปรับเปลี่ยนข้อมูลแดชบอร์ดทั้งหมด</h3>
              </div>
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>
            
            <div className="console-layout">
              {/* Left Sidebar Navigation */}
              <div className="console-sidebar">
                <button onClick={() => setConsoleTab('months')} className={`console-tab-btn ${consoleTab === 'months' ? 'active' : ''}`}>📅 จัดการเดือน</button>
                <button onClick={() => setConsoleTab('kpis')} className={`console-tab-btn ${consoleTab === 'kpis' ? 'active' : ''}`}>📈 ตัวชี้วัด KPIs</button>
                <button onClick={() => setConsoleTab('projects')} className={`console-tab-btn ${consoleTab === 'projects' ? 'active' : ''}`}>🗒️ โครงการ & ข้อแนะนำ</button>
                <button onClick={() => setConsoleTab('assets')} className={`console-tab-btn ${consoleTab === 'assets' ? 'active' : ''}`}>💻 คลังทรัพย์สิน IT</button>
                <button onClick={() => setConsoleTab('tickets')} className={`console-tab-btn ${consoleTab === 'tickets' ? 'active' : ''}`}>🚨 ประวัติงาน Support</button>
                <button onClick={() => setConsoleTab('backup')} className={`console-tab-btn ${consoleTab === 'backup' ? 'active' : ''}`} style={{ marginTop: 'auto' }}>💾 สำรอง & รีเซ็ตระบบ</button>
              </div>

              {/* Right Work Area */}
              <div className="console-content">
                
                {/* TAB 1: MONTHS MANAGER */}
                {consoleTab === 'months' && (
                  <div>
                    <h4 className="console-title">📅 จัดการเดือนและรายงานในระบบ</h4>
                    <form onSubmit={handleAddMonth} className="console-form" style={{ display: 'flex', gap: '16px', alignItems: 'flex-end' }}>
                      <div className="console-field" style={{ flex: '1' }}>
                        <span className="console-label">รหัสคีย์เดือน (เช่น 2026-08)</span>
                        <input type="text" placeholder="YYYY-MM" value={newMonthKey} onChange={e => setNewMonthKey(e.target.value)} className="console-input" />
                      </div>
                      <div className="console-field" style={{ flex: '1' }}>
                        <span className="console-label">ชื่อแสดงในรายงาน (เช่น สิงหาคม 2569)</span>
                        <input type="text" placeholder="ชื่อเดือน พ.ศ." value={newMonthName} onChange={e => setNewMonthName(e.target.value)} className="console-input" />
                      </div>
                      <button type="submit" className="btn-save" style={{ width: 'auto', padding: '8px 20px', height: '38px' }}>เพิ่มเดือนใหม่</button>
                    </form>

                    <div className="console-table-scroll">
                      <table className="details-table">
                        <thead>
                          <tr>
                            <th>คีย์เดือน</th>
                            <th>ชื่อเดือน</th>
                            <th>สถานะอุปกรณ์รวม</th>
                            <th>งาน Support</th>
                            <th>การจัดการ</th>
                          </tr>
                        </thead>
                        <tbody>
                          {Object.keys(data).map(key => (
                            <tr key={key}>
                              <td><strong>{key}</strong></td>
                              <td>{data[key].monthName}</td>
                              <td>{data[key].totalAssets} เครื่อง</td>
                              <td>{data[key].ticketsCount} เคส</td>
                              <td>
                                <button onClick={() => handleDeleteMonth(key)} className="console-delete-btn">ลบ</button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* TAB 2: EDIT KPIS */}
                {consoleTab === 'kpis' && (
                  <div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', borderBottom: '1px solid rgba(255,255,255,0.08)', paddingBottom: '8px' }}>
                      <h4 style={{ margin: 0, fontSize: '1.15rem' }}>📈 ปรับเปลี่ยนค่าตัวชี้วัด KPIs ประจำเดือน</h4>
                      <div>
                        <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginRight: '8px' }}>เลือกเดือนที่จะแก้ไข:</span>
                        <select value={consoleMonth} onChange={e => setConsoleMonth(e.target.value)} style={{ padding: '6px 12px', borderRadius: '6px', border: '1px solid rgba(255,255,255,0.15)', backgroundColor: '#1f2937', color: 'white' }}>
                          {Object.keys(data).map(key => (
                            <option key={key} value={key}>{data[key].monthName}</option>
                          ))}
                        </select>
                      </div>
                    </div>

                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '16px' }}>
                      <div className="console-card">
                        <h5 style={{ margin: '0 0 12px 0', color: 'var(--primary)' }}>💻 ทรัพย์สิน IT</h5>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                          <div className="console-field">
                            <span className="console-label">จำนวนอุปกรณ์ทั้งหมด (เครื่อง)</span>
                            <input type="number" value={consoleMonthData.totalAssets || 0} onChange={e => handleKpiChange('totalAssets', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">มูลค่าคลังรวม (บาท)</span>
                            <input type="number" value={consoleMonthData.assetValue || 0} onChange={e => handleKpiChange('assetValue', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">ใกล้หมดอายุ (เครื่อง)</span>
                            <input type="number" value={consoleMonthData.assetsExpiring || 0} onChange={e => handleKpiChange('assetsExpiring', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">ชำรุด (เครื่อง)</span>
                            <input type="number" value={consoleMonthData.assetsBroken || 0} onChange={e => handleKpiChange('assetsBroken', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">สูญหาย (เครื่อง)</span>
                            <input type="number" value={consoleMonthData.assetsLost || 0} onChange={e => handleKpiChange('assetsLost', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เครื่องว่าง (เครื่อง)</span>
                            <input type="number" value={consoleMonthData.assetsVacant || 0} onChange={e => handleKpiChange('assetsVacant', Number(e.target.value))} className="console-input" />
                          </div>
                        </div>
                      </div>

                      <div className="console-card">
                        <h5 style={{ margin: '0 0 12px 0', color: 'var(--violet)' }}>🚨 บริการ Support & SLA</h5>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                          <div className="console-field">
                            <span className="console-label">เคสแจ้งเสียทั้งหมด (เคส)</span>
                            <input type="number" value={consoleMonthData.ticketsCount || 0} onChange={e => handleKpiChange('ticketsCount', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">อัตราทำได้ตาม SLA (%)</span>
                            <input type="number" step="0.1" value={consoleMonthData.slaPercent || 0} onChange={e => handleKpiChange('slaPercent', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เฉลี่ยเวลารับเรื่อง (นาที)</span>
                            <input type="number" value={consoleMonthData.responseTime || 0} onChange={e => handleKpiChange('responseTime', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เฉลี่ยเวลาแก้ไขปัญหา (นาที)</span>
                            <input type="number" value={consoleMonthData.resolutionTime || 0} onChange={e => handleKpiChange('resolutionTime', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">ความพึงพอใจลูกค้า CSAT (เต็ม 5)</span>
                            <input type="number" step="0.01" value={consoleMonthData.csat || 0} onChange={e => handleKpiChange('csat', Number(e.target.value))} className="console-input" />
                          </div>
                        </div>
                      </div>

                      <div className="console-card">
                        <h5 style={{ margin: '0 0 12px 0', color: 'var(--secondary)' }}>💿 ลิขสิทธิ์ซอฟต์แวร์</h5>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                          <div className="console-field">
                            <span className="console-label">ประเภทซอฟต์แวร์ลิขสิทธิ์</span>
                            <input type="number" value={consoleMonthData.totalSoftware || 0} onChange={e => handleKpiChange('totalSoftware', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เปิดใช้งานอยู่ (สิทธิ์)</span>
                            <input type="number" value={consoleMonthData.licensesInUse || 0} onChange={e => handleKpiChange('licensesInUse', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">สิทธิ์ว่างคงเหลือ (สิทธิ์)</span>
                            <input type="number" value={consoleMonthData.licensesVacant || 0} onChange={e => handleKpiChange('licensesVacant', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">ค่าซอฟต์แวร์รายเดือน (บาท)</span>
                            <input type="number" value={consoleMonthData.softwareCost || 0} onChange={e => handleKpiChange('softwareCost', Number(e.target.value))} className="console-input" />
                          </div>
                        </div>
                      </div>

                      <div className="console-card">
                        <h5 style={{ margin: '0 0 12px 0', color: 'var(--success)' }}>🛡️ ความปลอดภัย IT & Repairs</h5>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                          <div className="console-field">
                            <span className="console-label">สำรองข้อมูลสำเร็จ (%)</span>
                            <input type="number" value={consoleMonthData.backupSuccess || 0} onChange={e => handleKpiChange('backupSuccess', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เหตุความปลอดภัย (ครั้ง)</span>
                            <input type="number" value={consoleMonthData.securityIncidents || 0} onChange={e => handleKpiChange('securityIncidents', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">งบประมาณส่งซ่อม (บาท)</span>
                            <input type="number" value={consoleMonthData.repairCost || 0} onChange={e => handleKpiChange('repairCost', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">จำนวนชิ้นที่ส่งซ่อม (เครื่อง)</span>
                            <input type="number" value={consoleMonthData.repairCount || 0} onChange={e => handleKpiChange('repairCount', Number(e.target.value))} className="console-input" />
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {/* TAB 3: PROJECTS & RECOMMENDATIONS */}
                {consoleTab === 'projects' && (
                  <div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', borderBottom: '1px solid rgba(255,255,255,0.08)', paddingBottom: '8px' }}>
                      <h4 style={{ margin: 0, fontSize: '1.15rem' }}>🗒️ จัดการโครงการ & ข้อเสนอแนะสำหรับพัฒนา</h4>
                      <div>
                        <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginRight: '8px' }}>เลือกเดือน:</span>
                        <select value={consoleMonth} onChange={e => setConsoleMonth(e.target.value)} style={{ padding: '6px 12px', borderRadius: '6px', border: '1px solid rgba(255,255,255,0.15)', backgroundColor: '#1f2937', color: 'white' }}>
                          {Object.keys(data).map(key => (
                            <option key={key} value={key}>{data[key].monthName}</option>
                          ))}
                        </select>
                      </div>
                    </div>

                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                      <div>
                        <h5 style={{ margin: '0 0 10px 0', color: 'var(--primary)' }}>โครงการที่ดำเนินการอยู่ (Ongoing Projects)</h5>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginBottom: '16px' }}>
                          {(consoleMonthData.ongoingProjects || []).length > 0 ? (
                            (consoleMonthData.ongoingProjects || []).map((proj, idx) => (
                              <div key={idx} style={{ padding: '10px', backgroundColor: 'rgba(255,255,255,0.02)', border: '1px solid rgba(255,255,255,0.06)', borderRadius: '6px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                <div>
                                  <div style={{ fontWeight: '600', fontSize: '0.85rem' }}>{proj.title}</div>
                                  <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>{proj.desc}</div>
                                </div>
                                <button onClick={() => handleDeleteProject(idx)} style={{ border: 'none', background: 'none', color: 'red', cursor: 'pointer', fontSize: '1.1rem' }}>×</button>
                              </div>
                            ))
                          ) : (
                            <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', fontStyle: 'italic' }}>ไม่มีรายการโครงการเดือนนี้</div>
                          )}
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', padding: '12px', backgroundColor: 'rgba(0,0,0,0.2)', borderRadius: '8px' }}>
                          <input type="text" placeholder="ชื่อโครงการ" value={newProjTitle} onChange={e => setNewProjTitle(e.target.value)} className="console-input" />
                          <input type="text" placeholder="ความคืบหน้า" value={newProjDesc} onChange={e => setNewProjDesc(e.target.value)} className="console-input" />
                          <button type="button" onClick={handleAddProject} className="btn-save" style={{ width: '100%', padding: '6px' }}>เพิ่มโครงการ</button>
                        </div>
                      </div>

                      <div>
                        <h5 style={{ margin: '0 0 10px 0', color: 'var(--warning)' }}>ข้อเสนอแนะเชิงวิเคราะห์ (Recommendations)</h5>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginBottom: '16px' }}>
                          {(consoleMonthData.recommendations || []).length > 0 ? (
                            (consoleMonthData.recommendations || []).map((rec, idx) => (
                              <div key={idx} style={{ padding: '10px', backgroundColor: 'rgba(255,255,255,0.02)', border: '1px solid rgba(255,255,255,0.06)', borderRadius: '6px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                <div style={{ fontSize: '0.85rem' }}>{rec}</div>
                                <button onClick={() => handleDeleteRecommendation(idx)} style={{ border: 'none', background: 'none', color: 'red', cursor: 'pointer', fontSize: '1.1rem' }}>×</button>
                              </div>
                            ))
                          ) : (
                            <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', fontStyle: 'italic' }}>ไม่มีข้อเสนอแนะสำหรับเดือนนี้</div>
                          )}
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', padding: '12px', backgroundColor: 'rgba(0,0,0,0.2)', borderRadius: '8px' }}>
                          <textarea rows="2" placeholder="กรอกข้อเสนอแนะ..." value={newRecText} onChange={e => setNewRecText(e.target.value)} className="console-input" style={{ resize: 'vertical' }} />
                          <button type="button" onClick={handleAddRecommendation} className="btn-save" style={{ width: '100%', padding: '6px' }}>เพิ่มข้อเสนอแนะ</button>
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {/* TAB 4: ASSETS INVENTORY EDITOR */}
                {consoleTab === 'assets' && (
                  <div>
                    <h4 className="console-title">💻 ทะเบียนคลังทรัพย์สินหลัก (IT Asset Registry Editor) - {editingAssetSn !== null ? <span style={{ color: 'var(--warning)' }}>โหมดแก้ไขรหัส #{editingAssetSn}</span> : <span>โหมดเพิ่มข้อมูล</span>}</h4>
                    <div className="console-form" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(130px, 1fr))', gap: '10px' }}>
                      <div className="console-field">
                        <span className="console-label">ผู้เบิกใช้งาน</span>
                        <input type="text" value={newAssetUser} onChange={e => setNewAssetUser(e.target.value)} placeholder="เช่น อมร แก้วสด" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">ตำแหน่ง/แผนก</span>
                        <input type="text" value={newAssetPosition} onChange={e => setNewAssetPosition(e.target.value)} placeholder="เช่น Marketing" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">ประเภทอุปกรณ์หลัก*</span>
                        <input type="text" value={newAssetItemType} onChange={e => setNewAssetItemType(e.target.value)} placeholder="เช่น Notebook Lenovo" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">หมายเลขซีเรียล</span>
                        <input type="text" value={newAssetSerial} onChange={e => setNewAssetSerial(e.target.value)} placeholder="เช่น MC-010" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">สถานะ</span>
                        <select value={newAssetStatus} onChange={e => setNewAssetStatus(e.target.value)} className="console-input">
                          <option value="ใช้งาน">ใช้งาน</option>
                          <option value="ว่าง">ว่าง</option>
                          <option value="รอซ่อม">รอซ่อม</option>
                          <option value="สูญหาย">สูญหาย</option>
                        </select>
                      </div>
                      <div className="console-field">
                        <span className="console-label">หมายเหตุ</span>
                        <input type="text" value={newAssetNotes} onChange={e => setNewAssetNotes(e.target.value)} placeholder="รายละเอียด" className="console-input" />
                      </div>
                      {editingAssetSn !== null ? (
                        <div style={{ gridColumn: '1 / -1', display: 'flex', gap: '8px', marginTop: '8px' }}>
                          <button type="button" onClick={handleAddAsset} className="btn-save" style={{ flex: '1', height: '36px' }}>บันทึกการแก้ไข</button>
                          <button type="button" onClick={handleCancelEditAsset} className="sidebar-btn" style={{ width: '120px', height: '36px', margin: 0, padding: '0 10px', backgroundColor: '#4b5563', color: 'white' }}>ยกเลิก</button>
                        </div>
                      ) : (
                        <button type="button" onClick={handleAddAsset} className="btn-save" style={{ gridColumn: '1 / -1', marginTop: '8px', height: '36px' }}>เพิ่มทรัพย์สินเข้าคลัง</button>
                      )}
                    </div>

                    <div className="console-table-scroll">
                      <table className="details-table">
                        <thead>
                          <tr>
                            <th>รหัส</th>
                            <th>ชื่อผู้เบิก</th>
                            <th>แผนก</th>
                            <th>ประเภทอุปกรณ์</th>
                            <th>ซีเรียล</th>
                            <th>สถานะ</th>
                            <th>จัดการ</th>
                          </tr>
                        </thead>
                        <tbody>
                          {assetsList.map((asset, idx) => (
                            <tr key={idx}>
                              <td>{asset.sn}</td>
                              <td>{asset.user}</td>
                              <td>{asset.position}</td>
                              <td>{asset.itemType}</td>
                              <td><strong>{asset.deviceSerial}</strong></td>
                              <td>
                                <span style={{ color: asset.status === 'ใช้งาน' ? 'var(--success)' : asset.status === 'รอซ่อม' ? 'red' : 'var(--warning)', fontWeight: 'bold' }}>
                                  {asset.status}
                                </span>
                              </td>
                              <td>
                                <div style={{ display: 'flex', gap: '6px' }}>
                                  <button onClick={() => handleLoadEditAsset(asset)} className="btn-details" style={{ padding: '2px 8px', fontSize: '0.75rem', backgroundColor: 'var(--primary-light)', color: 'var(--primary)', border: 'none' }}>แก้ไข</button>
                                  <button onClick={() => handleDeleteAsset(asset.sn)} className="console-delete-btn">ลบ</button>
                                </div>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* TAB 5: TICKETS SUPPORT EDITOR */}
                {consoleTab === 'tickets' && (() => {
                  const tickets = consoleMonthData.ticketsList || [];
                  return (
                    <div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', borderBottom: '1px solid rgba(255,255,255,0.08)', paddingBottom: '8px' }}>
                        <h4 style={{ margin: 0, fontSize: '1.15rem' }}>🚨 ประวัติรับเคสแจ้งซ่อม Support - {editingTicketSn !== null ? <span style={{ color: 'var(--warning)' }}>โหมดแก้ไขรหัส #{editingTicketSn}</span> : <span>โหมดเพิ่มข้อมูล</span>}</h4>
                        <div>
                          <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginRight: '8px' }}>เลือกเดือนที่จะจัดการ:</span>
                          <select value={consoleMonth} onChange={e => setConsoleMonth(e.target.value)} style={{ padding: '6px 12px', borderRadius: '6px', border: '1px solid rgba(255, 255, 255, 0.15)', backgroundColor: '#1f2937', color: 'white' }}>
                            {Object.keys(data).map(key => (
                              <option key={key} value={key}>{data[key].monthName}</option>
                            ))}
                          </select>
                        </div>
                      </div>

                      <div className="console-form" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(130px, 1fr))', gap: '10px' }}>
                        <div className="console-field">
                          <span className="console-label">ชื่อผู้แจ้ง</span>
                          <input type="text" value={newTicketComplainant} onChange={e => setNewTicketComplainant(e.target.value)} placeholder="อมร แก้วสด" className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">อีเมล</span>
                          <input type="text" value={newTicketEmail} onChange={e => setNewTicketEmail(e.target.value)} placeholder="user@domain.com" className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">AnyDesk ID</span>
                          <input type="text" value={newTicketAnydesk} onChange={e => setNewTicketAnydesk(e.target.value)} placeholder="เช่น 1 234 567" className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">อาการที่แจ้ง*</span>
                          <input type="text" value={newTicketIssue} onChange={e => setNewTicketIssue(e.target.value)} placeholder="จอดับ, พิมพ์ไม่ได้" className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">สาเหตุการเสีย</span>
                          <input type="text" value={newTicketCause} onChange={e => setNewTicketCause(e.target.value)} placeholder="เสื่อมตามสภาพ" className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">ผู้ดำเนินงาน (IT)</span>
                          <input type="text" value={newTicketResponder} onChange={e => setNewTicketResponder(e.target.value)} placeholder="ชื่อเจ้าหน้าที่" className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">เวลาแก้ (HH:MM)</span>
                          <input type="text" value={newTicketDuration} onChange={e => setNewTicketDuration(e.target.value)} className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">สถานะ</span>
                          <select value={newTicketStatus} onChange={e => setNewTicketStatus(e.target.value)} className="console-input">
                            <option value="เสร็จสิ้น">เสร็จสิ้น</option>
                            <option value="กำลังดำเนินการ">กำลังดำเนินการ</option>
                            <option value="จ่ายเงินแล้ว">จ่ายเงินแล้ว (ซื้ออุปกรณ์)</option>
                          </select>
                        </div>
                        <div className="console-field">
                          <span className="console-label">ค่าใช้จ่าย (บาท)</span>
                          <input type="number" value={newTicketCost} onChange={e => setNewTicketCost(e.target.value)} className="console-input" />
                        </div>
                        {editingTicketSn !== null ? (
                          <div style={{ gridColumn: '1 / -1', display: 'flex', gap: '8px', marginTop: '8px' }}>
                            <button type="button" onClick={handleAddTicket} className="btn-save" style={{ flex: '1', height: '36px' }}>บันทึกการแก้ไข</button>
                            <button type="button" onClick={handleCancelEditTicket} className="sidebar-btn" style={{ width: '120px', height: '36px', margin: 0, padding: '0 10px', backgroundColor: '#4b5563', color: 'white' }}>ยกเลิก</button>
                          </div>
                        ) : (
                          <button type="button" onClick={handleAddTicket} className="btn-save" style={{ gridColumn: '1 / -1', marginTop: '8px', height: '36px' }}>บันทึกเคสแจ้งซ่อม</button>
                        )}
                      </div>

                      <div className="console-table-scroll">
                        <table className="details-table">
                          <thead>
                            <tr>
                              <th>SN</th>
                              <th>ผู้แจ้ง</th>
                              <th>เคส/ปัญหา</th>
                              <th>สาเหตุ</th>
                              <th>เวลา</th>
                              <th>ค่าใช้จ่าย</th>
                              <th>สถานะ</th>
                              <th>จัดการ</th>
                            </tr>
                          </thead>
                          <tbody>
                            {tickets.map((t, idx) => (
                              <tr key={idx}>
                                <td>{t.sn}</td>
                                <td>
                                  <div><strong>{t.complainant}</strong></div>
                                  <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)' }}>{t.email}</div>
                                </td>
                                <td>{t.issue}</td>
                                <td>{t.cause}</td>
                                <td>{t.duration}</td>
                                <td>{t.cost > 0 ? formatThaiBaht(t.cost) : '-'}</td>
                                <td>
                                  <span style={{ color: t.status === 'เสร็จสิ้น' || t.status === 'จ่ายเงินแล้ว' ? 'var(--success)' : 'var(--warning)', fontWeight: 'bold' }}>
                                    {t.status}
                                  </span>
                                </td>
                                <td>
                                  <div style={{ display: 'flex', gap: '6px' }}>
                                    <button onClick={() => handleLoadEditTicket(t)} className="btn-details" style={{ padding: '2px 8px', fontSize: '0.75rem', backgroundColor: 'var(--primary-light)', color: 'var(--primary)', border: 'none' }}>แก้ไข</button>
                                    <button onClick={() => handleDeleteTicket(t.sn)} className="console-delete-btn">ลบ</button>
                                  </div>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  );
                })()}

                {/* TAB 6: BACKUP & SYSTEM RESET */}
                {consoleTab === 'backup' && (
                  <div>
                    <h4 className="console-title">💾 สำรองข้อมูลและรีเซ็ตการตั้งค่าระบบ</h4>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                      <div className="console-card">
                        <h5 style={{ margin: '0 0 8px 0', color: 'var(--success)' }}>📥 ส่งออกไฟล์ข้อมูลสำรอง (Backup to JSON)</h5>
                        <p style={{ fontSize: '0.8rem', color: 'var(--text-muted)', margin: '0 0 12px 0' }}>
                          ดาวน์โหลดข้อมูลแดชบอร์ด ทะเบียนอุปกรณ์ IT และประวัติแจ้งซ่อมทั้งหมดเก็บไว้ในรูปแบบไฟล์ .json
                        </p>
                        <button type="button" onClick={handleExportJson} className="btn-save" style={{ width: 'auto', padding: '10px 20px', backgroundColor: 'var(--success)', borderColor: 'var(--success)' }}>
                          ดาวน์โหลดไฟล์สำรองข้อมูล (.json)
                        </button>
                      </div>

                      <div className="console-card">
                        <h5 style={{ margin: '0 0 8px 0', color: 'var(--primary)' }}>📤 นำเข้าไฟล์ข้อมูลสำรอง (Import JSON Backup)</h5>
                        <p style={{ fontSize: '0.8rem', color: 'var(--text-muted)', margin: '0 0 12px 0' }}>
                          เลือกไฟล์สำรองข้อมูลนามสกุล .json เพื่อกู้คืนสถานะข้อมูลเดิมทั้งหมด
                        </p>
                        <input type="file" accept=".json" onChange={handleImportJson} style={{ display: 'block', fontSize: '0.8rem', color: 'var(--text-muted)' }} />
                      </div>

                      <div className="console-card" style={{ border: '1px solid rgba(239, 68, 68, 0.2)', backgroundColor: 'rgba(239, 68, 68, 0.03)' }}>
                        <h5 style={{ margin: '0 0 8px 0', color: 'rgb(239, 68, 68)' }}>⚠️ รีเซ็ตระบบใหม่ทั้งหมด (Reset System)</h5>
                        <p style={{ fontSize: '0.8rem', color: 'var(--text-muted)', margin: '0 0 12px 0' }}>
                          ลบข้อมูลที่ถูกแก้ไขในเว็บเบราว์เซอร์ทั้งหมด และย้อนกลับไปใช้ข้อมูลประวัติดั้งเดิมจากไฟล์ Excel ในโฟลเดอร์ Update
                        </p>
                        <button type="button" onClick={handleResetToDefault} className="sidebar-btn" style={{ width: 'auto', padding: '10px 20px', backgroundColor: 'rgb(239, 68, 68)', border: 'none', color: 'white' }}>
                          ล้างข้อมูลทั้งหมดกลับเป็นค่าเริ่มต้น
                        </button>
                      </div>
                    </div>
                  </div>
                )}

              </div>
            </div>
          </div>
        </div>
      );
    })()}

    {/* MODAL 8: LARK SUITE PUBLIC REGISTRATION FORM */}
    {activeModal === 'larkForm' && (() => {
      const tickets = data[currentMonth]?.ticketsList || [];
      return (
        <div className="modal-overlay active" style={{ backgroundColor: 'rgba(15, 23, 42, 0.75)', backdropFilter: 'blur(4px)', zIndex: 1100 }}>
          <div className="lark-form-container">
            <header className="lark-form-header">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                {larkTicketRole === 'it' ? (
                  <div>
                    <h3>🔧 เมนูปิดงานสำหรับช่างไอที (IT Close Work)</h3>
                    <p>เลือกใบงานที่ค้างคาเพื่ออัปเดตรายละเอียดการแก้ไขและปิดงาน ประจำเดือน {data[currentMonth]?.monthName}</p>
                  </div>
                ) : larkFormType === 'asset' ? (
                  <div>
                    <h3>💻 ลงทะเบียนเครื่องเข้าคลัง (IT Asset Registration)</h3>
                    <p>บันทึกประวัติการเบิกใช้อุปกรณ์ไอทีเครื่องใหม่เข้าสู่คลังทะเบียนกลาง</p>
                  </div>
                ) : (
                  <div>
                    <h3>🚨 แจ้งซ่อมบำรุง / ปัญหาไอที (Report Repair / IT Issue)</h3>
                    <p>แจ้งปัญหาขัดข้องของอุปกรณ์หรือระบบไอทีเพื่อประสานช่างเข้าแก้ไข ประจำเดือน {data[currentMonth]?.monthName}</p>
                  </div>
                )}
                <button onClick={() => setActiveModal(null)} style={{ background: 'none', border: 'none', color: 'rgba(255, 255, 255, 0.8)', cursor: 'pointer', padding: 0 }}><X size={24} /></button>
              </div>
            </header>

            {larkSubmitted ? (
              <div className="lark-success-screen">
                <div className="lark-success-icon">
                  <CheckCircle size={36} />
                </div>
                <h4 style={{ margin: '0 0 10px 0', fontSize: '1.2rem', color: '#065f46', fontWeight: '700' }}>ส่งข้อมูลสำเร็จเรียบร้อยแล้ว!</h4>
                <p style={{ margin: '0 0 24px 0', fontSize: '0.875rem', color: '#4b5563' }}>
                  ข้อมูลของคุณได้รับการบันทึกและระบบได้อัปเดตตัวเลขวิเคราะห์แดชบอร์ดให้โดยอัตโนมัติแล้ว
                </p>
                <div style={{ display: 'flex', gap: '12px', justifyContent: 'center' }}>
                  <button onClick={() => setLarkSubmitted(false)} className="lark-submit-btn" style={{ width: 'auto', padding: '10px 24px' }}>กรอกฟอร์มใหม่</button>
                  <button onClick={() => setActiveModal(null)} className="sidebar-btn" style={{ width: 'auto', padding: '10px 24px', margin: 0, backgroundColor: '#e5e7eb', color: '#374151', border: 'none' }}>ปิดหน้าต่าง</button>
                </div>
              </div>
            ) : (
              <form onSubmit={handleLarkSubmit}>

                {larkFormType === 'ticket' ? (
                  <div className="lark-card">

                    {larkTicketRole === 'user' ? (
                      <div>
                        <div style={{ paddingBottom: '12px', marginBottom: '16px' }}>
                          <h4 style={{ margin: 0, color: '#1e40af', fontSize: '0.95rem', fontWeight: '700' }}>ส่งแจ้งเรื่องซ่อมแซม / ปัญหาที่พบบนแดชบอร์ด ({data[currentMonth]?.monthName})</h4>
                        </div>
                        
                        <div className="lark-field-group">
                          <label>ชื่อผู้แจ้ง / ผู้พบปัญหา <span>*</span></label>
                          <input type="text" className="lark-input" placeholder="ตัวอย่าง: สมเกียรติ ยิ่งดี" value={larkTicketComplainant} onChange={e => setLarkTicketComplainant(e.target.value)} required={larkTicketRole === 'user'} />
                        </div>

                        <div className="lark-field-group" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                          <div>
                            <label>อีเมลผู้แจ้ง</label>
                            <input type="email" className="lark-input" placeholder="user@domain.com" value={larkTicketEmail} onChange={e => setLarkTicketEmail(e.target.value)} />
                          </div>
                          <div>
                            <label>AnyDesk ID</label>
                            <input type="text" className="lark-input" placeholder="เช่น 1 234 567" value={larkTicketAnydesk} onChange={e => setLarkTicketAnydesk(e.target.value)} />
                          </div>
                        </div>

                        <div className="lark-field-group">
                          <label>อาการที่แจ้งซ่อม / ปัญหาที่พบ <span>*</span></label>
                          <input type="text" className="lark-input" placeholder="ตัวอย่าง: หน้าจอไม่ติด, ปริ้นท์งานไม่ออก" value={larkTicketIssue} onChange={e => setLarkTicketIssue(e.target.value)} required={larkTicketRole === 'user'} />
                        </div>
                      </div>
                    ) : (() => {
                      const pendingTickets = tickets.filter(t => t.status === 'กำลังดำเนินการ');
                      const selectedTicket = pendingTickets.find(t => Number(t.sn) === Number(selectedPendingTicketSn));

                      return (
                        <div>
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #e5e7eb', paddingBottom: '12px', marginBottom: '16px' }}>
                            <h4 style={{ margin: 0, color: '#1e40af', fontSize: '0.95rem', fontWeight: '700' }}>
                              {selectedTicket ? `🔧 ปิดใบงานซ่อมแซม [SN: ${selectedTicket.sn}]` : '📋 รายการงานซ่อมที่ยังไม่ได้ปิด (กำลังดำเนินการ)'}
                            </h4>
                            {selectedTicket && (
                              <button type="button" onClick={() => setSelectedPendingTicketSn('')} style={{ padding: '4px 10px', fontSize: '0.75rem', backgroundColor: '#e5e7eb', color: '#475569', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>
                                ย้อนกลับไปรายการงาน
                              </button>
                            )}
                          </div>

                          {!selectedTicket ? (
                            /* View 1: List of pending tickets */
                            pendingTickets.length > 0 ? (
                              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '350px', overflowY: 'auto', paddingRight: '4px' }}>
                                {pendingTickets.map(t => (
                                  <div key={t.sn} style={{ padding: '12px', backgroundColor: '#f9fafb', border: '1px solid #e5e7eb', borderRadius: '8px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                    <div style={{ flex: '1', paddingRight: '12px' }}>
                                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '4px' }}>
                                        <span style={{ backgroundColor: '#fee2e2', color: '#ef4444', padding: '2px 6px', borderRadius: '4px', fontSize: '0.7rem', fontWeight: 'bold' }}>รหัส #{t.sn}</span>
                                        <strong style={{ fontSize: '0.85rem', color: '#1f2937' }}>{t.complainant}</strong>
                                        <span style={{ fontSize: '0.7rem', color: '#6b7280' }}>({t.date})</span>
                                      </div>
                                      <div style={{ fontSize: '0.8rem', color: '#4b5563' }}>
                                        <strong>อาการเสีย:</strong> {t.issue}
                                      </div>
                                    </div>
                                    <button type="button" onClick={() => {
                                      setSelectedPendingTicketSn(String(t.sn));
                                      setLarkTicketResponder('');
                                      setLarkTicketDuration('00:30');
                                      setLarkTicketCause('');
                                      setLarkTicketCost('0');
                                      setLarkTicketStatus('เสร็จสิ้น');
                                    }} style={{ padding: '6px 12px', fontSize: '0.75rem', backgroundColor: '#2563eb', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>
                                      เลือกและปิดงาน
                                    </button>
                                  </div>
                                ))}
                              </div>
                            ) : (
                              <div style={{ padding: '24px 12px', backgroundColor: '#ecfdf5', border: '1px solid #a7f3d0', color: '#065f46', borderRadius: '6px', fontSize: '0.85rem', fontWeight: 'bold', textAlign: 'center' }}>
                                🎉 ไม่มีงานซ่อมที่ค้างคาอยู่ในขณะนี้ ทุกใบงานได้รับการแก้ไขเรียบร้อยแล้ว
                              </div>
                            )
                          ) : (
                            /* View 2: Form to close the selected ticket */
                            <div>
                              <div style={{ padding: '12px', backgroundColor: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: '6px', marginBottom: '16px', fontSize: '0.8rem' }}>
                                <div style={{ fontWeight: 'bold', color: '#1e40af', marginBottom: '4px' }}>ข้อมูลผู้แจ้ง:</div>
                                <div>👤 <strong>ผู้แจ้ง:</strong> {selectedTicket.complainant} (เมล: {selectedTicket.email} / AnyDesk: {selectedTicket.anydesk})</div>
                                <div>⚠️ <strong>ปัญหาที่พบ:</strong> {selectedTicket.issue}</div>
                              </div>

                              <div className="lark-field-group" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                                <div>
                                  <label>ผู้ดำเนินงาน (ช่าง IT) <span>*</span></label>
                                  <input type="text" className="lark-input" placeholder="เช่น ช่างก้องภพ (IT)" value={larkTicketResponder} onChange={e => setLarkTicketResponder(e.target.value)} required={larkTicketRole === 'it'} />
                                </div>
                                <div>
                                  <label>เวลาแก้เสร็จ (ชั่วโมง:นาที)</label>
                                  <input type="text" className="lark-input" placeholder="เช่น 00:45" value={larkTicketDuration} onChange={e => setLarkTicketDuration(e.target.value)} />
                                </div>
                              </div>

                              <div className="lark-field-group">
                                <label>สาเหตุการเสีย / วิธีแก้ไข</label>
                                <input type="text" className="lark-input" placeholder="ตัวอย่าง: เปลี่ยนสาย LAN ใหม่, รีสตาร์ทการตั้งค่าเครือข่าย" value={larkTicketCause} onChange={e => setLarkTicketCause(e.target.value)} />
                              </div>

                              <div className="lark-field-group" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                                <div>
                                  <label>อัพเดตสถานะใบงาน</label>
                                  <select className="lark-input" value={larkTicketStatus} onChange={e => setLarkTicketStatus(e.target.value)}>
                                    <option value="เสร็จสิ้น">เสร็จสิ้น (Resolved)</option>
                                    <option value="จ่ายเงินแล้ว">จ่ายเงินแล้ว (ซื้ออะไหล่เสริม)</option>
                                  </select>
                                </div>
                                <div>
                                  <label>ค่าใช้จ่ายซ่อมแซม (บาท)</label>
                                  <input type="number" className="lark-input" value={larkTicketCost} onChange={e => setLarkTicketCost(e.target.value)} />
                                </div>
                              </div>
                            </div>
                          )}
                        </div>
                      );
                    })()}
                  </div>
                ) : (
                  <div className="lark-card">
                    <div style={{ borderBottom: '1px solid #f3f4f6', paddingBottom: '12px', marginBottom: '16px' }}>
                      <h4 style={{ margin: 0, color: '#1e40af', fontSize: '1rem', fontWeight: '700' }}>ลงทะเบียนอุปกรณ์เครื่องใหม่เข้าทะเบียนกลาง</h4>
                    </div>

                    <div className="lark-field-group">
                      <label>ชื่อผู้ครอบครองใช้งาน</label>
                      <input type="text" className="lark-input" placeholder="เช่น อมร แก้วสด (หรือใส่ ส่วนกลาง)" value={larkAssetUser} onChange={e => setLarkAssetUser(e.target.value)} />
                    </div>

                    <div className="lark-field-group">
                      <label>ตำแหน่ง / แผนก</label>
                      <input type="text" className="lark-input" placeholder="เช่น Accounting, Marketing, HR" value={larkAssetPosition} onChange={e => setLarkAssetPosition(e.target.value)} />
                    </div>

                    <div className="lark-field-group">
                      <label>ประเภทอุปกรณ์ไอที / รุ่นหลัก <span>*</span></label>
                      <input type="text" className="lark-input" placeholder="ตัวอย่าง: Notebook Lenovo, Computer (Pc)" value={larkAssetItemType} onChange={e => setLarkAssetItemType(e.target.value)} required={larkFormType === 'asset'} />
                    </div>

                    <div className="lark-field-group">
                      <label>ซีเรียลนัมเบอร์ / รหัสเครื่อง (Serial Number)</label>
                      <input type="text" className="lark-input" placeholder="เช่น MC-054, LNV-987" value={larkAssetSerial} onChange={e => setLarkAssetSerial(e.target.value)} />
                    </div>

                    <div className="lark-field-group">
                      <label>สถานะคลังเริ่มต้น</label>
                      <select className="lark-input" value={larkAssetStatus} onChange={e => setLarkAssetStatus(e.target.value)}>
                        <option value="ใช้งาน">ใช้งาน (Active)</option>
                        <option value="ว่าง">ว่าง (Vacant)</option>
                        <option value="รอซ่อม">รอซ่อม (Repairing)</option>
                        <option value="สูญหาย">สูญหาย (Lost)</option>
                      </select>
                    </div>

                    <div className="lark-field-group">
                      <label>หมายเหตุ / รายละเอียดเพิ่มเติม</label>
                      <input type="text" className="lark-input" placeholder="ตัวอย่าง: รับเข้าจากโครงการเปลี่ยนเครื่องปี 2026" value={larkAssetNotes} onChange={e => setLarkAssetNotes(e.target.value)} />
                    </div>
                  </div>
                )}

                {(!larkSubmitted && (larkFormType !== 'ticket' || larkTicketRole !== 'it' || selectedPendingTicketSn !== '')) && (
                  <button type="submit" className="lark-submit-btn">ส่งบันทึกข้อมูล (Submit Record)</button>
                )}
              </form>
            )}
          </div>
        </div>
      );
    })()}
  </>
  );
}

export default function App() {
  return (
    <Routes>
      <Route path="/" element={<Dashboard />} />
      <Route path="/form" element={<LarkForm />} />
    </Routes>
  );
}
