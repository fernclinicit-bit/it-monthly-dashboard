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
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:10",
        "responder": "ปัณณวิชญ์ สิริภานุพัฒน์ (บอย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 2,
        "date": "11/11/2568 14:04",
        "complainant": "นิธิดา รัตนอาภรณ์",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: ์Adobe",
        "duration": "00:30",
        "responder": "นิธิดา รัตนอาภรณ์ (เตย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 3,
        "date": "18/11/2568 15:49",
        "complainant": "พรพิมล เขียวจันทร์",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows",
        "duration": "00:50",
        "responder": "ผู้ใช้รับเชิญ 95166",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 4,
        "date": "19/11/2568 10:09",
        "complainant": "กฤษณา ลำเพ็ง",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: Printer WIFI",
        "duration": "00:50",
        "responder": "กฤษณา ลำเพ็ง (พลอย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 5,
        "date": "19/11/2568 10:35",
        "complainant": "นายกฤติน  วิชันดิษฐ ",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: Printer WIFI",
        "duration": "00:30",
        "responder": "กฤติน วิชัยดิษฐ  (อ้น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 6,
        "date": "20/11/2568 11:01",
        "complainant": "เนตรปรีญา ทัดศรี",
        "issue": "ฮาร์ดแวร์: Notebook, เน็ตเวิร์ค: เชื่อมต่อ บลูทูธ ไม่ได้",
        "duration": "00:30",
        "responder": "เนตรปรีญา ทัดศรี (แตงกวา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 7,
        "date": "26/11/2568 15:58",
        "complainant": "สุภาพ  แสนจันทร์ ",
        "issue": "ฮาร์ดแวร์: Notebook, Printer",
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
        "issue": "ฮาร์ดแวร์: Notebook, Printer, บัญชีผู้ใช้: Printer WIFI",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 77838",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 9,
        "date": "02/12/2568 13:21",
        "complainant": "โชตินันท์ ณ นคร",
        "issue": "ฮาร์ดแวร์: Notebook, Printer, บัญชีผู้ใช้: Printer WIFI",
        "duration": "00:40",
        "responder": "ผู้ใช้รับเชิญ 24619",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 10,
        "date": "04/12/2568 13:05",
        "complainant": "โชตินันท์ ณ นคร",
        "issue": "ฮาร์ดแวร์: PC Computer, บัญชีผู้ใช้: เชื่อมต่อ Nas ",
        "duration": "00:50",
        "responder": "ผู้ใช้รับเชิญ 81829",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 11,
        "date": "12/12/2568 10:25",
        "complainant": "นางสาวอัจฉรา เหรียญพิมาย",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: Printer WIFI",
        "duration": "00:50",
        "responder": "ผู้ใช้รับเชิญ 94307",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 12,
        "date": "12/12/2568 12:43",
        "complainant": "นภัสสร นาสวน",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: Printer WIFI",
        "duration": "00:40",
        "responder": "ผู้ใช้รับเชิญ 47307",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 13,
        "date": "15/12/2568 16:14",
        "complainant": "โชตินันท์ ณ นคร",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows",
        "duration": "00:40",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 14,
        "date": "19/12/2568 10:17",
        "complainant": "พรพิมล เขียวจันทร์",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: ์Adobe, Windows",
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
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: Email",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 31440",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 16,
        "date": "05/01/2569 11:06",
        "complainant": "นางสาวสุพรรษา อินทะเรืองรุ่ง",
        "issue": "ฮาร์ดแวร์: Printer, ซอฟต์แวร์: Windows, บัญชีผู้ใช้: Printer WIFI",
        "duration": "00:30",
        "responder": "สุพรรษา อินทะเรืองรุ่ง (ต่าย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 17,
        "date": "06/01/2569 09:52",
        "complainant": "นางสาวอังคณา ธงศรี",
        "issue": "ฮาร์ดแวร์: Notebook, บัญชีผู้ใช้: Email",
        "duration": "00:10",
        "responder": "Ampol",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 18,
        "date": "07/01/2569 10:50",
        "complainant": "ขวัญลออ นวลละออง",
        "issue": "ฮาร์ดแวร์: Notebook",
        "duration": "00:10",
        "responder": "Kwanlaoa Nuanlaong",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 19,
        "date": "08/01/2569 09:38",
        "complainant": "ชลธิชา ตาลพันธ์ นิชา",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows, Microsoft Office",
        "duration": "01:00",
        "responder": "ผู้ใช้รับเชิญ 65327",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 20,
        "date": "08/01/2569 10:03",
        "complainant": "ดลพร อุลุชาฎะ",
        "issue": "ฮาร์ดแวร์: Mac, ซอฟต์แวร์: Microsoft Office",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 48018",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 21,
        "date": "08/01/2569 10:42",
        "complainant": "อนุสรา สิมจันทา",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: IOS, บัญชีผู้ใช้: Tiktok",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 24821",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 22,
        "date": "08/01/2569 13:55",
        "complainant": "ปัณณวิชญ์   ทองวัน (บอย)",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "duration": "00:30",
        "responder": "ปัณณวิชญ์ สิริภานุพัฒน์ (บอย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 23,
        "date": "08/01/2569 14:52",
        "complainant": "กรรณิกา ค่ำคูณ",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: IOS, บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 35630",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 24,
        "date": "08/01/2569 14:54",
        "complainant": "ฐานิสา ศรีจันทร์โคตร",
        "issue": "ฮาร์ดแวร์: iPad",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 63311",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 25,
        "date": "12/01/2569 09:04",
        "complainant": "พิชยา ฮงทอง",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Google Link",
        "duration": "00:30",
        "responder": "พิชยา ฮงทอง (แนน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 26,
        "date": "12/01/2569 09:51",
        "complainant": "อังคณา ธงศรี (ใบเฟิร์น/บัญชี)",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Microsoft Office",
        "duration": "00:20",
        "responder": "Ampol",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 27,
        "date": "12/01/2569 10:44",
        "complainant": "ปาหนัน สุพรม",
        "issue": "ฮาร์ดแวร์: Notebook",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 55719",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 29,
        "date": "15/01/2569 15:01",
        "complainant": "ขวัญลออ นวลละออง",
        "issue": "ฮาร์ดแวร์: Notebook",
        "duration": "00:30",
        "responder": "Kwanlaoa Nuanlaong",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 30,
        "date": "16/01/2569 09:34",
        "complainant": "อัจฉรา เหรียญพิมาย (โบกี้)",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows",
        "duration": "00:30",
        "responder": "อัจฉรา เหรียญพิมาย (Bogie)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 31,
        "date": "16/01/2569 12:13",
        "complainant": "อนุสรา สิมจันทา",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: IOS",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 62685",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 32,
        "date": "19/01/2569 11:16",
        "complainant": "กฤติญา ทาระพันธ์",
        "issue": "ฮาร์ดแวร์: Notebook, Smartphone, ซอฟต์แวร์: Google Link, IOS, บัญชีผู้ใช้: Tiktok, Email",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 76985",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 33,
        "date": "20/01/2569 16:02",
        "complainant": "ณัฐณิชา ศรีวรอรรถิกุล",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: เชื่อมต่อ บลูทูธ ไม่ได้, บัญชีผู้ใช้: Printer WIFI",
        "duration": "00:40",
        "responder": "ผู้ใช้รับเชิญ 73645",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 34,
        "date": "21/01/2569 09:41",
        "complainant": "เอกรินทร์ จีนเพชร",
        "issue": "ฮาร์ดแวร์: PC Computer",
        "duration": "00:20",
        "responder": "Guest User 75941",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 35,
        "date": "21/01/2569 16:13",
        "complainant": "อนุสรา สิมจันทา",
        "issue": "ฮาร์ดแวร์: iPad",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 13173",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 36,
        "date": "22/01/2569 12:15",
        "complainant": "สุภาพ แสนจันทร์ ",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: Printer WIFI, Email",
        "duration": "00:50",
        "responder": "สุภาพ แสนจันทร์ (ส้ม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 40,
        "date": "22/01/2569 15:31",
        "complainant": "เฟิร์น",
        "issue": "ฮาร์ดแวร์: iPad, บัญชีผู้ใช้: Tiktok",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 42,
        "date": "22/01/2569 15:34",
        "complainant": "ต่าย HR",
        "issue": "ฮาร์ดแวร์: CCTV, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 43,
        "date": "22/01/2569 15:50",
        "complainant": "ครีม",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 44,
        "date": "22/01/2569 15:51",
        "complainant": "พี่ส้ม",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows, บัญชีผู้ใช้: Printer WIFI",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 46,
        "date": "23/01/2569 09:31",
        "complainant": "อำพล   แซ่แฮ",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 48,
        "date": "23/01/2569 09:38",
        "complainant": "ปภาวิน อักโขสุวรรณ",
        "issue": "ฮาร์ดแวร์: iPad, บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 16976",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 49,
        "date": "23/01/2569 10:25",
        "complainant": "นนทภัทร์  พึ่งพุ่ม",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: IOS, บัญชีผู้ใช้: Tiktok",
        "duration": "00:35",
        "responder": "Guest User 89423",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 50,
        "date": "23/01/2569 13:12",
        "complainant": "ขวัญลออ นวลละออง",
        "issue": "ฮาร์ดแวร์: Notebook",
        "duration": "00:50",
        "responder": "Kwanlaoa Nuanlaong",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 51,
        "date": "23/01/2569 13:19",
        "complainant": "กฤติญา ทาระพันธ์",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 76985",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 52,
        "date": "23/01/2569 13:31",
        "complainant": "ณัฏฐ์ชาวีร์ หิรัญรัชชากุล",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: Calenda",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 88787",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 53,
        "date": "23/01/2569 13:59",
        "complainant": "อำพล   แซ่แฮ",
        "issue": "ฮาร์ดแวร์: Mac, ซอฟต์แวร์: IOS, Windows",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 54,
        "date": "26/01/2569 08:31",
        "complainant": "อำพล  แซ่แฮ",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 55,
        "date": "26/01/2569 08:45",
        "complainant": "ภัทรศยา ไชยคุณ",
        "issue": "ฮาร์ดแวร์: iPad, บัญชีผู้ใช้: Email",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 94075",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 57,
        "date": "26/01/2569 16:06",
        "complainant": "สุดธิดา เผ่าหอม",
        "issue": "ฮาร์ดแวร์: Notebook, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "duration": "01:00",
        "responder": "สุดธิดา เผ่าหอม (ต่าย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 58,
        "date": "26/01/2569 16:56",
        "complainant": "พี่บี จัดซื้อ",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 59,
        "date": "27/01/2569 08:36",
        "complainant": "อำพล  แซ่แฮ",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 29120",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 60,
        "date": "27/01/2569 09:05",
        "complainant": "เบนซ์",
        "issue": "ฮาร์ดแวร์: Mac, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 61,
        "date": "27/01/2569 10:05",
        "complainant": "ทัศวรรณ วัลย์ดาว ",
        "issue": "ซอฟต์แวร์: CCTV",
        "duration": "03:00",
        "responder": "ผู้ใช้รับเชิญ 78271",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 62,
        "date": "27/01/2569 13:50",
        "complainant": "กฤติมา สอนพู",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 78411",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 63,
        "date": "27/01/2569 13:55",
        "complainant": "ภัทรศยา ไชยคุณ",
        "issue": "ฮาร์ดแวร์: Smartphone, บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 94075",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 64,
        "date": "27/01/2569 14:38",
        "complainant": "บอม",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 65,
        "date": "27/01/2569 14:48",
        "complainant": "พิชยา ฮงทอง (แนน)",
        "issue": "ฮาร์ดแวร์: Notebook, Printer, ซอฟต์แวร์: Microsoft Office, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้",
        "duration": "00:45",
        "responder": "พิชยา ฮงทอง (แนน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 66,
        "date": "28/01/2569 08:33",
        "complainant": "อำพล  แซ่แฮ",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 68,
        "date": "28/01/2569 08:38",
        "complainant": "ทีมไลฟ์",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: เชื่อมต่อ Nas ไม่ได้, บัญชีผู้ใช้: เชื่อมต่อ Nas , G-Suit",
        "duration": "03:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 69,
        "date": "28/01/2569 09:34",
        "complainant": "HR",
        "issue": "ซอฟต์แวร์: CCTV",
        "duration": "00:40",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 70,
        "date": "29/01/2569 14:14",
        "complainant": "Ing Admin",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 71,
        "date": "29/01/2569 14:16",
        "complainant": "Prem",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: IOS, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 72,
        "date": "29/01/2569 16:40",
        "complainant": "อภิสิทธิ์ พรจันทราวัฒน์",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: G-Suit",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 43809",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 73,
        "date": "29/01/2569 16:48",
        "complainant": "IT",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: google dive",
        "duration": "05:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 74,
        "date": "30/01/2569 12:11",
        "complainant": "ฐานิสา ศรีจีนทร์โคตร",
        "issue": "บัญชีผู้ใช้: google dive",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 86927",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 75,
        "date": "30/01/2569 13:14",
        "complainant": "กฤติน วิชัยดิษฐ  ",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: google dive",
        "duration": "00:20",
        "responder": "กฤติน วิชัยดิษฐ  (อ้น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 76,
        "date": "30/01/2569 13:32",
        "complainant": "IT",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: google dive",
        "duration": "03:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 77,
        "date": "30/01/2569 13:35",
        "complainant": "พี่พอส",
        "issue": "ฮาร์ดแวร์: Mac, บัญชีผู้ใช้: Click Up",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 79,
        "date": "31/01/2569 08:26",
        "complainant": "นนทภัทร์ พึ่งพุ่ม (หมู)",
        "issue": "ฮาร์ดแวร์: iPad, บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "Guest User 75073",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 80,
        "date": "31/01/2569 08:32",
        "complainant": "เอกรินทร์ จีนเพชร(เอก)",
        "issue": "เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "duration": "02:00",
        "responder": "ผู้ใช้รับเชิญ 90720",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 81,
        "date": "31/01/2569 09:29",
        "complainant": "Bam",
        "issue": "ฮาร์ดแวร์: Smartphone, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: Calenda, G-Suit, บัญชีผู้ใช้: google dive",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 82,
        "date": "31/01/2569 09:40",
        "complainant": "ฐิตารีย์ นรกุลศิริภักดี",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: เชื่อมต่อ Server ",
        "duration": "00:30",
        "responder": "ฐิตารีย์  นรกุลศิริภักดี (เฟิร์น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 83,
        "date": "31/01/2569 09:42",
        "complainant": "ต่าย กราฟิก",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: G-Suit, บัญชีผู้ใช้: G-Suit",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 84,
        "date": "31/01/2569 09:47",
        "complainant": "ทีม DATA",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 85,
        "date": "31/01/2569 09:52",
        "complainant": "พรพิมล เขียวจันทร์",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows",
        "duration": "03:00",
        "responder": "ผู้ใช้รับเชิญ 14598",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 86,
        "date": "31/01/2569 12:25",
        "complainant": "ชลธิชา ตาลพันธ์",
        "issue": "ฮาร์ดแวร์: Notebook, iPad, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้, Calenda, เชื่อมต่อ บลูทูธ ไม่ได้, บัญชีผู้ใช้: Email, google dive, เชื่อมต่อ Server ",
        "duration": "02:00",
        "responder": "ชลธิชา ตาลพันธ์ (นิชา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 87,
        "date": "31/01/2569 13:36",
        "complainant": "นนทภัทร์ พึ่งพุ่มา(หมู) TikTok content creators ",
        "issue": "บัญชีผู้ใช้: Capcut",
        "duration": "00:30",
        "responder": "Guest User 75513",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 88,
        "date": "31/01/2569 16:00",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
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
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "01:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 90,
        "date": "03/02/2569 09:24",
        "complainant": "",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: เชื่อมต่อ Nas ไม่ได้, บัญชีผู้ใช้: เชื่อมต่อ Server ",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 91,
        "date": "03/02/2569 09:28",
        "complainant": "ปาหนัน สุพรม",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "03:00",
        "responder": "ปาหนัน สุพรม (พู่กัน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 92,
        "date": "03/02/2569 12:29",
        "complainant": "สุภาพ  แสนจันทร์ (พี่ส่ม)ี",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: Printer WIFI",
        "duration": "01:00",
        "responder": "ผู้ใช้รับเชิญ 81286",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 93,
        "date": "04/02/2569 07:52",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 94,
        "date": "04/02/2569 09:10",
        "complainant": "สุภาพ แสนจันทร์ ",
        "issue": "ฮาร์ดแวร์: CCTV, ซอฟต์แวร์: CCTV, บัญชีผู้ใช้: CCTV",
        "duration": "01:30",
        "responder": "ผู้ใช้รับเชิญ 81286",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 95,
        "date": "04/02/2569 10:44",
        "complainant": "พิชยา ฮงทอง",
        "issue": "ฮาร์ดแวร์: Notebook, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "duration": "01:00",
        "responder": "Guest User 16356",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 96,
        "date": "04/02/2569 10:45",
        "complainant": "พิชยา ฮงทอง",
        "issue": "ฮาร์ดแวร์: Printer, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้, บัญชีผู้ใช้: Printer WIFI",
        "duration": "00:30",
        "responder": "Guest User 16356",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 98,
        "date": "05/02/2569 09:51",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 99,
        "date": "05/02/2569 09:52",
        "complainant": "",
        "issue": "ฮาร์ดแวร์: CCTV, ซอฟต์แวร์: CCTV",
        "duration": "48:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 101,
        "date": "05/02/2569 09:54",
        "complainant": "",
        "issue": "ฮาร์ดแวร์: CCTV, ซอฟต์แวร์: CCTV",
        "duration": "48:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 102,
        "date": "05/02/2569 15:11",
        "complainant": "ชลธิชา ตาลพันธ์",
        "issue": "ฮาร์ดแวร์: PC Computer, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: Calenda, G-Suit, บัญชีผู้ใช้: Email",
        "duration": "00:30",
        "responder": "ชลธิชา ตาลพันธ์ (นิชา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 103,
        "date": "06/02/2569 10:56",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 105,
        "date": "06/02/2569 14:01",
        "complainant": "สุรวิชญ์ โพธิ์ตาก",
        "issue": "ฮาร์ดแวร์: Mac, ซอฟต์แวร์: Windows, เน็ตเวิร์ค: เชื่อมต่อ Nas ไม่ได้, บัญชีผู้ใช้: เชื่อมต่อ Nas ",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 44149",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 106,
        "date": "09/02/2569 08:42",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 107,
        "date": "10/02/2569 09:06",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 108,
        "date": "10/02/2569 09:46",
        "complainant": "อัจฉรา เหรียญพิมาย (โบกี้)",
        "issue": "ฮาร์ดแวร์: Notebook",
        "duration": "00:10",
        "responder": "อัจฉรา เหรียญพิมาย (Bogie)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 109,
        "date": "10/02/2569 10:12",
        "complainant": "พิชยา ฮงทอง",
        "issue": "ฮาร์ดแวร์: Notebook, เน็ตเวิร์ค: Calenda",
        "duration": "00:30",
        "responder": "Guest User 22874",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 110,
        "date": "13/02/2569 10:07",
        "complainant": "นภัสสร นาสวน",
        "issue": "ฮาร์ดแวร์: Printer",
        "duration": "00:30",
        "responder": "Guest User 75198",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 111,
        "date": "16/02/2569 09:22",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 113,
        "date": "16/02/2569 10:58",
        "complainant": "",
        "issue": "ฮาร์ดแวร์: Printer, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้, บัญชีผู้ใช้: Printer WIFI",
        "duration": "04:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 114,
        "date": "17/02/2569 08:40",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 115,
        "date": "17/02/2569 08:56",
        "complainant": "กฤษณา ลำเพ็ง",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows",
        "duration": "00:30",
        "responder": "Guest User 67573",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 116,
        "date": "17/02/2569 09:13",
        "complainant": "อนุสรา สิมจันทา",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: IOS",
        "duration": "00:10",
        "responder": "ผู้ใช้รับเชิญ 52165",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 117,
        "date": "17/02/2569 13:57",
        "complainant": "กฤติมา สอนพูน",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Microsoft Office, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้, บัญชีผู้ใช้: google dive",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 14230",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 118,
        "date": "18/02/2569 08:28",
        "complainant": "",
        "issue": "โปรดให้เนื้อหาต้นฉบับที่ต้องการสรุปมาเพื่อให้ฉันสามารถสรุปเป็นภาษาไทยให้คุณได้ครับ",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 119,
        "date": "23/02/2569 09:38",
        "complainant": "",
        "issue": "โปรดให้เนื้อหาหลักที่ต้องการสรุปมา ฉันจะสรุปให้เป็นภาษาไทยตามที่ขอค่ะ",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 120,
        "date": "23/02/2569 09:39",
        "complainant": "",
        "issue": "ฮาร์ดแวร์: CCTV, ซอฟต์แวร์: CCTV",
        "duration": "72:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 121,
        "date": "23/02/2569 11:00",
        "complainant": "นางสาวสุพรรษา อินทะเรืองรุ่ง",
        "issue": "บัญชีผู้ใช้: เชื่อมต่อ Server , Email",
        "duration": "00:30",
        "responder": "สุพรรษา อินทะเรืองรุ่ง (ต่าย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 123,
        "date": "24/02/2569 08:47",
        "complainant": "",
        "issue": "กรุณาให้เนื้อหาหลัก (正文) ที่ต้องการสรุปเป็นภาษาไทยมาให้ฉันนะครับ/คะ ฉันจะสรุปเนื้อหานั้นให้ถูกต้องและกระชับตามที่คุณต้องการโดยทันที",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 132,
        "date": "25/02/2569 15:39",
        "complainant": "นิชา",
        "issue": "ฮาร์ดแวร์: Notebook",
        "duration": "00:20",
        "responder": "ผู้ใช้รับเชิญ 53591",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 136,
        "date": "26/02/2569 09:00",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 140,
        "date": "26/02/2569 11:12",
        "complainant": "ปาหนัน สุพรม พู่กัน",
        "issue": "ฮาร์ดแวร์: Mouse",
        "duration": "00:20",
        "responder": "ปาหนัน สุพรม (พู่กัน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 143,
        "date": "27/02/2569 09:04",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 144,
        "date": "28/02/2569 08:47",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
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
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 146,
        "date": "04/03/2569 08:30",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 147,
        "date": "10/03/2569 08:36",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 148,
        "date": "10/03/2569 09:51",
        "complainant": "",
        "issue": "เทรนนิ่ง Lark",
        "duration": "01:00",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 150,
        "date": "11/03/2569 08:28",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 151,
        "date": "12/03/2569 08:20",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 167,
        "date": "18/03/2569 09:06",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 169,
        "date": "18/03/2569 09:06",
        "complainant": "",
        "issue": "ฮาร์ดแวร์: Notebook",
        "duration": "00:20",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 171,
        "date": "19/03/2569 10:09",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 172,
        "date": "19/03/2569 10:09",
        "complainant": "",
        "issue": "ฮาร์ดแวร์: Notebook",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 173,
        "date": "19/03/2569 10:10",
        "complainant": "",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows",
        "duration": "00:30",
        "responder": "Base Assistant",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 175,
        "date": "19/03/2569 16:46",
        "complainant": "",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: CCTV",
        "duration": "00:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 177,
        "date": "19/03/2569 16:48",
        "complainant": "",
        "issue": "แจ้งติดตั้ง/อื่นๆ",
        "duration": "00:30",
        "responder": "Base Assistant",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 178,
        "date": "20/03/2569 12:54",
        "complainant": "พิชยา ฮงทอง (แนน)",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Microsoft Office",
        "duration": "00:20",
        "responder": "พิชยา ฮงทอง (แนน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 181,
        "date": "30/03/2569 08:47",
        "complainant": "ต่าย HR",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Microsoft Office",
        "duration": "0030",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 184,
        "date": "31/03/2569 13:06",
        "complainant": "ช่างภาพ ไมค์",
        "issue": "ฮาร์ดแวร์: Mac, ซอฟต์แวร์: IOS",
        "duration": "00:40",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 186,
        "date": "31/03/2569 15:42",
        "complainant": "กิ๊กจัดซื้อ",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Windows",
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
        "issue": "ติดตั้งเครื่องปริ้นบ้าน 18 ในโน้ตบุ้ค",
        "duration": "00:20",
        "responder": "กฤษณา ลำเพ็ง (พลอย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 189,
        "date": "07/04/2569 08:31",
        "complainant": "สุภาพ  แสนจันทร์ ",
        "issue": "ฮาร์ดแวร์: Printer, บัญชีผู้ใช้: เชื่อมต่อ Server , Printer WIFI",
        "duration": "02:00",
        "responder": "สุภาพ แสนจันทร์ (ส้ม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 190,
        "date": "07/04/2569 12:08",
        "complainant": "ชลธิชา สุวาส",
        "issue": "ฮาร์ดแวร์: Notebook, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้, บัญชีผู้ใช้: Email",
        "duration": "00:30",
        "responder": "ชลธิชา สุวาส (จิน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 191,
        "date": "09/04/2569 09:12",
        "complainant": "เอกรินทร์ จีนเพชร",
        "issue": "บัญชีผู้ใช้: Kumoo, Email",
        "duration": "02:00",
        "responder": "ผู้ใช้รับเชิญ 55420",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 192,
        "date": "22/04/2569 10:56",
        "complainant": "เนตรปรีญา ทัดศรี",
        "issue": "ฮาร์ดแวร์: Printer",
        "duration": "03:00",
        "responder": "เนตรปรีญา ทัดศรี(แตงกวา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 193,
        "date": "23/04/2569 15:21",
        "complainant": "ชลธิชา สุวาส",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Microsoft Office, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้",
        "duration": "01:00",
        "responder": "ผู้ใช้รับเชิญ 98623",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 194,
        "date": "24/04/2569 13:22",
        "complainant": "เนตรปรีญา ทัดศรี",
        "issue": "ฮาร์ดแวร์: Mac, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้",
        "duration": "01:00",
        "responder": "เนตรปรีญา ทัดศรี(แตงกวา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 195,
        "date": "30/04/2569 08:25",
        "complainant": "",
        "issue": "ซอฟต์แวร์: Config System, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
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
        "issue": "ฮาร์ดแวร์: Printer, เน็ตเวิร์ค: เชื่อมต่ออินเตอร์เน็ตไม่ได้",
        "duration": "01:00",
        "responder": "เนตรปรีญา ทัดศรี(แตงกวา)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 197,
        "date": "18/05/2569 14:04",
        "complainant": "กฤติญา ทาระพันธ์",
        "issue": "ฮาร์ดแวร์: iPad, Notebook, ซอฟต์แวร์: Google Link, เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: Email",
        "duration": "00:30",
        "responder": "ผู้ใช้รับเชิญ 41070",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 198,
        "date": "19/05/2569 13:30",
        "complainant": "ปาหนัน สุพรม",
        "issue": "ฮาร์ดแวร์: Mouse",
        "duration": "00:20",
        "responder": "ปาหนัน สุพรม (พู่กัน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 200,
        "date": "20/05/2569 13:22",
        "complainant": "ธนัชชา บุญมีมาก",
        "issue": "ฮาร์ดแวร์: Smartphone, บัญชีผู้ใช้: Email",
        "duration": "00:10",
        "responder": "ธนัชชา บุญมีมาก (ป๊อป)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 201,
        "date": "20/05/2569 14:49",
        "complainant": "ชนันพร อินขำ",
        "issue": "เน็ตเวิร์ค: Calenda, บัญชีผู้ใช้: google dive",
        "duration": "00:40",
        "responder": "ผู้ใช้รับเชิญ 40358",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 202,
        "date": "22/05/2569 09:02",
        "complainant": "ณัฏชณินภา กำจร",
        "issue": "บัญชีผู้ใช้: Lark, Chat GPT",
        "duration": "00:05",
        "responder": "ผู้ใช้รับเชิญ 57602",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 203,
        "date": "22/05/2569 10:16",
        "complainant": "ชนันพร อินขำ",
        "issue": "บัญชีผู้ใช้: Tiktok",
        "duration": "00:05",
        "responder": "ชนันพร อินขำ (ไอซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 204,
        "date": "22/05/2569 10:37",
        "complainant": "ปัญจมา สมบัติกำไร",
        "issue": "บัญชีผู้ใช้: Tiktok",
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
        "issue": "บัญชีผู้ใช้: Lark",
        "duration": "00.05",
        "responder": "ธันวา เเซ่เเฮ (ไนซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 206,
        "date": "01/06/2569 11:02",
        "complainant": "รามจิตติ ชินนะเกิด",
        "issue": "บัญชีผู้ใช้: Email",
        "duration": "00.05",
        "responder": "รามจิตติ ชินนะเกิดโชค (เบนซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 208,
        "date": "02/06/2569 08:46",
        "complainant": "กานต์ฑิตา ธีระพิบูลย์",
        "issue": "ฮาร์ดแวร์: Notebook, ซอฟต์แวร์: Google Link, Windows, Lark, เน็ตเวิร์ค: เชื่อมต่อ Server ไม่ได้, บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "กานต์ฑิตา ธีระพิบูลย์ (อิง)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 209,
        "date": "02/06/2569 11:27",
        "complainant": "รามจิตติ",
        "issue": "บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "รามจิตติ ชินนะเกิดโชค (เบนซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 210,
        "date": "02/06/2569 11:41",
        "complainant": "ชนันพร อินขำ",
        "issue": "บัญชีผู้ใช้: Chat GPT",
        "duration": "00.02",
        "responder": "Guest User 94269",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 211,
        "date": "02/06/2569 12:09",
        "complainant": "พิชชาพร คอทอง",
        "issue": "บัญชีผู้ใช้: Email",
        "duration": "00:30",
        "responder": "พิชชาพร คอทอง (พีเจ้น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 212,
        "date": "04/06/2569 09:24",
        "complainant": "เย็นฤดี มาระวัง",
        "issue": "ฮาร์ดแวร์: Notebook",
        "duration": "00.05",
        "responder": "เย็นฤดี มาระวัง (ฝ้าย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 213,
        "date": "05/06/2569 10:14",
        "complainant": "อาทิตยา มุมทอง",
        "issue": "บัญชีผู้ใช้: Email",
        "duration": "00:10",
        "responder": "อาทิตยา มุมทอง (ขมิ้น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 219,
        "date": "08/06/2569 11:01",
        "complainant": "เย็นฤดี มาระวัง",
        "issue": "ฮาร์ดแวร์: Notebook, บัญชีผู้ใช้: Printer WIFI",
        "duration": "00.10",
        "responder": "เย็นฤดี มาระวัง (ฝ้าย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 220,
        "date": "09/06/2569 09:34",
        "complainant": "เย็นฤดี มาระวัง",
        "issue": "ฮาร์ดแวร์: Mouse",
        "duration": "00.03",
        "responder": "ผู้ใช้รับเชิญ 66959",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 221,
        "date": "09/06/2569 09:51",
        "complainant": "พิสิษฐ์ มงคลสมบัติศิริ",
        "issue": "บัญชีผู้ใช้: Email",
        "duration": "00.03",
        "responder": "พิสิษฐ์ มงคลสมบัติศิริ (เจมส์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 223,
        "date": "09/06/2569 10:34",
        "complainant": "เอมปวีภร์ วัชระตระการพงศ์",
        "issue": "ฮาร์ดแวร์: Mac",
        "duration": "00.23",
        "responder": "เอมปวีภร์ วัชระตระการพงศ์ (กิ๊ฟ)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 224,
        "date": "10/06/2569 15:57",
        "complainant": "เนลินญาน์  ศิระไมตรีฉัตร (เนลิน)",
        "issue": "ฮาร์ดแวร์: Mac",
        "duration": "00.30",
        "responder": "เนลินญาน์ ศิระไมตรีฉัตร (เนลิน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 225,
        "date": "11/06/2569 08:21",
        "complainant": "อาทิตยา มุมทอง",
        "issue": "ฮาร์ดแวร์: Mac",
        "duration": "00:20",
        "responder": "อาทิตยา มุมทอง (ขมิ้น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 226,
        "date": "15/06/2569 09:08",
        "complainant": "มนัสนันท์ เทพแก้ว ",
        "issue": "บัญชีผู้ใช้: Email",
        "duration": "00.01",
        "responder": "Guest User 33774",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 227,
        "date": "15/06/2569 10:30",
        "complainant": "",
        "issue": "เช็คอุปกรณ์ Admin CRM บัญชี Producer Live",
        "duration": "02:30",
        "responder": "อำพล แซ่แฮ (โอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 228,
        "date": "15/06/2569 12:50",
        "complainant": "เนลินญาน์  ศิระไมตรีฉัตร ",
        "issue": "บัญชีผู้ใช้: google meet",
        "duration": "00:10",
        "responder": "เนลินญาน์ ศิระไมตรีฉัตร (เนลิน)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 257,
        "date": "16/06/2569 14:18",
        "complainant": "นภัสสร นาสวน",
        "issue": "ฮาร์ดแวร์: iPad",
        "duration": "00.02",
        "responder": "นภัสสร นาสวน (โบว์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 258,
        "date": "16/06/2569 15:40",
        "complainant": "อิศราภรณ์ ปิ่นงาม",
        "issue": "เข้า IG Fern Clinic ในคอมกับโทรศัพท์เลขาให้หน่อยค่ะ เพื่อให้ฝั่ง content ทำการปรับแก้ไอจีได้",
        "duration": "00:20",
        "responder": "อิศราภรณ์ ปิ่นงาม (เบนซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 259,
        "date": "16/06/2569 16:08",
        "complainant": "ชนันพร อินขำ",
        "issue": "บัญชีผู้ใช้: Tiktok",
        "duration": "00.15",
        "responder": "ชนันพร อินขำ (ไอซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 260,
        "date": "18/06/2569 08:15",
        "complainant": "อาทิตยา มุมทอง",
        "issue": "บัญชีผู้ใช้: Email",
        "duration": "00:20",
        "responder": "อาทิตยา มุมทอง (ขมิ้น)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 261,
        "date": "19/06/2569 10:34",
        "complainant": "อารยา ธนพันธุ์พาณิชย์",
        "issue": "ขอบัญชี",
        "duration": "00.10",
        "responder": "Guest User 72644",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 262,
        "date": "23/06/2569 10:02",
        "complainant": "ชัยธัช ชัยวัฒน์",
        "issue": "บัญชีผู้ใช้: Email, Lark",
        "duration": "00:30",
        "responder": "รามจิตติ ชินนะเกิดโชค (เบนซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 263,
        "date": "29/06/2569 09:41",
        "complainant": "ณัฐกานต์ ชิดปรางค์",
        "issue": "ฮาร์ดแวร์: Printer",
        "duration": "00:30",
        "responder": "ณัฐกานต์ ชิดปรางค์ (เตยหอม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 264,
        "date": "29/06/2569 13:09",
        "complainant": "บุษกร บัวสวรรค์",
        "issue": "กดล็อกอิน CapCut เข้าเมลที่ชื่อวิดีโอทีม แล้วมันให้กรอกวันเดือนปีเกิด กดออกจากหน้านี้ไม่ได้เลย น่าจะเป็นวันเดือนปีเกิดที่ตรงกับเมล โลกบังคับออกแล้วก็รีสตาร์ตเครื่องแล้วก็เป็นเหมือนเดิม",
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
        "sn": 265,
        "date": "01/07/2569 09:11",
        "complainant": "ชนันพร อินขำ",
        "issue": "บัญชีผู้ใช้: Chat GPT",
        "duration": "07.00",
        "responder": "ชนันพร อินขำ (ไอซ์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 266,
        "date": "01/07/2569 13:47",
        "complainant": "ศุภฤกษ์ ภายไธสง",
        "issue": "ฮาร์ดแวร์: iPad, ซอฟต์แวร์: IOS, บัญชีผู้ใช้: Email",
        "duration": "00.10",
        "responder": "ศุภฤกษ์ ภายไธสง (ดรีม)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 267,
        "date": "01/07/2569 16:14",
        "complainant": "สุรวิชญ์ โพธิ์ตาก (ไมค์)",
        "issue": "ท้องเสียบางวันเวลาเช้า ๆ ",
        "duration": "00:15",
        "responder": "สุรวิชญ์ โพธิ์ตาก (ไมค์เมโลดี้)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 268,
        "date": "03/07/2569 11:09",
        "complainant": "พิสิษฐ์ มงคลสมบัติศิริ",
        "issue": "ฮาร์ดแวร์: Notebook, บัญชีผู้ใช้: Printer WIFI",
        "duration": "00:20",
        "responder": "พิสิษฐ์ มงคลสมบัติศิริ (เจมส์)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 269,
        "date": "08/07/2569 15:57",
        "complainant": "วิลาสินี ทับทิม",
        "issue": "ไม่มีเครื่องปริ้นท์ในโปรแกรมพีคค่ะ/พี่นีปริ้นท์งานไม่ได้เลยค่ะ/SET ให้ด้วยค่ะ",
        "duration": "00:30",
        "responder": "Guest User 35635",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 271,
        "date": "13/07/2569 14:04",
        "complainant": "รามจิตติ ชินนะเกิดโชค",
        "issue": "ฮาร์ดแวร์: PC Computer, บัญชีผู้ใช้: Email, Kumoo, google dive",
        "duration": "00.25",
        "responder": "รามจิตติ ชินนะเกิดโชค (เบนซ์)",
        "status": "เสร็จสิ้น",
        "cost": 13900
      },
      {
        "sn": 272,
        "date": "14/07/2569 13:37",
        "complainant": "วิลาสินี ทับทิม",
        "issue": "ฮาร์ดแวร์: Keyboard , บัญชีผู้ใช้: Lark, Chat GPT",
        "duration": "00.30",
        "responder": "Guest User 35635",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 274,
        "date": "16/07/2569 08:45",
        "complainant": "เนลินญาน์  ศิระไมตรีฉัตร",
        "issue": "บัญชีผู้ใช้: Email",
        "duration": "00.05",
        "responder": "Guest User 32271",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 275,
        "date": "16/07/2569 09:06",
        "complainant": "ปณิศอร บุญจูบุตร",
        "issue": "บัญชีผู้ใช้: Email",
        "duration": "00.05",
        "responder": "ปณิศอร  บุญจูบุตร   (แต๊นซ์ )",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 276,
        "date": "16/07/2569 09:47",
        "complainant": "วิจิตราภรณ์ พึ่งจันดุม",
        "issue": "ฮาร์ดแวร์: PC Computer",
        "duration": "00.30",
        "responder": "วิจิตราภรณ์ พึ่งจันดุม (พลอย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 277,
        "date": "16/07/2569 12:34",
        "complainant": "วิจิตราภรณ์ พึ่งจันดุม",
        "issue": "ฮาร์ดแวร์: PC Computer",
        "duration": "00.30",
        "responder": "วิจิตราภรณ์ พึ่งจันดุม (พลอย)",
        "status": "เสร็จสิ้น",
        "cost": 0
      },
      {
        "sn": 278,
        "date": "16/07/2569 13:58",
        "complainant": "เนตรปรีญา ทัดศรี",
        "issue": "ฮาร์ดแวร์: Printer",
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







export default function App() {
  const [data, setData] = useState(initialDashboardData);
  const [assetsList, setAssetsList] = useState(initialAssetsData);
  const [assetSearch, setAssetSearch] = useState('');
  const [assetDeptFilter, setAssetDeptFilter] = useState('');
  const [assetStatusFilter, setAssetStatusFilter] = useState('');
  const [currentMonth, setCurrentMonth] = useState("2026-07");
  const [activeModal, setActiveModal] = useState(null); // 'edit', 'expiringAssets', 'expiringSoftware', 'topBrokenDevices', 'assetsList'
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
              notes: row['หมายเหตุ'] || ''
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
              issue: finalIssueText,
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
                      <th>ชื่อผู้แจ้ง</th>
                      <th>ปัญหา / อาการเสีย / รายการ</th>
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
                          <td>{ticket.complainant}</td>
                          <td>{ticket.issue}</td>
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
                        <td colSpan="8" style={{ textAlign: 'center' }}>ไม่มีข้อมูลประวัติประวัติงานซ่อมสำหรับเดือนนี้</td>
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
                        <th>สถานะ</th>
                        <th>หมายเหตุ</th>
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
                            <td>{asset.itemType}</td>
                            <td><strong>{asset.deviceSerial}</strong></td>
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
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan="8" style={{ textAlign: 'center' }}>ไม่พบคลังอุปกรณ์ที่ตรงตามเงื่อนไข</td>
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
    </>
  );
}
