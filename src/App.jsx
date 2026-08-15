import React, { Fragment, useState, useEffect, useRef, useCallback, useMemo } from 'react';
import { Routes, Route, useNavigate } from 'react-router-dom';
import LarkForm from './pages/LarkForm';
import lightItLogo from './assets/light_it_logo.jpg';
import Chart from 'chart.js/auto';
import * as XLSX from 'xlsx';
import { 
  Ticket, 
  Lightbulb, 
  Edit3, 
  Printer, 
  Database,
  CheckCircle,
  AlertTriangle,
  X,
  Upload,
  Download,
  FileSpreadsheet,
  ChevronDown,
  ChevronRight,
  Menu,
  Laptop,
  FileCode,
  ShieldCheck,
  Wrench,
  RotateCcw
} from 'lucide-react';

const API_BASE = window.location.hostname === 'localhost' ? 'http://localhost:5000' : '';
const ADMIN_PASSWORD_HASH = '1e630fe2c4c6fecd9f5181b3bd43242407c8efa7e6e7db16204dc447257224db';

// Initial blank data - use Excel import to load real data
const initialAssetsData = [
  {
    "sn": 23,
    "date": "20/05/2569",
    "user": "เธญเธณเธเธฅ เน€เน€เธเนเน€เน€เธฎ",
    "position": "IT",
    "itemType": "Computer (Pc)",
    "deviceSerial": "AS-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 83,
    "date": "06/12/2568",
    "user": "เธญเธณเธเธฅ    เนเธเนเนเธฎ",
    "position": "IT",
    "itemType": "Ipad",
    "deviceSerial": "iPad-010",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 217,
    "date": "",
    "user": "เธญเธณเธเธฅ  เนเธเนเนเธฎ",
    "position": "IT",
    "itemType": "Mornitor",
    "deviceSerial": "LG-005",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 213,
    "date": "18/05/2569",
    "user": "เธเธฑเธเธงเธฒ เน€เน€เธเนเน€เน€เธฎ",
    "position": "IT",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-010 ",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 138,
    "date": "20/11/2568",
    "user": "เธญเธณเธเธฅ   เนเธเนเนเธฎ",
    "position": "IT",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-033",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 239,
    "date": "",
    "user": "เธญเธณเธเธฅ  เนเธเนเนเธฎ",
    "position": "IT",
    "itemType": "External HDD",
    "deviceSerial": "ETN-003 WD My PassPort 1TB",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 246,
    "date": "",
    "user": "เธญเธณเธเธฅ  เนเธเนเนเธฎ",
    "position": "IT",
    "itemType": "Cable HDMI",
    "deviceSerial": "เนเธกเนเธกเธตเธซเธกเธฒเธขเน€เธฅเธ เธญเธธเธเธเธฃเธ“เน",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 248,
    "date": "10/07/2569",
    "user": "เธเธฑเธเธงเธฒ เน€เน€เธเนเน€เน€เธฎ",
    "position": "IT",
    "itemType": "Mouse",
    "deviceSerial": "MOS-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 252,
    "date": "",
    "user": "เธญเธณเธเธฅ  เนเธเนเนเธฎ",
    "position": "IT",
    "itemType": "Keyboard",
    "deviceSerial": "KBD-003",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 329,
    "date": "10/07/2569",
    "user": "เธเธฑเธเธงเธฒ เน€เน€เธเนเน€เน€เธฎ",
    "position": "IT",
    "itemType": "Keyboard",
    "deviceSerial": "KBD-007",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 272,
    "date": "",
    "user": "เธญเธณเธเธฅ   เนเธเนเนเธฎ",
    "position": "IT",
    "itemType": "Screwdriver",
    "deviceSerial": "เนเธกเนเธกเธตเธซเธกเธฒเธขเน€เธฅเธ เธญเธธเธเธเธฃเธ“เน",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 33,
    "date": "09/07/2569",
    "user": "เธงเธดเธฅเธฒเธชเธดเธเธต เธ—เธฑเธเธ—เธดเธก (เธเธต)",
    "position": "Accounting",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-006",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ เธงเธฑเธ เธ"
  },
  {
    "sn": 313,
    "date": "",
    "user": "เธชเธธเธ เธฒเธ เนเธชเธเธเธฑเธเธ—เธฃเน ( เธชเนเธก )",
    "position": "Accounting",
    "itemType": "Printer",
    "deviceSerial": "PT-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ เธงเธฑเธ เธ"
  },
  {
    "sn": 135,
    "date": "09/07/2569",
    "user": "เธงเธดเธฅเธฒเธชเธดเธเธต เธ—เธฑเธเธ—เธดเธก (เธเธต)",
    "position": "Accounting",
    "itemType": "Mornitor",
    "deviceSerial": "LG-004",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ เธงเธฑเธ เธ"
  },
  {
    "sn": 27,
    "date": "",
    "user": "เธชเธธเธ เธฒเธ เนเธชเธเธเธฑเธเธ—เธฃเน ( เธชเนเธก )",
    "position": "Accounting",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-005",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ เธงเธฑเธ เธ"
  },
  {
    "sn": 34,
    "date": "",
    "user": "เธ“เธฑเธเธเธฒ เธเธณเธชเธญเธเธเธฑเธเธเน (เธเธดเนเธ)",
    "position": "Procurement & Warehouse Officer",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-007",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 93,
    "date": "08/05/3111",
    "user": "เธจเธดเธฃเธดเธเธฃ เน€เธเธเธฃเธกเธนเธฅ ( เธเธต )",
    "position": "Procurement & Warehouse Officer",
    "itemType": "Ipad",
    "deviceSerial": "iPad-020",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 71,
    "date": "08/05/3111",
    "user": "เธจเธดเธฃเธดเธเธฃ เน€เธเธเธฃเธกเธนเธฅ ( เธเธต )",
    "position": "Procurement & Warehouse Officer",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-011",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 111,
    "date": "08/05/3111",
    "user": "เธจเธดเธฃเธดเธเธฃ เน€เธเธเธฃเธกเธนเธฅ ( เธเธต )",
    "position": "Procurement & Warehouse Officer",
    "itemType": "Apple Pancill",
    "deviceSerial": "Pencil-018",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 26,
    "date": "",
    "user": "เธจเธดเธฃเธดเธเธฃ เน€เธเธเธฃเธกเธนเธฅ ( เธเธต )",
    "position": "Procurement & Warehouse Officer",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-004",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 58,
    "date": "",
    "user": "เธเธเธฑเธเธเธฒ เธเธธเธเธกเธตเธกเธฒเธ (เธเนเธญเธ)",
    "position": "Content Creator",
    "itemType": "Macbook",
    "deviceSerial": "MacBookAir-028",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 325,
    "date": "10/07/2569",
    "user": "เธงเธฃเธดเธชเธฃเธฒ เธชเธเธงเธเธงเธเธฉเน (เธเธดเธเธ”เธตเน)",
    "position": "Content Creator",
    "itemType": "Macbook",
    "deviceSerial": "MacBook air-029  ",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 339,
    "date": "01/07/2569",
    "user": "เธ“เธฑเธเธเธเธเธฒเธ”เธฒ เธ•เธฃเธตเธงเธดเธงเธฑเธ’เธเนเธเธธเธฅ (Mac)",
    "position": "Content Creator",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-043",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 231,
    "date": "23/05/2569",
    "user": "เธเธเธฑเธเธเธฒ เธเธธเธเธกเธตเธกเธฒเธ (เธเนเธญเธ)",
    "position": "Content Creator",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-MKT-005",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 214,
    "date": "",
    "user": "เธงเธฃเธดเธชเธฃเธฒ เธชเธเธงเธเธงเธเธฉเน (เธเธดเธเธ”เธตเน)",
    "position": "Tiktok Content Creator",
    "itemType": "Notebook Asus",
    "deviceSerial": "Asus-031",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 265,
    "date": "11/12/2568",
    "user": "เธเธ เธฑเธชเธชเธฃ เธเธฒเธชเธงเธ ( เนเธเธงเน )",
    "position": "Tiktok Content Creator",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-037",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 86,
    "date": "",
    "user": "เธเธ เธฑเธชเธชเธฃ เธเธฒเธชเธงเธ ( เนเธเธงเน )",
    "position": "Tiktok Content Creator",
    "itemType": "Ipad",
    "deviceSerial": "iPad-013",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 87,
    "date": "18/07/2569",
    "user": "เธเธธเธฉเธเธฃ เธเธฑเธงเธชเธงเธฃเธฃเธเน",
    "position": "Tiktok Content Creator",
    "itemType": "Ipad",
    "deviceSerial": "iPad-014",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 282,
    "date": "15/07/2569",
    "user": "เธเธ“เธดเธจเธญเธฃ เธเธธเธเธเธนเธเธธเธ•เธฃ",
    "position": "Tiktok Content Creator",
    "itemType": "Ipad",
    "deviceSerial": "iPad-021",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 53,
    "date": "",
    "user": "เธ—เธตเธก",
    "position": "Tiktok Content Creator",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-013",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 62,
    "date": "",
    "user": "เธ—เธตเธก",
    "position": "Tiktok Content Creator",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 66,
    "date": "",
    "user": "เธ—เธตเธก",
    "position": "Tiktok Content Creator",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-006",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 72,
    "date": "",
    "user": "เธ—เธตเธก",
    "position": "Tiktok Content Creator",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-012",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 357,
    "date": "",
    "user": "เธ—เธตเธก",
    "position": "Tiktok Content Creator",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-014",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 365,
    "date": "30/06/2569",
    "user": "เธ—เธตเธก",
    "position": "Tiktok Content Creator",
    "itemType": "IPhone",
    "deviceSerial": "iphone-015",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 20,
    "date": "",
    "user": "เธเธ เธฑเธชเธชเธฃ เธเธฒเธชเธงเธ ( เนเธเธงเน )",
    "position": "Tiktok Content Creator",
    "itemType": "IMac",
    "deviceSerial": "MC-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 326,
    "date": "",
    "user": "เธเธฑเธขเธเธฑเธ เธเธฑเธขเธงเธฑเธ’เธเน (เธกเธฒเธฃเนเธ)",
    "position": "Graphic Designer",
    "itemType": "Computer (Pc)",
    "deviceSerial": "AS-003",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 305,
    "date": "08/06/2569",
    "user": "เธฃเธฒเธกเธเธดเธ•เธ•เธด เธเธดเธเธเธฐเน€เธเธดเธ”เนเธเธ(เน€เธเธเธเน)",
    "position": "Graphic Designer",
    "itemType": "Computer (Pc), Mornitor",
    "deviceSerial": "PC LG-004 ",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 294,
    "date": "08/06/2569",
    "user": "เธญเธฒเธ—เธดเธ•เธขเธฒ เธกเธธเธกเธ—เธญเธ (เธเธกเธดเนเธ)",
    "position": "Graphic Designer",
    "itemType": "IMac",
    "deviceSerial": "MC-008",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 319,
    "date": "19/05/2569",
    "user": "เธเธดเธเธเธฒเธเธฃ เธเธญเธ—เธญเธ(เธเธตเน€เธเนเธ)",
    "position": "Graphic Designer",
    "itemType": "IMac",
    "deviceSerial": "MC-009",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 218,
    "date": "08/06/2569",
    "user": "เธญเธฒเธ—เธดเธ•เธขเธฒ เธกเธธเธกเธ—เธญเธ (เธเธกเธดเนเธ)",
    "position": "Graphic Designer",
    "itemType": "Mornitor",
    "deviceSerial": "LG-003 (เธเธญ)",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 303,
    "date": "08/06/2569",
    "user": "เธฃเธฒเธกเธเธดเธ•เธ•เธด เธเธดเธเธเธฐเน€เธเธดเธ”เนเธเธ(เน€เธเธเธเน)",
    "position": "Graphic Designer",
    "itemType": "Mornitor",
    "deviceSerial": "AS-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 304,
    "date": "",
    "user": "เธเธดเธเธเธฒเธเธฃ เธเธญเธ—เธญเธ(เธเธตเน€เธเนเธ)",
    "position": "Graphic Designer",
    "itemType": "Mornitor",
    "deviceSerial": "AS-003",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 318,
    "date": "",
    "user": "เธเธฑเธขเธเธฑเธ เธเธฑเธขเธงเธฑเธ’เธเน (เธกเธฒเธฃเนเธ)",
    "position": "Graphic Designer",
    "itemType": "Mornitor",
    "deviceSerial": "LG-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 234,
    "date": "08/06/2569",
    "user": "เธญเธฒเธ—เธดเธ•เธขเธฒ เธกเธธเธกเธ—เธญเธ (เธเธกเธดเนเธ)",
    "position": "Graphic Designer",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-Meeting-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 301,
    "date": "",
    "user": "เธเธดเธเธเธฒเธเธฃ เธเธญเธ—เธญเธ(เธเธตเน€เธเนเธ)",
    "position": "Graphic Designer",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Meeting-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 238,
    "date": "08/06/2569",
    "user": "เธญเธฒเธ—เธดเธ•เธขเธฒ เธกเธธเธกเธ—เธญเธ (เธเธกเธดเนเธ)",
    "position": "Graphic Designer",
    "itemType": "External HDD",
    "deviceSerial": "ETN-002  WD My PassPort 4TB",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 247,
    "date": "",
    "user": "เธเธดเธเธเธฒเธเธฃ เธเธญเธ—เธญเธ(เธเธตเน€เธเนเธ)",
    "position": "Graphic Designer",
    "itemType": "Mouse",
    "deviceSerial": "MOS-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 249,
    "date": "08/06/2569",
    "user": "เธญเธฒเธ—เธดเธ•เธขเธฒ เธกเธธเธกเธ—เธญเธ (เธเธกเธดเนเธ)",
    "position": "Graphic Designer",
    "itemType": "Mouse",
    "deviceSerial": "MOS-003",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 250,
    "date": "08/06/2569",
    "user": "เธญเธฒเธ—เธดเธ•เธขเธฒ เธกเธธเธกเธ—เธญเธ (เธเธกเธดเนเธ)",
    "position": "Graphic Designer",
    "itemType": "Keyboard",
    "deviceSerial": "KBD-004",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 260,
    "date": "01/06/2569",
    "user": "เธเธฑเธขเธเธฑเธ เธเธฑเธขเธงเธฑเธ’เธเน (เธกเธฒเธฃเนเธ)",
    "position": "Graphic Designer",
    "itemType": "Keyboard",
    "deviceSerial": "KBD-006",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 302,
    "date": "08/06/2569",
    "user": "เธฃเธฒเธกเธเธดเธ•เธ•เธด เธเธดเธเธเธฐเน€เธเธดเธ”เนเธเธ(เน€เธเธเธเน)",
    "position": "Graphic Designer",
    "itemType": "Hub Lan",
    "deviceSerial": "เนเธกเนเธกเธตเธซเธกเธฒเธขเน€เธฅเธ",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 330,
    "date": "20/05/2569",
    "user": "เธเธดเธเธเธฒเธเธฃ เธเธญเธ—เธญเธ(เธเธตเน€เธเนเธ)",
    "position": "Graphic Designer",
    "itemType": "Hub Lan",
    "deviceSerial": "USB-C TO LAN ",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 229,
    "date": "",
    "user": "เธ—เธตเธก admin",
    "position": "Admin",
    "itemType": "Computer (Pc)",
    "deviceSerial": "MIS-002 เธกเธดเธเธดpc",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 36,
    "date": "",
    "user": "เธเธฒเธเธ•เนเธ‘เธดเธ•เธฒ เธเธตเธฃเธฐเธเธดเธเธนเธฅเธขเน ( เธญเธดเธ )",
    "position": "Admin",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-011",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 37,
    "date": "",
    "user": "เธจเธธเธ เธคเธเธฉเน เธ เธฒเธขเนเธเธชเธ (เธเธฃเธดเธก)",
    "position": "Admin",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-017",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธซเธฒ lark เนเธกเนเน€เธเธญ เธเธฃเธฑเธ"
  },
  {
    "sn": 76,
    "date": "30/06/2569",
    "user": "เธจเธธเธ เธคเธเธฉเน เธ เธฒเธขเนเธเธชเธ (เธ”เธฃเธตเธก)",
    "position": "Admin",
    "itemType": "Ipad",
    "deviceSerial": "iPad-003",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 358,
    "date": "",
    "user": "เธ—เธฑเธจเธงเธงเธฃเธ“ เธงเธฑเธฅเธขเนเธ”เธฒเธง (เนเธเธ•เธญเธ)",
    "position": "Admin",
    "itemType": "Ipad",
    "deviceSerial": "iPad-019",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 359,
    "date": "",
    "user": "เธ—เธฑเธจเธงเธงเธฃเธ“ เธงเธฑเธฅเธขเนเธ”เธฒเธง (เนเธเธ•เธญเธ)",
    "position": "Admin",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-004",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 360,
    "date": "",
    "user": "เธเธฒเธเธ•เนเธ‘เธดเธ•เธฒ เธเธตเธฃเธฐเธเธดเธเธนเธฅเธขเน ( เธญเธดเธ )",
    "position": "Admin",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-009",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 28,
    "date": "",
    "user": "เธเธฑเธเธเธกเธฒ เธชเธกเธเธฑเธ•เธดเธเธณเนเธฃ (เน€เธญเธดเธ)",
    "position": "Admin",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-014",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 47,
    "date": "",
    "user": "เธ—เธฑเธจเธงเธงเธฃเธ“ เธงเธฑเธฅเธขเนเธ”เธฒเธง (เนเธเธ•เธญเธ)",
    "position": "Admin",
    "itemType": "Notebook Acer, Iphone",
    "deviceSerial": "ACER-015",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 56,
    "date": "14/07/2569",
    "user": "เน€เธเธ•เธฃเธเธฃเธตเธเธฒ เธ—เธฑเธ”เธจเธฃเธต",
    "position": "Receptionist",
    "itemType": "Macbook",
    "deviceSerial": "MacBookAir-009",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 91,
    "date": "07/09/3111",
    "user": "Receptionist",
    "position": "Receptionist",
    "itemType": "Ipad",
    "deviceSerial": "iPad-018",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 67,
    "date": "",
    "user": "Receptionist",
    "position": "Receptionist",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-007",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 51,
    "date": "",
    "user": "Receptionist",
    "position": "Receptionist",
    "itemType": "IMac",
    "deviceSerial": "MC-003",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 52,
    "date": "",
    "user": "Receptionist",
    "position": "Receptionist",
    "itemType": "IMac",
    "deviceSerial": "MC-004",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 334,
    "date": "16/07/2569",
    "user": "เน€เธเธ•เธฃเธเธฃเธตเธเธฒ เธ—เธฑเธ”เธจเธฃเธต",
    "position": "Receptionist",
    "itemType": "Mouse",
    "deviceSerial": "MC-008",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 35,
    "date": "",
    "user": "เธงเธดเธเธฒเธเธ”เธฒ เธ—เธธเธกเธกเธเธ•เธฃเธต ( เนเธเนเธงเนเธซเธงเนเธง )",
    "position": "Payroll Officer",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-008",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 44,
    "date": "10/07/2569",
    "user": "เธญเธเธดเธฃเธธเธ•เธ•เน เธเธฒเธเนเธชเธเธเธฒ",
    "position": "Live Streamer",
    "itemType": "Ipad",
    "deviceSerial": "iPad-016",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 85,
    "date": "26/06/2569",
    "user": "เธเธเธฑเธเธงเธฅเธฑเธเธเธเน เธจเธฃเธตเธญเธญเธ (เธญเธฒเนเธ)",
    "position": "Live Streamer",
    "itemType": "Ipad",
    "deviceSerial": "iPad-012",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 314,
    "date": "",
    "user": "เธเธเธดเธชเธฃเธฒ เธ•เธฃเธตเธชเธฑเธ•เธขเธเธธเธฅ (เธญเธธเธเธญเธดเธ)",
    "position": "Live Streamer",
    "itemType": "Ipad",
    "deviceSerial": "\niPad 017\n",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 61,
    "date": "",
    "user": "เธเธเธดเธชเธฃเธฒ เธ•เธฃเธตเธชเธฑเธ•เธขเธเธธเธฅ  (เธญเธธเธเธญเธดเธ)",
    "position": "Live Streamer",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 63,
    "date": "26/06/2569",
    "user": "เธเธเธฑเธเธงเธฅเธฑเธเธเธเน เธจเธฃเธตเธญเธญเธ (เธญเธฒเนเธ)",
    "position": "Live Streamer",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-003",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 285,
    "date": "",
    "user": "เธเธเธดเธชเธฃเธฒ เธ•เธฃเธตเธชเธฑเธ•เธขเธเธธเธฅ  (เธญเธธเธเธญเธดเธ)",
    "position": "Live Streamer",
    "itemType": "Adapter Apple",
    "deviceSerial": "ADT-007",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 287,
    "date": "",
    "user": "เธเธเธดเธชเธฃเธฒ เธ•เธฃเธตเธชเธฑเธ•เธขเธเธธเธฅ  (เธญเธธเธเธญเธดเธ)",
    "position": "Live Streamer",
    "itemType": "Cable Apple",
    "deviceSerial": "-",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 40,
    "date": "",
    "user": "เธญเธชเธกเธฒเธ เธฃเธ“เน เธขเนเธฒเธเน€เธ”เธดเธก ( เน€เธเธก )",
    "position": "Personal Assistant to CEO",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-020",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 43,
    "date": "08/01/3111",
    "user": "เธชเธฃเธธเธ”เธ•เธฒ เธเนเธญเธเธฃเธดเนเธ (เธเธฒเธ)",
    "position": "Personal Assistant to CEO",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-023",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 57,
    "date": "16/07/2569",
    "user": "เน€เธเธฅเธดเธเธเธฒเธเน  เธจเธดเธฃเธฐเนเธกเธ•เธฃเธตเธเธฑเธ•เธฃ (เน€เธเธฃเธดเธ)",
    "position": "Personal Assistant to CEO",
    "itemType": "Macbook",
    "deviceSerial": "MacBookAir-027",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 7,
    "date": "",
    "user": "เธญเธชเธกเธฒเธ เธฃเธ“เน เธขเนเธฒเธเน€เธ”เธดเธก ( เน€เธเธก )",
    "position": "Personal Assistant to CEO",
    "itemType": "IPhone",
    "deviceSerial": "iPhone-010",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 1,
    "date": "",
    "user": "เธชเธธเธ”เธเธดเธ”เธฒ เน€เธเนเธฒเธซเธญเธก (เธ•เนเธฒเธข) (KT)",
    "position": "HR",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 41,
    "date": "",
    "user": "เธเธคเธฉเธ“เธฒ เธฅเธณเน€เธเนเธ ( เธเธฅเธญเธข )",
    "position": "HR",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-038",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 42,
    "date": "",
    "user": "เธเธฒเธฃเธตเธฃเธฑเธ•เธเน เธเธฑเธเธ—เธญเธ ( เธญเธญเธข )",
    "position": "HR",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-022",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 74,
    "date": "",
    "user": "เธเธฒเธฃเธตเธฃเธฑเธ•เธเน เธเธฑเธเธ—เธญเธ ( เธญเธญเธข )",
    "position": "HR",
    "itemType": "Ipad",
    "deviceSerial": "iPad-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 88,
    "date": "10/07/2569",
    "user": "เธจเธฃเธฑเธเธเธฒ เธเธฃเธฃเธกเน€เธเธตเธขเธกเธ เธฑเธเธ”เธต ( เนเธเธ )",
    "position": "HR",
    "itemType": "Ipad",
    "deviceSerial": "iPad-015",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 320,
    "date": "17/06/2569",
    "user": "เธเธฒเธฃเธตเธฃเธฑเธ•เธเน เธเธฑเธเธ—เธญเธ ( เธญเธญเธข )",
    "position": "HR",
    "itemType": "Ipad",
    "deviceSerial": "iPad-005",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 48,
    "date": "",
    "user": "เธจเธฃเธฑเธเธเธฒ เธเธฃเธฃเธกเน€เธเธตเธขเธกเธ เธฑเธเธ”เธต ( เนเธเธ )",
    "position": "HR",
    "itemType": "Notebook HP",
    "deviceSerial": "HP-016",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 49,
    "date": "",
    "user": "เธ“เธฑเธเธเธฒเธเธ•เน เธเธดเธ”เธเธฃเธฒเธเธเน (เน€เธ•เธขเธซเธญเธก)",
    "position": "Photographer",
    "itemType": "IMac",
    "deviceSerial": "MC-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 368,
    "date": "",
    "user": "เธงเธดเธเธดเธ•เธฃเธฒเธ เธฃเธ“เน เธเธถเนเธเธเธฑเธเธ”เธธเธก (เธเธฅเธญเธข)",
    "position": "Photographer",
    "itemType": "IMac",
    "deviceSerial": "MC-010",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 242,
    "date": "16/07/2569",
    "user": "เธ“เธฑเธเธเธฒเธเธ•เน เธเธดเธ”เธเธฃเธฒเธเธเน (เน€เธ•เธขเธซเธญเธก)",
    "position": "Photographer",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-MKT-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 75,
    "date": "",
    "user": "เธชเธฃเธงเธดเธเธเน เธชเธดเธ—เธเธด ( เธเธญเธก )",
    "position": "Sale",
    "itemType": "Ipad",
    "deviceSerial": "iPad-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 81,
    "date": "",
    "user": "เธญเธฒเธ—เธดเธ•เธขเน เธชเธกเธเธฒเธฃ ( เน€เธเนเธฒเธเธฒเธข )",
    "position": "Sale",
    "itemType": "Ipad",
    "deviceSerial": "iPad-008",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 82,
    "date": "",
    "user": "เธชเธธเธเธ—เธฃเธต เธเธธเธเธเธฒเธ ( เธ•เธฒเธฅ )",
    "position": "Sale",
    "itemType": "Ipad",
    "deviceSerial": "iPad-009",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 95,
    "date": "",
    "user": "เธชเธธเธเธ—เธฃเธต เธเธธเธเธเธฒเธ ( เธ•เธฒเธฅ )",
    "position": "Sale",
    "itemType": "Apple Pancill",
    "deviceSerial": "Pencil-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 100,
    "date": "",
    "user": "เธชเธฃเธงเธดเธเธเน เธชเธดเธ—เธเธด ( เธเธญเธก )",
    "position": "Sale",
    "itemType": "Apple Pancill",
    "deviceSerial": "Pencil-007",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 6,
    "date": "",
    "user": "เธเธคเธ•เธดเธ เธงเธดเธเธฑเธขเธ”เธดเธฉเธ (เธญเนเธ)crm",
    "position": "CRM Officer",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-019",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 46,
    "date": "",
    "user": "เธเธคเธ•เธดเธเธฒ เธ—เธฒเธฃเธฐเธเธฑเธเธเน (เนเธเธเธเธตเน) crm",
    "position": "CRM Officer",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-026",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 321,
    "date": "18/05/2569",
    "user": "เธเธเธฑเธเธเธฃ เธญเธดเธเธเธณ (เนเธญเธเน)crm",
    "position": "CRM Officer",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-003",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 361,
    "date": "",
    "user": "เธเธคเธ•เธดเธเธฒ เธ—เธฒเธฃเธฐเธเธฑเธเธเน (เนเธเธเธเธตเน) crm",
    "position": "CRM Officer",
    "itemType": "Samsung",
    "deviceSerial": "Samsung Galaxy-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 307,
    "date": "27/05/2569",
    "user": "เธเธฑเธ“เธ“เธงเธดเธเธเน เธ—เธญเธเธงเธฑเธ (เธเธญเธข)",
    "position": "Data Analysis",
    "itemType": "Computer (Pc)",
    "deviceSerial": "AS-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 309,
    "date": "",
    "user": "เธเธฑเธ“เธ“เธงเธดเธเธเน เธ—เธญเธเธงเธฑเธ (เธเธญเธข)",
    "position": "Data Analysis",
    "itemType": "Computer (Pc)",
    "deviceSerial": "LG-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 311,
    "date": "",
    "user": "เธเธฑเธ“เธ“เธงเธดเธเธเน เธ—เธญเธเธงเธฑเธ (เธเธญเธข)",
    "position": "Data Analysis",
    "itemType": "Computer (Pc)",
    "deviceSerial": "MSI-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 350,
    "date": "",
    "user": "เธเธฑเธ“เธ“เธงเธดเธเธเน เธ—เธญเธเธงเธฑเธ (เธเธญเธข)",
    "position": "Data Analysis",
    "itemType": "Notebook Asus",
    "deviceSerial": "Asus-024",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 223,
    "date": "",
    "user": "เธเธฑเธ“เธ“เธงเธดเธเธเน เธ—เธญเธเธงเธฑเธ (เธเธญเธข)",
    "position": "Data Analysis",
    "itemType": "Mornitor",
    "deviceSerial": "Asus-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 243,
    "date": "",
    "user": "เธเธฑเธ“เธ“เธงเธดเธเธเน เธ—เธญเธเธงเธฑเธ (เธเธญเธข)",
    "position": "Data Analysis",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-003",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 244,
    "date": "",
    "user": "เธเธฑเธ“เธ“เธงเธดเธเธเน เธ—เธญเธเธงเธฑเธ (เธเธญเธข)",
    "position": "Data Analysis",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-004",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 245,
    "date": "",
    "user": "เธเธฑเธ“เธ“เธงเธดเธเธเน เธ—เธญเธเธงเธฑเธ (เธเธญเธข)",
    "position": "Data Analysis",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-005",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 251,
    "date": "28/05/2569",
    "user": "เธเธฑเธ“เธ“เธงเธดเธเธเน เธ—เธญเธเธงเธฑเธ (เธเธญเธข)",
    "position": "Data Analysis",
    "itemType": "Keyboard",
    "deviceSerial": "KBD-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 262,
    "date": "",
    "user": "เธเธฑเธ“เธ“เธงเธดเธเธเน เธ—เธญเธเธงเธฑเธ (เธเธญเธข)",
    "position": "Data Analysis",
    "itemType": "Keyboard",
    "deviceSerial": "-",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 337,
    "date": "04/06/2569",
    "user": "เธญเธ เธดเธชเธดเธ—เธเธดเน เธเธฃเธเธฑเธเธ—เธฃเธฒเธงเธฑเธ’เธเน (เธเธธเนเธข)",
    "position": "Video Content Tiktok",
    "itemType": "IMac",
    "deviceSerial": "MC-005",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 310,
    "date": "",
    "user": "เธเธเธจเธเธฃ เธเนเธญเธเนเธช (เธเธญเธ)",
    "position": "Live Producer",
    "itemType": "Computer (Pc)",
    "deviceSerial": "LG-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 266,
    "date": "09/06/2569",
    "user": "เธเธดเธชเธดเธฉเธเน เธกเธเธเธฅเธชเธกเธเธฑเธ•เธดเธจเธดเธฃเธด (เน€เธเธกเธชเน)",
    "position": "Live Producer",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-021",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 280,
    "date": "",
    "user": "เธเธเธจเธเธฃ เธเนเธญเธเนเธช (เธเธญเธ)",
    "position": "Live Producer",
    "itemType": "Mornitor",
    "deviceSerial": "MSI-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 367,
    "date": "",
    "user": "เธเธฒเธขเธเธเธจเธเธฃ เธเนเธญเธเนเธช",
    "position": "Live Producer",
    "itemType": "Mornitor",
    "deviceSerial": "LG-006",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 30,
    "date": "01/06/2569",
    "user": "เน€เธขเนเธเธคเธ”เธต เธกเธฒเธฃเธฐเธงเธฑเธ (เธเนเธฒเธข)",
    "position": "Live Producer",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-013",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 31,
    "date": "05/05/2569",
    "user": "เธเธเธจเธเธฃ เธเนเธญเธเนเธช (เธเธญเธ)",
    "position": "Live Producer",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-039",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 343,
    "date": "",
    "user": "เธ—เธตเธก live Producer",
    "position": "Live Producer",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "usb-a to c",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 240,
    "date": "",
    "user": "เธเธเธจเธเธฃ เธเนเธญเธเนเธช (เธเธญเธ)",
    "position": "Live Producer",
    "itemType": "Capture Card",
    "deviceSerial": "VCS-LIVE-0001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 253,
    "date": "",
    "user": "เธเธเธจเธเธฃ เธเนเธญเธเนเธช (เธเธญเธ)",
    "position": "Live Producer",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-007",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 342,
    "date": "",
    "user": "เธ—เธตเธก live Producer",
    "position": "Live Producer",
    "itemType": "Cable HDMI",
    "deviceSerial": "hdmi Canon",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 315,
    "date": "",
    "user": "เธเธเธจเธเธฃ เธเนเธญเธเนเธช (เธเธญเธ)",
    "position": "Live Producer",
    "itemType": "Reez Live",
    "deviceSerial": "เนเธกเนเธกเธตเธซเธกเธฒเธขเน€เธฅเธ เธญเธธเธเธเธฃเธ“เน",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 344,
    "date": "",
    "user": "เธ—เธตเธก live Producer",
    "position": "Live Producer",
    "itemType": "เธญเธธเธเธเธฃเธ“เนเน€เธชเธฃเธดเธก/เธญเธทเนเธเน",
    "deviceSerial": "เธชเธฒเธข aux branding",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 345,
    "date": "",
    "user": "เธ—เธตเธก live Producer",
    "position": "Live Producer",
    "itemType": "เธญเธธเธเธเธฃเธ“เนเน€เธชเธฃเธดเธก/เธญเธทเนเธเน",
    "deviceSerial": "hdmi เธ•เนเธญเธเธญเธกเธญเธเธดเน€เธ•เธญเธฃเนเนเธฅเธเน",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 346,
    "date": "",
    "user": "เธ—เธตเธก live Producer",
    "position": "Live Producer",
    "itemType": "เธญเธธเธเธเธฃเธ“เนเน€เธชเธฃเธดเธก/เธญเธทเนเธเน",
    "deviceSerial": "hdmi it-002",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 347,
    "date": "",
    "user": "เธ—เธตเธก live Producer",
    "position": "Live Producer",
    "itemType": "เธญเธธเธเธเธฃเธ“เนเน€เธชเธฃเธดเธก/เธญเธทเนเธเน",
    "deviceSerial": "adc-006",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 348,
    "date": "",
    "user": "เธ—เธตเธก live Producer",
    "position": "Live Producer",
    "itemType": "เธญเธธเธเธเธฃเธ“เนเน€เธชเธฃเธดเธก/เธญเธทเนเธเน",
    "deviceSerial": "เนเธกเธเนเธฅเนเธญเธข dji-branding",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 349,
    "date": "",
    "user": "เธ—เธตเธก live Producer",
    "position": "Live Producer",
    "itemType": "เธญเธธเธเธเธฃเธ“เนเน€เธชเธฃเธดเธก/เธญเธทเนเธเน",
    "deviceSerial": "เนเธกเธเนเธฅเนเธญเธข dji-mkt",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 235,
    "date": "",
    "user": "เธซเนเธญเธเธเธฃเธฐเธเธธเธก เธเนเธฒเธ 18 เธเธฑเนเธ 4",
    "position": "Meeting Room",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-MKT-004",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 54,
    "date": "",
    "user": "เธญเธ เธดเธเธเธฒ เธจเธฃเธตเธ•เธฐเธงเธฑเธ ( เธเธญเธชเธ•เนเธ )",
    "position": "Boss",
    "itemType": "IMac",
    "deviceSerial": "MC-006",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 55,
    "date": "",
    "user": "เธญเธ เธดเธเธเธฒ เธจเธฃเธตเธ•เธฐเธงเธฑเธ ( เธเธญเธชเธ•เนเธ )",
    "position": "Boss",
    "itemType": "IMac",
    "deviceSerial": "MC-007",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 232,
    "date": "",
    "user": "เธญเธ เธดเธเธเธฒ เธจเธฃเธตเธ•เธฐเธงเธฑเธ ( เธเธญเธชเธ•เนเธ )",
    "position": "Boss",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-เธบBoss-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 241,
    "date": "",
    "user": "เธญเธ เธดเธเธเธฒ เธจเธฃเธตเธ•เธฐเธงเธฑเธ ( เธเธญเธชเธ•เนเธ )",
    "position": "Boss",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 233,
    "date": "",
    "user": "เธซเนเธญเธ Studio เธเนเธฒเธ 18 เธเธฑเนเธ 1",
    "position": "Studio",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-Studio-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 277,
    "date": "26/12/2568",
    "user": "เธเธเธ—เนเธชเธดเธฃเธต เธเธฅเธทเนเธกเธ—เธฃเธฑเธเธขเน (เนเธเธเธฅเธน)",
    "position": "Data Entry",
    "itemType": "Notebook Asus",
    "deviceSerial": "Asus-041",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 29,
    "date": "",
    "user": "เธเธ เธฒเธงเธต เธเธฑเธเธ—เธฃเนเธเธงเธฒเธ (เธเธฃเธตเธก)",
    "position": "Data Entry",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-012",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 125,
    "date": "16/06/2569",
    "user": "เธญเธฒเธฃเธขเธฒ เธเธเธเธฑเธเธเธธเนเธเธฒเธ“เธดเธเธขเน (เธซเธเธธเธเธซเธเธดเธ)",
    "position": "Sale Manager",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-032",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 363,
    "date": "19/01/2569",
    "user": "เธ“เธฑเธเธ“เธดเธเธฒ เธจเธฃเธตเธงเธฃเธญเธฃเธฃเธ–เธดเธเธธเธฅ (เนเธเธ•เธญเธ)",
    "position": "KOL & Event Marketing",
    "itemType": "Iphone",
    "deviceSerial": "IPhone -005",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 283,
    "date": "19/01/2569",
    "user": "เธ“เธฑเธเธ“เธดเธเธฒ เธจเธฃเธตเธงเธฃเธญเธฃเธฃเธ–เธดเธเธธเธฅ (เนเธเธ•เธญเธ)",
    "position": "KOL & Event Marketing",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "Lenovo-001",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "เธฃเธญเธ•เธฃเธงเธเธชเธญเธ"
  },
  {
    "sn": 45,
    "date": "",
    "user": "เธเธฑเธ—เธเธเธฑเธเธ—เน เธเธฒเน€เธ”เธตเธขเธฃเน  (เธเธฒเน€เธ”เธตเธขเธฃเน )",
    "position": "OR",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-025",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 60,
    "date": "",
    "user": "เธซเธกเธญเธเธฒเธ",
    "position": "CEO",
    "itemType": "Macbook",
    "deviceSerial": "MacBookAir-030",
    "status": "เนเธเนเธเธฒเธ",
    "notes": "     "
  },
  {
    "sn": 215,
    "date": "",
    "user": "เธซเธกเธญเธเธฒเธ",
    "position": "CEO",
    "itemType": "Macbook",
    "deviceSerial": "MAcbook Pro-034",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 269,
    "date": "",
    "user": "เธซเธกเธญเธเธฒเธ",
    "position": "CEO",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-MKT-005",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 236,
    "date": "",
    "user": "Connection Nas",
    "position": "Center Storage",
    "itemType": "External HDD",
    "deviceSerial": "Segate Station 10 TB",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 237,
    "date": "",
    "user": "Connection Nas",
    "position": "Center Storage",
    "itemType": "External HDD",
    "deviceSerial": "ETN-001  WD My Book 4TB",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 254,
    "date": "",
    "user": "Connection TV",
    "position": "TV Center",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-006",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 271,
    "date": "",
    "user": "Connection TV",
    "position": "TV Center",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-008",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 291,
    "date": "",
    "user": "Connection TV",
    "position": "TV Center",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-009",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 292,
    "date": "",
    "user": "Connection TV",
    "position": "TV Center",
    "itemType": "Cable HDMI",
    "deviceSerial": "Fernclinic-IT-010",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 340,
    "date": "06/07/2569",
    "user": "เธเธฒเธขเธชเธธเธ—เธเธดเธฃเธเธเน เน€เธเธฃเธทเธญเนเธเธเธนเธฅเธขเนเธเธธเธฅ ( เน€เธเธเธเน )",
    "position": "Senior Digital Marketing",
    "itemType": "Notebook Lenovo",
    "deviceSerial": "LENOVO-042",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 369,
    "date": "10/07/2569",
    "user": "เธชเธธเธฃเธงเธดเธเธเน เนเธเธเธดเนเธ•เธฒเธ (เนเธกเธเนเน€เธกเนเธฅเธ”เธตเน)",
    "position": "Operation Department",
    "itemType": "Macbook",
    "deviceSerial": "Macbook-044",
    "status": "เนเธเนเธเธฒเธ",
    "notes": ""
  },
  {
    "sn": 308,
    "date": "",
    "user": "เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ",
    "position": "-",
    "itemType": "Computer (Pc)",
    "deviceSerial": "LG-003",
    "status": "เธงเนเธฒเธ",
    "notes": ""
  },
  {
    "sn": 264,
    "date": "",
    "user": "เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ",
    "position": "-",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-036",
    "status": "เธงเนเธฒเธ",
    "notes": ""
  },
  {
    "sn": 316,
    "date": "",
    "user": "เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ",
    "position": "-",
    "itemType": "Notebook Asus",
    "deviceSerial": "ASUS-018",
    "status": "เธงเนเธฒเธ",
    "notes": ""
  },
  {
    "sn": 328,
    "date": "",
    "user": "เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ",
    "position": "-",
    "itemType": "Macbook",
    "deviceSerial": "MacBook air-035",
    "status": "เธงเนเธฒเธ",
    "notes": ""
  },
  {
    "sn": 50,
    "date": "",
    "user": "เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ",
    "position": "-",
    "itemType": "Ipad",
    "deviceSerial": "iPad-004",
    "status": "เธงเนเธฒเธ",
    "notes": ""
  },
  {
    "sn": 80,
    "date": "",
    "user": "เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ",
    "position": "-",
    "itemType": "Ipad",
    "deviceSerial": "iPad-007",
    "status": "เธงเนเธฒเธ",
    "notes": ""
  },
  {
    "sn": 84,
    "date": "",
    "user": "เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ",
    "position": "-",
    "itemType": "Ipad",
    "deviceSerial": "iPad-011",
    "status": "เธงเนเธฒเธ",
    "notes": ""
  },
  {
    "sn": 131,
    "date": "",
    "user": "เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ",
    "position": "-",
    "itemType": "Ipad",
    "deviceSerial": "iPad-006",
    "status": "เธงเนเธฒเธ",
    "notes": ""
  },
  {
    "sn": 117,
    "date": "",
    "user": "เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ",
    "position": "-",
    "itemType": "Printer",
    "deviceSerial": "PT-001",
    "status": "เธฃเธญเธเนเธญเธก",
    "notes": ""
  },
  {
    "sn": 230,
    "date": "",
    "user": "เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ",
    "position": "-",
    "itemType": "Hub USB-TypeC",
    "deviceSerial": "Fernclinic-MKT-003",
    "status": "เธงเนเธฒเธ",
    "notes": ""
  },
  {
    "sn": 335,
    "date": "",
    "user": "เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ",
    "position": "-",
    "itemType": "Mouse",
    "deviceSerial": "MC-009",
    "status": "เธงเนเธฒเธ",
    "notes": ""
  },
  {
    "sn": 261,
    "date": "",
    "user": "เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ",
    "position": "-",
    "itemType": "Keyboard",
    "deviceSerial": "KBD-001",
    "status": "เธงเนเธฒเธ",
    "notes": ""
  }
];

const initialDashboardData = {
  '2026-07': {
    monthName: 'กรกฎาคม 2569',
    totalAssets: 0,
    assetValue: 0,
    assetsExpiring: 0,
    assetsBroken: 0,
    assetsLost: 0,
    assetsVacant: 0,
    ticketsCount: 0,
    slaPercent: 100,
    responseTime: 0,
    resolutionTime: 0,
    csat: 0,
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

const AssetTagEditor = ({ value, onChange, single = false, placeholder = 'เธเธดเธกเธเน Tag เนเธฅเนเธงเธเธ” Enter' }) => {
  const [draft, setDraft] = useState('');
  const tags = String(value || '').split(/[,\n]+/).map((item) => item.trim()).filter(Boolean);

  const addDraft = () => {
    const nextTag = draft.trim();
    if (!nextTag) return;
    onChange((single ? [nextTag] : [...new Set([...tags, nextTag])]).join(', '));
    setDraft('');
  };

  return (
    <div className="asset-tag-editor">
      <div className="asset-tag-editor-tags">
        {tags.map((tag, index) => (
          <span key={`${tag}-${index}`} className={`asset-tag asset-tag-${[...tag].reduce((sum, char) => sum + char.charCodeAt(0), 0) % 6}`}>
            {tag}
            <button type="button" onClick={() => onChange(tags.filter((_, tagIndex) => tagIndex !== index).join(', '))} aria-label={`เธฅเธ Tag ${tag}`}>ร—</button>
          </span>
        ))}
      </div>
      <input
        type="text"
        value={draft}
        onChange={(event) => setDraft(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === 'Enter' || event.key === ',') {
            event.preventDefault();
            addDraft();
          }
        }}
        onBlur={addDraft}
        placeholder={tags.length ? 'เน€เธเธดเนเธก Tag...' : placeholder}
      />
    </div>
  );
};

const AssetTagPicker = ({ value, onChange, options, single = false, onClose }) => {
  const selected = String(value || '').split(/[,\n]+/).map((item) => item.trim()).filter(Boolean);
  const toggleOption = (option) => {
    if (single) {
      onChange(option);
      return;
    }
    onChange((selected.includes(option) ? selected.filter((tag) => tag !== option) : [...selected, option]).join(', '));
  };

  return (
    <div className="asset-tag-picker" onClick={(event) => event.stopPropagation()}>
      <div className="asset-tag-picker-title">เน€เธฅเธทเธญเธ Tag เธ—เธตเนเธกเธตเธญเธขเธนเน</div>
      <div className="asset-tag-picker-options">
        {options.map((option) => (
          <button
            type="button"
            key={option}
            className={selected.includes(option) ? 'selected' : ''}
            onClick={() => toggleOption(option)}
          >
            {option}
          </button>
        ))}
      </div>
      <div className="asset-tag-picker-title">เธซเธฃเธทเธญเน€เธเธดเนเธก Tag เนเธซเธกเน</div>
      <AssetTagEditor value={value} onChange={onChange} single={single} />
      <div className="asset-tag-picker-footer">
        <span>เธเธ”เธเธฑเธเธ—เธถเธเธ”เนเธฒเธเธเธเน€เธเธทเนเธญเธขเธทเธเธขเธฑเธ</span>
        <button type="button" onClick={onClose}>เน€เธชเธฃเนเธเธชเธดเนเธ</button>
      </div>
    </div>
  );
};

const seedSoftwareLicenses = [
  { name: 'Meitu', owner: 'Tiktok Content Creator', price: 1290, paymentChannel: 'Apple', paymentDate: 'เธฃเธฒเธขเธเธต', expiringDate: '', registeredEmail: 'drfernaesthetique@gmail.com', currentUsers: '' },
  { name: 'Adobe', owner: 'เธเธฃเธฒเธเธเธดเธ', price: 2592, paymentChannel: 'เธเธฑเธ•เธฃ', paymentDate: '02/07/2026', expiringDate: '2026-07-31', registeredEmail: 'drfernaesthetique@gmail.com', currentUsers: 'เธญเธฒเธ—เธดเธ•เธขเธฒ เธกเธธเธกเธ—เธญเธ (เธเธกเธดเนเธ), เธฃเธงเธกเธเธดเธ•เธ•เน เธเธฑเธเธ—เธฃเนเน€เธเธดเธ”เนเธเธ (เน€เธเธเธเน)' },
  { name: 'Freepik', owner: 'เธเธฃเธฒเธเธเธดเธ', price: 11250, paymentChannel: 'เธเธฑเธ•เธฃ', paymentDate: '11/05/2026-2027 (1 เธเธต)', expiringDate: '2027-06-11', registeredEmail: 'graphicfernclinic@gmail.com', currentUsers: 'เธญเธฒเธ—เธดเธ•เธขเธฒ เธกเธธเธกเธ—เธญเธ (เธเธกเธดเนเธ), เธฃเธงเธกเธเธดเธ•เธ•เน เธเธฑเธเธ—เธฃเนเน€เธเธดเธ”เนเธเธ (เน€เธเธเธเน), เธเธฑเธขเธเธฑเธ เธเธฑเธขเธงเธฑเธ’เธเน (เธกเธฒเธฃเนเธ)' },
  { name: 'Kumoo', owner: 'เธเธฃเธฒเธเธเธดเธ', price: 3077, paymentChannel: 'เธเธฑเธ•เธฃ', paymentDate: '06/01/2026', expiringDate: '2027-01-07', registeredEmail: 'drfernaesthetique@gmail.com', currentUsers: 'เธญเธฒเธ—เธดเธ•เธขเธฒ เธกเธธเธกเธ—เธญเธ (เธเธกเธดเนเธ), เธฃเธงเธกเธเธดเธ•เธ•เน เธเธฑเธเธ—เธฃเนเน€เธเธดเธ”เนเธเธ (เน€เธเธเธเน), เธเธฑเธขเธเธฑเธ เธเธฑเธขเธงเธฑเธ’เธเน (เธกเธฒเธฃเนเธ)' },
  { name: 'Cupcut', owner: 'Tiktok Content Creator', price: 345, paymentChannel: 'Apple', paymentDate: '31/07/2026', expiringDate: '2026-08-31', registeredEmail: 'drfernaesthetique@gmail.com', currentUsers: 'เธเธธเธฉเธเธฃ เธเธฑเธงเธชเธงเธฃเธฃเธเน (เน€เธฃเธเธเธตเน), เธญเธ เธดเธชเธดเธ—เธเธดเน เธเธฃเธเธฑเธเธ—เธฃเนเธงเธฑเธ’เธเน (เธเธธเนเธข)' },
  { name: 'Microsoft Office 365', owner: 'IT', price: 3690, paymentChannel: 'Microsoft Office', paymentDate: '', expiringDate: '', registeredEmail: '', currentUsers: '' },
  { name: 'Lark', owner: 'IT', price: 0, paymentChannel: '', paymentDate: '', expiringDate: '', registeredEmail: '', currentUsers: '' },
  { name: 'Google Suite', owner: 'IT', price: 0, paymentChannel: '', paymentDate: '', expiringDate: '', registeredEmail: '', currentUsers: '' },
  { name: 'Chat GPT', owner: 'IT', price: 0, paymentChannel: '', paymentDate: '', expiringDate: '', registeredEmail: '', currentUsers: '' },
  { name: 'Adobe', owner: 'เธเธฃเธฒเธเธเธดเธ', price: 11105, paymentChannel: 'เธเธฑเธ•เธฃ', paymentDate: '2025-12-01', expiringDate: '2026-12-01', registeredEmail: 'fernclinic.it@gmail.com', currentUsers: 'เธเธฑเธขเธเธฑเธ เธเธฑเธขเธงเธฑเธ’เธเน (เธกเธฒเธฃเนเธ), เธเธดเธเธเธฒเธเธฃ เธเธฅเธญเธงเธ (เธเธตเนเน€เธเธ)' },
  { name: 'Cupcut', owner: 'Tiktok Content Creator', price: 1810, paymentChannel: 'Apple', paymentDate: '25/06/2027', expiringDate: '2027-07-23', registeredEmail: 'drfernbussiness@gmail.com', currentUsers: '' },
  { name: 'Meitu', owner: 'Tiktok Content Creator', price: 1190, paymentChannel: 'Applepay', paymentDate: '22/01/2027', expiringDate: '2027-01-22', registeredEmail: 'drfernbussiness@gmail.com', currentUsers: '' },
  { name: 'PEAK', owner: 'Accounting', price: 12480, paymentChannel: '', paymentDate: 'เธฃเธฒเธขเธเธต', expiringDate: '2070-03-30', registeredEmail: '', currentUsers: '' },
  { name: 'empeo', owner: 'HR', price: 132515, paymentChannel: '', paymentDate: 'เธฃเธฒเธขเธเธต', expiringDate: '2070-06-09', registeredEmail: '', currentUsers: '' },
  { name: 'Chromecast Premium', owner: 'HR', price: 399, paymentChannel: '', paymentDate: 'เธฃเธฒเธขเน€เธ”เธทเธญเธ', expiringDate: '', registeredEmail: '', currentUsers: '' },
].map((item) => {
  const used = item.currentUsers ? item.currentUsers.split(',').filter(Boolean).length : 0;
  return { ...item, used, vacant: 0, licenses: used, monthlyCost: item.price, status: 'เนเธเนเธเธฒเธ', isLicenseRecord: true, sourceVersion: 'software-image-v2' };
});

const isValidDashboardData = (value) => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const monthKeys = Object.keys(value);
  return monthKeys.length > 0 && monthKeys.every((key) => /^\d{4}-\d{2}$/.test(key) && value[key] && typeof value[key] === 'object');
};

// Some legacy uploads were decoded as Windows-874 even though the workbook
// contained UTF-8 Thai text. Repair only strings that have the characteristic
// mojibake pattern and that round-trip to valid UTF-8; ordinary Thai is kept as-is.
let windows874ByteLookup;
const getWindows874ByteLookup = () => {
  if (windows874ByteLookup) return windows874ByteLookup;
  const decoder = new TextDecoder('windows-874');
  windows874ByteLookup = new Map();
  for (let byte = 0; byte <= 255; byte += 1) {
    windows874ByteLookup.set(decoder.decode(Uint8Array.of(byte)), byte);
  }
  return windows874ByteLookup;
};

const thaiMojibakeScore = (text) => {
  const value = String(text || '');
  const controlCount = (value.match(/[\u0080-\u009f]/g) || []).length;
  const thaiUtf8PrefixCount = (value.match(/\u0e18\u0e30|\u0e40\u0e19/g) || []).length;
  return (controlCount * 4) + thaiUtf8PrefixCount;
};

const repairThaiMojibake = (value) => {
  if (typeof value !== 'string' || thaiMojibakeScore(value) < 2) return value;
  try {
    const byteLookup = getWindows874ByteLookup();
    const bytes = [];
    for (const character of value) {
      const byte = byteLookup.get(character);
      if (byte === undefined) return value;
      bytes.push(byte);
    }
    const repaired = new TextDecoder('utf-8', { fatal: true }).decode(Uint8Array.from(bytes));
    if (!/[\u0e00-\u0e7f]/.test(repaired)) return value;
    return thaiMojibakeScore(repaired) < thaiMojibakeScore(value)
      ? repaired.normalize('NFC')
      : value;
  } catch {
    return value;
  }
};

const repairThaiTextDeep = (value) => {
  if (typeof value === 'string') return repairThaiMojibake(value);
  if (Array.isArray(value)) return value.map(repairThaiTextDeep);
  if (!value || typeof value !== 'object' || value instanceof Date) return value;
  return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, repairThaiTextDeep(item)]));
};

function Dashboard() {
  const navigate = useNavigate();
  const [data, setData] = useState(() => {
    const saved = localStorage.getItem('it_dashboard_data');
    if (!saved) return initialDashboardData;
    try {
      const parsed = JSON.parse(saved);
      return isValidDashboardData(parsed) ? repairThaiTextDeep(parsed) : initialDashboardData;
    } catch {
      return initialDashboardData;
    }
  });
  const [assetsList, setAssetsList] = useState(() => {
    const saved = localStorage.getItem('it_dashboard_assets');
    return saved ? repairThaiTextDeep(JSON.parse(saved)) : initialAssetsData;
  });
  const [assetSearch, setAssetSearch] = useState('');
  const [assetDeptFilter, setAssetDeptFilter] = useState('');
  const [assetStatusFilter, setAssetStatusFilter] = useState('');
  const [assetTypeFilter, setAssetTypeFilter] = useState('');
  const [assetRequests, setAssetRequests] = useState([]);
  const [assetRequestLoading, setAssetRequestLoading] = useState(false);
  const [assetRequestForm, setAssetRequestForm] = useState({ requester: '', department: '', itemType: '', purpose: '', requestedDate: new Date().toISOString().slice(0, 10), notes: '' });
  const [assetWorkflowRole, setAssetWorkflowRole] = useState('requester');
  const [assetRequesterSearch, setAssetRequesterSearch] = useState('');
  const [assetReturnSearch, setAssetReturnSearch] = useState('');
  const [assetReturnIdentity, setAssetReturnIdentity] = useState('');
  const [assetReturnView, setAssetReturnView] = useState('returns');
  const notificationInitializedRef = useRef(false);
  const [currentMonth, setCurrentMonth] = useState(() => {
    const keys = Object.keys(data || {}).sort((a, b) => b.localeCompare(a));
    return keys.length > 0 ? keys[0] : "2026-07";
  });
  const [activeModal, setActiveModal] = useState(null); // 'edit', 'expiringAssets', 'expiringSoftware', 'topBrokenDevices', 'assetsList', 'fullConsole'
  const [importStatus, setImportStatus] = useState(null); // { type: 'success' | 'error', message: string }
  
  // Console tab, month selectors, and editing state trackers
  const [consoleTab, setConsoleTab] = useState('months');
  const [consoleMonth, setConsoleMonth] = useState('2026-07');
  const [consoleSaving, setConsoleSaving] = useState(false);
  const [consoleSaveMessage, setConsoleSaveMessage] = useState('');
  const [consoleAssetSearch, setConsoleAssetSearch] = useState('');
  const [editingAssetTagField, setEditingAssetTagField] = useState(null);
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
  const [larkTicketAssetSerial, setLarkTicketAssetSerial] = useState('');
  const [larkTicketIssue, setLarkTicketIssue] = useState('');
  const [larkTicketCause, setLarkTicketCause] = useState('');
  const [larkTicketDuration, setLarkTicketDuration] = useState('00:30');
  const [larkTicketResponder, setLarkTicketResponder] = useState('');
  const [larkTicketStatus, setLarkTicketStatus] = useState('เน€เธชเธฃเนเธเธชเธดเนเธ');
  const [larkTicketCost, setLarkTicketCost] = useState('0');

  const [larkAssetUser, setLarkAssetUser] = useState('');
  const [larkAssetPosition, setLarkAssetPosition] = useState('');
  const [larkAssetItemType, setLarkAssetItemType] = useState('');
  const [larkAssetSerial, setLarkAssetSerial] = useState('');
  const [larkAssetStatus, setLarkAssetStatus] = useState('เนเธเนเธเธฒเธ');
  const [larkAssetNotes, setLarkAssetNotes] = useState('');

  const [larkSubmitted, setLarkSubmitted] = useState(false);
  const [larkTicketRole, setLarkTicketRole] = useState('user'); // 'user' | 'it'
  const [selectedPendingTicketSn, setSelectedPendingTicketSn] = useState('');
  const [sidebarExpanded, setSidebarExpanded] = useState({
    mgmt: false,
    excel: false,
    export: false
  });
  const [mobileSidebarOpen, setMobileSidebarOpen] = useState(false);
  const [isLoaded, setIsLoaded] = useState(false);
  
  const [externalDevices, setExternalDevices] = useState([]);
  const [isFetchingDevices, setIsFetchingDevices] = useState(false);
  const [externalDevicesLastSynced, setExternalDevicesLastSynced] = useState(null);
  const [externalDevicesSyncError, setExternalDevicesSyncError] = useState('');
  const [externalDevicesRefreshKey, setExternalDevicesRefreshKey] = useState(0);

  useEffect(() => {
    let cancelled = false;
    let activeController = null;

    const fetchDevices = async () => {
      if (activeController) return;
      activeController = new AbortController();
      setIsFetchingDevices(true);
      try {
        const response = await fetch('https://ios-device-monitor-lkxv.onrender.com/api/devices', {
          cache: 'no-store',
          signal: activeController.signal
        });
        if (!response.ok) throw new Error(`Device Monitor API ${response.status}`);
        const result = await response.json();
        if (!cancelled && Array.isArray(result?.devices)) {
          setExternalDevices(result.devices);
          setExternalDevicesLastSynced(new Date());
          setExternalDevicesSyncError('');
        }
      } catch (error) {
        if (error.name !== 'AbortError') {
          console.error('Error fetching external devices:', error);
          if (!cancelled) setExternalDevicesSyncError('เธเธดเธเธเนเนเธกเนเธชเธณเน€เธฃเนเธ');
        }
      } finally {
        activeController = null;
        if (!cancelled) setIsFetchingDevices(false);
      }
    };

    const syncWhenVisible = () => {
      if (document.visibilityState === 'visible') fetchDevices();
    };

    fetchDevices();
    const interval = window.setInterval(fetchDevices, 30 * 1000);
    window.addEventListener('focus', fetchDevices);
    document.addEventListener('visibilitychange', syncWhenVisible);
    return () => {
      cancelled = true;
      if (activeController) activeController.abort();
      window.clearInterval(interval);
      window.removeEventListener('focus', fetchDevices);
      document.removeEventListener('visibilitychange', syncWhenVisible);
    };
  }, [externalDevicesRefreshKey]);
  const verifyAdminPasswordLocally = async (password) => {
    const encodedPassword = new TextEncoder().encode(password);
    const digest = await window.crypto.subtle.digest('SHA-256', encodedPassword);
    return Array.from(new Uint8Array(digest))
      .map((byte) => byte.toString(16).padStart(2, '0'))
      .join('') === ADMIN_PASSWORD_HASH;
  };

  useEffect(() => {
    if (!activeModal) return undefined;
    setMobileSidebarOpen(false);
    document.body.classList.add('modal-open');
    return () => document.body.classList.remove('modal-open');
  }, [activeModal]);

  const requireAdminAccess = async (openMenu) => {
    const password = window.prompt('เธเธฃเธธเธ“เธฒเธเธฃเธญเธเธฃเธซเธฑเธชเธเนเธฒเธเน€เธเนเธฒเธซเธเนเธฒเธ—เธตเน IT');
    if (password === null) return;
    if (await verifyAdminPasswordLocally(password)) {
      openMenu();
      return;
    }
    const controller = new AbortController();
    const timeoutId = window.setTimeout(() => controller.abort(), 12000);
    try {
      const response = await fetch(`${API_BASE}/api/admin/verify`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ password }),
        signal: controller.signal
      });
      if (!response.ok) {
        alert('เธฃเธซเธฑเธชเธเนเธฒเธเนเธกเนเธ–เธนเธเธ•เนเธญเธ');
        return;
      }
      openMenu();
    } catch (error) {
      console.error('Admin verification failed:', error);
      if (await verifyAdminPasswordLocally(password)) {
        openMenu();
        return;
      }
      alert('เธฃเธซเธฑเธชเธเนเธฒเธเนเธกเนเธ–เธนเธเธ•เนเธญเธ');
    } finally {
      window.clearTimeout(timeoutId);
    }
  };

  const loadAssetRequests = useCallback(async () => {
    try {
      const response = await fetch(`${API_BASE}/api/asset-requests`);
      if (!response.ok) throw new Error(`API ${response.status}`);
      setAssetRequests(repairThaiTextDeep(await response.json()));
    } catch (err) {
      console.error('Failed to load asset requests:', err);
    }
  }, []);

  const refreshOperationalStateFromDb = async () => {
    const response = await fetch(`${API_BASE}/api/db-state`, { cache: 'no-store' });
    if (!response.ok) throw new Error(`API ${response.status}`);
    const result = await response.json();
    isPollingUpdateRef.current = true;
    if (result.data) setData(repairThaiTextDeep(result.data));
    if (result.assetsList) setAssetsList(repairThaiTextDeep(result.assetsList));
  };

  const submitAssetRequest = async (event) => {
    event.preventDefault();
    setAssetRequestLoading(true);
    try {
      const response = await fetch(`${API_BASE}/api/asset-requests`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(assetRequestForm)
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'เธชเนเธเธเธณเธเธญเนเธกเนเธชเธณเน€เธฃเนเธ');
      setAssetRequestForm({ requester: '', department: '', itemType: '', purpose: '', requestedDate: new Date().toISOString().slice(0, 10), notes: '' });
      await loadAssetRequests();
      alert(result.count > 1
        ? `เธชเนเธเธเธณเธเธญเน€เธเธดเธ ${result.count} เน€เธเธฃเธทเนเธญเธเธชเธณเน€เธฃเนเธ เน€เธฅเธเธ—เธตเน #${result.ids.join(', #')}`
        : `เธชเนเธเธเธณเธเธญเน€เธเธดเธเน€เธฅเธเธ—เธตเน #${result.id} เธชเธณเน€เธฃเนเธ`);
    } catch (err) {
      alert(err.message);
    } finally {
      setAssetRequestLoading(false);
    }
  };

  const editAssetRequest = async (request) => {
    const requester = window.prompt('เธเธทเนเธญเธเธนเนเธเธญ', request.requester || '');
    if (requester === null) return;
    const department = window.prompt('เนเธเธเธ', request.department || '');
    if (department === null) return;
    const itemType = window.prompt('เธเธฃเธฐเน€เธ เธ—เธญเธธเธเธเธฃเธ“เน', request.item_type || '');
    if (itemType === null) return;
    const purpose = window.prompt('เน€เธซเธ•เธธเธเธฅเธเธฒเธฃเนเธเนเธเธฒเธ', request.purpose || '');
    if (purpose === null) return;
    const dueDate = window.prompt('เธเธณเธซเธเธ”เธเธทเธ (YYYY-MM-DD เธซเธฃเธทเธญเน€เธงเนเธเธงเนเธฒเธ)', request.due_date ? String(request.due_date).slice(0, 10) : '');
    if (dueDate === null) return;
    const notes = window.prompt('เธซเธกเธฒเธขเน€เธซเธ•เธธ', request.notes || '');
    if (notes === null) return;

    if (![requester, department, itemType, purpose].every(value => value.trim())) {
      alert('เธเธฃเธธเธ“เธฒเธเธฃเธญเธเธเธทเนเธญเธเธนเนเธเธญ เนเธเธเธ เธเธฃเธฐเน€เธ เธ—เธญเธธเธเธเธฃเธ“เน เนเธฅเธฐเน€เธซเธ•เธธเธเธฅเนเธซเนเธเธฃเธ');
      return;
    }
    if (!window.confirm(`เธขเธทเธเธขเธฑเธเธเธฑเธเธ—เธถเธเธเธฒเธฃเนเธเนเนเธเธเธณเธเธญ #${request.id} เธซเธฃเธทเธญเนเธกเน?`)) return;

    setAssetRequestLoading(true);
    try {
      const response = await fetch(`${API_BASE}/api/asset-requests/${request.id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ requester, department, itemType, purpose, dueDate, notes })
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'เนเธเนเนเธเธเธณเธเธญเนเธกเนเธชเธณเน€เธฃเนเธ');
      await loadAssetRequests();
      alert(`เนเธเนเนเธเธเธณเธเธญ #${request.id} เธชเธณเน€เธฃเนเธ`);
    } catch (err) {
      alert(err.message);
    } finally {
      setAssetRequestLoading(false);
    }
  };

  const runAssetRequestAction = async (request, action) => {
    const payload = { action };
    if (action === 'approve') {
      const vacantAssets = assetsList.filter(asset => asset.status === 'เธงเนเธฒเธ');
      if (vacantAssets.length === 0) return alert('เนเธกเนเธกเธตเธญเธธเธเธเธฃเธ“เนเธชเธ–เธฒเธเธฐเธงเนเธฒเธเธชเธณเธซเธฃเธฑเธเธญเธเธธเธกเธฑเธ•เธด');
      const choices = vacantAssets.slice(0, 30).map(asset => `${asset.sn}: ${asset.itemType} (${asset.deviceSerial})`).join('\n');
      const selected = window.prompt(`เธเธฃเธญเธเธฅเธณเธ”เธฑเธเธญเธธเธเธเธฃเธ“เนเธ—เธตเนเธ•เนเธญเธเธเธฒเธฃเธเธญเธ\n\n${choices}`);
      if (selected === null) return;
      if (!vacantAssets.some(asset => Number(asset.sn) === Number(selected))) return alert('เธฅเธณเธ”เธฑเธเธญเธธเธเธเธฃเธ“เนเนเธกเนเธ–เธนเธเธ•เนเธญเธเธซเธฃเธทเธญเน€เธเธฃเธทเนเธญเธเนเธกเนเธงเนเธฒเธ');
      payload.assetSn = Number(selected);
      payload.reviewer = window.prompt('เธเธทเนเธญเธเธนเนเธญเธเธธเธกเธฑเธ•เธด / เน€เธเนเธฒเธซเธเนเธฒเธ—เธตเน IT') || 'IT';
    } else if (action === 'reject') {
      const note = window.prompt('เธฃเธฐเธเธธเน€เธซเธ•เธธเธเธฅเธ—เธตเนเนเธกเนเธญเธเธธเธกเธฑเธ•เธด');
      if (note === null) return;
      payload.note = note;
      payload.reviewer = window.prompt('เธเธทเนเธญเธเธนเนเธเธดเธเธฒเธฃเธ“เธฒ') || 'IT';
    } else if (action === 'issue') {
      if (!window.confirm(`เธขเธทเธเธขเธฑเธเธชเนเธเธกเธญเธเธญเธธเธเธเธฃเธ“เนเนเธซเน ${request.requester}?`)) return;
      payload.reviewer = window.prompt('เธเธทเนเธญเน€เธเนเธฒเธซเธเนเธฒเธ—เธตเนเธเธนเนเธชเนเธเธกเธญเธ') || request.reviewer || 'IT';
    } else if (action === 'request_return') {
      if (!assetReturnIdentity.trim()) return alert('เธเธฃเธธเธ“เธฒเธฃเธฐเธเธธเธเธทเนเธญเธเธนเนเธเธทเธเธเนเธญเธ');
      if (!window.confirm(`เธขเธทเธเธขเธฑเธเนเธเนเธเธเธญเธเธทเธเธญเธธเธเธเธฃเธ“เน ${request.device_serial || request.item_type} เนเธซเน IT เธ•เธฃเธงเธเธฃเธฑเธเนเธเนเธซเธฃเธทเธญเนเธกเน?`)) return;
      payload.requesterIdentity = assetReturnIdentity.trim();
      payload.reviewer = assetReturnIdentity.trim();
      payload.note = 'เธเธนเนเนเธเนเธเธฒเธเนเธเนเธเธเธญเธเธทเธเธญเธธเธเธเธฃเธ“เน เธฃเธญเน€เธเนเธฒเธซเธเนเธฒเธ—เธตเน IT เธ•เธฃเธงเธเธฃเธฑเธ';
    } else if (action === 'return') {
      const adminPassword = window.prompt('เธเธฃเธธเธ“เธฒเธเธฃเธญเธเธฃเธซเธฑเธชเธเนเธฒเธเน€เธเนเธฒเธซเธเนเธฒเธ—เธตเน IT เน€เธเธทเนเธญเธขเธทเธเธขเธฑเธเธฃเธฑเธเธเธทเธ');
      if (adminPassword === null) return;
      const condition = window.prompt('เธชเธ เธฒเธเธ•เธญเธเธเธทเธ: เธเธเธ•เธด, เธเธณเธฃเธธเธ” เธซเธฃเธทเธญ เธชเธนเธเธซเธฒเธข', 'เธเธเธ•เธด');
      if (condition === null) return;
      if (!['เธเธเธ•เธด', 'เธเธณเธฃเธธเธ”', 'เธชเธนเธเธซเธฒเธข'].includes(condition)) return alert('เธเธฃเธธเธ“เธฒเธฃเธฐเธเธธ เธเธเธ•เธด, เธเธณเธฃเธธเธ” เธซเธฃเธทเธญ เธชเธนเธเธซเธฒเธข');
      payload.adminPassword = adminPassword;
      payload.condition = condition;
      payload.reviewer = window.prompt('เธเธทเนเธญเน€เธเนเธฒเธซเธเนเธฒเธ—เธตเนเธเธนเนเธ•เธฃเธงเธเธฃเธฑเธ') || 'IT';
      payload.note = window.prompt('เธซเธกเธฒเธขเน€เธซเธ•เธธเธเธฒเธฃเธฃเธฑเธเธเธทเธ (เธ–เนเธฒเธกเธต)') || '';
    }

    setAssetRequestLoading(true);
    try {
      const response = await fetch(`${API_BASE}/api/asset-requests/${request.id}/action`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'เธ”เธณเน€เธเธดเธเธเธฒเธฃเนเธกเนเธชเธณเน€เธฃเนเธ');
      await Promise.all([loadAssetRequests(), refreshOperationalStateFromDb()]);
      if (action === 'request_return') alert('เนเธเนเธเธเธญเธเธทเธเธญเธธเธเธเธฃเธ“เนเธชเธณเน€เธฃเนเธ เธเธฃเธธเธ“เธฒเธเธณเธญเธธเธเธเธฃเธ“เนเนเธซเนเน€เธเนเธฒเธซเธเนเธฒเธ—เธตเน IT เธ•เธฃเธงเธเธฃเธฑเธ');
      if (action === 'return') alert('IT เธ•เธฃเธงเธเธฃเธฑเธเธญเธธเธเธเธฃเธ“เนเนเธฅเธฐเธญเธฑเธเน€เธ”เธ•เธชเธ–เธฒเธเธฐเธเธฅเธฑเธเธชเธณเน€เธฃเนเธ');
    } catch (err) {
      alert(err.message);
    } finally {
      setAssetRequestLoading(false);
    }
  };

  useEffect(() => {
    if (!['assetWorkflow', 'assetReturns'].includes(activeModal)) return undefined;
    loadAssetRequests();
    const timer = setInterval(loadAssetRequests, 3000);
    return () => clearInterval(timer);
  }, [activeModal, loadAssetRequests]);

  const isPollingUpdateRef = useRef(false);
  const isPendingSyncRef = useRef(false);
  const syncQueueRef = useRef(Promise.resolve());
  const latestDataRef = useRef(data);
  const latestAssetsRef = useRef(assetsList);
  const softwareSeededRef = useRef(false);

  useEffect(() => {
    latestDataRef.current = data;
    latestAssetsRef.current = assetsList;
  }, [data, assetsList]);

  const syncStateToDb = useCallback(async (updatedData, updatedAssets) => {
    const operation = syncQueueRef.current.catch(() => undefined).then(async () => {
      let lastError;
      for (let attempt = 1; attempt <= 2; attempt += 1) {
        const controller = new AbortController();
        const timeout = window.setTimeout(() => controller.abort(), 30000);
        try {
          const response = await fetch(`${API_BASE}/api/sync-all`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ data: updatedData, assetsList: updatedAssets }),
            signal: controller.signal
          });
          const result = await response.json().catch(() => ({}));
          if (response.ok) return result;
          lastError = new Error(result.error || `API server returned ${response.status}`);
          lastError.status = response.status;
          if (![500, 502, 503].includes(response.status)) break;
        } catch (error) {
          lastError = error;
          if (error.name === 'AbortError') break;
        } finally {
          window.clearTimeout(timeout);
        }
        if (attempt < 2) await new Promise(resolve => window.setTimeout(resolve, 1200));
      }
      throw lastError || new Error('เนเธกเนเธชเธฒเธกเธฒเธฃเธ–เน€เธเธทเนเธญเธกเธ•เนเธญเธเธฒเธเธเนเธญเธกเธนเธฅเนเธ”เน');
    });
    syncQueueRef.current = operation;
    try {
      return await operation;
    } catch (err) {
      console.error('Failed to sync state to PostgreSQL database:', err);
      if (err.name === 'AbortError') throw new Error('เน€เธเธดเธฃเนเธเน€เธงเธญเธฃเนเนเธเนเน€เธงเธฅเธฒเธเธฑเธเธ—เธถเธเธเธฒเธเน€เธเธดเธ 30 เธงเธดเธเธฒเธ—เธต เธเธฃเธธเธ“เธฒเธฅเธญเธเนเธซเธกเน');
      throw err;
    }
  }, []);

  // Load state from Render PostgreSQL database on mount
  useEffect(() => {
    async function loadDbState() {
      try {
        const res = await fetch(`${API_BASE}/api/db-state`, { cache: 'no-store' });
        if (!res.ok) throw new Error('API server returned error');
        const result = await res.json();
        
        if (result.data && Object.keys(result.data).length > 0) {
          isPollingUpdateRef.current = true;
          setData(repairThaiTextDeep(result.data));
          if (result.assetsList) {
            setAssetsList(repairThaiTextDeep(result.assetsList));
          }
          console.log('Successfully synced dashboard state with Render PostgreSQL database.');
        } else {
          // Database is empty. Seed it with the default initial data!
          console.log('PostgreSQL database is empty. Seeding initial baseline datasets...');
          await fetch(`${API_BASE}/api/sync-all`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ data: latestDataRef.current, assetsList: latestAssetsRef.current })
          });
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
        const pendingSyncRaw = localStorage.getItem('it_dashboard_pending_sync');
        if (pendingSyncRaw) {
          const pendingSync = JSON.parse(pendingSyncRaw);
          if (pendingSync?.type === 'asset_patch' && pendingSync?.sn && pendingSync?.payload) {
            isPendingSyncRef.current = true;
            const pendingResponse = await fetch(`${API_BASE}/api/assets/${pendingSync.sn}`, {
              method: 'PATCH',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(pendingSync.payload)
            });
            const pendingResult = await pendingResponse.json().catch(() => ({}));
            if (!pendingResponse.ok) throw new Error(pendingResult.error || `API ${pendingResponse.status}`);
            localStorage.removeItem('it_dashboard_pending_sync');
            setConsoleSaveMessage('เธเธดเธเธเนเธซเธกเธฒเธขเน€เธซเธ•เธธเธ—เธตเนเธฃเธญเธเธฑเธเธ—เธถเธเธเธถเนเธเธเธฒเธเธเนเธญเธกเธนเธฅเธชเธณเน€เธฃเนเธเนเธฅเนเธง');
          }
        }

        const res = await fetch(`${API_BASE}/api/db-state`, { cache: 'no-store' });
        if (!res.ok) return;
        const result = await res.json();

        // Check again after fetch completes in case sync was initiated during network roundtrip
        if (isPendingSyncRef.current) return;

        const currentLocalData = latestDataRef.current;
        const currentLocalAssets = latestAssetsRef.current;

        const repairedData = repairThaiTextDeep(result.data);
        const repairedAssets = repairThaiTextDeep(result.assetsList);
        const hasDataChanged = JSON.stringify(repairedData) !== JSON.stringify(currentLocalData);
        const hasAssetsChanged = JSON.stringify(repairedAssets) !== JSON.stringify(currentLocalAssets);

        if (hasDataChanged || hasAssetsChanged) {
          isPollingUpdateRef.current = true;
          if (hasDataChanged) setData(repairedData);
          if (hasAssetsChanged) setAssetsList(repairedAssets);
          console.log('Real-time update synced from Render PostgreSQL.');
        }
      } catch (err) {
        console.warn('Real-time sync poll failed:', err.message);
      } finally {
        isPendingSyncRef.current = false;
      }
    }, 3000);

    return () => clearInterval(interval);
  }, [isLoaded]);

  // Read-only notification watcher. Existing records become the initial baseline;
  // only records created afterwards are announced, so historical data is untouched.
  useEffect(() => {
    if (!isLoaded) return undefined;

    let cancelled = false;
    const seenTicketKey = 'it_dashboard_seen_ticket_ids';
    const seenRequestKey = 'it_dashboard_seen_asset_request_ids';

    const readSeenIds = (key) => {
      try {
        return new Set(JSON.parse(localStorage.getItem(key) || '[]').map(String));
      } catch {
        return new Set();
      }
    };

    const seenTickets = readSeenIds(seenTicketKey);
    const seenRequests = readSeenIds(seenRequestKey);

    const saveSeenIds = (key, values) => {
      localStorage.setItem(key, JSON.stringify(Array.from(values).slice(-1000)));
    };

    const announce = (item) => {
      if ('Notification' in window && window.Notification.permission === 'granted') {
        new window.Notification(item.title, {
          body: item.message,
          icon: lightItLogo,
          tag: item.key
        });
      }
    };

    const pollNewItems = async () => {
      try {
        const [dashboardResponse, requestsResponse] = await Promise.all([
          fetch(`${API_BASE}/api/db-state`, { cache: 'no-store' }),
          fetch(`${API_BASE}/api/asset-requests`, { cache: 'no-store' })
        ]);
        if (!dashboardResponse.ok || !requestsResponse.ok) return;

        const [dashboardResult, requestResult] = await Promise.all([
          dashboardResponse.json(),
          requestsResponse.json()
        ]);
        if (cancelled) return;

        const ticketRows = Object.values(dashboardResult.data || {})
          .flatMap(month => Array.isArray(month.ticketsList) ? month.ticketsList : []);
        const requestRows = Array.isArray(requestResult) ? requestResult : [];
        setAssetRequests(requestRows);

        if (!notificationInitializedRef.current && seenTickets.size === 0 && seenRequests.size === 0) {
          ticketRows.forEach(ticket => seenTickets.add(String(ticket.sn)));
          requestRows.forEach(request => seenRequests.add(String(request.id)));
          saveSeenIds(seenTicketKey, seenTickets);
          saveSeenIds(seenRequestKey, seenRequests);
          notificationInitializedRef.current = true;
          return;
        }

        ticketRows.forEach(ticket => {
          const id = String(ticket.sn);
          if (seenTickets.has(id)) return;
          seenTickets.add(id);
          announce({
            key: `ticket-${id}`,
            type: 'ticket',
            title: `เนเธเนเธ Ticket เนเธซเธกเน #${id}`,
            message: `${ticket.complainant || 'เธเธนเนเนเธเนเธเนเธกเนเธฃเธฐเธเธธเธเธทเนเธญ'}: ${ticket.issue || 'เนเธกเนเธกเธตเธฃเธฒเธขเธฅเธฐเน€เธญเธตเธขเธ”'}`,
            createdAt: new Date().toISOString()
          });
        });

        requestRows.forEach(request => {
          const id = String(request.id);
          if (seenRequests.has(id)) return;
          seenRequests.add(id);
          announce({
            key: `asset-request-${id}`,
            type: 'asset',
            title: `เธเธณเธเธญเน€เธเธดเธเธญเธธเธเธเธฃเธ“เนเนเธซเธกเน #${id}`,
            message: `${request.requester || 'เธเธนเนเธเธญเนเธกเนเธฃเธฐเธเธธเธเธทเนเธญ'} เธเธญ ${request.item_type || 'เธญเธธเธเธเธฃเธ“เน IT'}`,
            createdAt: new Date().toISOString()
          });
        });

        saveSeenIds(seenTicketKey, seenTickets);
        saveSeenIds(seenRequestKey, seenRequests);
        notificationInitializedRef.current = true;
      } catch (error) {
        console.warn('Notification polling failed:', error.message);
      }
    };

    pollNewItems();
    const notificationTimer = window.setInterval(pollNewItems, 5000);
    return () => {
      cancelled = true;
      window.clearInterval(notificationTimer);
    };
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
  }, [data, assetsList, isLoaded, syncStateToDb]);
  
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
  const [newAssetAdditionalEquipment, setNewAssetAdditionalEquipment] = useState('');
  const [newAssetSerial, setNewAssetSerial] = useState('');
  const [newAssetStatus, setNewAssetStatus] = useState('เนเธเนเธเธฒเธ');
  const [newAssetNotes, setNewAssetNotes] = useState('');
  const [newAssetSubmittedOn, setNewAssetSubmittedOn] = useState('');
  const [newAssetRespondent, setNewAssetRespondent] = useState('');
  const [newAssetDate, setNewAssetDate] = useState('');
  const [newAssetSoftwareApp, setNewAssetSoftwareApp] = useState('');
  const [newAssetRegisteredEmail, setNewAssetRegisteredEmail] = useState('');
  const [newAssetAdditionalSerial, setNewAssetAdditionalSerial] = useState('');
  const [newAssetReturnDueDate, setNewAssetReturnDueDate] = useState('');
  const [newAssetAuditDate, setNewAssetAuditDate] = useState('');
  const [newAssetPurchaseDate, setNewAssetPurchaseDate] = useState('');
  const [newAssetWarrantyExpiry, setNewAssetWarrantyExpiry] = useState('');
  const [newAssetCost, setNewAssetCost] = useState('');


  // New Ticket creation states
  const [newTicketComplainant, setNewTicketComplainant] = useState('');
  const [newTicketEmail, setNewTicketEmail] = useState('');
  const [newTicketAnydesk, setNewTicketAnydesk] = useState('');
  const [newTicketIssue, setNewTicketIssue] = useState('');
  const [newTicketCause, setNewTicketCause] = useState('');
  const [newTicketDuration, setNewTicketDuration] = useState('00:30');
  const [newTicketResponder, setNewTicketResponder] = useState('');
  const [newTicketStatus, setNewTicketStatus] = useState('เน€เธชเธฃเนเธเธชเธดเนเธ');
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
      if (cost > 0 || ticket.status === 'เธเนเธฒเธขเน€เธเธดเธเนเธฅเนเธง') {
        repairCount++;
      }

      // Track broken devices by category
      const issue = String(ticket.issue).toLowerCase();
      let matchedDevice = 'เธญเธทเนเธ เน';
      if (issue.includes('notebook') || issue.includes('lenovo') || issue.includes('asus') || issue.includes('hp')) matchedDevice = 'Notebook';
      else if (issue.includes('computer') || issue.includes('pc')) matchedDevice = 'PC';
      else if (issue.includes('ipad')) matchedDevice = 'iPad';
      else if (issue.includes('iphone')) matchedDevice = 'iPhone';
      else if (issue.includes('printer') || issue.includes('เธเธฃเธดเนเธเน€เธ•เธญเธฃเน')) matchedDevice = 'Printer';
      else if (issue.includes('mornitor') || issue.includes('เธเธญ')) matchedDevice = 'Monitor';
      else if (issue.includes('imac')) matchedDevice = 'iMac';
      else if (issue.includes('macbook')) matchedDevice = 'MacBook';
      else if (issue.includes('network') || issue.includes('lan') || issue.includes('wifi') || issue.includes('เน€เธเนเธ•')) matchedDevice = 'Network';

      if (ticket.status !== 'เน€เธชเธฃเนเธเธชเธดเนเธ' || cost > 0) {
        deviceCounts[matchedDevice] = (deviceCounts[matchedDevice] || 0) + 1;
      }

      // Department costs mapping
      const borrowerAsset = assets.find(a => String(a.user).trim() === String(ticket.complainant).trim());
      const dept = borrowerAsset ? borrowerAsset.position : 'เธชเนเธงเธเธเธฅเธฒเธ';
      if (cost > 0) {
        deptCosts[dept] = (deptCosts[dept] || 0) + cost;
      }
    });

    const calculatedSla = durationCount > 0 ? Math.round((slaCompliantCount / durationCount) * 1000) / 10 : 100;
    const resolutionTimeHours = durationCount > 0 ? Number((durationSum / durationCount / 60).toFixed(1)) : 0.5;
    const calculatedResponseTime = Math.max(5, Math.round(resolutionTimeHours * 12));
    const calculatedCsat = Number((4.5 + (calculatedSla / 100) * 0.4).toFixed(1));

    const brokenAssetsCount = assets.filter(a => a.status === 'เธฃเธญเธเนเธญเธก').length;
    const lostAssetsCount = assets.filter(a => a.status === 'เธชเธนเธเธซเธฒเธข').length;
    const vacantAssetsCount = assets.filter(a => a.status === 'เธงเนเธฒเธ').length;

    const topBrokenDevices = Object.entries(deviceCounts)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);

    const totalAssetsVal = assets.length;
    return {
      totalAssets: totalAssetsVal,
      assetsBroken: brokenAssetsCount,
      assetsLost: lostAssetsCount,
      assetsVacant: vacantAssetsCount,
      ticketsCount: tickets.length,
      slaPercent: calculatedSla,
      resolutionTime: resolutionTimeHours,
      responseTime: calculatedResponseTime,
      csat: calculatedCsat,
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
      alert('เธเธฃเธธเธ“เธฒเธเธฃเธญเธเธเนเธญเธกเธนเธฅเนเธซเนเธเธฃเธเธ–เนเธงเธ');
      return;
    }
    if (data[newMonthKey]) {
      alert('เธกเธตเธฃเธซเธฑเธชเน€เธ”เธทเธญเธเธเธตเนเนเธเธฃเธฐเธเธเธญเธขเธนเนเนเธฅเนเธง');
      return;
    }
    const latestMonthKey = Object.keys(data).sort((a, b) => b.localeCompare(a))[0];
    const baseData = latestMonthKey ? data[latestMonthKey] : {};

    const newMonthData = {
      ...baseData,
      monthName: newMonthName,
      totalAssets: assetsList.length,
      assetValue: baseData.assetValue || 0,
      
      // Reset monthly operational metrics
      ticketsCount: 0,
      slaPercent: 100,
      responseTime: 0,
      resolutionTime: 0,
      csat: 5.0,
      repairCount: 0,
      repairCost: 0,
      securityIncidents: 0,
      // If there was no baseData, provide some defaults for arrays
      topBrokenDevices: baseData.topBrokenDevices || [],
      deptCosts: baseData.deptCosts || {},
      softwareExpiringDetails: baseData.softwareExpiringDetails || [],
      assetsExpiringDetails: baseData.assetsExpiringDetails || [],
      ongoingProjects: baseData.ongoingProjects || [],
      recommendations: baseData.recommendations || [],
      ticketsList: [] // DO NOT copy tickets to avoid duplicate SN conflict in DB
    };

    const updatedData = {
      ...data,
      [newMonthKey]: newMonthData
    };
    
    setData(updatedData);
    setConsoleMonth(newMonthKey);
    setCurrentMonth(newMonthKey);
    setNewMonthKey('');
    setNewMonthName('');
    
    // Force immediate sync so it doesn't get lost on quick refresh
    syncStateToDb(updatedData, assetsList).catch(err => console.error(err));
    
    alert(`เน€เธเธดเนเธกเน€เธ”เธทเธญเธ ${newMonthName} เธชเธณเน€เธฃเนเธ!`);
  };

  const handleDeleteMonth = (key) => {
    const keys = Object.keys(data);
    if (keys.length <= 1) {
      alert('เนเธกเนเธชเธฒเธกเธฒเธฃเธ–เธฅเธเน€เธ”เธทเธญเธเธชเธธเธ”เธ—เนเธฒเธขเธเธญเธเธฃเธฐเธเธเนเธ”เน');
      return;
    }
    if (window.confirm(`เธเธธเธ“เนเธเนเนเธเธงเนเธฒเธ•เนเธญเธเธเธฒเธฃเธฅเธเน€เธ”เธทเธญเธ ${data[key].monthName} เนเธเนเธซเธฃเธทเธญเนเธกเน?`)) {
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
      alert('เธเธฃเธธเธ“เธฒเธเธฃเธญเธเธเธฃเธฐเน€เธ เธ—เธญเธธเธเธเธฃเธ“เนเธซเธฅเธฑเธ');
      return;
    }

    if (editingAssetSn !== null) {
      // Edit mode
      setAssetsList(prev => {
        const updated = prev.map(a => a.sn === editingAssetSn ? {
          ...a,
          user: newAssetUser || 'เธชเนเธงเธเธเธฅเธฒเธ',
          position: newAssetPosition || '-',
          itemType: newAssetItemType,
          additionalEquipment: newAssetAdditionalEquipment,
          deviceSerial: newAssetSerial || '-',
          status: newAssetStatus,
          notes: newAssetNotes,
          submittedOn: newAssetSubmittedOn,
          respondent: newAssetRespondent,
          date: newAssetDate,
          softwareApp: newAssetSoftwareApp,
          registeredEmail: newAssetRegisteredEmail,
          additionalSerial: newAssetAdditionalSerial,
          returnDueDate: newAssetReturnDueDate,
          auditDate: newAssetAuditDate,
          purchaseDate: newAssetPurchaseDate,
          warrantyExpiry: newAssetWarrantyExpiry,
          cost: newAssetCost
        } : a);
        runRecalculation(consoleMonth, data[consoleMonth]?.ticketsList || [], updated);
        return updated;
      });
      setEditingAssetSn(null);
      alert('เนเธเนเนเธเธเนเธญเธกเธนเธฅเธ—เธฃเธฑเธเธขเนเธชเธดเธเธชเธณเน€เธฃเนเธ!');
    } else {
      // Create mode
      const newAsset = {
        sn: assetsList.length > 0 ? Math.max(...assetsList.map(a => Number(a.sn) || 0)) + 1 : 1,
        user: newAssetUser || 'เธชเนเธงเธเธเธฅเธฒเธ',
        position: newAssetPosition || '-',
        itemType: newAssetItemType,
        additionalEquipment: newAssetAdditionalEquipment,
        deviceSerial: newAssetSerial || '-',
        status: newAssetStatus,
        notes: newAssetNotes,
        submittedOn: newAssetSubmittedOn,
        respondent: newAssetRespondent,
        date: newAssetDate || new Date().toLocaleDateString('th-TH'),
        softwareApp: newAssetSoftwareApp,
        registeredEmail: newAssetRegisteredEmail,
        additionalSerial: newAssetAdditionalSerial,
        returnDueDate: newAssetReturnDueDate,
        auditDate: newAssetAuditDate,
        purchaseDate: newAssetPurchaseDate,
        warrantyExpiry: newAssetWarrantyExpiry,
        cost: newAssetCost
      };
      setAssetsList(prev => {
        const updated = [...prev, newAsset];
        runRecalculation(consoleMonth, data[consoleMonth]?.ticketsList || [], updated);
        return updated;
      });
      alert('เน€เธเธดเนเธกเธ—เธฃเธฑเธเธขเนเธชเธดเธเน€เธเนเธฒเธเธฅเธฑเธเธชเธณเน€เธฃเนเธ!');
    }

    setNewAssetUser('');
    setNewAssetPosition('');
    setNewAssetItemType('');
    setNewAssetAdditionalEquipment('');
    setNewAssetSerial('');
    setNewAssetStatus('เนเธเนเธเธฒเธ');
    setNewAssetNotes('');
    setNewAssetSubmittedOn('');
    setNewAssetRespondent('');
    setNewAssetDate('');
    setNewAssetSoftwareApp('');
    setNewAssetRegisteredEmail('');
    setNewAssetAdditionalSerial('');
    setNewAssetReturnDueDate('');
    setNewAssetAuditDate('');
    setNewAssetPurchaseDate('');
    setNewAssetWarrantyExpiry('');
    setNewAssetCost('');
    setEditingAssetTagField(null);
  };

  const handleLoadEditAsset = (asset) => {
    setEditingAssetSn(asset.sn);
    setNewAssetUser(asset.user);
    setNewAssetPosition(asset.position);
    setNewAssetItemType(asset.itemType);
    setNewAssetAdditionalEquipment(asset.additionalEquipment || '');
    setNewAssetSerial(asset.deviceSerial || '');
    setNewAssetStatus(asset.status || 'เนเธเนเธเธฒเธ');
    setNewAssetNotes(asset.notes || '');
    setNewAssetSubmittedOn(asset.submittedOn || '');
    setNewAssetRespondent(asset.respondent || '');
    setNewAssetDate(asset.date || '');
    setNewAssetSoftwareApp(asset.softwareApp || '');
    setNewAssetRegisteredEmail(asset.registeredEmail || '');
    setNewAssetAdditionalSerial(asset.additionalSerial || '');
    setNewAssetReturnDueDate(asset.returnDueDate || '');
    setNewAssetAuditDate(asset.auditDate || '');
    setNewAssetPurchaseDate(asset.purchaseDate || '');
    setNewAssetWarrantyExpiry(asset.warrantyExpiry || '');
    setNewAssetCost(asset.cost || '');
  };

  const handleCancelEditAsset = () => {
    setEditingAssetSn(null);
    setEditingAssetTagField(null);
    setNewAssetUser('');
    setNewAssetPosition('');
    setNewAssetItemType('');
    setNewAssetAdditionalEquipment('');
    setNewAssetSerial('');
    setNewAssetStatus('เนเธเนเธเธฒเธ');
    setNewAssetNotes('');
    setNewAssetSubmittedOn('');
    setNewAssetRespondent('');
    setNewAssetDate('');
    setNewAssetSoftwareApp('');
    setNewAssetRegisteredEmail('');
    setNewAssetAdditionalSerial('');
    setNewAssetReturnDueDate('');
    setNewAssetAuditDate('');
    setNewAssetPurchaseDate('');
    setNewAssetWarrantyExpiry('');
    setNewAssetCost('');
  };

  const handleDeleteAsset = (sn) => {
    if (window.confirm('เธเธธเธ“เธ•เนเธญเธเธเธฒเธฃเธฅเธเธญเธธเธเธเธฃเธ“เนเธเธตเนเธญเธญเธเธเธฒเธเธ—เธฐเน€เธเธตเธขเธเธเธฅเธฑเธเนเธเนเธซเธฃเธทเธญเนเธกเน?')) {
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
      alert('เธเธฃเธธเธ“เธฒเธเธฃเธญเธเธญเธฒเธเธฒเธฃเน€เธชเธตเธข/เธเธฑเธเธซเธฒ');
      return;
    }

    const tickets = data[consoleMonth]?.ticketsList || [];

    if (editingTicketSn !== null) {
      // Edit mode
      const updatedTickets = tickets.map(t => t.sn === editingTicketSn ? {
        ...t,
        complainant: newTicketComplainant || 'เนเธกเนเธฃเธฐเธเธธเธเธทเนเธญ',
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
      alert('เนเธเนเนเธเธเนเธญเธกเธนเธฅเธเธฒเธเนเธเนเธเธเนเธญเธกเธชเธณเน€เธฃเนเธ!');
    } else {
      // Create mode
      const newTicket = {
        sn: tickets.length > 0 ? Math.max(...tickets.map(t => Number(t.sn) || 0)) + 1 : 1,
        date: new Date().toLocaleString('th-TH', { hour12: false }).replace(',', ''),
        complainant: newTicketComplainant || 'เนเธกเนเธฃเธฐเธเธธเธเธทเนเธญ',
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
      alert('เน€เธเธดเนเธกเธเธฃเธฐเธงเธฑเธ•เธดเธเธฒเธเนเธเนเธเธเนเธญเธกเธชเธณเน€เธฃเนเธ!');
    }

    setNewTicketComplainant('');
    setNewTicketEmail('');
    setNewTicketAnydesk('');
    setNewTicketIssue('');
    setNewTicketCause('');
    setNewTicketDuration('00:30');
    setNewTicketResponder('');
    setNewTicketStatus('เน€เธชเธฃเนเธเธชเธดเนเธ');
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
    setNewTicketStatus('เน€เธชเธฃเนเธเธชเธดเนเธ');
    setNewTicketCost('0');
  };

  const handleDeleteTicket = (sn) => {
    if (window.confirm('เธเธธเธ“เธ•เนเธญเธเธเธฒเธฃเธฅเธเธฃเธฒเธขเธเธฒเธฃเนเธเนเธเธเนเธญเธกเธเธตเนเนเธเนเธซเธฃเธทเธญเนเธกเน?')) {
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
          alert('เธเธณเน€เธเนเธฒเธเนเธญเธกเธนเธฅเธชเธณเธฃเธญเธเธชเธณเน€เธฃเนเธ!');
        } else {
          alert('เธฃเธนเธเนเธเธเนเธเธฅเนเธชเธณเธฃเธญเธเนเธกเนเธ–เธนเธเธ•เนเธญเธ');
        }
      } catch (err) {
        alert('เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”เนเธเธเธฒเธฃเธญเนเธฒเธเนเธเธฅเน: ' + err.message);
      }
    };
    reader.readAsText(file);
  };

  const handleResetToDefault = async () => {
    if (window.confirm('เธเธณเน€เธ•เธทเธญเธ: เธเธธเธ“เธ•เนเธญเธเธเธฒเธฃเธฅเธเธเนเธญเธกเธนเธฅเธ—เธฑเนเธเธซเธกเธ”เนเธเธฃเธฐเธเธเนเธเนเธซเธฃเธทเธญเนเธกเน? เธเนเธญเธกเธนเธฅเธ—เธฑเนเธเธซเธกเธ”เธ—เธตเนเธเธฑเธเธ—เธถเธเนเธงเนเนเธเธเธฒเธเธเนเธญเธกเธนเธฅเธเธฐเธ–เธนเธเธฅเนเธฒเธเนเธฅเธฐเนเธกเนเธชเธฒเธกเธฒเธฃเธ–เธเธนเนเธเธทเธเนเธ”เน!')) {
      try {
        const response = await fetch(`${API_BASE}/api/reset-database`, {
          method: 'POST'
        });
        if (!response.ok) {
          throw new Error('เนเธกเนเธชเธฒเธกเธฒเธฃเธ–เธฅเธเธเนเธญเธกเธนเธฅเธเธฒเธเธเนเธญเธกเธนเธฅเนเธ”เน');
        }
        localStorage.removeItem('it_dashboard_data');
        localStorage.removeItem('it_dashboard_assets');
        setData([]);
        setAssetsList([]);
        setCurrentMonth('2026-07');
        setConsoleMonth('2026-07');
        setEditingAssetSn(null);
        setEditingTicketSn(null);
        alert('เธฅเธเธเนเธญเธกเธนเธฅเธ—เธฑเนเธเธซเธกเธ”เน€เธฃเธตเธขเธเธฃเนเธญเธขเนเธฅเนเธง เธฃเธฐเธเธเธงเนเธฒเธเน€เธเธฅเนเธฒเน€เธซเธกเธทเธญเธเน€เธฃเธดเนเธกเธ•เนเธเนเธซเธกเน');
        window.location.reload();
      } catch (error) {
        console.error('Error resetting database:', error);
        alert('เน€เธเธดเธ”เธเนเธญเธเธดเธ”เธเธฅเธฒเธ”เนเธเธเธฒเธฃเธฅเธเธเนเธญเธกเธนเธฅ: ' + error.message);
      }
    }
  };

  const handleLarkSubmit = async (e) => {
    e.preventDefault();
    if (larkFormType === 'ticket') {
      if (larkTicketRole === 'it') {
        if (!selectedPendingTicketSn) {
          alert('เธเธฃเธธเธ“เธฒเน€เธฅเธทเธญเธเนเธเธเธฒเธเธ—เธตเนเธ•เนเธญเธเธเธฒเธฃเธเธดเธ”เธเธฒเธ');
          return;
        }
        if (!larkTicketResponder) {
          alert('เธเธฃเธธเธ“เธฒเธเธฃเธญเธเธเธทเนเธญเธเธนเนเธ”เธณเน€เธเธดเธเธเธฒเธ (เธเนเธฒเธ IT)');
          return;
        }
        
        // IT Close work mode - persist through the API before removing it from the queue.
        try {
          const response = await fetch(`${API_BASE}/api/tickets/${selectedPendingTicketSn}/close`, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              responder: larkTicketResponder,
              duration: larkTicketDuration || '00:30',
              cause: larkTicketCause || '-',
              cost: Number(larkTicketCost) || 0,
              status: larkTicketStatus
            })
          });
          const result = await response.json();
          if (!response.ok) throw new Error(result.error || 'เธเธดเธ”เธเธฒเธเนเธกเนเธชเธณเน€เธฃเนเธ');
          await refreshOperationalStateFromDb();
        } catch (error) {
          alert(error.message);
          return;
        }

        setSelectedPendingTicketSn('');
        setLarkTicketResponder('');
        setLarkTicketDuration('00:30');
        setLarkTicketCause('');
        setLarkTicketCost('0');
        setLarkTicketStatus('เน€เธชเธฃเนเธเธชเธดเนเธ');
        setLarkSubmitted(true);
      } else {
        // User Submit Mode
        if (!larkTicketIssue) {
          alert('เธเธฃเธธเธ“เธฒเธเธฃเธญเธเธญเธฒเธเธฒเธฃเน€เธชเธตเธข/เธเธฑเธเธซเธฒ');
          return;
        }
        const linkedAsset = larkTicketAssetSerial
          ? assetsList.find(asset => String(asset.deviceSerial || '').toLocaleLowerCase('th-TH') === larkTicketAssetSerial.trim().toLocaleLowerCase('th-TH'))
          : null;
        if (larkTicketAssetSerial && !linkedAsset) {
          alert('เนเธกเนเธเธเธซเธกเธฒเธขเน€เธฅเธเน€เธเธฃเธทเนเธญเธเธเธตเนเนเธเธ—เธฐเน€เธเธตเธขเธเธ—เธฃเธฑเธเธขเนเธชเธดเธ');
          return;
        }
        try {
          const response = await fetch(`${API_BASE}/api/tickets`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              name: larkTicketComplainant || 'เนเธกเนเธฃเธฐเธเธธเธเธทเนเธญ',
              department: linkedAsset?.position || 'เนเธกเนเธฃเธฐเธเธธ',
              date: `${currentMonth}-01`,
              deviceType: linkedAsset?.itemType || 'Other',
              assetSerial: larkTicketAssetSerial.trim(),
              issue: larkTicketIssue,
              priority: 'medium',
              email: larkTicketEmail,
              anydesk: larkTicketAnydesk
            })
          });
          const result = await response.json();
          if (!response.ok) throw new Error(result.error || 'เธเธฑเธเธ—เธถเธเธเธณเธฃเนเธญเธเนเธกเนเธชเธณเน€เธฃเนเธ');
          await Promise.all([refreshOperationalStateFromDb(), loadAssetRequests()]);
        } catch (error) {
          alert(error.message);
          return;
        }

        // Clear inputs
        setLarkTicketComplainant('');
        setLarkTicketEmail('');
        setLarkTicketAnydesk('');
        setLarkTicketAssetSerial('');
        setLarkTicketIssue('');
        setLarkSubmitted(true);
      }
    } else {
      if (!larkAssetItemType) {
        alert('เธเธฃเธธเธ“เธฒเธเธฃเธญเธเธเธฃเธฐเน€เธ เธ—เธญเธธเธเธเธฃเธ“เนเธซเธฅเธฑเธ');
        return;
      }
      const newAsset = {
        sn: assetsList.length > 0 ? Math.max(...assetsList.map(a => Number(a.sn) || 0)) + 1 : 1,
        date: new Date().toLocaleDateString('th-TH'),
        user: larkAssetUser || 'เธชเนเธงเธเธเธฅเธฒเธ',
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
      setLarkAssetStatus('เนเธเนเธเธฒเธ');
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
  const activeData = useMemo(() => ({
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
  }), [activeDataSource, fallbackMonthKey]);

  const saveConsoleChanges = async () => {
    setConsoleSaving(true);
    setConsoleSaveMessage('');
    isPendingSyncRef.current = true;
    let assetsToSave = assetsList;
    let dataToSave = data;
    let pendingAssetPatch = null;
    try {
      // If an asset is still open in edit mode, commit the visible form values
      // before sending the complete dashboard snapshot.
      if (editingAssetSn !== null && newAssetItemType) {
        pendingAssetPatch = {
          user: newAssetUser,
          position: newAssetPosition,
          itemType: newAssetItemType,
          additionalEquipment: newAssetAdditionalEquipment,
          deviceSerial: newAssetSerial,
          status: newAssetStatus,
          notes: newAssetNotes
        };
        assetsToSave = assetsList.map(asset => asset.sn === editingAssetSn ? {
          ...asset,
          user: newAssetUser || 'เธชเนเธงเธเธเธฅเธฒเธ',
          position: newAssetPosition || '-',
          itemType: newAssetItemType,
          additionalEquipment: newAssetAdditionalEquipment,
          deviceSerial: newAssetSerial || '-',
          status: newAssetStatus,
          notes: newAssetNotes,
          submittedOn: newAssetSubmittedOn,
          respondent: newAssetRespondent,
          date: newAssetDate,
          softwareApp: newAssetSoftwareApp,
          registeredEmail: newAssetRegisteredEmail,
          additionalSerial: newAssetAdditionalSerial,
          returnDueDate: newAssetReturnDueDate,
          auditDate: newAssetAuditDate,
          purchaseDate: newAssetPurchaseDate,
          warrantyExpiry: newAssetWarrantyExpiry,
          cost: newAssetCost
        } : asset);

        dataToSave = {
          ...data,
          [consoleMonth]: {
            ...data[consoleMonth],
            totalAssets: assetsToSave.length,
            assetsBroken: assetsToSave.filter(asset => asset.status === 'เธฃเธญเธเนเธญเธก').length,
            assetsLost: assetsToSave.filter(asset => asset.status === 'เธชเธนเธเธซเธฒเธข').length,
            assetsVacant: assetsToSave.filter(asset => asset.status === 'เธงเนเธฒเธ').length
          }
        };
      }

      await syncStateToDb(dataToSave, assetsToSave);
      if (editingAssetSn !== null && newAssetItemType) {
        const assetResponse = await fetch(`${API_BASE}/api/assets/${editingAssetSn}`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            user: newAssetUser,
            position: newAssetPosition,
            itemType: newAssetItemType,
            additionalEquipment: newAssetAdditionalEquipment,
            deviceSerial: newAssetSerial,
            status: newAssetStatus,
            notes: newAssetNotes
          })
        });
        const assetResult = await assetResponse.json();
        if (!assetResponse.ok) throw new Error(assetResult.error || 'เนเธเนเนเธเธชเธ–เธฒเธเธฐเธ—เธฃเธฑเธเธขเนเธชเธดเธเนเธกเนเธชเธณเน€เธฃเนเธ');
      }
      const response = await fetch(`${API_BASE}/api/db-state`, { cache: 'no-store' });
      if (!response.ok) throw new Error(`API server returned ${response.status}`);
      const synchronized = await response.json();
      if (synchronized.data) setData(synchronized.data);
      if (synchronized.assetsList) setAssetsList(synchronized.assetsList);
      setCurrentMonth(consoleMonth);
      if (editingAssetSn !== null) {
        setEditingAssetSn(null);
        setEditingAssetTagField(null);
        setNewAssetUser('');
        setNewAssetPosition('');
        setNewAssetItemType('');
        setNewAssetAdditionalEquipment('');
        setNewAssetSerial('');
        setNewAssetStatus('เนเธเนเธเธฒเธ');
        setNewAssetNotes('');
        setNewAssetSubmittedOn('');
        setNewAssetRespondent('');
        setNewAssetDate('');
        setNewAssetSoftwareApp('');
        setNewAssetRegisteredEmail('');
        setNewAssetAdditionalSerial('');
        setNewAssetReturnDueDate('');
        setNewAssetAuditDate('');
        setNewAssetPurchaseDate('');
        setNewAssetWarrantyExpiry('');
        setNewAssetCost('');
      }
      setConsoleSaveMessage('เธเธฑเธเธ—เธถเธเธชเธณเน€เธฃเนเธเนเธฅเธฐเธญเธฑเธเน€เธ”เธ•เนเธ”เธเธเธญเธฃเนเธ”เนเธฅเนเธง');
    } catch (error) {
      console.error('Failed to save console changes:', error);
      const databaseUnavailable = [500, 502, 503].includes(Number(error?.status)) ||
        /500|502|503|internal server error|database|connection|เน€เธเธทเนเธญเธกเธ•เนเธญ/i.test(String(error?.message || ''));
      if (databaseUnavailable && editingAssetSn !== null && pendingAssetPatch) {
        localStorage.setItem('it_dashboard_pending_sync', JSON.stringify({
          type: 'asset_patch',
          sn: editingAssetSn,
          payload: pendingAssetPatch,
          savedAt: new Date().toISOString()
        }));
        setData(dataToSave);
        setAssetsList(assetsToSave);
        if (editingAssetSn !== null) handleCancelEditAsset();
        setConsoleSaveMessage('เธเธฑเธเธ—เธถเธเธซเธกเธฒเธขเน€เธซเธ•เธธเนเธงเนเนเธเน€เธเธฃเธทเนเธญเธเนเธฅเนเธง เนเธฅเธฐเธเธฐเธเธดเธเธเนเธเธถเนเธเธเธฒเธเธเนเธญเธกเธนเธฅเธญเธฑเธ•เนเธเธกเธฑเธ•เธดเน€เธกเธทเนเธญเธฃเธฐเธเธเธเธฃเนเธญเธก');
      } else if (databaseUnavailable) {
        setConsoleSaveMessage('เธขเธฑเธเนเธกเนเธเธฑเธเธ—เธถเธ: เธเธฒเธเธเนเธญเธกเธนเธฅเนเธกเนเธเธฃเนเธญเธกเนเธเนเธเธฒเธ เธเธฃเธธเธ“เธฒเธ•เธฃเธงเธเธชเธญเธ Render PostgreSQL เนเธฅเนเธงเธฅเธญเธเนเธซเธกเน');
      } else {
        setConsoleSaveMessage(`เธเธฑเธเธ—เธถเธเนเธกเนเธชเธณเน€เธฃเนเธ: ${error?.message || 'เธเธฃเธธเธ“เธฒเธฅเธญเธเนเธซเธกเน'}`);
      }
    } finally {
      isPendingSyncRef.current = false;
      setConsoleSaving(false);
    }
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

  const vacantStockAssets = assetsList.filter((asset) => asset.status === 'เธงเนเธฒเธ');
  const vacantStockCount = vacantStockAssets.length;
  const vacantStockBreakdown = Array.from(vacantStockAssets.reduce((groups, asset) => {
    const type = String(asset.itemType || 'เนเธกเนเธฃเธฐเธเธธ').trim() || 'เนเธกเนเธฃเธฐเธเธธ';
    const mainCategory = mainAssetCategories.find((category) => category.match(type));
    const label = mainCategory?.label || type;
    groups.set(label, (groups.get(label) || 0) + 1);
    return groups;
  }, new Map()), ([label, count]) => ({ label, count }))
    .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, 'th'));

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
    if (!category || !startDate || asset.status === 'เธชเธนเธเธซเธฒเธข') return false;

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
      status: 'เนเธเนเธเธฒเธ',
      isLicenseRecord: true,
    };
    if (!license.name) return;

    setData((previous) => {
      const monthData = previous[currentMonth] || { ...initialDashboardData['2026-07'] };
      const rows = [...(monthData.softwareExpiringDetails || [])];
      if (editingSoftwareIndex === null) rows.push(license);
      else rows[editingSoftwareIndex] = license;
      return {
        ...previous,
        [currentMonth]: { ...monthData, softwareExpiringDetails: rows },
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
    if (!window.confirm('เธ•เนเธญเธเธเธฒเธฃเธฅเธเธฃเธฒเธขเธฅเธฐเน€เธญเธตเธขเธ” License เธเธตเนเนเธเนเธซเธฃเธทเธญเนเธกเน?')) return;
    setData((previous) => {
      const monthData = previous[currentMonth] || { ...initialDashboardData['2026-07'] };
      const currentList = monthData.softwareExpiringDetails || [];
      return {
        ...previous,
        [currentMonth]: {
          ...monthData,
          softwareExpiringDetails: currentList.filter((_, rowIndex) => rowIndex !== index),
        },
      };
    });
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
          labels: ['เธเธเธ•เธด', 'เนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ', 'เธเธณเธฃเธธเธ”', 'เธชเธนเธเธซเธฒเธข'],
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
          labels: ['เธชเธดเธ—เธเธดเนเนเธเนเธเธฒเธ (Licenses)'],
          datasets: [
            {
              label: 'เนเธเนเธเธฒเธเธญเธขเธนเน (In Use)',
              data: [calculatedLicensesInUse],
              backgroundColor: 'rgba(59, 130, 246, 0.75)',
              borderColor: '#3b82f6',
              borderWidth: 1
            },
            {
              label: 'เธงเนเธฒเธ (Vacant)',
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
            label: 'เธเนเธฒเนเธเนเธเนเธฒเธขเธเนเธญเธก (เธเธฒเธ—)',
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
  }, [activeData, primaryExpiringAssets, calculatedLicensesInUse, calculatedLicensesVacant]);

  // ========================================
  // XLSX IMPORT / EXPORT / TEMPLATE FUNCTIONS
  // ========================================

  // Define the field mapping for the main "Dashboard" sheet
  const FIELD_MAP = [
    { key: 'monthName', header: 'เน€เธ”เธทเธญเธ (Month Name)', example: 'เธเธฃเธเธเธฒเธเธก 2026' },
    { key: 'monthKey', header: 'เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ (Month Key, เน€เธเนเธ 2026-07)', example: '2026-07' },
    // Asset
    { key: 'totalAssets', header: 'เธเธณเธเธงเธเธญเธธเธเธเธฃเธ“เนเธ—เธฑเนเธเธซเธกเธ”', example: 1450 },
    { key: 'assetValue', header: 'เธกเธนเธฅเธเนเธฒเธ—เธฃเธฑเธเธขเนเธชเธดเธ IT (เธเธฒเธ—)', example: 85200000 },
    { key: 'assetsExpiring', header: 'เธญเธธเธเธเธฃเธ“เนเนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ', example: 38 },
    { key: 'assetsBroken', header: 'เธญเธธเธเธเธฃเธ“เนเธเธณเธฃเธธเธ”', example: 8 },
    { key: 'assetsLost', header: 'เธญเธธเธเธเธฃเธ“เนเธชเธนเธเธซเธฒเธข', example: 1 },
    // Support
    { key: 'ticketsCount', header: 'เธเธณเธเธงเธ Ticket', example: 280 },
    { key: 'slaPercent', header: 'SLA Compliance (%)', example: 98.8 },
    { key: 'responseTime', header: 'Response Time เน€เธเธฅเธตเนเธข (เธเธฒเธ—เธต)', example: 8 },
    { key: 'resolutionTime', header: 'Resolution Time เน€เธเธฅเธตเนเธข (เธเธก.)', example: 1.8 },
    { key: 'csat', header: 'CSAT เธเธฐเนเธเธเธเธงเธฒเธกเธเธถเธเธเธญเนเธ (เธเธฒเธ 5)', example: 4.9 },
    // Software
    { key: 'totalSoftware', header: 'เนเธเธฃเนเธเธฃเธกเธ—เธฑเนเธเธซเธกเธ”', example: 45 },
    { key: 'licensesInUse', header: 'License เนเธเนเธเธฒเธ', example: 2450 },
    { key: 'licensesVacant', header: 'License เธงเนเธฒเธ', example: 350 },
    { key: 'softwareCost', header: 'เธเนเธฒเนเธเนเธเนเธฒเธข Software (เธเธฒเธ—/เน€เธ”เธทเธญเธ)', example: 1280000 },
    { key: 'softwareExpiring', header: 'เนเธเธฃเนเธเธฃเธกเนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ', example: 3 },
    // Security
    { key: 'backupSuccess', header: 'Backup เธชเธณเน€เธฃเนเธ (%)', example: 99.98 },
    { key: 'securityIncidents', header: 'Security Incident (เธเธฃเธฑเนเธ)', example: 0 },
    { key: 'antivirusCoverage', header: 'Antivirus Coverage (%)', example: 100 },
    { key: 'mfaCoverage', header: 'MFA Coverage (%)', example: 100 },
    // Repair
    { key: 'repairCount', header: 'เธเธณเธเธงเธเธเธฒเธฃเธเนเธญเธก', example: 12 },
    { key: 'repairCost', header: 'เธเนเธฒเนเธเนเธเนเธฒเธขเธเธฒเธฃเธเนเธญเธก (เธเธฒเธ—)', example: 145000 },
    // Improvement
    { key: 'automationsDone', header: 'Automation เธ—เธตเนเธ—เธณเน€เธชเธฃเนเธ', example: 5 },
    { key: 'aiApps', header: 'AI เธ—เธตเนเธเธณเธกเธฒเนเธเน', example: 4 },
    { key: 'hoursSaved', header: 'เธเธฑเนเธงเนเธกเธเธ—เธตเนเธฅเธ”เธฅเธเธเธฒเธ Automation', example: 320 },
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

    // Sheet 2: Top 10 เธญเธธเธเธเธฃเธ“เนเน€เธชเธตเธขเธเนเธญเธข
    const repairHeaders = ['เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ (Month Key)', 'เธเธทเนเธญเธญเธธเธเธเธฃเธ“เน', 'เธเธณเธเธงเธเธเธฃเธฑเนเธเธ—เธตเนเน€เธชเธตเธข', 'เธเนเธฒเนเธเนเธเนเธฒเธขเธเนเธญเธก (เธเธฒเธ—)'];
    const repairExample = ['2026-07', 'Google TPU v5e Node', 4, 80000];
    const repairWs = XLSX.utils.aoa_to_sheet([repairHeaders, repairExample]);
    repairWs['!cols'] = repairHeaders.map(h => ({ wch: Math.max(h.length + 4, 18) }));
    XLSX.utils.book_append_sheet(wb, repairWs, 'Top10 เธญเธธเธเธเธฃเธ“เนเน€เธชเธตเธขเธเนเธญเธข');

    // Sheet 3: เธเนเธฒเนเธเนเธเนเธฒเธขเธ•เนเธญเนเธเธเธ
    const deptHeaders = ['เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ (Month Key)', 'เธเธทเนเธญเนเธเธเธ', 'เธเนเธฒเนเธเนเธเนเธฒเธขเธเนเธญเธก (เธเธฒเธ—)'];
    const deptExample = ['2026-07', 'AI Research', 85000];
    const deptWs = XLSX.utils.aoa_to_sheet([deptHeaders, deptExample]);
    deptWs['!cols'] = deptHeaders.map(h => ({ wch: Math.max(h.length + 4, 18) }));
    XLSX.utils.book_append_sheet(wb, deptWs, 'เธเนเธฒเนเธเนเธเนเธฒเธขเธ•เนเธญเนเธเธเธ');

    // Sheet 4: เนเธเธฃเธเธเธฒเธฃเธ—เธตเนเธเธณเธฅเธฑเธเธ”เธณเน€เธเธดเธเธเธฒเธฃ
    const projHeaders = ['เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ (Month Key)', 'เธเธทเนเธญเนเธเธฃเธเธเธฒเธฃ', 'เธฃเธฒเธขเธฅเธฐเน€เธญเธตเธขเธ”เธเธงเธฒเธกเธเธทเธเธซเธเนเธฒ'];
    const projExample = ['2026-07', 'Gemini Auto-IT Agent', 'เธเธณเนเธกเน€เธ”เธฅ Gemini เธกเธฒเธเนเธงเธขเธ•เธญเธเนเธฅเธฐเนเธเนเธเธฑเธเธซเธฒเนเธญเธ—เธต เธเธทเธเธซเธเนเธฒ 85%'];
    const projWs = XLSX.utils.aoa_to_sheet([projHeaders, projExample]);
    projWs['!cols'] = projHeaders.map(h => ({ wch: Math.max(h.length + 4, 25) }));
    XLSX.utils.book_append_sheet(wb, projWs, 'เนเธเธฃเธเธเธฒเธฃเธ”เธณเน€เธเธดเธเธเธฒเธฃ');

    // Sheet 5: Recommendation
    const recHeaders = ['เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ (Month Key)', 'เธเนเธญเน€เธชเธเธญเนเธเธฐ'];
    const recExample = ['2026-07', 'เนเธเธฐเธเธณเธเธฑเธ”เธ—เธณเนเธเธเธเธเธเธฃเธฐเธกเธฒเธ“เน€เธเธทเนเธญเน€เธเธฅเธตเนเธขเธเธเนเธฒเธเธเธฒเธ TPU v4 Nodes'];
    const recWs = XLSX.utils.aoa_to_sheet([recHeaders, recExample]);
    recWs['!cols'] = recHeaders.map(h => ({ wch: Math.max(h.length + 4, 40) }));
    XLSX.utils.book_append_sheet(wb, recWs, 'Recommendation');

    // Sheet 6: เธญเธธเธเธเธฃเธ“เนเนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ (Expiring Assets Details)
    const expAssetHeaders = ['เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ (Month Key)', 'เธฃเธซเธฑเธชเธ—เธฃเธฑเธเธขเนเธชเธดเธ', 'เธเธฃเธฐเน€เธ เธ—', 'เธฃเธธเนเธ/เนเธกเน€เธ”เธฅ', 'เนเธเธเธ', 'เธงเธฑเธเธ—เธตเนเธซเธกเธ”เธญเธฒเธขเธธ'];
    const expAssetExample = ['2026-07', 'AST-TPU-042', 'Server Node', 'Google TPU v4 Node', 'AI Research', '10 เธช.เธ. 2026'];
    const expAssetWs = XLSX.utils.aoa_to_sheet([expAssetHeaders, expAssetExample]);
    expAssetWs['!cols'] = expAssetHeaders.map(h => ({ wch: Math.max(h.length + 4, 20) }));
    XLSX.utils.book_append_sheet(wb, expAssetWs, 'เธญเธธเธเธเธฃเธ“เนเนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ');

    // Sheet 7: เนเธเธฃเนเธเธฃเธกเนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ (Expiring Software Details)
    const expSwHeaders = ['เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ (Month Key)', 'เธเธทเนเธญเธเธญเธเธ•เนเนเธงเธฃเน', 'เธเธณเธเธงเธ Licenses', 'เธงเธฑเธเธซเธกเธ”เธชเธฑเธเธเธฒ', 'เธชเธ–เธฒเธเธฐ'];
    const expSwExample = ['2026-07', 'Google Cloud Platform', 500, '15 เธช.เธ. 2026', 'เนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ'];
    const expSwWs = XLSX.utils.aoa_to_sheet([expSwHeaders, expSwExample]);
    expSwWs['!cols'] = expSwHeaders.map(h => ({ wch: Math.max(h.length + 4, 22) }));
    XLSX.utils.book_append_sheet(wb, expSwWs, 'เนเธเธฃเนเธเธฃเธกเนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ');

    XLSX.writeFile(wb, 'IT_Dashboard_Template.xlsx');
  };

  const exportAssetsToExcel = (assetsToExport, monthToExport) => {
    if (!assetsToExport || assetsToExport.length === 0) {
      alert('เนเธกเนเธกเธตเธเนเธญเธกเธนเธฅเธชเธณเธซเธฃเธฑเธ Export');
      return;
    }
    const wb = XLSX.utils.book_new();
    const headers = ['Number', 'เธงเธฑเธเธ—เธตเน Submit', 'เธเธนเนเธฃเธฑเธเธเธดเธ”เธเธญเธ', 'เธงเธฑเธเธ—เธตเนเน€เธเธดเธเนเธเนเธเธฒเธ', 'เธเธนเนเน€เธเธดเธเนเธเนเธเธฒเธ', 'เธ•เธณเนเธซเธเนเธ/เนเธเธเธ', 'เธฃเธฒเธขเธเธฒเธฃเธญเธธเธเธเธฃเธ“เนเธซเธฅเธฑเธ', 'เธญเธธเธเธเธฃเธ“เนเน€เธเธดเนเธกเน€เธ•เธดเธก', 'เธเธญเธเธ•เนเนเธงเธฃเน/App', 'เธญเธตเน€เธกเธฅเธ—เธตเนเธฅเธเธ—เธฐเน€เธเธตเธขเธ', 'เธซเธกเธฒเธขเน€เธฅเธเธญเธธเธเธเธฃเธ“เน', 'เธซเธกเธฒเธขเน€เธฅเธเธญเธธเธเธเธฃเธ“เนเน€เธเธดเนเธกเน€เธ•เธดเธก', 'เธเธณเธซเธเธ”เธเธทเธ', 'เธชเธ–เธฒเธเธฐ', 'เธซเธกเธฒเธขเน€เธซเธ•เธธ', 'เธงเธฑเธเธ—เธตเนเธ•เธฃเธงเธเธชเธญเธ', 'เธงเธฑเธเธ—เธตเนเธเธทเนเธญ', 'เธงเธฑเธเธซเธกเธ”เธเธฃเธฐเธเธฑเธ', 'เธเนเธฒเนเธเนเธเนเธฒเธข'];
    
    const rows = assetsToExport.map((asset, idx) => [
      idx + 1,
      asset.submittedOn || '',
      asset.respondent || '',
      asset.date || '',
      asset.user || '',
      asset.position || '',
      asset.itemType || '',
      Array.isArray(asset.additionalEquipment) ? asset.additionalEquipment.join(', ') : (asset.additionalEquipment || ''),
      Array.isArray(asset.softwareApp) ? asset.softwareApp.join(', ') : (asset.softwareApp || ''),
      asset.registeredEmail || '',
      asset.deviceSerial || '',
      asset.additionalSerial || '',
      asset.returnDueDate || '',
      asset.status || '',
      asset.notes || '',
      asset.inspectionDate || '',
      asset.purchaseDate || '',
      asset.warrantyEndDate || '',
      asset.expense || ''
    ]);

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    ws['!cols'] = headers.map(h => ({ wch: Math.max(h.length + 4, 15) }));
    XLSX.utils.book_append_sheet(wb, ws, 'Asset_Registry');
    XLSX.writeFile(wb, `Asset_Registry_${monthToExport}.xlsx`);
  };

  const exportAssetsToPDF = () => {
    // We can use a simple window.print() and hide the form elements using CSS.
    // By adding a temporary class to body to indicate asset PDF export.
    document.body.classList.add('printing-assets');
    window.print();
    setTimeout(() => {
      document.body.classList.remove('printing-assets');
    }, 1000);
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
    const repairHeaders = ['เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ', 'เธเธทเนเธญเธญเธธเธเธเธฃเธ“เน', 'เธเธณเธเธงเธเธเธฃเธฑเนเธเธ—เธตเนเน€เธชเธตเธข', 'เธเนเธฒเนเธเนเธเนเธฒเธขเธเนเธญเธก (เธเธฒเธ—)'];
    const repairRows = [];
    Object.entries(data).forEach(([monthKey, d]) => {
      (d.topBrokenDevices || []).forEach(dev => {
        repairRows.push([monthKey, dev.name, dev.count, dev.cost]);
      });
    });
    const repairWs = XLSX.utils.aoa_to_sheet([repairHeaders, ...repairRows]);
    repairWs['!cols'] = repairHeaders.map(h => ({ wch: Math.max(h.length + 4, 18) }));
    XLSX.utils.book_append_sheet(wb, repairWs, 'Top10 เธญเธธเธเธเธฃเธ“เนเน€เธชเธตเธขเธเนเธญเธข');

    // Sheet 3: Department costs (all months)
    const deptHeaders = ['เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ', 'เธเธทเนเธญเนเธเธเธ', 'เธเนเธฒเนเธเนเธเนเธฒเธขเธเนเธญเธก (เธเธฒเธ—)'];
    const deptRows = [];
    Object.entries(data).forEach(([monthKey, d]) => {
      Object.entries(d.deptCosts || {}).forEach(([dept, cost]) => {
        deptRows.push([monthKey, dept, cost]);
      });
    });
    const deptWs = XLSX.utils.aoa_to_sheet([deptHeaders, ...deptRows]);
    deptWs['!cols'] = deptHeaders.map(h => ({ wch: Math.max(h.length + 4, 18) }));
    XLSX.utils.book_append_sheet(wb, deptWs, 'เธเนเธฒเนเธเนเธเนเธฒเธขเธ•เนเธญเนเธเธเธ');

    // Sheet 4: Ongoing projects (all months)
    const projHeaders = ['เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ', 'เธเธทเนเธญเนเธเธฃเธเธเธฒเธฃ', 'เธฃเธฒเธขเธฅเธฐเน€เธญเธตเธขเธ”เธเธงเธฒเธกเธเธทเธเธซเธเนเธฒ'];
    const projRows = [];
    Object.entries(data).forEach(([monthKey, d]) => {
      (d.ongoingProjects || []).forEach(proj => {
        projRows.push([monthKey, proj.title, proj.desc]);
      });
    });
    const projWs = XLSX.utils.aoa_to_sheet([projHeaders, ...projRows]);
    projWs['!cols'] = projHeaders.map(h => ({ wch: Math.max(h.length + 4, 25) }));
    XLSX.utils.book_append_sheet(wb, projWs, 'เนเธเธฃเธเธเธฒเธฃเธ”เธณเน€เธเธดเธเธเธฒเธฃ');

    // Sheet 5: Recommendations (all months)
    const recHeaders = ['เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ', 'เธเนเธญเน€เธชเธเธญเนเธเธฐ'];
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
    const expAssetHeaders = ['เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ', 'เธฃเธซเธฑเธชเธ—เธฃเธฑเธเธขเนเธชเธดเธ', 'เธเธฃเธฐเน€เธ เธ—', 'เธฃเธธเนเธ/เนเธกเน€เธ”เธฅ', 'เนเธเธเธ', 'เธงเธฑเธเธ—เธตเนเธซเธกเธ”เธญเธฒเธขเธธ'];
    const expAssetRows = [];
    Object.entries(data).forEach(([monthKey, d]) => {
      (d.assetsExpiringDetails || []).forEach(a => {
        expAssetRows.push([monthKey, a.id, a.type, a.model, a.dept, a.expDate]);
      });
    });
    const expAssetWs = XLSX.utils.aoa_to_sheet([expAssetHeaders, ...expAssetRows]);
    expAssetWs['!cols'] = expAssetHeaders.map(h => ({ wch: Math.max(h.length + 4, 20) }));
    XLSX.utils.book_append_sheet(wb, expAssetWs, 'เธญเธธเธเธเธฃเธ“เนเนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ');

    // Sheet 7: Expiring software details (all months)
    const expSwHeaders = ['เธฃเธซเธฑเธชเน€เธ”เธทเธญเธ', 'เธเธทเนเธญเธเธญเธเธ•เนเนเธงเธฃเน', 'เธเธณเธเธงเธ Licenses', 'เธงเธฑเธเธซเธกเธ”เธชเธฑเธเธเธฒ', 'เธชเธ–เธฒเธเธฐ'];
    const expSwRows = [];
    Object.entries(data).forEach(([monthKey, d]) => {
      (d.softwareExpiringDetails || []).forEach(s => {
        expSwRows.push([monthKey, s.name, s.licenses, s.expiringDate, s.status]);
      });
    });
    const expSwWs = XLSX.utils.aoa_to_sheet([expSwHeaders, ...expSwRows]);
    expSwWs['!cols'] = expSwHeaders.map(h => ({ wch: Math.max(h.length + 4, 22) }));
    XLSX.utils.book_append_sheet(wb, expSwWs, 'เนเธเธฃเนเธเธฃเธกเนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ');

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
            if (row['เธงเธฑเธเธ—เธตเนเน€เธเธดเธเนเธเนเธเธฒเธ']) {
              const dObj = parseExcelDate(row['เธงเธฑเธเธ—เธตเนเน€เธเธดเธเนเธเนเธเธฒเธ']);
              if (dObj) dateIssued = formatDate(dObj);
            }
            
            return {
              sn: row['Nember'] || (i + 1),
              date: dateIssued,
              user: row['เธเธธเธเธเธฅเน€เธเธดเธเนเธเนเธญเธธเธเธเธฃเธ“เน'] || 'เธชเนเธงเธเธเธฅเธฒเธ/เนเธกเนเธฃเธฐเธเธธ',
              position: row['เธ•เธณเนเธซเธเนเธ'] || '-',
              itemType: row['เธฃเธฒเธขเธเธฒเธฃเธญเธธเธเธเธฃเธ“เนเธซเธฅเธฑเธ'] || 'เธญเธธเธเธเธฃเธ“เนเน€เธชเธฃเธดเธก/เธญเธทเนเธเน',
              deviceSerial: row['เธซเธกเธฒเธขเน€เธฅเธเธญเธธเธเธเธฃเธ“เน (เน€เธเนเธ  Ipad 016)'] || '-',
              status: row['เธชเธ–เธฒเธเธฐ'] || 'เนเธเนเธเธฒเธ',
              notes: row['เธซเธกเธฒเธขเน€เธซเธ•เธธ'] || '',
              submittedOn: row['Submitted on'] || '',
              respondent: row['Respondents'] || '',
              additionalEquipment: row['เธญเธธเธเธเธฃเธ“เนเน€เธเธดเนเธกเน€เธ•เธดเธกเธ—เธตเนเธ•เนเธญเธเธเธฒเธฃเน€เธเธดเธ'] || '',
              softwareApp: row['เธเธญเธ•เนเธเนเธงเธฃเน/ App'] || '',
              registeredEmail: row['เน€เธกเธฅเธ—เธตเนเธฅเธเธ—เธฐเน€เธเธตเธขเธ'] || '',
              additionalSerial: row['เธซเธกเธฒเธขเน€เธฅเธเธญเธธเธเธเธฃเธ“เน เน€เธเธดเนเธกเน€เธ•เธดเธก  (เน€เธเนเธ  เธชเธฒเธข เธญเธฐเน€เน€เธ”เธเน€เธ•เธญเธฃเน ipad-011))'] || '',
              returnDueDate: row['เธเธณเธซเธเธ”เธเธทเธเธญเธธเธเธเธฃเธ“เน'] || '',
              inspectionDate: row['เธงเธฑเธเธ—เธตเนเธ•เธฃเธงเธเธชเธญเธ'] || '',
              purchaseDate: row['เธงเธฑเธเธ—เธตเนเธเธทเนเธญ'] || '',
              warrantyEndDate: row['เธงเธฑเธเธซเธกเธ”เธเธฃเธฐเธเธฑเธ'] || '',
              expense: Number(row['เธเนเธฒเนเธเนเธเนเธฒเธข']) || 0
            };
          }).map(repairThaiTextDeep);

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
            const cat = String(row['เธฃเธฒเธขเธเธฒเธฃเธญเธธเธเธเธฃเธ“เนเธซเธฅเธฑเธ'] || '').trim();
            calculatedAssetValue += CATEGORY_VALUES[cat] || 1500;
          });

          const brokenInventory = parsedAssets.filter(a => a.status === 'เธฃเธญเธเนเธญเธก').length;
          const lostInventory = parsedAssets.filter(a => a.status === 'เธชเธนเธเธซเธฒเธข').length;

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
          setImportStatus({ type: 'success', message: `เธเธณเน€เธเนเธฒเธเนเธญเธกเธนเธฅเธเธฅเธฑเธเธญเธธเธเธเธฃเธ“เนเธชเธณเน€เธฃเนเธ! เธเธเธญเธธเธเธเธฃเธ“เน ${parsedAssets.length} เธฃเธฒเธขเธเธฒเธฃ` });
          setTimeout(() => setImportStatus(null), 4000);
          return;
        }

        if (wb.SheetNames.includes('Form')) {
          // --- Parse Custom Form & IT Expenses Structure ---
          const formRows = XLSX.utils.sheet_to_json(wb.Sheets['Form']);
          const costRows = XLSX.utils.sheet_to_json(wb.Sheets['เธเนเธฒเนเธเนเธเนเธฒเธข IT'] || wb.Sheets[wb.SheetNames[1]]);
          
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
            'เธกเธเธฃเธฒเธเธก', 'เธเธธเธกเธ เธฒเธเธฑเธเธเน', 'เธกเธตเธเธฒเธเธก', 'เน€เธกเธฉเธฒเธขเธ', 'เธเธคเธฉเธ เธฒเธเธก', 'เธกเธดเธ–เธธเธเธฒเธขเธ',
            'เธเธฃเธเธเธฒเธเธก', 'เธชเธดเธเธซเธฒเธเธก', 'เธเธฑเธเธขเธฒเธขเธ', 'เธ•เธธเธฅเธฒเธเธก', 'เธเธคเธจเธเธดเธเธฒเธขเธ', 'เธเธฑเธเธงเธฒเธเธก'
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
                  { name: "Microsoft 365 Copilot", licenses: 50, expiringDate: "30 เธช.เธ. 2026", status: "เนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ" },
                  { name: "Adobe Creative Cloud", licenses: 15, expiringDate: "12 เธ.เธข. 2026", status: "เนเธเนเธเน€เธ•เธทเธญเธเธฅเนเธงเธเธซเธเนเธฒ" }
                ],
                assetsExpiringDetails: [
                  { id: "AST-NB-001", type: "Laptop", model: "Lenovo ThinkPad L14", dept: "Operations", expDate: "15 เธช.เธ. 2026" },
                  { id: "AST-PR-004", type: "Printer", model: "HP LaserJet Pro M404", dept: "Accounting", expDate: "22 เธช.เธ. 2026" }
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

            const durationMins = parseDurationToMinutes(row['เน€เธงเธฅเธฒเธ—เธตเนเนเธเนเนเธเธเธฒเธฃเธ—เธณเธเธฒเธ']);
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
            const hwIssue = row['เนเธเนเธ เธฎเธฒเธฃเนเธ”เนเธงเธฃเน เธเธฑเธ”เธเนเธญเธ'];
            const swIssue = row['เนเธเนเธ เธเธญเธ•เนเธเนเธงเธฃเน เธเธฑเธ”เธเนเธญเธ'] || row['เนเธเนเธ เธะพเธ•เนเธเนเธงเธฃเน เธเธฑเธ”เธเนเธญเธ'];
            const netIssue = row['เนเธเนเธ เธฃเธฐเธเธเน€เธเนเธ•เน€เธงเธดเธฃเนเธ เธเธฑเธ”เธเนเธญเธ'];
            const acctIssue = row['เนเธเนเธเธเธญเธเธฑเธเธเธตเธเธฒเธฃเนเธเนเธเธฒเธเธ•เนเธฒเธ / เธเธฑเธ”เธเนเธญเธ'];
            
            let issueSummary = [];
            if (hwIssue) issueSummary.push(`เธฎเธฒเธฃเนเธ”เนเธงเธฃเน: ${hwIssue}`);
            if (swIssue) issueSummary.push(`เธเธญเธเธ•เนเนเธงเธฃเน: ${swIssue}`);
            if (netIssue) issueSummary.push(`เน€เธเนเธ•เน€เธงเธดเธฃเนเธ: ${netIssue}`);
            if (acctIssue) issueSummary.push(`เธเธฑเธเธเธตเธเธนเนเนเธเน: ${acctIssue}`);
            
            const finalIssueText = issueSummary.join(', ') || row['เธญเธฒเธเธฒเธฃเน€เธชเธตเธขเธ•เนเธฒเธเน'] || 'เนเธเนเธเธ•เธดเธ”เธ•เธฑเนเธ/เธญเธทเนเธเน';

            monthData.ticketsList.push({
              sn: row['SN'] || (rowIndex + 1),
              date: formatDate(dateObj),
              complainant: row['เธเธทเนเธญ-เธเธฒเธกเธชเธเธธเธฅ'] || '',
              email: row['Email'] || '-',
              anydesk: row['เน€เธฅเธเธ—เธตเน Any Desk'] || '-',
              issue: finalIssueText,
              cause: row['เธชเธฒเน€เธซเธ•เธธเธเธฒเธฃเน€เธชเธตเธข'] || '-',
              duration: row['เน€เธงเธฅเธฒเธ—เธตเนเนเธเนเนเธเธเธฒเธฃเธ—เธณเธเธฒเธ'] || '-',
              responder: row['Respondents'] || '-',
              status: row['เธเธงเธฒเธกเธเธทเธเธซเธเนเธฒ'] || 'เน€เธชเธฃเนเธเธชเธดเนเธ',
              cost: Number(row['เธเธณเธเธงเธเน€เธเธดเธ']) || 0
            });

            const hwField = row['เนเธเนเธ เธฎเธฒเธฃเนเธ”เนเธงเธฃเน เธเธฑเธ”เธเนเธญเธ'] || row['เนเธเนเธเธ•เธดเธ”เธ•เธฑเนเธ เธฎเธฒเธฃเนเธ”เนเธงเธฃเน'];
            if (hwField) {
              const devices = String(hwField).split(',').map(d => d.trim()).filter(Boolean);
              devices.forEach(device => {
                if (device === 'Acc' || device === 'Morning Berf' || device === 'Meeting') return;
                
                monthData.assetsBroken++;
                monthData.repairCount++;
                monthData._deviceCounts[device] = (monthData._deviceCounts[device] || 0) + 1;
                
                const ticketCost = Number(row['เธเธณเธเธงเธเน€เธเธดเธ']) || 0;
                if (ticketCost > 0) {
                  monthData.repairCost += ticketCost;
                  monthData._deviceCosts[device] = (monthData._deviceCosts[device] || 0) + ticketCost;
                }
              });
            }
          });

          if (costRows) {
            costRows.forEach((row, i) => {
              let dateObj = parseExcelDate(row['เธงเธฑเธเธ—เธตเน']);
              if (!dateObj && row['เน€เธ”เธทเธญเธ']) {
                const mIdx = THAI_MONTHS.indexOf(row['เน€เธ”เธทเธญเธ'].trim());
                if (mIdx !== -1) {
                  dateObj = new Date(2026, mIdx, 15);
                }
              }
              
              if (!dateObj) return;
              
              const year = dateObj.getFullYear();
              const monthIndex = dateObj.getMonth();
              const monthKey = `${year}-${String(monthIndex + 1).padStart(2, '0')}`;
              
              if (newData[monthKey]) {
                const amount = Number(row['เธเธณเธเธงเธเน€เธเธดเธ']) || 0;
                newData[monthKey].repairCost += amount;
                newData[monthKey].repairCount++;
                
                newData[monthKey].ticketsList.push({
                  sn: `EXP-${row['เธฅเธณเธ”เธฑเธ'] || (i + 1)}`,
                  date: formatDate(dateObj),
                  complainant: row['เธเธทเนเธญ-เธเธฒเธกเธชเธเธธเธฅ'] || 'IT Dept',
                  issue: `เธเนเธฒเนเธเนเธเนเธฒเธข IT: ${row['เธเนเธฒเนเธเนเธเนเธฒเธข'] || 'เธเธทเนเธญเธญเธธเธเธเธฃเธ“เน'} (${row['เธชเธฒเน€เธซเธ•เธธเธเธฒเธฃเน€เธชเธตเธข'] || 'เน€เธชเธทเนเธญเธกเธ•เธฒเธกเธชเธ เธฒเธ'})`,
                  duration: '-',
                  responder: '-',
                  status: 'เธเนเธฒเธขเน€เธเธดเธเนเธฅเนเธง',
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
              { title: "IT Ticket Automation", desc: `เธเธฑเธ’เธเธฒเธชเธเธฃเธดเธเธ•เนเธเนเธงเธขเธเธฑเธ”เธเธฒเธฃเธเธฑเธเธซเธฒเธเนเธณเธเธฒเธ เธเธทเธเธซเธเนเธฒ ${Math.min(95, 40 + monthData.automationsDone * 5)}%` },
              { title: "Asset Management System", desc: "เธฃเธฐเธเธเน€เธเนเธเธญเธดเธ-เน€เธเนเธเน€เธญเธฒเธ—เนเธญเธธเธเธเธฃเธ“เนเนเธญเธ—เธต เธเธทเธเธซเธเนเธฒ 60%" }
            ];

            if (monthData.assetsBroken > 5) {
              monthData.recommendations.push(`เธเธเธเธฑเธเธซเธฒเธญเธธเธเธเธฃเธ“เนเธเธฑเธ”เธเนเธญเธเธชเธนเธเธ–เธถเธ ${monthData.assetsBroken} เธเธฃเธฑเนเธเนเธเน€เธ”เธทเธญเธเธเธตเน เนเธเธฐเธเธณเธเธฑเธ”เธฃเธญเธเธเธณเธฃเธธเธเธฃเธฑเธเธฉเธฒเน€เธเธดเธเธเนเธญเธเธเธฑเธ (Preventive Maintenance) เนเธ”เธขเน€เธเธเธฒเธฐเธญเธธเธเธเธฃเธ“เนเธเธฃเธฐเน€เธ เธ— ${monthData.topBrokenDevices[0]?.name || 'Notebook'}`);
            }
            if (monthData.slaPercent < 95) {
              monthData.recommendations.push(`เธญเธฑเธ•เธฃเธฒเธเธฒเธฃเธเธฃเธฃเธฅเธธเน€เธเนเธฒเธซเธกเธฒเธข SLA เธฅเธ”เธฅเธเน€เธซเธฅเธทเธญ ${monthData.slaPercent}% เนเธเธฐเธเธณเนเธซเนเธเธฃเธฑเธเธเธฃเธฐเธเธงเธเธเธฒเธฃเธเธฑเธ”เธเธฃเธญเธ Ticket เน€เธเธทเนเธญเน€เธเธดเนเธกเธเธงเธฒเธกเธฃเธงเธ”เน€เธฃเนเธงเนเธเธเธฒเธฃเนเธเนเธเธฑเธเธซเธฒ`);
            } else {
              monthData.recommendations.push("เธเธฒเธฃเธชเธเธฑเธเธชเธเธธเธเธเธนเนเนเธเนเธฃเธฐเธเธ IT เธญเธขเธนเนเนเธเน€เธเธ“เธ‘เนเธ”เธตเน€เธขเธตเนเธขเธก เธชเธฒเธกเธฒเธฃเธ–เธฃเธฑเธเธฉเธฒเธกเธฒเธ•เธฃเธเธฒเธ SLA เนเธ”เนเธ•เธฒเธกเน€เธเนเธฒเธซเธกเธฒเธข");
            }
            monthData.recommendations.push("เนเธเธฐเธเธณเนเธซเนเธเธนเนเนเธเนเธญเธฑเธเน€เธเธฃเธ”เธเธงเธฒเธกเธเธฅเธญเธ”เธ เธฑเธขเธเธญเธเธเธฑเธเธเธตเธเนเธฒเธเธเธฒเธฃเน€เธเธดเธ”เนเธเนเธเธฒเธ MFA เธเธฃเธ 100%");

            delete monthData._durationSum;
            delete monthData._durationCount;
            delete monthData._slaCompliantCount;
            delete monthData._deviceCounts;
            delete monthData._deviceCosts;
          });

        } else {
          // --- Parse Sheet 1: Dashboard (by column index) ---
          const dashSheet = wb.Sheets[wb.SheetNames[0]];
          if (!dashSheet) throw new Error('เนเธกเนเธเธ Sheet เนเธฃเธ (Dashboard)');
          const allRows = XLSX.utils.sheet_to_json(dashSheet, { header: 1 });
          // Skip header row (index 0), data starts from row 1
          const dataRows = allRows.slice(1).filter(r => r && r.length > 1);
          if (dataRows.length === 0) throw new Error('เนเธกเนเธเธเธเนเธญเธกเธนเธฅเนเธ Sheet Dashboard');

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
        if (monthKeys.length === 0) throw new Error('เนเธกเนเธเธเธเนเธญเธกเธนเธฅเน€เธ”เธทเธญเธเธ—เธตเนเธชเธฒเธกเธฒเธฃเธ–เธเธณเน€เธเนเธฒเนเธ”เน');

        setData(newData);
        setCurrentMonth(monthKeys[0]);
        setImportStatus({ type: 'success', message: `เธเธณเน€เธเนเธฒเธชเธณเน€เธฃเนเธ! เธเธเธเนเธญเธกเธนเธฅ ${monthKeys.length} เน€เธ”เธทเธญเธ` });
        setTimeout(() => setImportStatus(null), 4000);
      } catch (err) {
        setImportStatus({ type: 'error', message: `เธเธณเน€เธเนเธฒเธฅเนเธกเน€เธซเธฅเธง: ${err.message}` });
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

  const monitoredTotal = externalDevices.length;
  // Active includes devices that are also close to their due date. The warning
  // metric is a subset of Active, matching the Device Monitor source dashboard.
  const monitoredActive = externalDevices.filter(d => d.status === 'active').length;
  const monitoredWarning = externalDevices.filter(d => d.status === 'active' && d.daysRemaining <= 7).length;
  const monitoredUnverified = externalDevices.filter(d => d.status !== 'active').length;
  const pendingAssetApprovalCount = assetRequests.filter(request =>
    ['pending', 'need_info', 'approved'].includes(request.status)
  ).length;
  const pendingTicketCloseCount = (data[currentMonth]?.ticketsList || []).filter(ticket =>
    ticket.status === 'เธเธณเธฅเธฑเธเธ”เธณเน€เธเธดเธเธเธฒเธฃ'
  ).length;

  const pendingTicketsCount = pendingTicketCloseCount;
  const pendingAssetRequestsCount = (assetRequests || []).filter(req => req.status === 'pending').length;
  return (
    <>
      {/* SIDEBAR NAVIGATION CONTROL PANEL */}
      <aside className={`sidebar no-print ${mobileSidebarOpen ? 'mobile-active' : ''}`}>
        <div className="logo-container" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', width: '100%' }}>
          <img className="fern-brand-logo" src={lightItLogo} alt="Light IT" />
          <button 
            onClick={() => setMobileSidebarOpen(false)} 
            className="mobile-menu-close"
            style={{ display: 'none', background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', padding: '4px' }}
          >
            <X size={20} />
          </button>
        </div>

        <div className="sidebar-primary-actions">

          <button
            onClick={() => {
              setMobileSidebarOpen(false);
              navigate('/form');
            }}
            className="sidebar-form-btn"
            style={{ position: 'relative' }}
          >
            เนเธเนเธ Ticket
            <span
              className={`menu-count-badge sidebar-notification-count ${pendingTicketsCount > 0 ? 'has-items' : ''}`}
              aria-label={`Ticket เธ—เธตเนเธฃเธญเธ”เธณเน€เธเธดเธเธเธฒเธฃ ${pendingTicketsCount} เธฃเธฒเธขเธเธฒเธฃ`}
            >
              {pendingTicketsCount}
            </span>
          </button>
          <button onClick={() => {
            setAssetWorkflowRole('requester');
            setMobileSidebarOpen(false);
            setActiveModal('assetWorkflow');
          }} className="sidebar-btn" style={{ backgroundColor: '#7c3aed', border: 'none', color: 'white', position: 'relative' }}>
            <Ticket size={16} />
            เธเธญเน€เธเธดเธเธญเธธเธเธเธฃเธ“เน
            <span
              className={`menu-count-badge sidebar-notification-count ${pendingAssetRequestsCount > 0 ? 'has-items' : ''}`}
              aria-label={`เธเธณเธเธญเน€เธเธดเธเธญเธธเธเธเธฃเธ“เนเธ—เธตเนเธฃเธญเธญเธเธธเธกเธฑเธ•เธด ${pendingAssetRequestsCount} เธฃเธฒเธขเธเธฒเธฃ`}
            >
              {pendingAssetRequestsCount}
            </span>
          </button>
          <button onClick={() => {
            setAssetReturnView('returns');
            setAssetReturnSearch('');
            setAssetReturnIdentity('');
            setMobileSidebarOpen(false);
            setActiveModal('assetReturns');
          }} className="sidebar-btn asset-return-menu-btn">
            <RotateCcw size={16} />
            เธเธทเธเธญเธธเธเธเธฃเธ“เน
          </button>

        </div>

        {/* Month Dropdown Selection */}
        <div className="control-group">
          <label className="control-label">เน€เธฅเธทเธญเธเน€เธ”เธทเธญเธเธ—เธตเนเธ•เนเธญเธเธเธฒเธฃเธฃเธฒเธขเธเธฒเธ</label>
          <select 
            className="month-selector"
            value={currentMonth}
            onChange={(e) => setCurrentMonth(e.target.value)}
          >
            {Object.keys(data).sort((a, b) => b.localeCompare(a)).map(key => (
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
            <label className="control-label" style={{ margin: 0, cursor: 'pointer' }}>๐ ๏ธ เธเธฒเธฃเธเธฑเธ”เธเธฒเธฃเธเนเธญเธกเธนเธฅ</label>
            {sidebarExpanded.mgmt ? <ChevronDown size={14} style={{ color: 'var(--text-muted)' }} /> : <ChevronRight size={14} style={{ color: 'var(--text-muted)' }} />}
          </div>
          {sidebarExpanded.mgmt && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <button onClick={() => requireAdminAccess(openEditModal)} className="sidebar-btn">
                <Edit3 size={16} />
                เนเธเนเนเธเธ•เธฑเธงเน€เธฅเธเน€เธ”เธทเธญเธเธเธตเน
              </button>
              <button onClick={() => requireAdminAccess(() => {
                setConsoleMonth(currentMonth);
                setActiveModal('fullConsole');
              })} className="sidebar-btn secondary">
                <Database size={16} />
                เธเธฃเธฑเธเน€เธเธฅเธตเนเธขเธเธเนเธญเธกเธนเธฅเธ—เธฑเนเธเธซเธกเธ”
              </button>

              <button onClick={() => requireAdminAccess(() => {
                setLarkFormType('asset');
                setLarkTicketRole('user');
                setLarkSubmitted(false);
                setActiveModal('larkForm');
              })} className="sidebar-btn" style={{ backgroundColor: '#06b6d4', border: 'none', color: 'white' }}>
                <Laptop size={16} />
                เธฅเธเธ—เธฐเน€เธเธตเธขเธเน€เธเธฃเธทเนเธญเธเน€เธเนเธฒเธเธฅเธฑเธ
              </button>
              <button onClick={() => requireAdminAccess(() => {
                setAssetWorkflowRole('it');
                setActiveModal('assetWorkflow');
              })} className="sidebar-btn" style={{ backgroundColor: '#4338ca', border: 'none', color: 'white' }}>
                <ShieldCheck size={16} />
                IT เธญเธเธธเธกเธฑเธ•เธดเธเธฒเธฃเนเธเนเธเธฒเธ
                <span className={`menu-count-badge ${pendingAssetApprovalCount > 0 ? 'has-items' : ''}`}>{pendingAssetApprovalCount}</span>
              </button>
              <button onClick={() => requireAdminAccess(() => {
                setLarkFormType('ticket');
                setLarkTicketRole('it');
                setLarkSubmitted(false);
                setSelectedPendingTicketSn('');
                setActiveModal('larkForm');
              })} className="sidebar-btn" style={{ backgroundColor: '#f59e0b', border: 'none', color: 'white' }}>
                <Wrench size={16} />
                เน€เธกเธเธนเธเธดเธ”เธเธฒเธ (IT Close)
                <span className={`menu-count-badge ${pendingTicketCloseCount > 0 ? 'has-items' : ''}`}>{pendingTicketCloseCount}</span>
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
            <label className="control-label" style={{ margin: 0, cursor: 'pointer' }}>๐“ เธเธฒเธเธเนเธญเธกเธนเธฅ Excel (.xlsx)</label>
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
              <button onClick={() => requireAdminAccess(() => fileInputRef.current?.click())} className="sidebar-btn" style={{ backgroundColor: '#059669' }}>
                <Upload size={16} />
                เธเธณเน€เธเนเธฒเธเนเธญเธกเธนเธฅเธเธฒเธ Excel
              </button>
              <button onClick={() => requireAdminAccess(exportToXlsx)} className="sidebar-btn secondary">
                <Download size={16} />
                เธชเนเธเธญเธญเธเธเนเธญเธกเธนเธฅเน€เธเนเธ Excel
              </button>
              <button onClick={downloadTemplate} className="sidebar-btn secondary">
                <FileSpreadsheet size={16} />
                เธ”เธฒเธงเธเนเนเธซเธฅเธ”เน€เธ—เธกเน€เธเธฅเธ• Excel
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
            <label className="control-label" style={{ margin: 0, cursor: 'pointer' }}>๐“ เธชเนเธเธญเธญเธเน€เธญเธเธชเธฒเธฃ</label>
            {sidebarExpanded.export ? <ChevronDown size={14} style={{ color: 'var(--text-muted)' }} /> : <ChevronRight size={14} style={{ color: 'var(--text-muted)' }} />}
          </div>
          {sidebarExpanded.export && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <button onClick={() => window.print()} className="sidebar-btn secondary">
                <Printer size={16} />
                เธเธฑเธเธ—เธถเธเน€เธเนเธ PDF / เธเธดเธกเธเน
              </button>
            </div>
          )}
        </div>

        <div className="org-info">
          <p><strong>เธซเธเนเธงเธขเธเธฒเธ:</strong> เธเนเธฒเธขเน€เธ—เธเนเธเนเธฅเธขเธตเธชเธฒเธฃเธชเธเน€เธ—เธจ (IT)</p>
          <p><strong>เธญเธเธเนเธเธฃ:</strong> Fern Aesthetique</p>
          <p style={{ marginTop: '8px', fontSize: '0.7rem', color: 'var(--text-muted)' }}>
            เธเธฃเธฑเธเธเธฃเธธเธเธเนเธญเธกเธนเธฅเธฅเนเธฒเธชเธธเธ”: <br />18 เธเธฃเธเธเธฒเธเธก 2026
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
              <h2>เธฃเธฒเธขเธเธฒเธเธชเธฃเธธเธเธเธฒเธฃเธ”เธณเน€เธเธดเธเธเธฒเธเน€เธ—เธเนเธเนเธฅเธขเธตเธชเธฒเธฃเธชเธเน€เธ—เธจ (IT Monthly Dashboard)</h2>
              <p>เธเธฃเธฐเธเธณเน€เธ”เธทเธญเธ {activeData.monthName}</p>
            </div>
          </div>
          <div className="header-status status-indicator" style={{ display: 'flex', gap: '15px', alignItems: 'center', background: 'none' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', background: 'rgba(255, 255, 255, 0.1)', padding: '5px 12px', borderRadius: '20px' }}>
              <span className="status-dot"></span>
              <span>เธฃเธฐเธเธเธฃเธฒเธขเธเธฒเธเธเธฃเนเธญเธกเธ—เธณเธเธฒเธ</span>
            </div>
          </div>
        </header>

        <section className="dashboard-grid">
          {/* CARD 0: DEVICE MONITOR */}
          <article className="card asset-card" style={{ borderColor: monitoredUnverified > 0 || monitoredWarning > 0 ? 'var(--warning)' : 'var(--border-color)' }}>
            <div className="card-header">
              <h3 className="card-title">
                <span className="card-icon"><ShieldCheck size={18} style={{ color: 'var(--primary)' }} /></span>
                เธ•เธดเธ”เธ•เธฒเธกเธชเธ–เธฒเธเธฐเธญเธธเธเธเธฃเธ“เน iOS (Device Monitor)
              </h3>
              <div style={{ display: 'flex', gap: '7px', alignItems: 'center' }}>
                <span style={{ fontSize: '0.8rem', color: externalDevicesSyncError ? 'var(--danger)' : isFetchingDevices ? 'var(--warning)' : 'var(--text-muted)', display: 'flex', alignItems: 'center' }}>
                  {externalDevicesSyncError || (isFetchingDevices
                    ? 'เธเธณเธฅเธฑเธเธเธดเธเธเน...'
                    : externalDevicesLastSynced
                      ? `เธเธดเธเธเนเธฅเนเธฒเธชเธธเธ” ${externalDevicesLastSynced.toLocaleTimeString('th-TH')}`
                      : 'เธฃเธญเธเธดเธเธเนเธเนเธญเธกเธนเธฅ')}
                </span>
                <button
                  type="button"
                  title="เธเธดเธเธเน Device Monitor เธ•เธญเธเธเธตเน"
                  aria-label="เธเธดเธเธเน Device Monitor เธ•เธญเธเธเธตเน"
                  onClick={() => setExternalDevicesRefreshKey(previous => previous + 1)}
                  disabled={isFetchingDevices}
                  style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '30px', height: '30px', padding: 0, borderRadius: '8px', border: '1px solid var(--border-color)', background: 'var(--card-bg)', color: 'var(--text-primary)', cursor: isFetchingDevices ? 'wait' : 'pointer' }}
                >
                  <RotateCcw size={14} className={isFetchingDevices ? 'device-sync-spinning' : ''} />
                </button>
              </div>
            </div>
            <div className="metrics-row" style={{ marginTop: '15px' }}>
              <div 
                className="metric-item" 
                style={{ cursor: 'pointer' }}
                onClick={() => window.open('https://ios-device-monitor-46w9.onrender.com/', '_blank')}
              >
                <div className="metric-label">เธเธณเธฅเธฑเธเธ•เธดเธ”เธ•เธฒเธกเธฃเธงเธก</div>
                <div className="metric-value highlight-primary">{monitoredTotal} เน€เธเธฃเธทเนเธญเธ</div>
              </div>
              <div 
                className="metric-item"
                style={{ cursor: 'pointer' }}
                onClick={() => window.open('https://ios-device-monitor-46w9.onrender.com/', '_blank')}
              >
                <div className="metric-label">เธ•เธฃเธงเธเธชเธญเธเนเธฅเนเธง (Active)</div>
                <div className="metric-value highlight-success">{monitoredActive} เน€เธเธฃเธทเนเธญเธ</div>
              </div>
              <div 
                className="metric-item"
                style={{ cursor: 'pointer' }}
                onClick={() => window.open('https://ios-device-monitor-46w9.onrender.com/', '_blank')}
              >
                <div className="metric-label">เนเธเธฅเนเธเธฃเธเธเธณเธซเธเธ” (&le;7 เธงเธฑเธ)</div>
                <div className="metric-value highlight-warning">{monitoredWarning} เน€เธเธฃเธทเนเธญเธ</div>
              </div>
              <div 
                className="metric-item"
                style={{ cursor: 'pointer' }}
                onClick={() => window.open('https://ios-device-monitor-46w9.onrender.com/', '_blank')}
              >
                <div className="metric-label">เธเนเธฒเธเธเธฒเธฃเธ•เธฃเธงเธเธชเธญเธ (Action Required)</div>
                <div className="metric-value highlight-danger">{monitoredUnverified} เน€เธเธฃเธทเนเธญเธ</div>
              </div>
            </div>
          </article>

          {/* CARD 1: ASSETS */}
          <article className="card asset-card">
            <div className="card-header">
              <h3 className="card-title">
                <span className="card-icon"><Laptop size={18} style={{ color: 'var(--primary)' }} /></span>
                เธ—เธฃเธฑเธเธขเนเธชเธดเธเนเธฅเธฐเธญเธธเธเธเธฃเธ“เน (Asset)
              </h3>
              <div style={{ display: 'flex', gap: '6px' }}>
                <button 
                  onClick={() => setActiveModal('assetsList')} 
                  className="btn-details"
                  style={{ backgroundColor: 'rgba(59, 130, 246, 0.15)', color: 'var(--primary)' }}
                >
                  เธ”เธนเธ—เธฐเน€เธเธตเธขเธเธญเธธเธเธเธฃเธ“เน
                </button>
                <button 
                  onClick={() => setActiveModal('expiringAssets')} 
                  className="btn-details"
                >
                  เธ”เธนเธเนเธญเธกเธนเธฅเธซเธกเธ”เธญเธฒเธขเธธ
                </button>
              </div>
            </div>
            <div className="metrics-row">
              <div className="metric-item full-width">
                <div className="asset-total-summary">
                  <div>
                    <div className="metric-label">เธเธณเธเธงเธเธญเธธเธเธเธฃเธ“เนเธ—เธฑเนเธเธซเธกเธ”</div>
                    <div className="metric-value highlight-primary">{activeData.totalAssets.toLocaleString()} เน€เธเธฃเธทเนเธญเธ</div>
                  </div>
                  <div className="asset-vacant-summary">
                    <div className="metric-label">เน€เธเธฃเธทเนเธญเธเธงเนเธฒเธ</div>
                    <div className="metric-value highlight-success">{vacantStockCount} เน€เธเธฃเธทเนเธญเธ</div>
                    <div className="vacant-assets-breakdown">
                      {vacantStockBreakdown.length > 0 ? vacantStockBreakdown.map((asset) => (
                        <span key={asset.label}>{asset.label} <strong>{asset.count}</strong></span>
                      )) : (
                        <span>เนเธกเนเธกเธตเน€เธเธฃเธทเนเธญเธเธงเนเธฒเธ</span>
                      )}
                    </div>
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
                <div className="metric-label">เธกเธนเธฅเธเนเธฒเธ—เธฃเธฑเธเธขเนเธชเธดเธ IT เธฃเธงเธก</div>
                <div className="metric-value">{formatThaiBaht(activeData.assetValue)}</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">เนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ</div>
                <div className="metric-value highlight-warning">{primaryExpiringAssets} เน€เธเธฃเธทเนเธญเธ</div>
                <div className="metric-note">เน€เธซเธฅเธทเธญเธญเธฒเธขเธธเธฃเธธเนเธเนเธกเนเน€เธเธดเธ 1 เธเธต</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">เน€เธเธฃเธทเนเธญเธเธงเนเธฒเธ (เธเธฃเนเธญเธกเนเธเน)</div>
                <div className="metric-value highlight-success">{vacantStockCount} เน€เธเธฃเธทเนเธญเธ</div>
                <div className="metric-note">เธ•เธฒเธกเธชเธ–เธฒเธเธฐเธงเนเธฒเธเนเธเธ—เธฐเน€เธเธตเธขเธเธเธฅเธฑเธ</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">เธเธณเธฃเธธเธ”</div>
                <div className="metric-value highlight-danger">{activeData.assetsBroken} เน€เธเธฃเธทเนเธญเธ</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">เธชเธนเธเธซเธฒเธข</div>
                <div className="metric-value highlight-danger">{activeData.assetsLost} เน€เธเธฃเธทเนเธญเธ</div>
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
                เธเธฒเธฃเธชเธเธฑเธเธชเธเธธเธเธเธนเนเนเธเน (Support)
              </h3>
              <button 
                onClick={() => setActiveModal('ticketsList')} 
                className="btn-details"
              >
                เธ”เธนเธฃเธฒเธขเธฅเธฐเน€เธญเธตเธขเธ”
              </button>
            </div>
            <div className="metrics-row">
              <div className="metric-item">
                <div className="metric-label">เธเธณเธเธงเธ Ticket เธ—เธฑเนเธเธซเธกเธ”</div>
                <div className="metric-value highlight-primary">{activeData.ticketsCount.toLocaleString()} เนเธ</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">เธชเธ–เธดเธ•เธด SLA Compliance</div>
                <div className="metric-value highlight-success">{activeData.slaPercent}%</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">Response Time เน€เธเธฅเธตเนเธข</div>
                <div className="metric-value">{activeData.responseTime} เธเธฒเธ—เธต</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">Resolution Time เน€เธเธฅเธตเนเธข</div>
                <div className="metric-value">{activeData.resolutionTime} เธเธก.</div>
              </div>
              <div className="metric-item full-width">
                <div className="metric-label">เธเธฐเนเธเธเธเธงเธฒเธกเธเธถเธเธเธญเนเธเธเธนเนเนเธเน (CSAT)</div>
                <div className="metric-value highlight-warning">{activeData.csat.toFixed(1)} / 5.0</div>
              </div>
            </div>
          </article>

          {/* CARD 3: SOFTWARE */}
          <article className="card software-card">
            <div className="card-header">
              <h3 className="card-title">
                <span className="card-icon"><FileCode size={18} style={{ color: 'var(--secondary)' }} /></span>
                เธเธญเธเธ•เนเนเธงเธฃเนเนเธฅเธฐเธฅเธดเธเธชเธดเธ—เธเธดเน (Software)
              </h3>
              <button 
                onClick={() => requireAdminAccess(() => setActiveModal('expiringSoftware'))}
                className="btn-details"
              >
                เธ”เธนเธฃเธฒเธขเธฅเธฐเน€เธญเธตเธขเธ” / เนเธเนเนเธ
              </button>
            </div>
            <div className="metrics-row">
              <div className="metric-item">
                <div className="metric-label">เนเธเธฃเนเธเธฃเธกเธ—เธฑเนเธเธซเธกเธ”</div>
                <div className="metric-value">{calculatedTotalSoftware} เนเธเธฃเนเธเธฃเธก</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">เนเธเธฃเนเธเธฃเธกเนเธเธฅเนเธซเธกเธ”เธชเธฑเธเธเธฒ</div>
                <div className="metric-value highlight-danger">{activeData.softwareExpiring} เนเธเธฃเนเธเธฃเธก</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">License เนเธเนเธเธฒเธ</div>
                <div className="metric-value highlight-primary">{calculatedLicensesInUse.toLocaleString()} Core/User</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">License เธงเนเธฒเธ</div>
                <div className="metric-value highlight-secondary">{calculatedLicensesVacant.toLocaleString()} Core/User</div>
              </div>
              <div className="metric-item full-width">
                <div className="metric-label">เธเนเธฒเนเธเนเธเนเธฒเธขเธเธญเธเธ•เนเนเธงเธฃเนเธฃเธงเธกเธฃเธฒเธขเน€เธ”เธทเธญเธ</div>
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
                เธเธงเธฒเธกเธเธฅเธญเธ”เธ เธฑเธขเธเนเธญเธกเธนเธฅ (Security)
              </h3>
            </div>
            <div className="metrics-row">
              <div className="metric-item full-width">
                <div className="metric-label">Security Incident (เน€เธซเธ•เธธเธเธฒเธฃเธ“เนเธเธธเธเธเธฒเธก)</div>
                <div className={`metric-value ${activeData.securityIncidents > 0 ? 'highlight-danger' : 'highlight-success'}`}>
                  {activeData.securityIncidents} เธเธฃเธฑเนเธ
                </div>
              </div>
            </div>
            <div className="gauges-container">
              <CircularProgress value={activeData.backupSuccess} label="Backup เธชเธณเน€เธฃเนเธ" />
              <CircularProgress value={activeData.antivirusCoverage} label="Antivirus Coverage" />
              <CircularProgress value={activeData.mfaCoverage} label="MFA Coverage" />
            </div>
          </article>

          {/* CARD 5: REPAIR & COST */}
          <article className="card repair-card">
            <div className="card-header">
              <h3 className="card-title">
                <span className="card-icon"><Wrench size={18} style={{ color: 'var(--warning)' }} /></span>
                เธเธฒเธฃเธเนเธญเธกเธเธณเธฃเธธเธเนเธฅเธฐเนเธเธเธ (Repair)
              </h3>
              <button 
                onClick={() => setActiveModal('topBrokenDevices')} 
                className="btn-details"
              >
                เธ”เธน Top 10 เน€เธชเธตเธขเธเนเธญเธข
              </button>
            </div>
            <div className="metrics-row">
              <div className="metric-item">
                <div className="metric-label">เธเธณเธเธงเธเธเธฒเธเธเนเธญเธก</div>
                <div className="metric-value highlight-warning">{activeData.repairCount} เธเธฃเธฑเนเธ</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">เธเนเธฒเธเนเธญเธกเธเธณเธฃเธธเธเธชเธฐเธชเธก</div>
                <div className="metric-value">{formatThaiBaht(activeData.repairCost)}</div>
              </div>
            </div>
            <div className="card-chart-container" style={{ height: '100px' }}>
              <canvas ref={repairCanvasRef}></canvas>
            </div>
            <div className="repair-list-summary">
              <div className="metric-label" style={{ marginBottom: '2px' }}>เธญเธธเธเธเธฃเธ“เนเนเธเนเธเธเนเธญเธกเธชเธนเธเธชเธธเธ” 3 เธญเธฑเธเธ”เธฑเธเนเธฃเธ:</div>
              {activeData.topBrokenDevices.filter(d => d.count > 0).slice(0, 3).map((device, idx) => (
                <div key={idx} className="repair-list-item">
                  <span className="repair-item-name">{device.name}</span>
                  <span className="repair-item-count">{device.count} เน€เธเธฃเธทเนเธญเธ</span>
                </div>
              ))}
              {activeData.topBrokenDevices.filter(d => d.count > 0).length === 0 && (
                <div className="repair-list-item">
                  <span className="repair-item-name">เนเธกเนเธกเธตเธเธฃเธฐเธงเธฑเธ•เธดเธญเธธเธเธเธฃเธ“เนเธเธณเธฃเธธเธ”เนเธเน€เธ”เธทเธญเธเธเธตเน</span>
                </div>
              )}
            </div>
          </article>

          {/* CARD 6: IMPROVEMENT */}
          <article className="card improvement-card">
            <div className="card-header">
              <h3 className="card-title">
                <span className="card-icon"><Lightbulb size={18} style={{ color: 'var(--danger)' }} /></span>
                เธเธฒเธฃเธเธฃเธฑเธเธเธฃเธธเธเนเธฅเธฐเน€เธ—เธเนเธเนเธฅเธขเธต (Improvement)
              </h3>
            </div>
            <div className="metrics-row">
              <div className="metric-item">
                <div className="metric-label">Automation เธ—เธตเนเธ—เธณเน€เธชเธฃเนเธ</div>
                <div className="metric-value highlight-primary">{activeData.automationsDone} เธฃเธฒเธขเธเธฒเธฃ</div>
              </div>
              <div className="metric-item">
                <div className="metric-label">AI เธ—เธตเนเธเธณเธกเธฒเธเธฃเธฐเธขเธธเธเธ•เนเนเธเน</div>
                <div className="metric-value highlight-secondary">{activeData.aiApps} เนเธกเน€เธ”เธฅ</div>
              </div>
              <div className="metric-item full-width">
                <div className="metric-label">เธเธฃเธฐเธซเธขเธฑเธ”เธเธฑเนเธงเนเธกเธเธเธฒเธฃเธ—เธณเธเธฒเธ</div>
                <div className="metric-value highlight-success">{activeData.hoursSaved} เธเธฑเนเธงเนเธกเธ/เน€เธ”เธทเธญเธ</div>
              </div>
            </div>
            <div className="project-list">
              <div className="metric-label" style={{ marginBottom: '2px' }}>เนเธเธฃเธเธเธฒเธฃเธซเธฅเธฑเธเธ—เธตเนเธเธณเธฅเธฑเธเธ”เธณเน€เธเธดเธเธเธฒเธฃ:</div>
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
            เธเนเธญเน€เธชเธเธญเนเธเธฐเนเธฅเธฐเนเธเธงเธ—เธฒเธเธเธเธดเธเธฑเธ•เธด (Recommendation)
          </h3>
          <div className="recommendation-content">
            {activeData.recommendations.length > 0 ? (
              <ul className="recommendation-list">
                {activeData.recommendations.map((rec, idx) => (
                  <li key={idx}>{rec}</li>
                ))}
              </ul>
            ) : (
              <p>เนเธกเนเธกเธตเธเนเธญเน€เธชเธเธญเนเธเธฐเน€เธเธดเนเธกเน€เธ•เธดเธกเธชเธณเธซเธฃเธฑเธเน€เธ”เธทเธญเธเธเธตเน เธฃเธฐเธเธเธญเธขเธนเนเนเธเน€เธเธ“เธ‘เนเธเธเธ•เธด</p>
            )}
          </div>
        </section>
      </main>

      {activeModal === 'assetWorkflow' && (() => {
        const statusLabels = {
          pending: 'เธฃเธญเธญเธเธธเธกเธฑเธ•เธด', approved: 'เธญเธเธธเธกเธฑเธ•เธดเนเธฅเนเธง', rejected: 'เนเธกเนเธญเธเธธเธกเธฑเธ•เธด',
          issued: 'เธชเนเธเธกเธญเธเนเธฅเนเธง', overdue: 'เน€เธเธดเธเธเธณเธซเธเธ”', return_requested: 'เธฃเธญ IT เธ•เธฃเธงเธเธฃเธฑเธ', returned: 'เธเธทเธเนเธฅเนเธง', need_info: 'เธฃเธญเธเนเธญเธกเธนเธฅเน€เธเธดเนเธก'
        };
        const requesterSearch = assetRequesterSearch.trim().toLocaleLowerCase('th-TH');
        const itUsageStatuses = ['pending', 'need_info', 'approved', 'issued', 'overdue'];
        const displayedRequests = assetWorkflowRole === 'it'
          ? assetRequests.filter(request => itUsageStatuses.includes(request.status))
          : (!requesterSearch
              ? assetRequests
              : assetRequests.filter(request => request.requester?.toLocaleLowerCase('th-TH').includes(requesterSearch)));
        return (
          <div className="modal-overlay active">
            <div className="modal large dashboard-fullscreen-modal asset-workflow-modal">
              <header className="modal-header">
                <div>
                  <h3>{assetWorkflowRole === 'it' ? '๐ก๏ธ IT เธญเธเธธเธกเธฑเธ•เธดเธเธฒเธฃเนเธเนเธเธฒเธเนเธฅเธฐเธชเนเธเธกเธญเธ' : '๐ เธเธนเนเธเธญเนเธเนเธเธฃเธดเธเธฒเธฃเธญเธธเธเธเธฃเธ“เน IT'}</h3>
                  <p className="workflow-subtitle">{assetWorkflowRole === 'it' ? 'เธ•เธฃเธงเธเธชเธญเธเธเธณเธเธญ โ’ เธญเธเธธเธกเธฑเธ•เธดเนเธฅเธฐเน€เธฅเธทเธญเธเน€เธเธฃเธทเนเธญเธ โ’ เธชเนเธเธกเธญเธเน€เธเธทเนเธญเนเธเนเธเธฒเธ' : 'เธชเนเธเธเธณเธเธญเนเธซเธกเนเนเธฅเธฐเธ•เธดเธ”เธ•เธฒเธกเธชเธ–เธฒเธเธฐเธเธฒเธฃเนเธเนเธเธฒเธ'}</p>
                </div>
                <div className="workflow-header-actions">
                  <div className="workflow-role-switch">
                    <button className={assetWorkflowRole === 'requester' ? 'active' : ''} onClick={() => setAssetWorkflowRole('requester')}>เธเธนเนเธเธญเนเธเนเธเธฃเธดเธเธฒเธฃ</button>
                  </div>
                  <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
                </div>
              </header>
              <div className={`modal-body workflow-body workflow-role-${assetWorkflowRole}`}>
                {assetWorkflowRole === 'requester' && <form className="workflow-request-form" onSubmit={submitAssetRequest}>
                  <div className="workflow-section-heading">
                    <h4>เธชเธฃเนเธฒเธเธเธณเธเธญเน€เธเธดเธเธญเธธเธเธเธฃเธ“เน</h4>
                    <span>เธฃเธฐเธเธธเนเธ”เนเธซเธฅเธฒเธขเน€เธเธฃเธทเนเธญเธ เนเธ”เธขเธเธฑเนเธเธซเธกเธฒเธขเน€เธฅเธเธ”เนเธงเธขเธเธธเธฅเธ เธฒเธเธซเธฃเธทเธญเธเธถเนเธเธเธฃเธฃเธ—เธฑเธ”เนเธซเธกเน</span>
                  </div>
                  <div className="workflow-form-grid">
                    <label>เธเธทเนเธญเธเธนเนเธเธญ <input required value={assetRequestForm.requester} onChange={e => setAssetRequestForm(p => ({ ...p, requester: e.target.value }))} /></label>
                    <label>เนเธเธเธ <input required value={assetRequestForm.department} onChange={e => setAssetRequestForm(p => ({ ...p, department: e.target.value }))} /></label>
                    <label>เธซเธกเธฒเธขเน€เธฅเธเน€เธเธฃเธทเนเธญเธ / เธเธฃเธฐเน€เธ เธ—เธญเธธเธเธเธฃเธ“เน <textarea required rows="2" placeholder="เน€เธเนเธ iPad-006, iPad-007" value={assetRequestForm.itemType} onChange={e => setAssetRequestForm(p => ({ ...p, itemType: e.target.value }))} /></label>
                    <label>เธงเธฑเธเธ—เธตเนเน€เธเธดเธ <input required type="date" value={assetRequestForm.requestedDate} onChange={e => setAssetRequestForm(p => ({ ...p, requestedDate: e.target.value }))} /></label>
                    <label className="workflow-span-2">เน€เธซเธ•เธธเธเธฅเธเธฒเธฃเนเธเนเธเธฒเธ <textarea required value={assetRequestForm.purpose} onChange={e => setAssetRequestForm(p => ({ ...p, purpose: e.target.value }))} /></label>
                    <label className="workflow-span-2">เธซเธกเธฒเธขเน€เธซเธ•เธธ <input value={assetRequestForm.notes} onChange={e => setAssetRequestForm(p => ({ ...p, notes: e.target.value }))} /></label>
                  </div>
                  <button className="workflow-primary-btn" type="submit" disabled={assetRequestLoading}>{assetRequestLoading ? 'เธเธณเธฅเธฑเธเธเธฑเธเธ—เธถเธ...' : 'เธชเนเธเธเธณเธเธญเน€เธเธดเธ'}</button>
                </form>}

                <section className="workflow-list-section">
                  <div className="workflow-section-heading">
                    <div>
                      <h4>{assetWorkflowRole === 'it' ? 'เธฃเธฒเธขเธเธฒเธฃเธฃเธญเธญเธเธธเธกเธฑเธ•เธดเนเธฅเธฐเธเธฒเธฃเนเธเนเธเธฒเธ' : 'เธ•เธดเธ”เธ•เธฒเธกเธเธณเธเธญเธเธญเธเธเธนเนเนเธเนเธเธฃเธดเธเธฒเธฃ'}</h4>
                      {assetWorkflowRole === 'requester' && <input className="workflow-requester-search" placeholder="เธเนเธเธซเธฒเธ”เนเธงเธขเธเธทเนเธญเธเธนเนเธเธญ" value={assetRequesterSearch} onChange={event => setAssetRequesterSearch(event.target.value)} />}
                    </div>
                    <span>{displayedRequests.length} เธฃเธฒเธขเธเธฒเธฃ{assetWorkflowRole === 'it' ? ` ยท เน€เธเธฃเธทเนเธญเธเธงเนเธฒเธ ${assetsList.filter(asset => asset.status === 'เธงเนเธฒเธ').length} เน€เธเธฃเธทเนเธญเธ` : ''}</span>
                  </div>
                  <div className="workflow-table-wrap">
                    <table className="details-table workflow-table">
                      <thead><tr><th>เน€เธฅเธเธ—เธตเน</th><th>เธเธนเนเธเธญ/เนเธเธเธ</th><th>เธญเธธเธเธเธฃเธ“เน/เน€เธซเธ•เธธเธเธฅ</th><th>เธเธณเธซเธเธ”เธเธทเธ</th><th>เน€เธเธฃเธทเนเธญเธเธ—เธตเนเธเธฑเธ”เธชเธฃเธฃ</th><th>เธชเธ–เธฒเธเธฐ</th><th>เธเธฑเธ”เธเธฒเธฃ</th><th>เนเธเนเนเธ</th></tr></thead>
                      <tbody>
                        {displayedRequests.length === 0 ? (
                          <tr><td colSpan="8" className="workflow-empty">เธขเธฑเธเนเธกเนเธเธเธเธณเธเธญเน€เธเธดเธเธญเธธเธเธเธฃเธ“เน</td></tr>
                        ) : displayedRequests.map(request => (
                          <tr key={request.id}>
                            <td><strong>#{request.id}</strong><small>{request.requested_date || request.created_at ? new Date(request.requested_date || request.created_at).toLocaleDateString('th-TH') : '-'}</small></td>
                            <td><strong>{request.requester}</strong><small>{request.department}</small></td>
                            <td><strong>{request.item_type}</strong><small>{request.purpose}</small></td>
                            <td>{request.due_date ? String(request.due_date).slice(0, 10) : '-'}</td>
                            <td>{request.device_serial ? <><strong>{request.device_serial}</strong><small>{request.assigned_item_type}</small></> : '-'}</td>
                            <td><span className={`workflow-status status-${request.status}`}>{statusLabels[request.status] || request.status}</span></td>
                            <td>
                              {assetWorkflowRole === 'requester' ? (
                                <span className="workflow-done">-</span>
                              ) : (
                              <div className="workflow-actions">
                                {['pending', 'need_info'].includes(request.status) && <><button onClick={() => runAssetRequestAction(request, 'approve')} disabled={assetRequestLoading}>เธญเธเธธเธกเธฑเธ•เธด/เน€เธฅเธทเธญเธเน€เธเธฃเธทเนเธญเธ</button><button className="danger" onClick={() => runAssetRequestAction(request, 'reject')} disabled={assetRequestLoading}>เนเธกเนเธญเธเธธเธกเธฑเธ•เธด</button></>}
                                {request.status === 'approved' && <button onClick={() => runAssetRequestAction(request, 'issue')} disabled={assetRequestLoading}>เธชเนเธเธกเธญเธ</button>}
                                {['returned', 'rejected'].includes(request.status) && <span className="workflow-done">เน€เธชเธฃเนเธเธชเธดเนเธ</span>}
                                {['issued', 'overdue'].includes(request.status) && <span className="workflow-done">เธเธณเธฅเธฑเธเนเธเนเธเธฒเธ</span>}
                              </div>
                              )}
                            </td>
                            <td>
                              <button
                                className="workflow-edit-btn"
                                onClick={() => editAssetRequest(request)}
                                disabled={assetRequestLoading}
                              >
                                Edit
                              </button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>
              </div>
            </div>
          </div>
        );
      })()}

      {activeModal === 'assetReturns' && (() => {
        const statusLabels = {
          pending: 'เธฃเธญเธญเธเธธเธกเธฑเธ•เธด',
          approved: 'เธฃเธญเธชเนเธเธกเธญเธ',
          issued: 'เธชเนเธเธกเธญเธเนเธฅเนเธง',
          rejected: 'เนเธกเนเธญเธเธธเธกเธฑเธ•เธด',
          overdue: 'เน€เธเธดเธเธเธณเธซเธเธ”',
          return_requested: 'เธฃเธญ IT เธ•เธฃเธงเธเธฃเธฑเธ',
          returned: 'เธเธทเธเนเธฅเนเธง'
        };
        const search = assetReturnSearch.trim().toLocaleLowerCase('th-TH');
        const normalizeReturnIdentity = value => String(value || '')
          .replace(/[\u200B-\u200D\uFEFF]/g, '')
          .trim()
          .replace(/\s+/g, ' ')
          .toLocaleLowerCase('th-TH');
        const getReturnIdentityKeys = value => {
          const raw = String(value || '').replace(/[\u200B-\u200D\uFEFF]/g, '');
          const keys = [normalizeReturnIdentity(raw)];
          const nameBeforeNickname = raw.split(/[๏ผ(]/, 1)[0];
          keys.push(normalizeReturnIdentity(nameBeforeNickname));
          for (const match of raw.matchAll(/[๏ผ(]([^()๏ผ๏ผ]+)[)๏ผ]/g)) {
            keys.push(normalizeReturnIdentity(match[1]));
          }
          return new Set(keys.filter(Boolean));
        };
        const identity = normalizeReturnIdentity(assetReturnIdentity);
        const returnRequests = assetRequests
          .filter(request => ['issued', 'overdue', 'return_requested', 'returned'].includes(request.status))
          .filter(request => {
            if (identity) return getReturnIdentityKeys(request.requester).has(identity);
            if (!search) return false;
            return [
              request.requester,
              request.department,
              request.item_type,
              request.device_serial,
              request.assigned_item_type
            ].some(value => String(value || '').toLocaleLowerCase('th-TH').includes(search));
          })
          .sort((a, b) => {
            const aDone = a.status === 'returned' ? 1 : 0;
            const bDone = b.status === 'returned' ? 1 : 0;
            return aDone - bDone;
          });
        const waitingCount = returnRequests.filter(request => ['issued', 'overdue'].includes(request.status)).length;
        const inspectionRequests = assetRequests.filter(request => request.status === 'return_requested');
        const uniquePositions = Array.from(new Set(assetsList.map(asset => asset.position).filter(Boolean))).sort();
        const uniqueStatuses = Array.from(new Set(assetsList.map(asset => asset.status).filter(Boolean))).sort();
        const registryAssets = assetsList.filter(asset => {
          const query = assetSearch.toLocaleLowerCase('th-TH');
          const matchesSearch = [
            asset.sn,
            asset.date,
            asset.user,
            asset.position,
            asset.itemType,
            asset.deviceSerial,
            asset.additionalEquipment,
            asset.additionalSerial,
            asset.softwareApp,
            asset.registeredEmail,
            asset.returnDueDate,
            asset.status,
            asset.notes,
            asset.inspectionDate,
            asset.purchaseDate,
            asset.warrantyEndDate,
            asset.expense
          ].some(value => String(value || '').toLocaleLowerCase('th-TH').includes(query));
          return matchesSearch
            && (!assetDeptFilter || asset.position === assetDeptFilter)
            && (!assetStatusFilter || asset.status === assetStatusFilter);
        });

        return (
          <div className="modal-overlay active">
            <div className="modal large dashboard-fullscreen-modal asset-workflow-modal asset-returns-modal">
              <header className="modal-header">
                <div>
                  <h3>โฉ๏ธ เธเธทเธเธญเธธเธเธเธฃเธ“เน</h3>
                  <p className="workflow-subtitle">เธเนเธเธซเธฒเธญเธธเธเธเธฃเธ“เนเธ—เธตเนเธฃเธฑเธเนเธเนเธฅเนเธง เธขเธทเธเธขเธฑเธเธเธฒเธฃเธชเนเธเธเธทเธ เนเธฅเธฐเธ•เธฃเธงเธเธชเธญเธเธเธฃเธฐเธงเธฑเธ•เธดเธเธฒเธฃเธเธทเธ</p>
                </div>
                <div className="workflow-header-actions">
                  <div className="return-view-switch">
                    <button className={assetReturnView === 'returns' ? 'active' : ''} onClick={() => setAssetReturnView('returns')}>เนเธเนเธเธเธญเธเธทเธ</button>
                    <button className={assetReturnView === 'registry' ? 'active' : ''} onClick={() => setAssetReturnView('registry')}>เธ—เธฐเน€เธเธตเธขเธเธ—เธฃเธฑเธเธขเนเธชเธดเธเธ—เธฑเนเธเธซเธกเธ”</button>
                    <button
                      className={assetReturnView === 'inspection' ? 'active it-inspection' : 'it-inspection'}
                      onClick={() => requireAdminAccess(() => setAssetReturnView('inspection'))}
                    >
                      IT เธ•เธฃเธงเธเธฃเธฑเธ
                    </button>
                  </div>
                  <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
                </div>
              </header>
              <div className="modal-body">
                {assetReturnView === 'returns' && <section className="workflow-list-section">
                  <div className="workflow-section-heading">
                    <div>
                      <h4>เธฃเธฒเธขเธเธฒเธฃเธญเธธเธเธเธฃเธ“เนเธเธญเธเธเธนเนเธเธทเธ</h4>
                      <label className="return-identity-field">
                        <span>เธขเธทเธเธขเธฑเธเธ•เธฑเธงเธเธนเนเธเธทเธเธ”เนเธงเธขเธเธทเนเธญ-เธเธฒเธกเธชเธเธธเธฅ เธซเธฃเธทเธญเธเธทเนเธญเน€เธฅเนเธเนเธเธงเธเน€เธฅเนเธ</span>
                        <input
                          className="workflow-requester-search"
                          placeholder="เธเธฃเธญเธเธเธทเนเธญ-เธเธฒเธกเธชเธเธธเธฅ เธซเธฃเธทเธญเธเธทเนเธญเน€เธฅเนเธ เน€เธเนเธ เธเธญเธ"
                          value={assetReturnIdentity}
                          onChange={event => setAssetReturnIdentity(event.target.value)}
                        />
                      </label>
                      <input
                        className="workflow-requester-search"
                        placeholder="เธเนเธเธซเธฒเนเธเธเธเธซเธฃเธทเธญเธซเธกเธฒเธขเน€เธฅเธเน€เธเธฃเธทเนเธญเธเธเธญเธเธเธธเธ“"
                        value={assetReturnSearch}
                        onChange={event => setAssetReturnSearch(event.target.value)}
                      />
                    </div>
                    <span>เธฃเธญเธเธทเธ {waitingCount} เธฃเธฒเธขเธเธฒเธฃ ยท เธ—เธฑเนเธเธซเธกเธ” {returnRequests.length} เธฃเธฒเธขเธเธฒเธฃ</span>
                  </div>
                  <div className="workflow-table-wrap">
                    <table className="details-table workflow-table asset-return-table">
                      <thead>
                        <tr>
                          <th>เน€เธฅเธเธ—เธตเน</th>
                          <th>เธเธนเนเธเธญ/เนเธเธเธ</th>
                          <th>เธญเธธเธเธเธฃเธ“เนเธ—เธตเนเธฃเธฑเธเนเธ</th>
                          <th>เธเธณเธซเธเธ”เธเธทเธ</th>
                          <th>เธชเธ–เธฒเธเธฐ</th>
                          <th>เธเธฑเธ”เธเธฒเธฃเธเธทเธ</th>
                        </tr>
                      </thead>
                      <tbody>
                        {returnRequests.length === 0 ? (
                          <tr><td colSpan="6" className="workflow-empty">{
                            identity
                              ? 'เนเธกเนเธเธเธญเธธเธเธเธฃเธ“เนเธ—เธตเนเธ•เธฃเธเธเธฑเธเธเธทเนเธญเธเธนเนเธเธทเธ'
                              : search
                                ? 'เนเธกเนเธเธเธฃเธฒเธขเธเธฒเธฃเธ—เธตเนเธ•เธฃเธเธเธฑเธเธเธณเธเนเธเธซเธฒ'
                                : 'เธเธฃเธญเธเธเธทเนเธญเธเธนเนเธเธทเธ เธซเธฃเธทเธญเธเนเธเธซเธฒเธ”เนเธงเธขเธเธทเนเธญ เนเธเธเธ เนเธฅเธฐเธซเธกเธฒเธขเน€เธฅเธเน€เธเธฃเธทเนเธญเธ'
                          }</td></tr>
                        ) : returnRequests.map(request => (
                          <tr key={request.id}>
                            <td>
                              <strong>#{request.id}</strong>
                              <small>{request.requested_date || request.created_at ? new Date(request.requested_date || request.created_at).toLocaleDateString('th-TH') : '-'}</small>
                            </td>
                            <td><strong>{request.requester}</strong><small>{request.department}</small></td>
                            <td>
                              <strong>{request.item_type} {request.device_serial ? `(${request.device_serial})` : ''}</strong>
                              <small>{request.assigned_item_type || request.purpose}</small>
                            </td>
                            <td>{request.due_date ? String(request.due_date).slice(0, 10) : '-'}</td>
                            <td><span className={`workflow-status status-${request.status}`}>{statusLabels[request.status]}</span></td>
                            <td>
                              {['issued', 'overdue'].includes(request.status) ? (
                                <div className="workflow-actions">
                                  <button
                                    className="return"
                                    onClick={() => runAssetRequestAction(request, 'request_return')}
                                    disabled={assetRequestLoading}
                                  >
                                    เนเธเนเธเธเธญเธเธทเธเธญเธธเธเธเธฃเธ“เน
                                  </button>
                                </div>
                              ) : request.status === 'return_requested' ? (
                                <span className="workflow-done">เธฃเธญเน€เธเนเธฒเธซเธเนเธฒเธ—เธตเน IT เธ•เธฃเธงเธเธฃเธฑเธเน€เธเธฃเธทเนเธญเธ</span>
                              ) : (
                                <span className="workflow-done">
                                  เธเธทเธเน€เธกเธทเนเธญ {request.return_date ? new Date(request.return_date).toLocaleDateString('th-TH') : '-'}
                                </span>
                              )}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>}

                {assetReturnView === 'inspection' && <section className="workflow-list-section return-inspection-section">
                  <div className="workflow-section-heading">
                    <div>
                      <h4>เธฃเธฒเธขเธเธฒเธฃเธฃเธญเน€เธเนเธฒเธซเธเนเธฒเธ—เธตเน IT เธ•เธฃเธงเธเธฃเธฑเธ</h4>
                      <p className="workflow-subtitle">เธ•เธฃเธงเธเธชเธ เธฒเธเน€เธเธฃเธทเนเธญเธเธเนเธญเธเธขเธทเธเธขเธฑเธ เธฃเธฐเธเธเธเธถเธเธเธฐเน€เธเธฅเธตเนเธขเธเธชเธ–เธฒเธเธฐเธเธฅเธฑเธเน€เธเนเธ เธงเนเธฒเธ / เธฃเธญเธเนเธญเธก / เธชเธนเธเธซเธฒเธข</p>
                    </div>
                    <span>เธฃเธญเธ•เธฃเธงเธเธฃเธฑเธ {inspectionRequests.length} เธฃเธฒเธขเธเธฒเธฃ</span>
                  </div>
                  <div className="workflow-table-wrap">
                    <table className="details-table workflow-table asset-return-table">
                      <thead>
                        <tr>
                          <th>เน€เธฅเธเธ—เธตเน</th>
                          <th>เธเธนเนเธเธทเธ/เนเธเธเธ</th>
                          <th>เธญเธธเธเธเธฃเธ“เน</th>
                          <th>เธเธณเธซเธเธ”เธเธทเธ</th>
                          <th>เธชเธ–เธฒเธเธฐ</th>
                          <th>เธ•เธฃเธงเธเธฃเธฑเธ</th>
                        </tr>
                      </thead>
                      <tbody>
                        {inspectionRequests.length === 0 ? (
                          <tr><td colSpan="6" className="workflow-empty">เนเธกเนเธกเธตเธญเธธเธเธเธฃเธ“เนเธฃเธญเธ•เธฃเธงเธเธฃเธฑเธ</td></tr>
                        ) : inspectionRequests.map(request => (
                          <tr key={request.id}>
                            <td><strong>#{request.id}</strong><small>{request.requested_date || request.created_at ? new Date(request.requested_date || request.created_at).toLocaleDateString('th-TH') : '-'}</small></td>
                            <td><strong>{request.requester}</strong><small>{request.department}</small></td>
                            <td><strong>{request.device_serial || request.item_type}</strong><small>{request.assigned_item_type || request.purpose}</small></td>
                            <td>{request.due_date ? String(request.due_date).slice(0, 10) : '-'}</td>
                            <td><span className="workflow-status status-return_requested">เธฃเธญ IT เธ•เธฃเธงเธเธฃเธฑเธ</span></td>
                            <td>
                              <div className="workflow-actions">
                                <button className="return" onClick={() => runAssetRequestAction(request, 'return')} disabled={assetRequestLoading}>
                                  เธ•เธฃเธงเธเธชเธ เธฒเธเนเธฅเธฐเธฃเธฑเธเธเธทเธ
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>}

                {assetReturnView === 'registry' && <section className="workflow-list-section return-registry-section">
                  <div className="workflow-section-heading">
                    <div>
                      <h4>เธ—เธฐเน€เธเธตเธขเธเธเธฅเธฑเธเธ—เธฃเธฑเธเธขเนเธชเธดเธเนเธฅเธฐเธญเธธเธเธเธฃเธ“เน IT (Asset Registry)</h4>
                      <input
                        className="workflow-requester-search"
                        placeholder="เธเนเธเธซเธฒเธญเธธเธเธเธฃเธ“เน เธเธทเนเธญเธเธนเนเนเธเน เนเธเธเธ เธซเธฃเธทเธญเธซเธกเธฒเธขเน€เธฅเธเน€เธเธฃเธทเนเธญเธ"
                        value={assetSearch}
                        onChange={event => setAssetSearch(event.target.value)}
                      />
                    </div>
                    <span>เนเธชเธ”เธ {registryAssets.length} เธเธฒเธเธ—เธฑเนเธเธซเธกเธ” {assetsList.length} เธญเธธเธเธเธฃเธ“เน</span>
                  </div>
                  <div className="return-registry-filters">
                    <select value={assetDeptFilter} onChange={event => setAssetDeptFilter(event.target.value)}>
                      <option value="">เธ—เธฑเนเธเธซเธกเธ”เนเธเธเธ</option>
                      {uniquePositions.map(position => <option key={position} value={position}>{position}</option>)}
                    </select>
                    <select value={assetStatusFilter} onChange={event => setAssetStatusFilter(event.target.value)}>
                      <option value="">เธ—เธฑเนเธเธซเธกเธ”เธชเธ–เธฒเธเธฐ</option>
                      {uniqueStatuses.map(status => <option key={status} value={status}>{status}</option>)}
                    </select>
                  </div>
                  <div className="workflow-table-wrap return-registry-table-wrap">
                    <table className="details-table return-registry-table">
                      <thead>
                        <tr>
                          <th>เธฅเธณเธ”เธฑเธเธ—เธตเน</th>
                          <th>เธงเธฑเธเธ—เธตเนเน€เธเธดเธเนเธเนเธเธฒเธ</th>
                          <th>เธเธนเนเน€เธเธดเธเนเธเนเธเธฒเธ</th>
                          <th>เธ•เธณเนเธซเธเนเธ/เนเธเธเธ</th>
                          <th>เธฃเธฒเธขเธเธฒเธฃเธญเธธเธเธเธฃเธ“เนเธซเธฅเธฑเธ</th>
                          <th>เธซเธกเธฒเธขเน€เธฅเธเธญเธธเธเธเธฃเธ“เน (Serial)</th>
                          <th>เธญเธธเธเธเธฃเธ“เนเน€เธเธดเนเธกเน€เธ•เธดเธก</th>
                          <th>เธซเธกเธฒเธขเน€เธฅเธเธญเธธเธเธเธฃเธ“เนเน€เธเธดเนเธกเน€เธ•เธดเธก</th>
                          <th>เธเธญเธเธ•เนเนเธงเธฃเน / App</th>
                          <th>เธญเธตเน€เธกเธฅเธ—เธตเนเธฅเธเธ—เธฐเน€เธเธตเธขเธ</th>
                          <th>เธเธณเธซเธเธ”เธเธทเธ</th>
                          <th>เธชเธ–เธฒเธเธฐ</th>
                          <th>เธซเธกเธฒเธขเน€เธซเธ•เธธ</th>
                          <th>เธงเธฑเธเธ—เธตเนเธ•เธฃเธงเธเธชเธญเธ</th>
                          <th>เธงเธฑเธเธ—เธตเนเธเธทเนเธญ</th>
                          <th>เธงเธฑเธเธซเธกเธ”เธเธฃเธฐเธเธฑเธ</th>
                          <th>เธเนเธฒเนเธเนเธเนเธฒเธข</th>
                        </tr>
                      </thead>
                      <tbody>
                        {registryAssets.length === 0 ? (
                          <tr><td colSpan="17" className="workflow-empty">เนเธกเนเธเธเธญเธธเธเธเธฃเธ“เนเธ—เธตเนเธ•เธฃเธเธ•เธฒเธกเน€เธเธทเนเธญเธเนเธ</td></tr>
                        ) : registryAssets.map((asset, index) => (
                          <tr key={asset.sn ?? index}>
                            <td><strong>{index + 1}</strong></td>
                            <td>{asset.date || '-'}</td>
                            <td>{asset.user || '-'}</td>
                            <td>{asset.position || '-'}</td>
                            <td><AssetTags value={asset.itemType} /></td>
                            <td><strong>{asset.deviceSerial || '-'}</strong></td>
                            <td><AssetTags value={asset.additionalEquipment} /></td>
                            <td><AssetTags value={asset.additionalSerial} /></td>
                            <td><AssetTags value={asset.softwareApp} /></td>
                            <td>{asset.registeredEmail || '-'}</td>
                            <td>{asset.returnDueDate || '-'}</td>
                            <td><span className="registry-status">{asset.status || '-'}</span></td>
                            <td>{asset.notes || '-'}</td>
                            <td>{asset.inspectionDate || '-'}</td>
                            <td>{asset.purchaseDate || '-'}</td>
                            <td>{asset.warrantyEndDate || '-'}</td>
                            <td>{Number(asset.expense || 0).toLocaleString('th-TH')} เธเธฒเธ—</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>}
              </div>
            </div>
          </div>
        );
      })()}

      {/* MODAL 1: EDIT FORM CONSOLE */}
      {activeModal === 'edit' && (
        <div className="modal-overlay active">
          <div className="modal large">
            <header className="modal-header">
              <h3>โ๏ธ เนเธเนเนเธเธเนเธญเธกเธนเธฅเธเธฃเธฐเธเธณเน€เธ”เธทเธญเธ <span style={{ color: 'var(--primary)' }}>{activeData.monthName}</span></h3>
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>
            <div className="modal-body">
              <form onSubmit={handleEditSubmit}>
                
                {/* Hardware inputs */}
                <div className="form-section">
                  <h4 className="form-section-title">1. เธ—เธฃเธฑเธเธขเนเธชเธดเธเนเธฅเธฐเธฎเธฒเธฃเนเธ”เนเธงเธฃเน (Assets)</h4>
                  <div className="form-grid">
                    <div className="form-group">
                      <label>เธเธณเธเธงเธเธญเธธเธเธเธฃเธ“เนเธ—เธฑเนเธเธซเธกเธ” (เน€เธเธฃเธทเนเธญเธ)</label>
                      <input 
                        type="number" 
                        value={formInputs.totalAssets ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, totalAssets: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>เธกเธนเธฅเธเนเธฒเธ—เธฃเธฑเธเธขเนเธชเธดเธเนเธญเธ—เธตเธฃเธงเธก (เธเธฒเธ—)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetValue ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetValue: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>เธญเธธเธเธเธฃเธ“เนเนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ (เน€เธเธฃเธทเนเธญเธ)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetsExpiring ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetsExpiring: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>เธญเธธเธเธเธฃเธ“เนเธเธณเธฃเธธเธ” (เน€เธเธฃเธทเนเธญเธ)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetsBroken ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetsBroken: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>เธญเธธเธเธเธฃเธ“เนเธชเธนเธเธซเธฒเธข (เน€เธเธฃเธทเนเธญเธ)</label>
                      <input 
                        type="number" 
                        value={formInputs.assetsLost ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, assetsLost: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>เธญเธธเธเธเธฃเธ“เนเธงเนเธฒเธ/เธเธฃเนเธญเธกเนเธเน (เน€เธเธฃเธทเนเธญเธ)</label>
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
                  <h4 className="form-section-title">2. เธเธฃเธดเธเธฒเธฃเธเนเธงเธขเน€เธซเธฅเธทเธญเธเธนเนเนเธเน (Support)</h4>
                  <div className="form-grid">
                    <div className="form-group">
                      <label>เธเธณเธเธงเธ Ticket เธ—เธฑเนเธเธซเธกเธ” (เนเธ)</label>
                      <input 
                        type="number" 
                        value={formInputs.ticketsCount ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, ticketsCount: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>เธชเธ–เธดเธ•เธดเธเธฒเธฃเธเธฃเธฃเธฅเธธเธเนเธญเธ•เธเธฅเธ SLA (%)</label>
                      <input 
                        type="number" 
                        step="0.1" 
                        value={formInputs.slaPercent ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, slaPercent: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>Response Time เน€เธเธฅเธตเนเธข (เธเธฒเธ—เธต)</label>
                      <input 
                        type="number" 
                        value={formInputs.responseTime ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, responseTime: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>Resolution Time เน€เธเธฅเธตเนเธข (เธเธฑเนเธงเนเธกเธ)</label>
                      <input 
                        type="number" 
                        step="0.1" 
                        value={formInputs.resolutionTime ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, resolutionTime: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group full-width">
                      <label>เธเธฐเนเธเธเธเธงเธฒเธกเธเธถเธเธเธญเนเธเธเธนเนเนเธเน CSAT (เธเธฐเนเธเธเน€เธ•เนเธก 5)</label>
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
                  <h4 className="form-section-title">3. เธฅเธดเธเธชเธดเธ—เธเธดเนเธเธญเธเธ•เนเนเธงเธฃเน (Software)</h4>
                  <div className="form-grid">
                    <div className="form-group">
                      <label>เนเธเธฃเนเธเธฃเธกเธเธญเธเธ•เนเนเธงเธฃเนเธ—เธฑเนเธเธซเธกเธ”</label>
                      <input 
                        type="number" 
                        value={formInputs.totalSoftware ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, totalSoftware: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>เนเธเธฃเนเธเธฃเธกเนเธเธฅเนเธชเธฑเธเธเธฒเธซเธกเธ”เธชเธฑเธเธเธฒ</label>
                      <input 
                        type="number" 
                        value={formInputs.softwareExpiring ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, softwareExpiring: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>เธชเธดเธ—เธเธดเน/เธเธฑเธเธเธตเนเธเนเธเธฒเธเธญเธขเธนเน (Licenses)</label>
                      <input 
                        type="number" 
                        value={formInputs.licensesInUse ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, licensesInUse: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>เธชเธดเธ—เธเธดเน/เธเธฑเธเธเธตเธงเนเธฒเธ (Licenses)</label>
                      <input 
                        type="number" 
                        value={formInputs.licensesVacant ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, licensesVacant: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group full-width">
                      <label>เธเนเธฒเนเธเนเธเนเธฒเธขเธเธญเธเธ•เนเนเธงเธฃเนเธฃเธงเธก (เธเธฒเธ—)</label>
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
                  <h4 className="form-section-title">4. เธเธงเธฒเธกเธเธฅเธญเธ”เธ เธฑเธขเธเนเธญเธกเธนเธฅเนเธฅเธฐเธเธฒเธฃเธเธนเนเธเธทเธ (Security)</h4>
                  <div className="form-grid">
                    <div className="form-group">
                      <label>เธเธฒเธฃเธชเธณเธฃเธญเธเธเนเธญเธกเธนเธฅ (Backup) เธชเธณเน€เธฃเนเธ (%)</label>
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
                      <label>เธ เธฑเธขเธเธธเธเธเธฒเธก Security Incident (เธเธฃเธฑเนเธ)</label>
                      <input 
                        type="number" 
                        value={formInputs.securityIncidents ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, securityIncidents: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>เธเธงเธฒเธกเธเธธเนเธกเธเธฃเธญเธเธฃเธฐเธเธ Antivirus (%)</label>
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
                      <label>เธเธฒเธฃเน€เธเธดเธ”เนเธเนเธเธฒเธ MFA เนเธเธฃเธฐเธเธเธซเธฅเธฑเธ (%)</label>
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
                  <h4 className="form-section-title">5. เธเธฒเธฃเนเธเนเธเธชเนเธเธเนเธญเธกเนเธฅเธฐเธเนเธฒเธเธณเธฃเธธเธเธฃเธฑเธเธฉเธฒ (Repair)</h4>
                  <div className="form-grid">
                    <div className="form-group">
                      <label>เธเธณเธเธงเธเธเธฃเธฑเนเธเธเธฒเธฃเธชเนเธเธเนเธญเธก</label>
                      <input 
                        type="number" 
                        value={formInputs.repairCount ?? ''} 
                        onChange={(e) => setFormInputs(p => ({ ...p, repairCount: e.target.value }))}
                        required 
                      />
                    </div>
                    <div className="form-group">
                      <label>เธเนเธฒเธเนเธญเธกเนเธเธกเนเธฅเธฐเธเธณเธฃเธธเธเธฃเธฑเธเธฉเธฒเธญเธธเธเธเธฃเธ“เนเธฃเธงเธก (เธเธฒเธ—)</label>
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
                  <h4 className="form-section-title">6. เธเนเธญเน€เธชเธเธญเนเธเธฐเนเธฅเธฐเธชเธฃเธธเธเน€เธเธดเธเธเธนเนเธเธฃเธดเธซเธฒเธฃ (Recommendations)</h4>
                  <div className="form-group full-width">
                    <label>เธเนเธญเน€เธชเธเธญเนเธเธฐ (เนเธชเนเธซเธเธถเนเธเธเนเธญเน€เธชเธเธญเนเธเธฐเธ•เนเธญ 1 เธเธฃเธฃเธ—เธฑเธ”)</label>
                    <textarea 
                      rows="4" 
                      value={formInputs.recommendations || ''} 
                      onChange={(e) => setFormInputs(p => ({ ...p, recommendations: e.target.value }))}
                      placeholder="เนเธเธฐเธเธณเนเธซเนเธเธฃเธฑเธเธเธฃเธธเธเธเธฒเธฃ..."
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
                    เธขเธเน€เธฅเธดเธ
                  </button>
                  <button 
                    type="submit" 
                    className="sidebar-btn" 
                    style={{ width: 'auto', padding: '10px 24px' }}
                  >
                    เธเธฑเธเธ—เธถเธเธเธฒเธฃเนเธเนเนเธ
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
          <div className="modal large dashboard-fullscreen-modal">
            <header className="modal-header">
              <h3>เธฃเธฒเธขเธฅเธฐเน€เธญเธตเธขเธ”เธญเธธเธเธเธฃเธ“เนเนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ ({activeData.monthName})</h3>
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>
            <div className="modal-body">
              <div style={{ overflowX: 'auto' }}>
                <table className="details-table">
                  <thead>
                    <tr>
                      <th>เธฃเธซเธฑเธชเธ—เธฃเธฑเธเธขเนเธชเธดเธ</th>
                      <th>เธเธฃเธฐเน€เธ เธ—</th>
                      <th>เธฃเธธเนเธ</th>
                      <th>เนเธเธเธ</th>
                      <th>เธงเธฑเธเธ—เธตเนเธซเธกเธ”เธญเธฒเธขเธธ</th>
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
                        <td colSpan="5" style={{ textAlign: 'center' }}>เนเธกเนเธกเธตเธเนเธญเธกเธนเธฅเธเธฒเธฃเน€เธ•เธทเธญเธเธซเธกเธ”เธญเธฒเธขเธธเธเธญเธเธฎเธฒเธฃเนเธ”เนเธงเธฃเน</td>
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
          <div className="modal large software-license-modal dashboard-fullscreen-modal">
            <header className="modal-header">
              <h3>เธ—เธฐเน€เธเธตเธขเธเนเธเธฃเนเธเธฃเธกเนเธฅเธฐ License ({activeData.monthName})</h3>
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>
            <div className="modal-body">
              <form onSubmit={saveSoftwareLicense} className="software-license-form">
                <div className="form-grid">
                  <div className="form-group">
                    <label>เธเธทเนเธญเธเธญเธเธ•เนเนเธงเธฃเน/เนเธเธฃเนเธเธฃเธก</label>
                    <input value={softwareName} onChange={(event) => setSoftwareName(event.target.value)} required />
                  </div>
                  <div className="form-group">
                    <label>เธงเธฑเธเธซเธกเธ”เธญเธฒเธขเธธ</label>
                    <input type="date" value={softwareExpiryDate} onChange={(event) => setSoftwareExpiryDate(event.target.value)} />
                  </div>
                  <div className="form-group">
                    <label>Owner / เนเธเธเธ</label>
                    <input value={softwareOwner} onChange={(event) => setSoftwareOwner(event.target.value)} />
                  </div>
                  <div className="form-group">
                    <label>เธเนเธญเธเธ—เธฒเธเธเธณเธฃเธฐเน€เธเธดเธ</label>
                    <input value={softwarePaymentChannel} onChange={(event) => setSoftwarePaymentChannel(event.target.value)} />
                  </div>
                  <div className="form-group">
                    <label>เธงเธฑเธเธ—เธตเน/เธฃเธญเธเธเธณเธฃเธฐเน€เธเธดเธ</label>
                    <input value={softwarePaymentDate} onChange={(event) => setSoftwarePaymentDate(event.target.value)} />
                  </div>
                  <div className="form-group">
                    <label>เธญเธตเน€เธกเธฅเธ—เธตเนเธชเธกเธฑเธเธฃ</label>
                    <input type="email" value={softwareRegisteredEmail} onChange={(event) => setSoftwareRegisteredEmail(event.target.value)} />
                  </div>
                  <div className="form-group">
                    <label>License เนเธเนเธเธฒเธ</label>
                    <input type="number" min="0" value={softwareUsed} onChange={(event) => setSoftwareUsed(event.target.value)} required />
                  </div>
                  <div className="form-group">
                    <label>License เธงเนเธฒเธ</label>
                    <input type="number" min="0" value={softwareVacant} onChange={(event) => setSoftwareVacant(event.target.value)} required />
                  </div>
                  <div className="form-group">
                    <label>เธเนเธฒเนเธเนเธเนเธฒเธขเธ•เนเธญเน€เธ”เธทเธญเธ (เธเธฒเธ—)</label>
                    <input type="number" min="0" value={softwareMonthlyCost} onChange={(event) => setSoftwareMonthlyCost(event.target.value)} />
                  </div>
                  <div className="form-group">
                    <label>เธเธนเนเนเธเนเธเธฒเธเธเธฑเธเธเธธเธเธฑเธ (เธเธฑเนเธเธ”เนเธงเธขเธเธธเธฅเธ เธฒเธ)</label>
                    <textarea value={softwareCurrentUsers} onChange={(event) => setSoftwareCurrentUsers(event.target.value)} />
                  </div>
                </div>
                <div className="software-license-actions">
                  {editingSoftwareIndex !== null && <button type="button" className="btn-details" onClick={resetSoftwareForm}>เธขเธเน€เธฅเธดเธ</button>}
                  <button type="submit" className="btn-save">{editingSoftwareIndex === null ? 'เน€เธเธดเนเธก License' : 'เธเธฑเธเธ—เธถเธเธเธฒเธฃเนเธเนเนเธ'}</button>
                </div>
              </form>
              <div style={{ overflowX: 'auto' }}>
                <table className="details-table">
                  <thead>
                    <tr>
                      <th>เธเธทเนเธญเธเธญเธเธ•เนเนเธงเธฃเน/เนเธเธฃเนเธเธฃเธก</th>
                      <th>Owner</th>
                      <th>เนเธเนเธเธฒเธ</th>
                      <th>เธงเนเธฒเธ</th>
                      <th>เธฃเธงเธก</th>
                      <th>เธฃเธฒเธเธฒ</th>
                      <th>เธเนเธญเธเธ—เธฒเธเธเธณเธฃเธฐ</th>
                      <th>เธงเธฑเธเธ—เธตเนเธเธณเธฃเธฐ</th>
                      <th>เธงเธฑเธเธซเธกเธ”เธชเธฑเธเธเธฒ</th>
                      <th>เธญเธตเน€เธกเธฅเธชเธกเธฑเธเธฃ</th>
                      <th>เธเธนเนเนเธเนเธเธฒเธเธเธฑเธเธเธธเธเธฑเธ</th>
                      <th>เธเธฑเธ”เธเธฒเธฃ</th>
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
                              <button type="button" className="btn-details" onClick={() => editSoftwareLicense(soft, idx)}>เนเธเนเนเธ</button>
                              <button type="button" className="console-delete-btn" onClick={() => deleteSoftwareLicense(idx)}>เธฅเธ</button>
                            </div>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan="12" style={{ textAlign: 'center' }}>เธขเธฑเธเนเธกเนเธกเธตเธฃเธฒเธขเธฅเธฐเน€เธญเธตเธขเธ” License</td>
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
              <h3>เธ—เธณเน€เธเธตเธขเธเธญเธธเธเธเธฃเธ“เนเธเธณเธฃเธธเธ” (Top 10 เธญเธธเธเธเธฃเธ“เนเน€เธชเธตเธขเธเนเธญเธข)</h3>
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>
            <div className="modal-body">
              <div style={{ overflowX: 'auto' }}>
                <table className="details-table">
                  <thead>
                    <tr>
                      <th>เธญเธฑเธเธ”เธฑเธ</th>
                      <th>เธญเธธเธเธเธฃเธ“เน/เธฃเธธเนเธ</th>
                      <th>เธเธณเธเธงเธเธเธฃเธฑเนเธเธ—เธตเนเน€เธชเธตเธข</th>
                      <th>เธเนเธฒเนเธเนเธเนเธฒเธขเนเธเธเธฒเธฃเธเนเธญเธกเธฃเธงเธก (เนเธ”เธขเธเธฃเธฐเธกเธฒเธ“)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {[...activeData.topBrokenDevices].sort((a,b) => b.count - a.count).map((device, idx) => (
                      <tr key={idx}>
                        <td><strong>{idx + 1}</strong></td>
                        <td>{device.name}</td>
                        <td><span className="repair-item-count">{device.count} เธเธฃเธฑเนเธ</span></td>
                        <td>{device.cost > 0 ? formatThaiBaht(device.cost) : '0 เธเธฒเธ—'}</td>
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
          <div className="modal large dashboard-fullscreen-modal">
            <header className="modal-header">
              <h3>เธฃเธฒเธขเธฅเธฐเน€เธญเธตเธขเธ”เธเธฃเธฐเธงเธฑเธ•เธดเธเธฒเธเธเนเธญเธกเนเธฅเธฐเธเธฃเธดเธเธฒเธฃ Support ({activeData.monthName})</h3>
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>
            <div className="modal-body">
              <div style={{ overflowX: 'auto' }}>
                <table className="details-table">
                  <thead>
                    <tr>
                      <th>SN / เธฃเธซเธฑเธช</th>
                      <th>เธงเธฑเธ-เน€เธงเธฅเธฒ</th>
                      <th>เธเธนเนเนเธเนเธ / เธ•เธดเธ”เธ•เนเธญ</th>
                      <th>เธเธฑเธเธซเธฒ / เธญเธฒเธเธฒเธฃเน€เธชเธตเธข / เธฃเธฒเธขเธเธฒเธฃ</th>
                      <th>เธชเธฒเน€เธซเธ•เธธเธเธฒเธฃเน€เธชเธตเธข</th>
                      <th>เธเธนเนเธฃเธฑเธเธเธดเธ”เธเธญเธ (IT)</th>
                      <th>เน€เธงเธฅเธฒเธ—เธณเธเธฒเธ</th>
                      <th>เธชเธ–เธฒเธเธฐ</th>
                      <th>เธเนเธฒเนเธเนเธเนเธฒเธข (เธเธฒเธ—)</th>
                      <th>เธฃเธนเธเนเธเธ</th>
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
                                color: ticket.status === 'เน€เธชเธฃเนเธเธชเธดเนเธ' || ticket.status === 'เธเนเธฒเธขเน€เธเธดเธเนเธฅเนเธง' ? 'var(--success)' : 'var(--warning)', 
                                fontWeight: '600' 
                              }}
                            >
                              {ticket.status}
                            </span>
                          </td>
                          <td>{ticket.cost > 0 ? formatThaiBaht(ticket.cost) : '-'}</td>
                          <td>
                            {(ticket.hasAttachment || ticket.attachmentData) ? (
                              <a
                                className="ticket-attachment-link"
                                href={ticket.attachmentData || `${API_BASE}/api/tickets/${ticket.sn}/attachment?v=2`}
                                target="_blank"
                                rel="noopener noreferrer"
                                title={ticket.attachmentName || 'เน€เธเธดเธ”เธฃเธนเธเนเธเธ'}
                              >
                                <img
                                  src={ticket.attachmentData || `${API_BASE}/api/tickets/${ticket.sn}/attachment?v=2`}
                                  alt={ticket.attachmentName || `เธฃเธนเธเนเธเธ Ticket ${ticket.sn}`}
                                />
                                <span>เธ”เธนเธฃเธนเธ</span>
                              </a>
                            ) : '-'}
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan="10" style={{ textAlign: 'center' }}>เนเธกเนเธกเธตเธเนเธญเธกเธนเธฅเธเธฃเธฐเธงเธฑเธ•เธดเธเธฃเธฐเธงเธฑเธ•เธดเธเธฒเธเธเนเธญเธกเธชเธณเธซเธฃเธฑเธเน€เธ”เธทเธญเธเธเธตเน</td>
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
        const uniqueTypes = Array.from(new Set(assetsList.map(a => a.itemType).filter(Boolean))).sort();

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
          const matchesType = !assetTypeFilter || asset.itemType === assetTypeFilter;
          
          return matchesSearch && matchesDept && matchesStatus && matchesType;
        });

        return (
          <div className="modal-overlay active">
            <div className="modal large dashboard-fullscreen-modal asset-list-modal">
              <header className="modal-header">
                <h3>เธ—เธฐเน€เธเธตเธขเธเธเธฅเธฑเธเธ—เธฃเธฑเธเธขเนเธชเธดเธเนเธฅเธฐเธญเธธเธเธเธฃเธ“เน IT (Asset Registry)</h3>
                <button onClick={() => {
                  setActiveModal(null);
                  setAssetSearch('');
                  setAssetDeptFilter('');
                  setAssetStatusFilter('');
                  setAssetTypeFilter('');
                }} className="modal-close"><X size={20} /></button>
              </header>
              <div className="modal-body asset-list-modal-body">
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
                    <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '4px' }}>เธเนเธเธซเธฒเธญเธธเธเธเธฃเธ“เน / เธเธทเนเธญเธเธนเนเน€เธเธดเธ / เธซเธกเธฒเธขเน€เธฅเธเน€เธเธฃเธทเนเธญเธ</label>
                    <input 
                      type="text"
                      value={assetSearch}
                      onChange={(e) => setAssetSearch(e.target.value)}
                      placeholder="เน€เธเนเธ Lenovo, เธเธทเนเธญเธเธเธฑเธเธเธฒเธ, LENOVO-010..."
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
                    <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '4px' }}>เธเธฃเธญเธเธ•เธฒเธกเธเธฃเธฐเน€เธ เธ—</label>
                    <select
                      value={assetTypeFilter}
                      onChange={(e) => setAssetTypeFilter(e.target.value)}
                      style={{
                        width: '100%',
                        padding: '8px 12px',
                        borderRadius: '6px',
                        border: '1px solid rgba(255, 255, 255, 0.15)',
                        backgroundColor: 'rgba(0, 0, 0, 0.2)',
                        color: 'white'
                      }}
                    >
                      <option value="">เธ—เธฑเนเธเธซเธกเธ”เธเธฃเธฐเน€เธ เธ—</option>
                      {uniqueTypes.map((type, idx) => (
                        <option key={idx} value={type}>{type}</option>
                      ))}
                    </select>
                  </div>
                  <div style={{ width: '200px' }}>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '4px' }}>เธเธฃเธญเธเธ•เธฒเธกเนเธเธเธ/เธ•เธณเนเธซเธเนเธ</label>
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
                      <option value="">เธ—เธฑเนเธเธซเธกเธ”เนเธเธเธ</option>
                      {uniquePositions.map((pos, idx) => (
                        <option key={idx} value={pos}>{pos}</option>
                      ))}
                    </select>
                  </div>
                  <div style={{ width: '150px' }}>
                    <label style={{ display: 'block', fontSize: '0.75rem', color: 'var(--text-muted)', marginBottom: '4px' }}>เธเธฃเธญเธเธ•เธฒเธกเธชเธ–เธฒเธเธฐ</label>
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
                      <option value="">เธ—เธฑเนเธเธซเธกเธ”เธชเธ–เธฒเธเธฐ</option>
                      {uniqueStatuses.map((stat, idx) => (
                        <option key={idx} value={stat}>{stat}</option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="asset-list-table-wrap">
                  <table className="details-table">
                    <thead>
                      <tr>
                        <th>เธฅเธณเธ”เธฑเธเธ—เธตเน</th>
                        <th>เธงเธฑเธเธ—เธตเนเน€เธเธดเธเนเธเนเธเธฒเธ</th>
                        <th>เธเธนเนเน€เธเธดเธเนเธเนเธเธฒเธ</th>
                        <th>เธ•เธณเนเธซเธเนเธ/เนเธเธเธ</th>
                        <th>เธฃเธฒเธขเธเธฒเธฃเธญเธธเธเธเธฃเธ“เนเธซเธฅเธฑเธ</th>
                        <th>เธซเธกเธฒเธขเน€เธฅเธเธญเธธเธเธเธฃเธ“เน (Serial)</th>
                        <th>เธญเธธเธเธเธฃเธ“เนเน€เธเธดเนเธกเน€เธ•เธดเธก</th>
                        <th>เธซเธกเธฒเธขเน€เธฅเธเธญเธธเธเธเธฃเธ“เนเน€เธเธดเนเธกเน€เธ•เธดเธก</th>
                        <th>เธเธญเธเธ•เนเนเธงเธฃเน / App</th>
                        <th>เธญเธตเน€เธกเธฅเธ—เธตเนเธฅเธเธ—เธฐเน€เธเธตเธขเธ</th>
                        <th>เธเธณเธซเธเธ”เธเธทเธ</th>
                        <th>เธชเธ–เธฒเธเธฐ</th>
                        <th>เธซเธกเธฒเธขเน€เธซเธ•เธธ</th>
                        <th>เธงเธฑเธเธ—เธตเนเธ•เธฃเธงเธเธชเธญเธ</th>
                        <th>เธงเธฑเธเธ—เธตเนเธเธทเนเธญ</th>
                        <th>เธงเธฑเธเธซเธกเธ”เธเธฃเธฐเธเธฑเธ</th>
                        <th>เธเนเธฒเนเธเนเธเนเธฒเธข</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredAssetsList.length > 0 ? (
                        filteredAssetsList.map((asset, idx) => (
                          <tr key={idx}>
                            <td><strong>{idx + 1}</strong></td>
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
                                  color: asset.status === 'เนเธเนเธเธฒเธ' ? 'var(--success)' : asset.status === 'เธฃเธญเธเนเธญเธก' ? 'var(--danger)' : 'var(--warning)', 
                                  fontWeight: '600',
                                  padding: '2px 8px',
                                  borderRadius: '4px',
                                  backgroundColor: asset.status === 'เนเธเนเธเธฒเธ' ? 'rgba(16, 185, 129, 0.1)' : asset.status === 'เธฃเธญเธเนเธญเธก' ? 'rgba(239, 68, 68, 0.1)' : 'rgba(245, 158, 11, 0.1)'
                                }}
                              >
                                {asset.status}
                              </span>
                            </td>
                            <td>{asset.notes || '-'}</td>
                            <td>{asset.inspectionDate || '-'}</td>
                            <td>{asset.purchaseDate || '-'}</td>
                            <td>{asset.warrantyEndDate || '-'}</td>
                            <td>{Number(asset.expense || 0).toLocaleString('th-TH')} เธเธฒเธ—</td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan="18" style={{ textAlign: 'center' }}>เนเธกเนเธเธเธเธฅเธฑเธเธญเธธเธเธเธฃเธ“เนเธ—เธตเนเธ•เธฃเธเธ•เธฒเธกเน€เธเธทเนเธญเธเนเธ</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
              <div className="asset-list-table-summary">
                เนเธชเธ”เธ {filteredAssetsList.length} เธเธฒเธเธ—เธฑเนเธเธซเธกเธ” {assetsList.length} เธญเธธเธเธเธฃเธ“เน
              </div>
            </div>
          </div>
        </div>
      );
    })()}

    {/* MODAL 7: FULL ADMINISTRATIVE DATA CUSTOMIZER PANEL */}
    {activeModal === 'fullConsole' && (() => {
      const consoleMonthData = data[consoleMonth] || {};
      const consoleAssetQuery = consoleAssetSearch.trim().toLocaleLowerCase('th-TH');
      const consoleAssets = assetsList
        .filter((asset) => !consoleAssetQuery || [
          asset.sn, asset.submittedOn, asset.respondent, asset.date, asset.user, asset.position,
          asset.itemType, asset.additionalEquipment, asset.softwareApp, asset.registeredEmail,
          asset.deviceSerial, asset.additionalSerial, asset.returnDueDate, asset.status,
          asset.notes, asset.inspectionDate, asset.purchaseDate, asset.warrantyEndDate, asset.expense
        ].some((value) => String(value ?? '').toLocaleLowerCase('th-TH').includes(consoleAssetQuery)))
        .sort((a, b) => String(a.position || '').localeCompare(String(b.position || ''), 'th')
          || Number(a.sn || 0) - Number(b.sn || 0));
      const primaryTagOptions = [...new Set(assetsList.map((asset) => String(asset.itemType || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'th'));
      const additionalTagOptions = [...new Set(assetsList.flatMap((asset) => String(asset.additionalEquipment || '').split(/[,\n]+/).map((tag) => tag.trim()).filter(Boolean)))].sort((a, b) => a.localeCompare(b, 'th'));
      
      return (
        <div className="modal-overlay active full-console-overlay">
          <div className={`modal large full-console-modal ${consoleTab === 'assets' ? 'asset-console-mode' : ''}`}>
            <header className="modal-header console-topbar">
              <div className="console-topbar-title">
                <Database size={22} style={{ color: 'var(--primary)' }} />
                <h3>
                  {consoleTab === 'assets'
                    ? <>เธ—เธฐเน€เธเธตเธขเธเธเธฅเธฑเธเธ—เธฃเธฑเธเธขเนเธชเธดเธเธซเธฅเธฑเธ (IT Asset Registry Editor) - {editingAssetSn !== null ? <span style={{ color: 'var(--warning)' }}>เนเธซเธกเธ”เนเธเนเนเธเธฃเธซเธฑเธช #{editingAssetSn}</span> : <span>เนเธซเธกเธ”เน€เธเธดเนเธกเธเนเธญเธกเธนเธฅ</span>}</>
                    : 'เธฃเธฐเธเธเธเธฑเธ”เธเธฒเธฃเนเธฅเธฐเธเธฃเธฑเธเน€เธเธฅเธตเนเธขเธเธเนเธญเธกเธนเธฅเนเธ”เธเธเธญเธฃเนเธ”เธ—เธฑเนเธเธซเธกเธ”'}
                </h3>
              </div>
              <nav className="console-topbar-nav" aria-label="เน€เธกเธเธนเธเธฑเธ”เธเธฒเธฃเธเนเธญเธกเธนเธฅเนเธ”เธเธเธญเธฃเนเธ”">
                <button onClick={() => setConsoleTab('months')} className={`console-tab-btn ${consoleTab === 'months' ? 'active' : ''}`}>๐“… เธเธฑเธ”เธเธฒเธฃเน€เธ”เธทเธญเธ</button>
                <button onClick={() => setConsoleTab('kpis')} className={`console-tab-btn ${consoleTab === 'kpis' ? 'active' : ''}`}>๐“ เธ•เธฑเธงเธเธตเนเธงเธฑเธ” KPIs</button>
                <button onClick={() => setConsoleTab('projects')} className={`console-tab-btn ${consoleTab === 'projects' ? 'active' : ''}`}>๐—’๏ธ เนเธเธฃเธเธเธฒเธฃ & เธเนเธญเนเธเธฐเธเธณ</button>
                <button onClick={() => setConsoleTab('assets')} className={`console-tab-btn ${consoleTab === 'assets' ? 'active' : ''}`}>๐’ป เธเธฅเธฑเธเธ—เธฃเธฑเธเธขเนเธชเธดเธ IT</button>
                <button onClick={() => setConsoleTab('tickets')} className={`console-tab-btn ${consoleTab === 'tickets' ? 'active' : ''}`}>๐จ เธเธฃเธฐเธงเธฑเธ•เธดเธเธฒเธ Support</button>
                <button onClick={() => setConsoleTab('backup')} className={`console-tab-btn ${consoleTab === 'backup' ? 'active' : ''}`}>๐’พ เธชเธณเธฃเธญเธ & เธฃเธตเน€เธเนเธ•เธฃเธฐเธเธ</button>
              </nav>
              {consoleTab === 'assets' && editingAssetSn === null && (
                <button type="button" onClick={handleAddAsset} className="btn-save asset-add-topbar-btn">เน€เธเธดเนเธกเธ—เธฃเธฑเธเธขเนเธชเธดเธเน€เธเนเธฒเธเธฅเธฑเธ</button>
              )}
              <button onClick={() => setActiveModal(null)} className="modal-close"><X size={20} /></button>
            </header>

            <div className="console-layout">
              {/* Work Area */}
              <div className="console-content">
                
                {/* TAB 1: MONTHS MANAGER */}
                {consoleTab === 'months' && (
                  <div>
                    <h4 className="console-title">๐“… เธเธฑเธ”เธเธฒเธฃเน€เธ”เธทเธญเธเนเธฅเธฐเธฃเธฒเธขเธเธฒเธเนเธเธฃเธฐเธเธ</h4>
                    <form onSubmit={handleAddMonth} className="console-form" style={{ display: 'flex', gap: '16px', alignItems: 'flex-end' }}>
                      <div className="console-field" style={{ flex: '1' }}>
                        <span className="console-label">เธฃเธซเธฑเธชเธเธตเธขเนเน€เธ”เธทเธญเธ (เน€เธเนเธ 2026-08)</span>
                        <input type="text" placeholder="YYYY-MM" value={newMonthKey} onChange={e => setNewMonthKey(e.target.value)} className="console-input" />
                      </div>
                      <div className="console-field" style={{ flex: '1' }}>
                        <span className="console-label">เธเธทเนเธญเนเธชเธ”เธเนเธเธฃเธฒเธขเธเธฒเธ (เน€เธเนเธ เธชเธดเธเธซเธฒเธเธก 2569)</span>
                        <input type="text" placeholder="เธเธทเนเธญเน€เธ”เธทเธญเธ เธ.เธจ." value={newMonthName} onChange={e => setNewMonthName(e.target.value)} className="console-input" />
                      </div>
                      <button type="submit" className="btn-save" style={{ width: 'auto', padding: '8px 20px', height: '38px' }}>เน€เธเธดเนเธกเน€เธ”เธทเธญเธเนเธซเธกเน</button>
                    </form>

                    <div className="console-table-scroll">
                      <table className="details-table">
                        <thead>
                          <tr>
                            <th>เธเธตเธขเนเน€เธ”เธทเธญเธ</th>
                            <th>เธเธทเนเธญเน€เธ”เธทเธญเธ</th>
                            <th>เธชเธ–เธฒเธเธฐเธญเธธเธเธเธฃเธ“เนเธฃเธงเธก</th>
                            <th>เธเธฒเธ Support</th>
                            <th>เธเธฒเธฃเธเธฑเธ”เธเธฒเธฃ</th>
                          </tr>
                        </thead>
                        <tbody>
                          {Object.keys(data).map(key => (
                            <tr key={key}>
                              <td><strong>{key}</strong></td>
                              <td>{data[key].monthName}</td>
                              <td>{data[key].totalAssets} เน€เธเธฃเธทเนเธญเธ</td>
                              <td>{data[key].ticketsCount} เน€เธเธช</td>
                              <td>
                                <button onClick={() => handleDeleteMonth(key)} className="console-delete-btn">เธฅเธ</button>
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
                      <h4 style={{ margin: 0, fontSize: '1.15rem' }}>๐“ เธเธฃเธฑเธเน€เธเธฅเธตเนเธขเธเธเนเธฒเธ•เธฑเธงเธเธตเนเธงเธฑเธ” KPIs เธเธฃเธฐเธเธณเน€เธ”เธทเธญเธ</h4>
                      <div>
                        <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginRight: '8px' }}>เน€เธฅเธทเธญเธเน€เธ”เธทเธญเธเธ—เธตเนเธเธฐเนเธเนเนเธ:</span>
                        <select value={consoleMonth} onChange={e => setConsoleMonth(e.target.value)} style={{ padding: '6px 12px', borderRadius: '6px', border: '1px solid rgba(255,255,255,0.15)', backgroundColor: '#1f2937', color: 'white' }}>
                          {Object.keys(data).map(key => (
                            <option key={key} value={key}>{data[key].monthName}</option>
                          ))}
                        </select>
                      </div>
                    </div>

                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '16px' }}>
                      <div className="console-card">
                        <h5 style={{ margin: '0 0 12px 0', color: 'var(--primary)' }}>๐’ป เธ—เธฃเธฑเธเธขเนเธชเธดเธ IT</h5>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                          <div className="console-field">
                            <span className="console-label">เธเธณเธเธงเธเธญเธธเธเธเธฃเธ“เนเธ—เธฑเนเธเธซเธกเธ” (เน€เธเธฃเธทเนเธญเธ)</span>
                            <input type="number" value={consoleMonthData.totalAssets || 0} onChange={e => handleKpiChange('totalAssets', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เธกเธนเธฅเธเนเธฒเธเธฅเธฑเธเธฃเธงเธก (เธเธฒเธ—)</span>
                            <input type="number" value={consoleMonthData.assetValue || 0} onChange={e => handleKpiChange('assetValue', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เนเธเธฅเนเธซเธกเธ”เธญเธฒเธขเธธ (เน€เธเธฃเธทเนเธญเธ)</span>
                            <input type="number" value={consoleMonthData.assetsExpiring || 0} onChange={e => handleKpiChange('assetsExpiring', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เธเธณเธฃเธธเธ” (เน€เธเธฃเธทเนเธญเธ)</span>
                            <input type="number" value={consoleMonthData.assetsBroken || 0} onChange={e => handleKpiChange('assetsBroken', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เธชเธนเธเธซเธฒเธข (เน€เธเธฃเธทเนเธญเธ)</span>
                            <input type="number" value={consoleMonthData.assetsLost || 0} onChange={e => handleKpiChange('assetsLost', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เน€เธเธฃเธทเนเธญเธเธงเนเธฒเธ (เน€เธเธฃเธทเนเธญเธ)</span>
                            <input type="number" value={consoleMonthData.assetsVacant || 0} onChange={e => handleKpiChange('assetsVacant', Number(e.target.value))} className="console-input" />
                          </div>
                        </div>
                      </div>

                      <div className="console-card">
                        <h5 style={{ margin: '0 0 12px 0', color: 'var(--violet)' }}>๐จ เธเธฃเธดเธเธฒเธฃ Support & SLA</h5>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                          <div className="console-field">
                            <span className="console-label">เน€เธเธชเนเธเนเธเน€เธชเธตเธขเธ—เธฑเนเธเธซเธกเธ” (เน€เธเธช)</span>
                            <input type="number" value={consoleMonthData.ticketsCount || 0} onChange={e => handleKpiChange('ticketsCount', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เธญเธฑเธ•เธฃเธฒเธ—เธณเนเธ”เนเธ•เธฒเธก SLA (%)</span>
                            <input type="number" step="0.1" value={consoleMonthData.slaPercent || 0} onChange={e => handleKpiChange('slaPercent', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เน€เธเธฅเธตเนเธขเน€เธงเธฅเธฒเธฃเธฑเธเน€เธฃเธทเนเธญเธ (เธเธฒเธ—เธต)</span>
                            <input type="number" value={consoleMonthData.responseTime || 0} onChange={e => handleKpiChange('responseTime', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เน€เธเธฅเธตเนเธขเน€เธงเธฅเธฒเนเธเนเนเธเธเธฑเธเธซเธฒ (เธเธฒเธ—เธต)</span>
                            <input type="number" value={consoleMonthData.resolutionTime || 0} onChange={e => handleKpiChange('resolutionTime', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เธเธงเธฒเธกเธเธถเธเธเธญเนเธเธฅเธนเธเธเนเธฒ CSAT (เน€เธ•เนเธก 5)</span>
                            <input type="number" step="0.01" value={consoleMonthData.csat || 0} onChange={e => handleKpiChange('csat', Number(e.target.value))} className="console-input" />
                          </div>
                        </div>
                      </div>

                      <div className="console-card">
                        <h5 style={{ margin: '0 0 12px 0', color: 'var(--secondary)' }}>๐’ฟ เธฅเธดเธเธชเธดเธ—เธเธดเนเธเธญเธเธ•เนเนเธงเธฃเน</h5>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                          <div className="console-field">
                            <span className="console-label">เธเธฃเธฐเน€เธ เธ—เธเธญเธเธ•เนเนเธงเธฃเนเธฅเธดเธเธชเธดเธ—เธเธดเน</span>
                            <input type="number" value={consoleMonthData.totalSoftware || 0} onChange={e => handleKpiChange('totalSoftware', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เน€เธเธดเธ”เนเธเนเธเธฒเธเธญเธขเธนเน (เธชเธดเธ—เธเธดเน)</span>
                            <input type="number" value={consoleMonthData.licensesInUse || 0} onChange={e => handleKpiChange('licensesInUse', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เธชเธดเธ—เธเธดเนเธงเนเธฒเธเธเธเน€เธซเธฅเธทเธญ (เธชเธดเธ—เธเธดเน)</span>
                            <input type="number" value={consoleMonthData.licensesVacant || 0} onChange={e => handleKpiChange('licensesVacant', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เธเนเธฒเธเธญเธเธ•เนเนเธงเธฃเนเธฃเธฒเธขเน€เธ”เธทเธญเธ (เธเธฒเธ—)</span>
                            <input type="number" value={consoleMonthData.softwareCost || 0} onChange={e => handleKpiChange('softwareCost', Number(e.target.value))} className="console-input" />
                          </div>
                        </div>
                      </div>

                      <div className="console-card">
                        <h5 style={{ margin: '0 0 12px 0', color: 'var(--success)' }}>๐ก๏ธ เธเธงเธฒเธกเธเธฅเธญเธ”เธ เธฑเธข IT & Repairs</h5>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                          <div className="console-field">
                            <span className="console-label">เธชเธณเธฃเธญเธเธเนเธญเธกเธนเธฅเธชเธณเน€เธฃเนเธ (%)</span>
                            <input type="number" value={consoleMonthData.backupSuccess || 0} onChange={e => handleKpiChange('backupSuccess', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เน€เธซเธ•เธธเธเธงเธฒเธกเธเธฅเธญเธ”เธ เธฑเธข (เธเธฃเธฑเนเธ)</span>
                            <input type="number" value={consoleMonthData.securityIncidents || 0} onChange={e => handleKpiChange('securityIncidents', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เธเธเธเธฃเธฐเธกเธฒเธ“เธชเนเธเธเนเธญเธก (เธเธฒเธ—)</span>
                            <input type="number" value={consoleMonthData.repairCost || 0} onChange={e => handleKpiChange('repairCost', Number(e.target.value))} className="console-input" />
                          </div>
                          <div className="console-field">
                            <span className="console-label">เธเธณเธเธงเธเธเธดเนเธเธ—เธตเนเธชเนเธเธเนเธญเธก (เน€เธเธฃเธทเนเธญเธ)</span>
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
                      <h4 style={{ margin: 0, fontSize: '1.15rem' }}>๐—’๏ธ เธเธฑเธ”เธเธฒเธฃเนเธเธฃเธเธเธฒเธฃ & เธเนเธญเน€เธชเธเธญเนเธเธฐเธชเธณเธซเธฃเธฑเธเธเธฑเธ’เธเธฒ</h4>
                      <div>
                        <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginRight: '8px' }}>เน€เธฅเธทเธญเธเน€เธ”เธทเธญเธ:</span>
                        <select value={consoleMonth} onChange={e => setConsoleMonth(e.target.value)} style={{ padding: '6px 12px', borderRadius: '6px', border: '1px solid rgba(255,255,255,0.15)', backgroundColor: '#1f2937', color: 'white' }}>
                          {Object.keys(data).map(key => (
                            <option key={key} value={key}>{data[key].monthName}</option>
                          ))}
                        </select>
                      </div>
                    </div>

                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                      <div>
                        <h5 style={{ margin: '0 0 10px 0', color: 'var(--primary)' }}>เนเธเธฃเธเธเธฒเธฃเธ—เธตเนเธ”เธณเน€เธเธดเธเธเธฒเธฃเธญเธขเธนเน (Ongoing Projects)</h5>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginBottom: '16px' }}>
                          {(consoleMonthData.ongoingProjects || []).length > 0 ? (
                            (consoleMonthData.ongoingProjects || []).map((proj, idx) => (
                              <div key={idx} style={{ padding: '10px', backgroundColor: 'rgba(255,255,255,0.02)', border: '1px solid rgba(255,255,255,0.06)', borderRadius: '6px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                <div>
                                  <div style={{ fontWeight: '600', fontSize: '0.85rem' }}>{proj.title}</div>
                                  <div style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>{proj.desc}</div>
                                </div>
                                <button onClick={() => handleDeleteProject(idx)} style={{ border: 'none', background: 'none', color: 'red', cursor: 'pointer', fontSize: '1.1rem' }}>ร—</button>
                              </div>
                            ))
                          ) : (
                            <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', fontStyle: 'italic' }}>เนเธกเนเธกเธตเธฃเธฒเธขเธเธฒเธฃเนเธเธฃเธเธเธฒเธฃเน€เธ”เธทเธญเธเธเธตเน</div>
                          )}
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', padding: '12px', backgroundColor: 'rgba(0,0,0,0.2)', borderRadius: '8px' }}>
                          <input type="text" placeholder="เธเธทเนเธญเนเธเธฃเธเธเธฒเธฃ" value={newProjTitle} onChange={e => setNewProjTitle(e.target.value)} className="console-input" />
                          <input type="text" placeholder="เธเธงเธฒเธกเธเธทเธเธซเธเนเธฒ" value={newProjDesc} onChange={e => setNewProjDesc(e.target.value)} className="console-input" />
                          <button type="button" onClick={handleAddProject} className="btn-save" style={{ width: '100%', padding: '6px' }}>เน€เธเธดเนเธกเนเธเธฃเธเธเธฒเธฃ</button>
                        </div>
                      </div>

                      <div>
                        <h5 style={{ margin: '0 0 10px 0', color: 'var(--warning)' }}>เธเนเธญเน€เธชเธเธญเนเธเธฐเน€เธเธดเธเธงเธดเน€เธเธฃเธฒเธฐเธซเน (Recommendations)</h5>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginBottom: '16px' }}>
                          {(consoleMonthData.recommendations || []).length > 0 ? (
                            (consoleMonthData.recommendations || []).map((rec, idx) => (
                              <div key={idx} style={{ padding: '10px', backgroundColor: 'rgba(255,255,255,0.02)', border: '1px solid rgba(255,255,255,0.06)', borderRadius: '6px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                <div style={{ fontSize: '0.85rem' }}>{rec}</div>
                                <button onClick={() => handleDeleteRecommendation(idx)} style={{ border: 'none', background: 'none', color: 'red', cursor: 'pointer', fontSize: '1.1rem' }}>ร—</button>
                              </div>
                            ))
                          ) : (
                            <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', fontStyle: 'italic' }}>เนเธกเนเธกเธตเธเนเธญเน€เธชเธเธญเนเธเธฐเธชเธณเธซเธฃเธฑเธเน€เธ”เธทเธญเธเธเธตเน</div>
                          )}
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', padding: '12px', backgroundColor: 'rgba(0,0,0,0.2)', borderRadius: '8px' }}>
                          <textarea rows="2" placeholder="เธเธฃเธญเธเธเนเธญเน€เธชเธเธญเนเธเธฐ..." value={newRecText} onChange={e => setNewRecText(e.target.value)} className="console-input" style={{ resize: 'vertical' }} />
                          <button type="button" onClick={handleAddRecommendation} className="btn-save" style={{ width: '100%', padding: '6px' }}>เน€เธเธดเนเธกเธเนเธญเน€เธชเธเธญเนเธเธฐ</button>
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {/* TAB 4: ASSETS INVENTORY EDITOR */}
                {consoleTab === 'assets' && (
                  <div className="lark-registry">
                    <div className="console-form asset-registry-form" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))', gap: '10px' }}>
                      <div className="console-field">
                        <span className="console-label">เธเธนเนเน€เธเธดเธเนเธเนเธเธฒเธ</span>
                        <input type="text" value={newAssetUser} onChange={e => setNewAssetUser(e.target.value)} placeholder="เน€เธเนเธ เธญเธกเธฃ เนเธเนเธงเธชเธ”" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธ•เธณเนเธซเธเนเธ/เนเธเธเธ</span>
                        <input type="text" value={newAssetPosition} onChange={e => setNewAssetPosition(e.target.value)} placeholder="เน€เธเนเธ Marketing" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธฃเธฒเธขเธเธฒเธฃเธญเธธเธเธเธฃเธ“เนเธซเธฅเธฑเธ (Tag)*</span>
                        <AssetTagEditor value={newAssetItemType} onChange={setNewAssetItemType} single placeholder="เน€เธเนเธ Notebook Lenovo" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธญเธธเธเธเธฃเธ“เนเน€เธเธดเนเธกเน€เธ•เธดเธก (Tags)</span>
                        <AssetTagEditor value={newAssetAdditionalEquipment} onChange={setNewAssetAdditionalEquipment} placeholder="เน€เธเนเธ เธชเธฒเธข HDMI" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธซเธกเธฒเธขเน€เธฅเธเธเธตเน€เธฃเธตเธขเธฅ</span>
                        <input type="text" value={newAssetSerial} onChange={e => setNewAssetSerial(e.target.value)} placeholder="เน€เธเนเธ MC-010" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธชเธ–เธฒเธเธฐ</span>
                        <select value={newAssetStatus} onChange={e => setNewAssetStatus(e.target.value)} className="console-input">
                          <option value="เนเธเนเธเธฒเธ">เนเธเนเธเธฒเธ</option>
                          <option value="เธงเนเธฒเธ">เธงเนเธฒเธ</option>
                          <option value="เธฃเธญเธเนเธญเธก">เธฃเธญเธเนเธญเธก</option>
                          <option value="เธชเธนเธเธซเธฒเธข">เธชเธนเธเธซเธฒเธข</option>
                        </select>
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธซเธกเธฒเธขเน€เธซเธ•เธธ</span>
                        <input type="text" value={newAssetNotes} onChange={e => setNewAssetNotes(e.target.value)} placeholder="เธฃเธฒเธขเธฅเธฐเน€เธญเธตเธขเธ”" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธงเธฑเธเธ—เธตเน Submit</span>
                        <input type="text" value={newAssetSubmittedOn} onChange={e => setNewAssetSubmittedOn(e.target.value)} placeholder="YYYY-MM-DD" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">Respondents</span>
                        <input type="text" value={newAssetRespondent} onChange={e => setNewAssetRespondent(e.target.value)} placeholder="เธเธนเนเธ•เธญเธเธเธฅเธฑเธ" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธงเธฑเธเธ—เธตเนเน€เธเธดเธเนเธเนเธเธฒเธ</span>
                        <input type="text" value={newAssetDate} onChange={e => setNewAssetDate(e.target.value)} placeholder="เธงเธฑเธเธ—เธตเนเน€เธเธดเธเนเธเนเธเธฒเธ" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธเธญเธเธ•เนเนเธงเธฃเน / App</span>
                        <AssetTagEditor value={newAssetSoftwareApp} onChange={setNewAssetSoftwareApp} placeholder="เนเธเธฃเนเธเธฃเธก" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เน€เธกเธฅเธ—เธตเนเธฅเธเธ—เธฐเน€เธเธตเธขเธ</span>
                        <input type="text" value={newAssetRegisteredEmail} onChange={e => setNewAssetRegisteredEmail(e.target.value)} placeholder="เธญเธตเน€เธกเธฅ" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธซเธกเธฒเธขเน€เธฅเธเธญเธธเธเธเธฃเธ“เนเน€เธเธดเนเธกเน€เธ•เธดเธก</span>
                        <input type="text" value={newAssetAdditionalSerial} onChange={e => setNewAssetAdditionalSerial(e.target.value)} placeholder="เน€เธเนเธ Pencil-001, Cable-018" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธเธณเธซเธเธ”เธเธทเธ</span>
                        <input type="text" value={newAssetReturnDueDate} onChange={e => setNewAssetReturnDueDate(e.target.value)} placeholder="เธงเธฑเธเธ—เธตเนเธเธณเธซเธเธ”เธเธทเธ" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธงเธฑเธเธ—เธตเนเธ•เธฃเธงเธเธชเธญเธ</span>
                        <input type="text" value={newAssetAuditDate} onChange={e => setNewAssetAuditDate(e.target.value)} placeholder="เธงเธฑเธเธ—เธตเนเธ•เธฃเธงเธเธชเธญเธ" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธงเธฑเธเธ—เธตเนเธเธทเนเธญ</span>
                        <input type="text" value={newAssetPurchaseDate} onChange={e => setNewAssetPurchaseDate(e.target.value)} placeholder="เธงเธฑเธเธ—เธตเนเธเธทเนเธญ" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธงเธฑเธเธซเธกเธ”เธเธฃเธฐเธเธฑเธ</span>
                        <input type="text" value={newAssetWarrantyExpiry} onChange={e => setNewAssetWarrantyExpiry(e.target.value)} placeholder="เธงเธฑเธเธซเธกเธ”เธเธฃเธฐเธเธฑเธ" className="console-input" />
                      </div>
                      <div className="console-field">
                        <span className="console-label">เธเนเธฒเนเธเนเธเนเธฒเธข</span>
                        <input type="text" value={newAssetCost} onChange={e => setNewAssetCost(e.target.value)} placeholder="เธเนเธฒเนเธเนเธเนเธฒเธข" className="console-input" />
                      </div>
                      
                      {/* Export Actions */}
                      <div className="console-field" style={{ display: 'flex', gap: '8px', alignItems: 'flex-end', minWidth: '220px' }}>
                        <button type="button" onClick={() => exportAssetsToExcel(consoleAssets, consoleMonth)} className="sidebar-btn" style={{ flex: 1, backgroundColor: '#059669', color: 'white', padding: '6px', height: '42px', margin: 0, border: 'none' }}>
                          <Download size={16} style={{ marginRight: '6px' }} /> Excel
                        </button>
                        <button type="button" onClick={exportAssetsToPDF} className="sidebar-btn" style={{ flex: 1, backgroundColor: '#dc2626', color: 'white', padding: '6px', height: '42px', margin: 0, border: 'none' }}>
                          <Printer size={16} style={{ marginRight: '6px' }} /> PDF
                        </button>
                      </div>
                      <div className={`asset-inline-save-status ${consoleSaveMessage ? (consoleSaveMessage.startsWith('เธเธฑเธเธ—เธถเธเธชเธณเน€เธฃเนเธ') ? 'success' : 'error') : ''}`} style={{ gridColumn: '1 / -1', marginTop: '10px' }}>
                        {consoleSaveMessage || `เธเธฃเนเธญเธกเธเธฑเธเธ—เธถเธเธเนเธญเธกเธนเธฅ ${data[consoleMonth]?.monthName || consoleMonth} เธเธถเนเธเนเธ”เธเธเธญเธฃเนเธ”`}
                      </div>
                      <button type="button" className="console-save-dashboard-btn asset-inline-save-btn" onClick={saveConsoleChanges} disabled={consoleSaving} style={{ gridColumn: '1 / -1' }}>
                        {consoleSaving ? 'เธเธณเธฅเธฑเธเธเธฑเธเธ—เธถเธเนเธฅเธฐเธเธดเธเธเน...' : '๐’พ เธเธฑเธเธ—เธถเธเนเธฅเธฐเธญเธฑเธเน€เธ”เธ•เนเธ”เธเธเธญเธฃเนเธ”'}
                      </button>
                      {editingAssetSn !== null ? (
                        <div style={{ gridColumn: '1 / -1', display: 'flex', gap: '8px', marginTop: '8px' }}>
                          <button type="button" onClick={handleAddAsset} className="btn-save" style={{ flex: '1', height: '36px' }}>เธเธฑเธเธ—เธถเธเธเธฒเธฃเนเธเนเนเธ</button>
                          <button type="button" onClick={handleCancelEditAsset} className="sidebar-btn" style={{ width: '120px', height: '36px', margin: 0, padding: '0 10px', backgroundColor: '#4b5563', color: 'white' }}>เธขเธเน€เธฅเธดเธ</button>
                        </div>
                      ) : null}
                    </div>

                    <div className="lark-registry-toolbar">
                      <input
                        type="search"
                        value={consoleAssetSearch}
                        onChange={(event) => setConsoleAssetSearch(event.target.value)}
                        placeholder="เธเนเธเธซเธฒเน€เธฅเธเธฃเธฒเธขเธเธฒเธฃ เธเธทเนเธญ เนเธเธเธ เธญเธธเธเธเธฃเธ“เน เธซเธกเธฒเธขเน€เธฅเธเน€เธเธฃเธทเนเธญเธ เธซเธฃเธทเธญเธชเธ–เธฒเธเธฐ..."
                        aria-label="เธเนเธเธซเธฒเธ—เธฐเน€เธเธตเธขเธเธ—เธฃเธฑเธเธขเนเธชเธดเธ"
                      />
                      <span>เนเธชเธ”เธ {consoleAssets.length} เธเธฒเธ {assetsList.length} เธฃเธฒเธขเธเธฒเธฃ ยท เธเธฑเธ”เธเธฅเธธเนเธกเธ•เธฒเธกเนเธเธเธ</span>
                    </div>

                    <div className="console-table-scroll lark-registry-scroll">
                      <table className="details-table lark-registry-table">
                        <thead>
                          <tr>
                            <th>Number</th>
                            <th>Submitted on</th>
                            <th>Respondents</th>
                            <th>เธงเธฑเธเธ—เธตเนเน€เธเธดเธเนเธเนเธเธฒเธ</th>
                            <th>เธเธธเธเธเธฅเน€เธเธดเธเนเธเนเธญเธธเธเธเธฃเธ“เน</th>
                            <th>เธ•เธณเนเธซเธเนเธ</th>
                            <th>เธฃเธฒเธขเธเธฒเธฃเธญเธธเธเธเธฃเธ“เนเธซเธฅเธฑเธ</th>
                            <th>เธญเธธเธเธเธฃเธ“เนเน€เธเธดเนเธกเน€เธ•เธดเธกเธ—เธตเนเธ•เนเธญเธเธเธฒเธฃเน€เธเธดเธ</th>
                            <th>เธเธญเธเธ•เนเนเธงเธฃเน / App</th>
                            <th>เน€เธกเธฅเธ—เธตเนเธฅเธเธ—เธฐเน€เธเธตเธขเธ</th>
                            <th>เธซเธกเธฒเธขเน€เธฅเธเธญเธธเธเธเธฃเธ“เน</th>
                            <th>เธซเธกเธฒเธขเน€เธฅเธเธญเธธเธเธเธฃเธ“เนเน€เธเธดเนเธกเน€เธ•เธดเธก</th>
                            <th>เธเธณเธซเธเธ”เธเธทเธเธญเธธเธเธเธฃเธ“เน</th>
                            <th>เธชเธ–เธฒเธเธฐ</th>
                            <th>เธซเธกเธฒเธขเน€เธซเธ•เธธ</th>
                            <th>เธงเธฑเธเธ—เธตเนเธ•เธฃเธงเธเธชเธญเธ</th>
                            <th>เธงเธฑเธเธ—เธตเนเธเธทเนเธญ</th>
                            <th>เธงเธฑเธเธซเธกเธ”เธเธฃเธฐเธเธฑเธ</th>
                            <th>เธเนเธฒเนเธเนเธเนเธฒเธข</th>
                            <th>เธเธฑเธ”เธเธฒเธฃ</th>
                          </tr>
                        </thead>
                        <tbody>
                          {consoleAssets.map((asset, idx) => {
                            const previous = consoleAssets[idx - 1];
                            const showGroup = !previous || String(previous.position || '-') !== String(asset.position || '-');
                            return (
                            <Fragment key={asset.sn ?? idx}>
                              {showGroup && (
                                <tr className="lark-group-row">
                                  <td colSpan="20">
                                    <span>{asset.position || 'เนเธกเนเธฃเธฐเธเธธเนเธเธเธ'}</span>
                                    <small>{consoleAssets.filter((row) => String(row.position || '-') === String(asset.position || '-')).length} เธฃเธฒเธขเธเธฒเธฃ</small>
                                  </td>
                                </tr>
                              )}
                            <tr>
                              <td>{idx + 1}</td>
                              <td>{asset.submittedOn || '-'}</td>
                              <td>{asset.respondent || '-'}</td>
                              <td>{asset.date || '-'}</td>
                              <td>{asset.user}</td>
                              <td>{asset.position}</td>
                              <td>
                                <div className="asset-tag-cell">
                                  <span className="lark-pill lark-pill-device">{asset.itemType || '-'}</span>
                                  <button type="button" onClick={() => {
                                    handleLoadEditAsset(asset);
                                    setEditingAssetTagField({ sn: asset.sn, field: 'itemType' });
                                  }}>เนเธเนเนเธ Tag</button>
                                  {editingAssetTagField?.sn === asset.sn && editingAssetTagField?.field === 'itemType' && (
                                    <AssetTagPicker
                                      value={newAssetItemType}
                                      onChange={setNewAssetItemType}
                                      options={primaryTagOptions}
                                      single
                                      onClose={() => setEditingAssetTagField(null)}
                                    />
                                  )}
                                </div>
                              </td>
                              <td>
                                <div className="asset-tag-cell">
                                  <AssetTags value={asset.additionalEquipment} />
                                  <button type="button" onClick={() => {
                                    handleLoadEditAsset(asset);
                                    setEditingAssetTagField({ sn: asset.sn, field: 'additionalEquipment' });
                                  }}>เนเธเนเนเธ Tag</button>
                                  {editingAssetTagField?.sn === asset.sn && editingAssetTagField?.field === 'additionalEquipment' && (
                                    <AssetTagPicker
                                      value={newAssetAdditionalEquipment}
                                      onChange={setNewAssetAdditionalEquipment}
                                      options={additionalTagOptions}
                                      onClose={() => setEditingAssetTagField(null)}
                                    />
                                  )}
                                </div>
                              </td>
                              <td><AssetTags value={asset.softwareApp} /></td>
                              <td>{asset.registeredEmail || '-'}</td>
                              <td>
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                  <strong>{asset.deviceSerial}</strong>
                                  {(() => {
                                    const match = externalDevices.find(d => 
                                      (d.deviceNumber || '').toLowerCase().replace(/\s/g, '') === (asset.deviceSerial || '').toLowerCase().replace(/\s/g, '')
                                      || (d.deviceNumber || '').toLowerCase().replace(/\s/g, '') === (asset.additionalSerial || '').toLowerCase().replace(/\s/g, '')
                                    );
                                    if (!match) return null;
                                    
                                    let badgeColor = 'var(--text-muted)';
                                    let statusText = 'Unverified';
                                    if (match.status === 'active') {
                                      if (match.daysRemaining > 7) { badgeColor = 'var(--success)'; statusText = 'Verified'; }
                                      else { badgeColor = 'var(--warning)'; statusText = 'Expiring Soon'; }
                                    }
                                    return (
                                      <span style={{ fontSize: '0.7rem', padding: '2px 6px', borderRadius: '4px', backgroundColor: badgeColor, color: '#fff', alignSelf: 'flex-start' }} title={`เธ•เธฃเธงเธเธชเธญเธเธฅเนเธฒเธชเธธเธ”: ${match.lastVerifiedAt ? new Date(match.lastVerifiedAt).toLocaleString('th-TH') : '-'}`}>
                                        {statusText} ({match.daysRemaining} เธงเธฑเธ)
                                      </span>
                                    );
                                  })()}
                                </div>
                              </td>
                              <td>{asset.additionalSerial || '-'}</td>
                              <td>{asset.returnDueDate || '-'}</td>
                              <td>
                                <span className={`lark-status lark-status-${asset.status === 'เนเธเนเธเธฒเธ' ? 'active' : asset.status === 'เธฃเธญเธเนเธญเธก' ? 'repair' : asset.status === 'เธงเนเธฒเธ' ? 'vacant' : 'other'}`}>
                                  {asset.status || '-'}
                                </span>
                              </td>
                              <td>{asset.notes || '-'}</td>
                              <td>{asset.inspectionDate || '-'}</td>
                              <td>{asset.purchaseDate || '-'}</td>
                              <td>{asset.warrantyEndDate || '-'}</td>
                              <td className="lark-number">{Number(asset.expense || 0).toLocaleString('th-TH')}</td>
                              <td>
                                <div style={{ display: 'flex', gap: '6px' }}>
                                  <button onClick={() => handleLoadEditAsset(asset)} className="btn-details" style={{ padding: '2px 8px', fontSize: '0.75rem', backgroundColor: 'var(--primary-light)', color: 'var(--primary)', border: 'none' }}>เนเธเนเนเธ</button>
                                  <button onClick={() => handleDeleteAsset(asset.sn)} className="console-delete-btn">เธฅเธ</button>
                                </div>
                              </td>
                            </tr>
                            </Fragment>
                          )})}
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
                        <h4 style={{ margin: 0, fontSize: '1.15rem' }}>๐จ เธเธฃเธฐเธงเธฑเธ•เธดเธฃเธฑเธเน€เธเธชเนเธเนเธเธเนเธญเธก Support - {editingTicketSn !== null ? <span style={{ color: 'var(--warning)' }}>เนเธซเธกเธ”เนเธเนเนเธเธฃเธซเธฑเธช #{editingTicketSn}</span> : <span>เนเธซเธกเธ”เน€เธเธดเนเธกเธเนเธญเธกเธนเธฅ</span>}</h4>
                        <div>
                          <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginRight: '8px' }}>เน€เธฅเธทเธญเธเน€เธ”เธทเธญเธเธ—เธตเนเธเธฐเธเธฑเธ”เธเธฒเธฃ:</span>
                          <select value={consoleMonth} onChange={e => setConsoleMonth(e.target.value)} style={{ padding: '6px 12px', borderRadius: '6px', border: '1px solid rgba(255, 255, 255, 0.15)', backgroundColor: '#1f2937', color: 'white' }}>
                            {Object.keys(data).map(key => (
                              <option key={key} value={key}>{data[key].monthName}</option>
                            ))}
                          </select>
                        </div>
                      </div>

                      <div className="console-form" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(130px, 1fr))', gap: '10px' }}>
                        <div className="console-field">
                          <span className="console-label">เธเธทเนเธญเธเธนเนเนเธเนเธ</span>
                          <input type="text" value={newTicketComplainant} onChange={e => setNewTicketComplainant(e.target.value)} placeholder="เธญเธกเธฃ เนเธเนเธงเธชเธ”" className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">เธญเธตเน€เธกเธฅ</span>
                          <input type="text" value={newTicketEmail} onChange={e => setNewTicketEmail(e.target.value)} placeholder="user@domain.com" className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">AnyDesk ID</span>
                          <input type="text" value={newTicketAnydesk} onChange={e => setNewTicketAnydesk(e.target.value)} placeholder="เน€เธเนเธ 1 234 567" className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">เธญเธฒเธเธฒเธฃเธ—เธตเนเนเธเนเธ*</span>
                          <input type="text" value={newTicketIssue} onChange={e => setNewTicketIssue(e.target.value)} placeholder="เธเธญเธ”เธฑเธ, เธเธดเธกเธเนเนเธกเนเนเธ”เน" className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">เธชเธฒเน€เธซเธ•เธธเธเธฒเธฃเน€เธชเธตเธข</span>
                          <input type="text" value={newTicketCause} onChange={e => setNewTicketCause(e.target.value)} placeholder="เน€เธชเธทเนเธญเธกเธ•เธฒเธกเธชเธ เธฒเธ" className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">เธเธนเนเธ”เธณเน€เธเธดเธเธเธฒเธ (IT)</span>
                          <input type="text" value={newTicketResponder} onChange={e => setNewTicketResponder(e.target.value)} placeholder="เธเธทเนเธญเน€เธเนเธฒเธซเธเนเธฒเธ—เธตเน" className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">เน€เธงเธฅเธฒเนเธเน (HH:MM)</span>
                          <input type="text" value={newTicketDuration} onChange={e => setNewTicketDuration(e.target.value)} className="console-input" />
                        </div>
                        <div className="console-field">
                          <span className="console-label">เธชเธ–เธฒเธเธฐ</span>
                          <select value={newTicketStatus} onChange={e => setNewTicketStatus(e.target.value)} className="console-input">
                            <option value="เน€เธชเธฃเนเธเธชเธดเนเธ">เน€เธชเธฃเนเธเธชเธดเนเธ</option>
                            <option value="เธเธณเธฅเธฑเธเธ”เธณเน€เธเธดเธเธเธฒเธฃ">เธเธณเธฅเธฑเธเธ”เธณเน€เธเธดเธเธเธฒเธฃ</option>
                            <option value="เธเนเธฒเธขเน€เธเธดเธเนเธฅเนเธง">เธเนเธฒเธขเน€เธเธดเธเนเธฅเนเธง (เธเธทเนเธญเธญเธธเธเธเธฃเธ“เน)</option>
                          </select>
                        </div>
                        <div className="console-field">
                          <span className="console-label">เธเนเธฒเนเธเนเธเนเธฒเธข (เธเธฒเธ—)</span>
                          <input type="number" value={newTicketCost} onChange={e => setNewTicketCost(e.target.value)} className="console-input" />
                        </div>
                        {editingTicketSn !== null ? (
                          <div style={{ gridColumn: '1 / -1', display: 'flex', gap: '8px', marginTop: '8px' }}>
                            <button type="button" onClick={handleAddTicket} className="btn-save" style={{ flex: '1', height: '36px' }}>เธเธฑเธเธ—เธถเธเธเธฒเธฃเนเธเนเนเธ</button>
                            <button type="button" onClick={handleCancelEditTicket} className="sidebar-btn" style={{ width: '120px', height: '36px', margin: 0, padding: '0 10px', backgroundColor: '#4b5563', color: 'white' }}>เธขเธเน€เธฅเธดเธ</button>
                          </div>
                        ) : (
                          <button type="button" onClick={handleAddTicket} className="btn-save" style={{ gridColumn: '1 / -1', marginTop: '8px', height: '36px' }}>เธเธฑเธเธ—เธถเธเน€เธเธชเนเธเนเธเธเนเธญเธก</button>
                        )}
                      </div>

                      <div className="console-table-scroll">
                        <table className="details-table">
                          <thead>
                            <tr>
                              <th>SN</th>
                              <th>เธเธนเนเนเธเนเธ</th>
                              <th>เน€เธเธช/เธเธฑเธเธซเธฒ</th>
                              <th>เธชเธฒเน€เธซเธ•เธธ</th>
                              <th>เน€เธงเธฅเธฒ</th>
                              <th>เธเนเธฒเนเธเนเธเนเธฒเธข</th>
                              <th>เธชเธ–เธฒเธเธฐ</th>
                              <th>เธเธฑเธ”เธเธฒเธฃ</th>
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
                                  <span style={{ color: t.status === 'เน€เธชเธฃเนเธเธชเธดเนเธ' || t.status === 'เธเนเธฒเธขเน€เธเธดเธเนเธฅเนเธง' ? 'var(--success)' : 'var(--warning)', fontWeight: 'bold' }}>
                                    {t.status}
                                  </span>
                                </td>
                                <td>
                                  <div style={{ display: 'flex', gap: '6px' }}>
                                    <button onClick={() => handleLoadEditTicket(t)} className="btn-details" style={{ padding: '2px 8px', fontSize: '0.75rem', backgroundColor: 'var(--primary-light)', color: 'var(--primary)', border: 'none' }}>เนเธเนเนเธ</button>
                                    <button onClick={() => handleDeleteTicket(t.sn)} className="console-delete-btn">เธฅเธ</button>
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
                    <h4 className="console-title">๐’พ เธชเธณเธฃเธญเธเธเนเธญเธกเธนเธฅเนเธฅเธฐเธฃเธตเน€เธเนเธ•เธเธฒเธฃเธ•เธฑเนเธเธเนเธฒเธฃเธฐเธเธ</h4>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                      <div className="console-card">
                        <h5 style={{ margin: '0 0 8px 0', color: 'var(--success)' }}>๐“ฅ เธชเนเธเธญเธญเธเนเธเธฅเนเธเนเธญเธกเธนเธฅเธชเธณเธฃเธญเธ (Backup to JSON)</h5>
                        <p style={{ fontSize: '0.8rem', color: 'var(--text-muted)', margin: '0 0 12px 0' }}>
                          เธ”เธฒเธงเธเนเนเธซเธฅเธ”เธเนเธญเธกเธนเธฅเนเธ”เธเธเธญเธฃเนเธ” เธ—เธฐเน€เธเธตเธขเธเธญเธธเธเธเธฃเธ“เน IT เนเธฅเธฐเธเธฃเธฐเธงเธฑเธ•เธดเนเธเนเธเธเนเธญเธกเธ—เธฑเนเธเธซเธกเธ”เน€เธเนเธเนเธงเนเนเธเธฃเธนเธเนเธเธเนเธเธฅเน .json
                        </p>
                        <button type="button" onClick={handleExportJson} className="btn-save" style={{ width: 'auto', padding: '10px 20px', backgroundColor: 'var(--success)', borderColor: 'var(--success)' }}>
                          เธ”เธฒเธงเธเนเนเธซเธฅเธ”เนเธเธฅเนเธชเธณเธฃเธญเธเธเนเธญเธกเธนเธฅ (.json)
                        </button>
                      </div>

                      <div className="console-card">
                        <h5 style={{ margin: '0 0 8px 0', color: 'var(--primary)' }}>๐“ค เธเธณเน€เธเนเธฒเนเธเธฅเนเธเนเธญเธกเธนเธฅเธชเธณเธฃเธญเธ (Import JSON Backup)</h5>
                        <p style={{ fontSize: '0.8rem', color: 'var(--text-muted)', margin: '0 0 12px 0' }}>
                          เน€เธฅเธทเธญเธเนเธเธฅเนเธชเธณเธฃเธญเธเธเนเธญเธกเธนเธฅเธเธฒเธกเธชเธเธธเธฅ .json เน€เธเธทเนเธญเธเธนเนเธเธทเธเธชเธ–เธฒเธเธฐเธเนเธญเธกเธนเธฅเน€เธ”เธดเธกเธ—เธฑเนเธเธซเธกเธ”
                        </p>
                        <input type="file" accept=".json" onChange={handleImportJson} style={{ display: 'block', fontSize: '0.8rem', color: 'var(--text-muted)' }} />
                      </div>

                      <div className="console-card" style={{ border: '1px solid rgba(239, 68, 68, 0.2)', backgroundColor: 'rgba(239, 68, 68, 0.03)' }}>
                        <h5 style={{ margin: '0 0 8px 0', color: 'rgb(239, 68, 68)' }}>โ ๏ธ เธฃเธตเน€เธเนเธ•เธฃเธฐเธเธเนเธซเธกเนเธ—เธฑเนเธเธซเธกเธ” (Wipe Database)</h5>
                        <p style={{ fontSize: '0.8rem', color: 'var(--text-muted)', margin: '0 0 12px 0' }}>
                          เธฅเธเธเนเธญเธกเธนเธฅเธ—เธฑเนเธเธซเธกเธ”เนเธเธเธฒเธเธเนเธญเธกเธนเธฅ เธฃเธฐเธเธเธเธฐเธ–เธนเธเธฅเนเธฒเธเธเนเธฒเนเธซเนเธงเนเธฒเธเน€เธเธฅเนเธฒเน€เธซเธกเธทเธญเธเน€เธฃเธดเนเธกเธ•เนเธเนเธซเธกเน (เธเธณเน€เธ•เธทเธญเธ: เธเนเธญเธกเธนเธฅเธ—เธตเนเธ–เธนเธเธฅเธเธเธฐเนเธกเนเธชเธฒเธกเธฒเธฃเธ–เธเธนเนเธเธทเธเนเธ”เน)
                        </p>
                        <button type="button" onClick={handleResetToDefault} className="sidebar-btn" style={{ width: 'auto', padding: '10px 20px', backgroundColor: 'rgb(239, 68, 68)', border: 'none', color: 'white' }}>
                          เธฅเธเธเนเธญเธกเธนเธฅเธ—เธฑเนเธเธซเธกเธ” (Wipe Database)
                        </button>
                      </div>
                    </div>
                  </div>
                )}

              </div>
            </div>
            {consoleTab !== 'assets' && (
              <div className="console-save-bar">
                <span className={consoleSaveMessage ? (consoleSaveMessage.startsWith('เธเธฑเธเธ—เธถเธเธชเธณเน€เธฃเนเธ') ? 'success' : 'error') : ''}>
                  {consoleSaveMessage || `เธเธฃเนเธญเธกเธเธฑเธเธ—เธถเธเธเนเธญเธกเธนเธฅ ${data[consoleMonth]?.monthName || consoleMonth} เธเธถเนเธเนเธ”เธเธเธญเธฃเนเธ”`}
                </span>
                <button type="button" className="console-save-dashboard-btn" onClick={saveConsoleChanges} disabled={consoleSaving}>
                  {consoleSaving ? 'เธเธณเธฅเธฑเธเธเธฑเธเธ—เธถเธเนเธฅเธฐเธเธดเธเธเน...' : '๐’พ เธเธฑเธเธ—เธถเธเนเธฅเธฐเธญเธฑเธเน€เธ”เธ•เนเธ”เธเธเธญเธฃเนเธ”'}
                </button>
              </div>
            )}
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
                    <h3>๐”ง เน€เธกเธเธนเธเธดเธ”เธเธฒเธเธชเธณเธซเธฃเธฑเธเธเนเธฒเธเนเธญเธ—เธต (IT Close Work)</h3>
                    <p>เน€เธฅเธทเธญเธเนเธเธเธฒเธเธ—เธตเนเธเนเธฒเธเธเธฒเน€เธเธทเนเธญเธญเธฑเธเน€เธ”เธ•เธฃเธฒเธขเธฅเธฐเน€เธญเธตเธขเธ”เธเธฒเธฃเนเธเนเนเธเนเธฅเธฐเธเธดเธ”เธเธฒเธ เธเธฃเธฐเธเธณเน€เธ”เธทเธญเธ {data[currentMonth]?.monthName}</p>
                  </div>
                ) : larkFormType === 'asset' ? (
                  <div>
                    <h3>๐’ป เธฅเธเธ—เธฐเน€เธเธตเธขเธเน€เธเธฃเธทเนเธญเธเน€เธเนเธฒเธเธฅเธฑเธ (IT Asset Registration)</h3>
                    <p>เธเธฑเธเธ—เธถเธเธเธฃเธฐเธงเธฑเธ•เธดเธเธฒเธฃเน€เธเธดเธเนเธเนเธญเธธเธเธเธฃเธ“เนเนเธญเธ—เธตเน€เธเธฃเธทเนเธญเธเนเธซเธกเนเน€เธเนเธฒเธชเธนเนเธเธฅเธฑเธเธ—เธฐเน€เธเธตเธขเธเธเธฅเธฒเธ</p>
                  </div>
                ) : (
                  <div>
                    <h3>๐จ เนเธเนเธเธเนเธญเธกเธเธณเธฃเธธเธ / เธเธฑเธเธซเธฒเนเธญเธ—เธต (Report Repair / IT Issue)</h3>
                    <p>เนเธเนเธเธเธฑเธเธซเธฒเธเธฑเธ”เธเนเธญเธเธเธญเธเธญเธธเธเธเธฃเธ“เนเธซเธฃเธทเธญเธฃเธฐเธเธเนเธญเธ—เธตเน€เธเธทเนเธญเธเธฃเธฐเธชเธฒเธเธเนเธฒเธเน€เธเนเธฒเนเธเนเนเธ เธเธฃเธฐเธเธณเน€เธ”เธทเธญเธ {data[currentMonth]?.monthName}</p>
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
                <h4 style={{ margin: '0 0 10px 0', fontSize: '1.2rem', color: '#065f46', fontWeight: '700' }}>เธชเนเธเธเนเธญเธกเธนเธฅเธชเธณเน€เธฃเนเธเน€เธฃเธตเธขเธเธฃเนเธญเธขเนเธฅเนเธง!</h4>
                <p style={{ margin: '0 0 24px 0', fontSize: '0.875rem', color: '#4b5563' }}>
                  เธเนเธญเธกเธนเธฅเธเธญเธเธเธธเธ“เนเธ”เนเธฃเธฑเธเธเธฒเธฃเธเธฑเธเธ—เธถเธเนเธฅเธฐเธฃเธฐเธเธเนเธ”เนเธญเธฑเธเน€เธ”เธ•เธ•เธฑเธงเน€เธฅเธเธงเธดเน€เธเธฃเธฒเธฐเธซเนเนเธ”เธเธเธญเธฃเนเธ”เนเธซเนเนเธ”เธขเธญเธฑเธ•เนเธเธกเธฑเธ•เธดเนเธฅเนเธง
                </p>
                <div style={{ display: 'flex', gap: '12px', justifyContent: 'center' }}>
                  <button onClick={() => setLarkSubmitted(false)} className="lark-submit-btn" style={{ width: 'auto', padding: '10px 24px' }}>เธเธฃเธญเธเธเธญเธฃเนเธกเนเธซเธกเน</button>
                  <button onClick={() => setActiveModal(null)} className="sidebar-btn" style={{ width: 'auto', padding: '10px 24px', margin: 0, backgroundColor: '#e5e7eb', color: '#374151', border: 'none' }}>เธเธดเธ”เธซเธเนเธฒเธ•เนเธฒเธ</button>
                </div>
              </div>
            ) : (
              <form onSubmit={handleLarkSubmit}>

                {larkFormType === 'ticket' ? (
                  <div className="lark-card">

                    {larkTicketRole === 'user' ? (
                      <div>
                        <div style={{ paddingBottom: '12px', marginBottom: '16px' }}>
                          <h4 style={{ margin: 0, color: '#1e40af', fontSize: '0.95rem', fontWeight: '700' }}>เธชเนเธเนเธเนเธเน€เธฃเธทเนเธญเธเธเนเธญเธกเนเธเธก / เธเธฑเธเธซเธฒเธ—เธตเนเธเธเธเธเนเธ”เธเธเธญเธฃเนเธ” ({data[currentMonth]?.monthName})</h4>
                        </div>
                        
                        <div className="lark-field-group">
                          <label>เธเธทเนเธญเธเธนเนเนเธเนเธ / เธเธนเนเธเธเธเธฑเธเธซเธฒ <span>*</span></label>
                          <input type="text" className="lark-input" placeholder="เธ•เธฑเธงเธญเธขเนเธฒเธ: เธชเธกเน€เธเธตเธขเธฃเธ•เธด เธขเธดเนเธเธ”เธต" value={larkTicketComplainant} onChange={e => setLarkTicketComplainant(e.target.value)} required={larkTicketRole === 'user'} />
                        </div>

                        <div className="lark-field-group" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                          <div>
                            <label>เธญเธตเน€เธกเธฅเธเธนเนเนเธเนเธ</label>
                            <input type="email" className="lark-input" placeholder="user@domain.com" value={larkTicketEmail} onChange={e => setLarkTicketEmail(e.target.value)} />
                          </div>
                          <div>
                            <label>AnyDesk ID</label>
                            <input type="text" className="lark-input" placeholder="เน€เธเนเธ 1 234 567" value={larkTicketAnydesk} onChange={e => setLarkTicketAnydesk(e.target.value)} />
                          </div>
                        </div>

                        <div className="lark-field-group">
                          <label>เธซเธกเธฒเธขเน€เธฅเธเน€เธเธฃเธทเนเธญเธเธเธฒเธเธ—เธฐเน€เธเธตเธขเธ (เธ–เนเธฒเธกเธต)</label>
                          <input
                            type="text"
                            className="lark-input"
                            placeholder="เน€เธเนเธ ASUS-019, MC-002"
                            value={larkTicketAssetSerial}
                            onChange={e => setLarkTicketAssetSerial(e.target.value)}
                          />
                          <small style={{ color: '#64748b' }}>เน€เธกเธทเนเธญเธฃเธฐเธเธธเธซเธกเธฒเธขเน€เธฅเธเน€เธเธฃเธทเนเธญเธ เธฃเธฐเธเธเธเธฐเน€เธเธฅเธตเนเธขเธเธชเธ–เธฒเธเธฐเน€เธเนเธ โ€เธฃเธญเธเนเธญเธกโ€ เนเธฅเธฐเธเธดเธ”เธฃเธฒเธขเธเธฒเธฃเน€เธเธดเธเธเธญเธเน€เธเธฃเธทเนเธญเธเธเธฑเนเธเธญเธฑเธ•เนเธเธกเธฑเธ•เธด</small>
                        </div>

                        <div className="lark-field-group">
                          <label>เธญเธฒเธเธฒเธฃเธ—เธตเนเนเธเนเธเธเนเธญเธก / เธเธฑเธเธซเธฒเธ—เธตเนเธเธ <span>*</span></label>
                          <input type="text" className="lark-input" placeholder="เธ•เธฑเธงเธญเธขเนเธฒเธ: เธซเธเนเธฒเธเธญเนเธกเนเธ•เธดเธ”, เธเธฃเธดเนเธเธ—เนเธเธฒเธเนเธกเนเธญเธญเธ" value={larkTicketIssue} onChange={e => setLarkTicketIssue(e.target.value)} required={larkTicketRole === 'user'} />
                        </div>
                      </div>
                    ) : (() => {
                      const pendingTickets = tickets.filter(t => t.status === 'เธเธณเธฅเธฑเธเธ”เธณเน€เธเธดเธเธเธฒเธฃ');
                      const selectedTicket = pendingTickets.find(t => Number(t.sn) === Number(selectedPendingTicketSn));

                      return (
                        <div>
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #e5e7eb', paddingBottom: '12px', marginBottom: '16px' }}>
                            <h4 style={{ margin: 0, color: '#1e40af', fontSize: '0.95rem', fontWeight: '700' }}>
                              {selectedTicket ? `๐”ง เธเธดเธ”เนเธเธเธฒเธเธเนเธญเธกเนเธเธก [SN: ${selectedTicket.sn}]` : '๐“ เธฃเธฒเธขเธเธฒเธฃเธเธฒเธเธเนเธญเธกเธ—เธตเนเธขเธฑเธเนเธกเนเนเธ”เนเธเธดเธ” (เธเธณเธฅเธฑเธเธ”เธณเน€เธเธดเธเธเธฒเธฃ)'}
                            </h4>
                            {selectedTicket && (
                              <button type="button" onClick={() => setSelectedPendingTicketSn('')} style={{ padding: '4px 10px', fontSize: '0.75rem', backgroundColor: '#e5e7eb', color: '#475569', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>
                                เธขเนเธญเธเธเธฅเธฑเธเนเธเธฃเธฒเธขเธเธฒเธฃเธเธฒเธ
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
                                        <span style={{ backgroundColor: '#fee2e2', color: '#ef4444', padding: '2px 6px', borderRadius: '4px', fontSize: '0.7rem', fontWeight: 'bold' }}>เธฃเธซเธฑเธช #{t.sn}</span>
                                        <strong style={{ fontSize: '0.85rem', color: '#1f2937' }}>{t.complainant}</strong>
                                        <span style={{ fontSize: '0.7rem', color: '#6b7280' }}>({t.date})</span>
                                      </div>
                                      <div style={{ fontSize: '0.8rem', color: '#4b5563' }}>
                                        <strong>เธญเธฒเธเธฒเธฃเน€เธชเธตเธข:</strong> {t.issue}
                                      </div>
                                    </div>
                                    <button type="button" onClick={() => {
                                      setSelectedPendingTicketSn(String(t.sn));
                                      setLarkTicketResponder('');
                                      setLarkTicketDuration('00:30');
                                      setLarkTicketCause('');
                                      setLarkTicketCost('0');
                                      setLarkTicketStatus('เน€เธชเธฃเนเธเธชเธดเนเธ');
                                    }} style={{ padding: '6px 12px', fontSize: '0.75rem', backgroundColor: '#2563eb', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>
                                      เน€เธฅเธทเธญเธเนเธฅเธฐเธเธดเธ”เธเธฒเธ
                                    </button>
                                  </div>
                                ))}
                              </div>
                            ) : (
                              <div style={{ padding: '24px 12px', backgroundColor: '#ecfdf5', border: '1px solid #a7f3d0', color: '#065f46', borderRadius: '6px', fontSize: '0.85rem', fontWeight: 'bold', textAlign: 'center' }}>
                                ๐ เนเธกเนเธกเธตเธเธฒเธเธเนเธญเธกเธ—เธตเนเธเนเธฒเธเธเธฒเธญเธขเธนเนเนเธเธเธ“เธฐเธเธตเน เธ—เธธเธเนเธเธเธฒเธเนเธ”เนเธฃเธฑเธเธเธฒเธฃเนเธเนเนเธเน€เธฃเธตเธขเธเธฃเนเธญเธขเนเธฅเนเธง
                              </div>
                            )
                          ) : (
                            /* View 2: Form to close the selected ticket */
                            <div>
                              <div style={{ padding: '12px', backgroundColor: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: '6px', marginBottom: '16px', fontSize: '0.8rem' }}>
                                <div style={{ fontWeight: 'bold', color: '#1e40af', marginBottom: '4px' }}>เธเนเธญเธกเธนเธฅเธเธนเนเนเธเนเธ:</div>
                                <div>๐‘ค <strong>เธเธนเนเนเธเนเธ:</strong> {selectedTicket.complainant} (เน€เธกเธฅ: {selectedTicket.email} / AnyDesk: {selectedTicket.anydesk})</div>
                                <div>โ ๏ธ <strong>เธเธฑเธเธซเธฒเธ—เธตเนเธเธ:</strong> {selectedTicket.issue}</div>
                              </div>

                              <div className="lark-field-group" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                                <div>
                                  <label>เธเธนเนเธ”เธณเน€เธเธดเธเธเธฒเธ (เธเนเธฒเธ IT) <span>*</span></label>
                                  <input type="text" className="lark-input" placeholder="เน€เธเนเธ เธเนเธฒเธเธเนเธญเธเธ เธ (IT)" value={larkTicketResponder} onChange={e => setLarkTicketResponder(e.target.value)} required={larkTicketRole === 'it'} />
                                </div>
                                <div>
                                  <label>เน€เธงเธฅเธฒเนเธเนเน€เธชเธฃเนเธ (เธเธฑเนเธงเนเธกเธ:เธเธฒเธ—เธต)</label>
                                  <input type="text" className="lark-input" placeholder="เน€เธเนเธ 00:45" value={larkTicketDuration} onChange={e => setLarkTicketDuration(e.target.value)} />
                                </div>
                              </div>

                              <div className="lark-field-group">
                                <label>เธชเธฒเน€เธซเธ•เธธเธเธฒเธฃเน€เธชเธตเธข / เธงเธดเธเธตเนเธเนเนเธ</label>
                                <input type="text" className="lark-input" placeholder="เธ•เธฑเธงเธญเธขเนเธฒเธ: เน€เธเธฅเธตเนเธขเธเธชเธฒเธข LAN เนเธซเธกเน, เธฃเธตเธชเธ•เธฒเธฃเนเธ—เธเธฒเธฃเธ•เธฑเนเธเธเนเธฒเน€เธเธฃเธทเธญเธเนเธฒเธข" value={larkTicketCause} onChange={e => setLarkTicketCause(e.target.value)} />
                              </div>

                              <div className="lark-field-group" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                                <div>
                                  <label>เธญเธฑเธเน€เธ”เธ•เธชเธ–เธฒเธเธฐเนเธเธเธฒเธ</label>
                                  <select className="lark-input" value={larkTicketStatus} onChange={e => setLarkTicketStatus(e.target.value)}>
                                    <option value="เน€เธชเธฃเนเธเธชเธดเนเธ">เน€เธชเธฃเนเธเธชเธดเนเธ (Resolved)</option>
                                    <option value="เธเนเธฒเธขเน€เธเธดเธเนเธฅเนเธง">เธเนเธฒเธขเน€เธเธดเธเนเธฅเนเธง (เธเธทเนเธญเธญเธฐเนเธซเธฅเนเน€เธชเธฃเธดเธก)</option>
                                  </select>
                                </div>
                                <div>
                                  <label>เธเนเธฒเนเธเนเธเนเธฒเธขเธเนเธญเธกเนเธเธก (เธเธฒเธ—)</label>
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
                      <h4 style={{ margin: 0, color: '#1e40af', fontSize: '1rem', fontWeight: '700' }}>เธฅเธเธ—เธฐเน€เธเธตเธขเธเธญเธธเธเธเธฃเธ“เนเน€เธเธฃเธทเนเธญเธเนเธซเธกเนเน€เธเนเธฒเธ—เธฐเน€เธเธตเธขเธเธเธฅเธฒเธ</h4>
                    </div>

                    <div className="lark-field-group">
                      <label>เธเธทเนเธญเธเธนเนเธเธฃเธญเธเธเธฃเธญเธเนเธเนเธเธฒเธ</label>
                      <input type="text" className="lark-input" placeholder="เน€เธเนเธ เธญเธกเธฃ เนเธเนเธงเธชเธ” (เธซเธฃเธทเธญเนเธชเน เธชเนเธงเธเธเธฅเธฒเธ)" value={larkAssetUser} onChange={e => setLarkAssetUser(e.target.value)} />
                    </div>

                    <div className="lark-field-group">
                      <label>เธ•เธณเนเธซเธเนเธ / เนเธเธเธ</label>
                      <input type="text" className="lark-input" placeholder="เน€เธเนเธ Accounting, Marketing, HR" value={larkAssetPosition} onChange={e => setLarkAssetPosition(e.target.value)} />
                    </div>

                    <div className="lark-field-group">
                      <label>เธเธฃเธฐเน€เธ เธ—เธญเธธเธเธเธฃเธ“เนเนเธญเธ—เธต / เธฃเธธเนเธเธซเธฅเธฑเธ <span>*</span></label>
                      <input type="text" className="lark-input" placeholder="เธ•เธฑเธงเธญเธขเนเธฒเธ: Notebook Lenovo, Computer (Pc)" value={larkAssetItemType} onChange={e => setLarkAssetItemType(e.target.value)} required={larkFormType === 'asset'} />
                    </div>

                    <div className="lark-field-group">
                      <label>เธเธตเน€เธฃเธตเธขเธฅเธเธฑเธกเน€เธเธญเธฃเน / เธฃเธซเธฑเธชเน€เธเธฃเธทเนเธญเธ (Serial Number)</label>
                      <input type="text" className="lark-input" placeholder="เน€เธเนเธ MC-054, LNV-987" value={larkAssetSerial} onChange={e => setLarkAssetSerial(e.target.value)} />
                    </div>

                    <div className="lark-field-group">
                      <label>เธชเธ–เธฒเธเธฐเธเธฅเธฑเธเน€เธฃเธดเนเธกเธ•เนเธ</label>
                      <select className="lark-input" value={larkAssetStatus} onChange={e => setLarkAssetStatus(e.target.value)}>
                        <option value="เนเธเนเธเธฒเธ">เนเธเนเธเธฒเธ (Active)</option>
                        <option value="เธงเนเธฒเธ">เธงเนเธฒเธ (Vacant)</option>
                        <option value="เธฃเธญเธเนเธญเธก">เธฃเธญเธเนเธญเธก (Repairing)</option>
                        <option value="เธชเธนเธเธซเธฒเธข">เธชเธนเธเธซเธฒเธข (Lost)</option>
                      </select>
                    </div>

                    <div className="lark-field-group">
                      <label>เธซเธกเธฒเธขเน€เธซเธ•เธธ / เธฃเธฒเธขเธฅเธฐเน€เธญเธตเธขเธ”เน€เธเธดเนเธกเน€เธ•เธดเธก</label>
                      <input type="text" className="lark-input" placeholder="เธ•เธฑเธงเธญเธขเนเธฒเธ: เธฃเธฑเธเน€เธเนเธฒเธเธฒเธเนเธเธฃเธเธเธฒเธฃเน€เธเธฅเธตเนเธขเธเน€เธเธฃเธทเนเธญเธเธเธต 2026" value={larkAssetNotes} onChange={e => setLarkAssetNotes(e.target.value)} />
                    </div>
                  </div>
                )}

                {(!larkSubmitted && (larkFormType !== 'ticket' || larkTicketRole !== 'it' || selectedPendingTicketSn !== '')) && (
                  <button type="submit" className="lark-submit-btn">เธชเนเธเธเธฑเธเธ—เธถเธเธเนเธญเธกเธนเธฅ (Submit Record)</button>
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
