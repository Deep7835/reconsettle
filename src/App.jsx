import { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx-js-style";

const CHARGE_RATE = 0.35;
const GST_RATE = 18;
const GST_OPTIONS = [0, 5, 12, 18, 28];

// Bank settlement cycle presets — chronological numbering. Cycle 3 settles the next day.
const RECON_CYCLE_PRESETS = [
  { id: 1, label: "Cycle 1", range: "12 AM → 6 AM",     start: "00:00", end: "06:00", startLabel: "12:00 AM", endLabel: "06:00 AM", settleNextDay: false },
  { id: 2, label: "Cycle 2", range: "6 AM → 4 PM",      start: "06:00", end: "16:00", startLabel: "06:00 AM", endLabel: "04:00 PM", settleNextDay: false },
  { id: 3, label: "Cycle 3", range: "4 PM → 11:59 PM",  start: "16:00", end: "23:59", startLabel: "04:00 PM", endLabel: "11:59 PM", settleNextDay: true  },
];

// Detect the cycle number from a bank file's name. Accepts "1.xlsx", "Cycle 1.xlsx",
// "C1 settlement.csv", "settlement 2 2026-05-16.xlsx", etc.
function detectBankFileCycle(filename) {
  if (!filename) return null;
  const base = filename.replace(/\.(csv|xlsx?|xls)$/i, "");
  const patterns = [
    /^([123])(?:[\b.\-_ ]|$)/,         // leading 1/2/3
    /\bcycle[\s_-]?([123])\b/i,        // "Cycle 1", "cycle_2"
    /\bc([123])\b/i,                    // "C1"
    /\b([123])(?:st|nd|rd)\b/i,         // "1st", "2nd", "3rd"
    /\b([123])\b/,                      // bare 1/2/3 anywhere
  ];
  for (const re of patterns) {
    const m = base.match(re);
    if (m) {
      const n = parseInt(m[1], 10);
      if (n >= 1 && n <= 3) return n;
    }
  }
  return null;
}

// Add days to an ISO yyyy-mm-dd date string. Timezone-safe (uses UTC throughout)
// so a date like "2026-05-16" + 1 day always returns "2026-05-17", regardless of the
// browser's local timezone.
function addDaysToIso(iso, days) {
  if (!iso) return "";
  const parts = iso.split("-").map((s) => parseInt(s, 10));
  if (parts.length !== 3 || parts.some(isNaN)) return iso;
  const [y, m, d] = parts;
  const date = new Date(Date.UTC(y, m - 1, d));
  date.setUTCDate(date.getUTCDate() + days);
  const yy = date.getUTCFullYear();
  const mm = String(date.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(date.getUTCDate()).padStart(2, "0");
  return `${yy}-${mm}-${dd}`;
}

// "2026-05-15" → "15-May-2026"
function formatDateLong(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const mi = parseInt(m, 10) - 1;
  if (mi < 0 || mi > 11) return iso;
  return `${d}-${months[mi]}-${y}`;
}

const MASTER_DATA = [
  { mid: "MER0000000030641", merchant: "Aryan", company: "ECOMPANTHER SOLUTION PRIVATE LIMITED", moa: "DONE", incorporation: "19-05-2025", gstNumber: "05AAICE9033A1Z1" },
  { mid: "MER0000000030642", merchant: "Aryan", company: "ECOMNEXTGEN TECHNOLOGY PRIVATE LIMITED", moa: "DONE", incorporation: "30-04-2025", gstNumber: "27AAICE8602M1Z9" },
  { mid: "MER0000000030643", merchant: "Aryan", company: "ZIPAT PEOPLE SERVICES AND CONSULTING PRIVATE LIMITED", moa: "DONE", incorporation: "24-07-2023", gstNumber: "27AACCZ2414J1ZA" },
  { mid: "MER0000000030644", merchant: "Aryan", company: "ALGOBLOOM TECHNHOLOGY PRIVATE LIMITED", moa: "DONE", incorporation: "17-12-2024", gstNumber: "27ABBCA8209Q1Z7" },
  { mid: "MER0000000030645", merchant: "Aryan", company: "LIBEERTON SOFTWARE SOLUTIONS PRIVATE LIMITED", moa: "DONE", incorporation: "18-12-2024", gstNumber: "27AAGCL0724P1Z7" },
  { mid: "MER0000000030648", merchant: "Aryan", company: "VILUMA INFOTECH PRIVATE LIMITED", moa: "DONE", incorporation: "25-07-2025", gstNumber: "07AALCV2680H1Z0" },
  { mid: "MER0000000030652", merchant: "Aryan", company: "LOGICGROVE TECHNOLOGY PRIVATE LIMITED", moa: "DONE", incorporation: "20-11-2024", gstNumber: "27AAGCL0406J1ZO" },
  { mid: "MER0000000030653", merchant: "Aryan", company: "ZOHA STORE PRIVATE LIMITED", moa: "DONE", incorporation: "05-09-2024", gstNumber: "32AACCZ4631P1Z0" },
  { mid: "MER0000000030657", merchant: "Aryan", company: "MOBIMUDRA TECH SOLUTION PRIVATE LIMITED", moa: "DONE", incorporation: "04-04-2025", gstNumber: "24AATCM1849H1Z4" },
  { mid: "MER0000000030658", merchant: "Aryan", company: "ENLIL TECHNOLOGY PRIVATE LIMITED", moa: "DONE", incorporation: "04-10-2023", gstNumber: "27AAHCE8083F1ZE" },
  { mid: "MER0000000030659", merchant: "Aryan", company: "CAMIDANO SOLUTIONS PRIVATE LIMITED", moa: "DONE", incorporation: "10-10-2024", gstNumber: "32AAMCC1729K1ZM" },
  { mid: "MER0000000030663", merchant: "Aryan", company: "TRENDTIDE DIGIMART PRIVATE LIMITED", moa: "DONE", incorporation: "18-09-2025", gstNumber: "27AAMCT3076C1Z9" },
  { mid: "MER0000000030664", merchant: "Aryan", company: "NEPIT TECHNOLOGY PRIVATE LIMITED", moa: "DONE", incorporation: "21-11-2024", gstNumber: "27AAJCN9961D1Z2" },
  { mid: "MER0000000030668", merchant: "Aryan", company: "PAYMARK SOLUTIONS PRIVATE LIMITED", moa: "DONE", incorporation: "19-06-2024", gstNumber: "32AAOCP8341E1ZE" },
  { mid: "MER0000000030671", merchant: "Rishabh", company: "RAPZAP TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "15-06-2024", gstNumber: "33AANCR727731Z2" },
  { mid: "MER0000000030673", merchant: "Rishabh", company: "COINNEX HUB PRIVATE LIMITED", moa: "", incorporation: "26-04-2024", gstNumber: "07AALCC8149A1ZP" },
  { mid: "MER0000000030683", merchant: "Nitin", company: "PRIME SOLUTIONS", moa: "", incorporation: "", gstNumber: "07NGWPS8968E1ZZ" },
  { mid: "MER0000000030708", merchant: "Aryan", company: "RKY ENTERTAINMENT PRIVATE LIMITED", moa: "DONE", incorporation: "10-12-2024", gstNumber: "27AAOCR1410N1Z4" },
  { mid: "MER0000000030709", merchant: "Aryan", company: "TRUEELEGENCE PRIVATE LIMITED", moa: "DONE", incorporation: "03-08-2025", gstNumber: "27AAMCT1327N1ZW" },
  { mid: "MER0000000030710", merchant: "Nitin", company: "ANIDEV MULTITRADING PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "09ABDCA5004F1Z3" },
  { mid: "MER0000000030746", merchant: "Aryan", company: "DORTEX IT SOLUTION PRIVATE LIMITED", moa: "DONE", incorporation: "14-10-2025", gstNumber: "07AALCD5415C1ZX" },
  { mid: "MER0000000030755", merchant: "Aryan", company: "ATHENA GOLD PRIVATE LIMITED", moa: "DONE", incorporation: "19-01-2024", gstNumber: "32ABACA0271H1ZA" },
  { mid: "MER0000000030756", merchant: "Aryan", company: "HUMBLE HAUTE ECOM SERVICES PVT LTD", moa: "DONE", incorporation: "11-05-2025", gstNumber: "05AAHCH9463L1Z2" },
  { mid: "MER0000000030784", merchant: "Sachin", company: "JSS Global Fintech Pvt Ltd", moa: "", incorporation: "16-08-2023", gstNumber: "07AAGCJ1223M1ZL" },
  { mid: "MER0000000030785", merchant: "Nitin", company: "PAY WANTAGE INDIA PVT LTD", moa: "", incorporation: "", gstNumber: "07AAKCP9575A1Z5" },
  { mid: "MER0000000030795", merchant: "Aryan", company: "Clickhorizon Traders Private Limited", moa: "DONE", incorporation: "10-12-2025", gstNumber: "07AANCC3226E1ZV" },
  { mid: "MER0000000030796", merchant: "Aryan", company: "PREMIERWEB TECHNOLOGY PVT LTD", moa: "DONE", incorporation: "12-11-2025", gstNumber: "07AAQCP3608L1ZY" },
  { mid: "MER0000000030797", merchant: "Aryan", company: "STOMBEO VENTURES PVT LTD", moa: "DONE", incorporation: "24-03-2025", gstNumber: "07ABPCS6776A1ZY" },
  { mid: "MER0000000030798", merchant: "Aryan", company: "KRISMON SOFTTECH PRIVATE LIMITED", moa: "DONE", incorporation: "05-08-2025", gstNumber: "07AAMCK0722E1ZT" },
  { mid: "MER0000000030801", merchant: "Nitin", company: "CANDYFLOW PRIVATE LIMITED", moa: "", incorporation: "22-08-2023", gstNumber: "09AALCC2617F1ZQ" },
  { mid: "MER0000000030803", merchant: "Aryan", company: "WICHARIT ECOM PRIVATE LIMITED", moa: "DONE", incorporation: "11-12-2025", gstNumber: "24AAECW3124C1ZV" },
  { mid: "MER0000000030823", merchant: "Aryan", company: "DOVIRA FOODS PRIVATE LIMITED", moa: "DONE", incorporation: "22-12-2020", gstNumber: "27AAICD2148K1ZH" },
  { mid: "MER0000000030848", merchant: "Aryan", company: "TRIVOXA GLOBAL PRIVATE LIMITED", moa: "DONE", incorporation: "11-12-2025", gstNumber: "07AAMCT5494N1ZE" },
  { mid: "MER0000000030849", merchant: "Aryan", company: "YOVZENA TECHNOLOGY PRIVATE LIMITED", moa: "DONE", incorporation: "6-12-2025", gstNumber: "07AACCY1203G1ZR" },
  { mid: "MER0000000030852", merchant: "Nitin", company: "SWASTIK SARV SEVA", moa: "", incorporation: "", gstNumber: "07ALEPR0215M2Z2" },
  { mid: "MER0000000030853", merchant: "Nitin", company: "WAREPRO TECH PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "09AAECW0394G1Z5" },
  { mid: "MER0000000030854", merchant: "Nitin", company: "ZENKAI PAYTECH OPC PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "09AACCZ4711A1ZN" },
  { mid: "MER0000000030855", merchant: "Aryan", company: "TIS THE INGNITE SPECTRUM PVT LTD", moa: "DONE", incorporation: "20-03-2025", gstNumber: "27AALCT7148J1ZQ" },
  { mid: "MER0000000030866", merchant: "Aryan", company: "SILNOZ FINTECH PRIVATE LIMITED", moa: "", incorporation: "23-05-2025", gstNumber: "24ABQCS3247P1ZM" },
  { mid: "MER0000000030873", merchant: "Rishabh", company: "AARUSHA PAYMENTS PRIVATE LIMITED", moa: "", incorporation: "02-04-2024", gstNumber: "09ABACA4545L1ZH" },
  { mid: "MER0000000030874", merchant: "Aryan", company: "LYSTNE ECOM PRIVATE LIMITED", moa: "DONE", incorporation: "08-12-2025", gstNumber: "24AAGCL6052L1ZC" },
  { mid: "MER0000000030875", merchant: "Aryan", company: "BRAVINO GLOBAL PRIVATE LIMITED", moa: "", incorporation: "12-11-2025", gstNumber: "07AAOCB1927C1ZV" },
  { mid: "MER0000000030894", merchant: "Aryan", company: "ZEBITECH INFO SOLUTIONS PRIVATE LIMITED", moa: "", incorporation: "08-08-2025", gstNumber: "07AACCZ7054J1ZY" },
  { mid: "MER0000000030895", merchant: "Aryan", company: "MEFOY TECH SERVICES PRIVATE LIMITED", moa: "", incorporation: "11-04-2025", gstNumber: "27AASCM6256B1Z8" },
  { mid: "MER0000000030921", merchant: "Rishabh", company: "SHREEANSH HANUMAT TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "03-11-2025", gstNumber: "20ABNCS8990E1Z0" },
  { mid: "MER0000000030980", merchant: "Aryan", company: "ERYNTAL TRADING PRIVATE LIMITED", moa: "", incorporation: "12-12-2025", gstNumber: "07AAJCE3347E1ZR" },
  { mid: "MER0000000030984", merchant: "Shashi", company: "SALESDOME TECHNOLOGIES PVT LTD", moa: "DONE", incorporation: "04-06-2025", gstNumber: "07ABNCS6674F1ZT" },
  { mid: "MER0000000030986", merchant: "Shashi", company: "YAZO INFOTECH PRIVATE LIMITED", moa: "DONE", incorporation: "23-07-2024", gstNumber: "09AABCY8713E1Z7" },
  { mid: "MER0000000030988", merchant: "Shashi", company: "SIREPLES INDIA PRIVATE LIMITED", moa: "DONE", incorporation: "27-02-2025", gstNumber: "08ABPCS3115B1ZJ" },
  { mid: "MER0000000030989", merchant: "Aryan", company: "PAYMOND SERVICES PRIVATE LIMITED", moa: "DONE", incorporation: "24-05-2022", gstNumber: "10AAMCP9293R1Z1" },
  { mid: "MER0000000031040", merchant: "Shashi", company: "TRUEWIN INFOTECH PRIVATE LIMITED", moa: "", incorporation: "14-08-2025", gstNumber: "09AAMCT1689B1Z2" },
  { mid: "MER0000000031041", merchant: "Aryan", company: "ALTRIX AGRO FARM PRIVATE LIMITED", moa: "", incorporation: "03-02-2025", gstNumber: "27ABCCA1114D1ZF" },
  { mid: "MER0000000031042", merchant: "Nilesh", company: "HAVISTA VENTURES PRIVATE LIMITED", moa: "", incorporation: "10-09-2025", gstNumber: "08AAICH1862Q1ZY" },
  { mid: "MER0000000031043", merchant: "Nilesh", company: "HEALLINGMIND RETAIL PRIVATE LIMITED", moa: "", incorporation: "11-12-2024", gstNumber: "19AAHCH6513R1ZW" },
  { mid: "MER0000000031044", merchant: "Shashi", company: "SPHRAN TECHNOLOGIES PRIVATE LIMITED", moa: "DONE", incorporation: "06-08-2024", gstNumber: "36ABCCS8088G2ZW" },
  { mid: "MER0000000031060", merchant: "Shashi", company: "OPULENT FUELS SERVICES PRIVATE LIMITED", moa: "DONE", incorporation: "11-11-2020", gstNumber: "09AADCO3166B1ZO" },
  { mid: "MER0000000031061", merchant: "Shashi", company: "OMAIMZ MARKETING PRIVATE LIMITED", moa: "", incorporation: "18-02-2025", gstNumber: "27AAECO6117H1ZG" },
  { mid: "MER0000000031062", merchant: "Shashi", company: "INNOVATORS MARINE PRIVATE LIMITED", moa: "DONE", incorporation: "19-12-2025", gstNumber: "27A ACCI1924J1ZM" },
  { mid: "MER0000000031063", merchant: "Shashi", company: "CODEVA SOLUTIONS PRIVATE LIMITED", moa: "DONE", incorporation: "25-07-2025", gstNumber: "27AAMCC9338A1ZL" },
  { mid: "MER0000000031064", merchant: "Shashi", company: "FELUNOR SUPPLIER PRIVATED LIMITED", moa: "DONE", incorporation: "17-09-2025", gstNumber: "07AAGCF6262L1ZA" },
  { mid: "MER0000000031065", merchant: "Shashi", company: "WEARHOUSE PRIVATE LIMITED", moa: "", incorporation: "04-12-2025", gstNumber: "09AAECW3023G1ZH" },
  { mid: "MER0000000031068", merchant: "Nilesh", company: "TRANSACTGRID TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "10-12-2025", gstNumber: "27AAMCT5447R1ZB" },
  { mid: "MER0000000031069", merchant: "Nilesh", company: "WEB STORE PRIVATE LIMITED", moa: "", incorporation: "18-02-2025", gstNumber: "24AAECW0363M1Z7" },
  { mid: "MER0000000031070", merchant: "Nilesh", company: "AXIOMIXX PRIVATE LIMITED", moa: "", incorporation: "15-12-2025", gstNumber: "09ABECA3526F1ZV" },
  { mid: "MER0000000031077", merchant: "Nilesh", company: "PAYMIQUE STUDIOS PRIVATE LIMITED", moa: "", incorporation: "06-05-2025", gstNumber: "37AAPCP7452B1Z7" },
  { mid: "MER0000000031085", merchant: "AJ", company: "METALLO TRADERS PRIVATE LIMITED", moa: "", incorporation: "12-09-2025", gstNumber: "07AASCM5493N1ZG" },
  { mid: "MER0000000031086", merchant: "AJ", company: "CARTOPIA COLLECTIONS PRIVATE LIMITED", moa: "", incorporation: "30-09-2024", gstNumber: "07AAMCC1530G2ZW" },
  { mid: "MER0000000031098", merchant: "Nilesh", company: "RVMPAY TECHNOLOGIES PVT LTD", moa: "", incorporation: "01-03-2025", gstNumber: "24AAOCR3256L1ZY" },
  { mid: "", merchant: "Sachin", company: "BARRINGER", moa: "", incorporation: "23-04-2008", gstNumber: "" },
  { mid: "", merchant: "Maheeps", company: "CYROTECH PRIVATE LIMITED", moa: "", incorporation: "18-12-2024", gstNumber: "27AAMCC3002B1ZA" },
  { mid: "", merchant: "Maheeps", company: "MINIZONE FINTECH LLP.", moa: "", incorporation: "17-11-2024", gstNumber: "09ACCFM1690E1ZA" },
  { mid: "", merchant: "Nilesh", company: "ONE STOP SHOPPING STATION PVT LTD", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan", company: "OSILA SOLUTION PRIVATE LIMITED", moa: "DONE", incorporation: "18-04-2025", gstNumber: "27AAECO4832D1ZM" },
  { mid: "", merchant: "Aryan", company: "RJ PAPILLON MULTITRADE PRIVATE LIMITED", moa: "DONE", incorporation: "01-11-2018", gstNumber: "27AAJCR2324N1Z2" },
  { mid: "", merchant: "Maheeps", company: "RUPEECARE", moa: "", incorporation: "30-08-2022", gstNumber: "27ABFFR6242P1ZL" },
  { mid: "", merchant: "Aryan", company: "ZENTRITECH INFOTECH PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan-Sasi", company: "APEXKART PVT LTD", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan-Sasi", company: "KRAYBAZAR PVT LTD", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan-Sasi", company: "SPACEMAZE PVT LTD", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan-Sasi", company: "COPLOST ENTERPRISES PVT LTD", moa: "", incorporation: "", gstNumber: "" },
  { mid: "MER0000000031108", merchant: "Aryan", company: "ECOJUST ECOM PRIVATE LIMITED", moa: "DONE", incorporation: "25-09-2024", gstNumber: "07AAICE4188J1Z8" },
  { mid: "", merchant: "Shashi", company: "KONNECT INFOCOM (OPC) PRIVATE LIMITED", moa: "DONE", incorporation: "05-04-2024", gstNumber: "36AAKCK9358C1Z8" },
  { mid: "", merchant: "Aryan", company: "DEPAM TECHNOLOGY PRIVATE LIMITED", moa: "DONE", incorporation: "12-01-2024", gstNumber: "27AAKCD5632N1Z6" },
  { mid: "MER0000000031122", merchant: "Nilesh", company: "HOSJA TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "07-07-2025", gstNumber: "19AAICH0527J1ZH" },
  { mid: "MER0000000031126", merchant: "Nilesh", company: "KUWARINU INTERNATIONAL TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "01-07-2025", gstNumber: "09AALCK9876C1ZX" },
  { mid: "MER0000000031176", merchant: "Shashi", company: "PRAHITO VOYAGES INDIA PRIVATE LIMITED", moa: "DONE", incorporation: "12-03-2025", gstNumber: "07AAPCP5546C1Z9" },
  { mid: "MER0000000031206", merchant: "Shashi", company: "PRADYUT ENTERPRISES PRIVATE LIMITED", moa: "DONE", incorporation: "03-10-2025", gstNumber: "06AAQCP2533A1ZP" },
  { mid: "MER0000000031192", merchant: "Shashi", company: "CKLS CLOTHING PRIVATE LIMITED", moa: "DONE", incorporation: "12-06-2024", gstNumber: "07AAMCC5579K1ZZ" },
  { mid: "MER0000000031125", merchant: "Nilesh", company: "ZENCODE AI PRIVATE LIMITED", moa: "DONE", incorporation: "25-07-2025", gstNumber: "33AACCZ6928M1ZS" },
  { mid: "MER0000000031114", merchant: "Nilesh", company: "TRADELINKER INNOVATIVE SOLUTIONS PRIVATE LIMITED", moa: "DONE", incorporation: "22-11-2025", gstNumber: "33AAMCT4805F1ZF" },
  { mid: "MER0000000031124", merchant: "Nilesh", company: "KATROR SYSTEMS TECHNOLOGY PRIVATE LIMITED", moa: "DONE", incorporation: "06-08-2025", gstNumber: "33AAMCK0734A1Z2" },
  { mid: "MER0000000031113", merchant: "Nilesh", company: "DIGIPOUCH TECHNOLOGY PRIVATE LIMITED", moa: "DONE", incorporation: "19-04-2025", gstNumber: "33AALCD6226H1ZP" },
  { mid: "MER0000000031112", merchant: "Nilesh", company: "CODESNAP TECHNOLOGY PRIVATE LIMITED", moa: "DONE", incorporation: "07-05-2025", gstNumber: "34AAMCC7082P1ZY" },
  { mid: "MER0000000031115", merchant: "Nilesh", company: "KOPYSYNK TECHNOLOGY PRIVATE LIMITED", moa: "DONE", incorporation: "14-05-2025", gstNumber: "33AALCK8632N1ZZ" },
  { mid: "MER0000000031118", merchant: "Nilesh", company: "AKBARLINK E-SERVICES PRIVATE LIMITED", moa: "DONE", incorporation: "02-12-2025", gstNumber: "33ABECA2317E1ZB" },
  { mid: "MER0000000031130", merchant: "AJ", company: "STARPAYOUT SOLUTIONNS PRIVATE LIMITED", moa: "DONE", incorporation: "24-11-2023", gstNumber: "09ABMCS1757J1ZR" },
  { mid: "MER0000000031129", merchant: "AJ", company: "PERFAITE TECHNOLOGIES PRIVATE LIMITED", moa: "DONE", incorporation: "28-12-2023", gstNumber: "07AAOCP3482C1ZB" },
  { mid: "MER0000000031119", merchant: "Nilesh", company: "PRATIMDE SOLUTIONS PRIVATE LIMITED", moa: "DONE", incorporation: "28-08-2025", gstNumber: "19AAQCP1281P1ZK" },
  { mid: "MER0000000031117", merchant: "Nilesh", company: "ZENWEI TECHNOLOGIES PRIVATE LIMITED", moa: "DONE", incorporation: "29-04-2025", gstNumber: "33AACCZ6280A1ZK" },
  { mid: "MER0000000031120", merchant: "Nilesh", company: "NINMU TECHNOLOGY PRIVATE LIMITED", moa: "DONE", incorporation: "01-07-2025", gstNumber: "09AAKCN6077H1ZY" },
  { mid: "MER0000000031116", merchant: "Nilesh", company: "DEALDISHA ENTERPRISE PRIVATE LIMITED", moa: "DONE", incorporation: "20-04-2025", gstNumber: "19AALCD6240P1Z0" },
  { mid: "MER0000000031121", merchant: "Nilesh", company: "PUDERO TECH PRIVATE LIMITED", moa: "DONE", incorporation: "21-01-2025", gstNumber: "07AAPCP3612D1ZK" },
  { mid: "MER0000000031123", merchant: "Nilesh", company: "VELOGENIX TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "24-07-2025", gstNumber: "09AALCV2639Q1ZE" },
  { mid: "MER0000000031131", merchant: "Nilesh", company: "ORIVENTA PRIVATE LIMITED", moa: "", incorporation: "17-10-2025", gstNumber: "08AAECO9143M1ZX" },
  { mid: "MER0000000031170", merchant: "Nilesh", company: "ZELTRIUM TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "30-12-2025", gstNumber: "33AACCZ8531E1ZE" },
  { mid: "MER0000000031128", merchant: "Nilesh", company: "ARKABYTE DATA SOLUTIONS PRIVATE LIMITED", moa: "DONE", incorporation: "16-05-2025", gstNumber: "19ABCCA9267H1ZA" },
  { mid: "MER0000000031153", merchant: "AK", company: "JOVLERA TRADING PRIVATE LIMITED", moa: "", incorporation: "06-12-2025", gstNumber: "07AAHCJ1595M1Z1" },
  { mid: "MER0000000031136", merchant: "AJ", company: "STPK SERVICES PRIVATE LIMITED", moa: "", incorporation: "25-02-2025", gstNumber: "09ABPCS3882F1ZR" },
  { mid: "MER0000000031159", merchant: "Nilesh", company: "AIVEX DIGITAL SOLUTIONS PRIVATE LIMITED", moa: "", incorporation: "10-02-2026", gstNumber: "33ABECA8639F1ZO" },
  { mid: "MER0000000031158", merchant: "Nilesh", company: "GAYTI INFRATECH PRIVATE LIMITED", moa: "DONE", incorporation: "22-09-2025", gstNumber: "09AAMCG3937B2ZJ" },
  { mid: "MER0000000031162", merchant: "Nilesh", company: "HYPERNOVA TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "11-08-2025", gstNumber: "09AAICH1132P1ZB" },
  { mid: "MER0000000031161", merchant: "Nilesh", company: "KARANARJUN TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "28-07-2025", gstNumber: "09AAMCK0549H1ZA" },
  { mid: "MER0000000031163", merchant: "Nilesh", company: "NISHASOLUTION TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "4-12-2025", gstNumber: "07AALCN0612R1Z3" },
  { mid: "MER0000000031160", merchant: "Nilesh", company: "FASAFIN TECKY SOLUTION PRIVATE LIMITED", moa: "", incorporation: "29-01-2026", gstNumber: "33AAGCF8471P1Z0" },
  { mid: "MER0000000031155", merchant: "Nilesh", company: "ARAB WINGS TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "30-11-2023", gstNumber: "32AAZCA7632A2ZP" },
  { mid: "MER0000000031171", merchant: "Nilesh", company: "APPAREL ROUTE PRIVATE LIMITED", moa: "", incorporation: "09-12-2025", gstNumber: "29ABECA2948R1ZW" },
  { mid: "MER0000000031137", merchant: "Nilesh", company: "RASHEEYA TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "20-11-2025", gstNumber: "33AAPCR2254P1ZU" },
  { mid: "MER0000000031157", merchant: "Nilesh", company: "LUDIK TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "15-01-2026", gstNumber: "06AAGCL6791A1ZJ" },
  { mid: "MER0000000031156", merchant: "Nilesh", company: "BACHAR TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "30-05-2024", gstNumber: "19AAMCB6439K1Z2" },
  { mid: "MER0000000031166", merchant: "Nilesh", company: "EYEBRAWN TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "05-08-2024", gstNumber: "37AAICE3363K1ZC" },
  { mid: "MER0000000031179", merchant: "AK", company: "TROVINA GLOBAL PRIVATE LIMITED", moa: "", incorporation: "12-12-2025", gstNumber: "07AAMCT5561E1Z5" },
  { mid: "MER0000000031152", merchant: "AK", company: "VISTOWISE TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "06-12-2025", gstNumber: "07AALCV6340E1Z9" },
  { mid: "MER0000000031151", merchant: "AK", company: "WELYRON TRADING PRIVATE LIMITED", moa: "", incorporation: "06-12-2025", gstNumber: "07AAECW3050B1ZS" },
  { mid: "MER0000000031167", merchant: "Nilesh", company: "VANNITAAY TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "31-12-2024", gstNumber: "19AAKCV6513G1Z2" },
  { mid: "MER0000000031165", merchant: "Nilesh", company: "CREDI REVIVE TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "21-06-2025", gstNumber: "09AAMCC8478P1ZH" },
  { mid: "", merchant: "Nilesh", company: "FIRST BUZZ TECH PRIVATE LIMITED", moa: "", incorporation: "16-04-2025", gstNumber: "09AAGCF3554M1Z7" },
  { mid: "MER0000000031164", merchant: "Nilesh", company: "NEATNODE INNOVATIONS PRIVATE LIMITED", moa: "", incorporation: "27-01-2025", gstNumber: "19AAKCN1311Q1Z3" },
  { mid: "MER0000000031168", merchant: "Nilesh", company: "METAFLECK TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "09-05-2025", gstNumber: "27AATCRΙ3556H1ZY" },
  { mid: "MER0000000031169", merchant: "Nilesh", company: "FABRIVIO INNOVATIONS PRIVATE LIMITED", moa: "", incorporation: "13-10-2025", gstNumber: "27AAGCTF6610G1ZR" },
  { mid: "MER0000000031172", merchant: "Nilesh", company: "SHOPZYARA ENTERPRISED PRIVATE LIMITED", moa: "", incorporation: "30-11-2025", gstNumber: "27ABRCS9155R1Z0" },
  { mid: "MER0000000031173", merchant: "Nilesh", company: "BMJB ECOM PRIVATE LIMITED", moa: "", incorporation: "03-12-2025", gstNumber: "19AAOCB1572F1ZJ" },
  { mid: "MER0000000031174", merchant: "Nilesh", company: "PROTAY TRADERS PRIVATE LIMITED", moa: "", incorporation: "17-11-2024", gstNumber: "19AAPCP2014H1ZD" },
  { mid: "MER0000000031175", merchant: "Nilesh", company: "ZENOPAY SOLUTIONS PRIVATE LIMITED", moa: "", incorporation: "01-09-2025", gstNumber: "19AACCZ7300F1ZD" },
  { mid: "", merchant: "Nilesh", company: "RVMP OPC SERVICE", moa: "", incorporation: "", gstNumber: "" },
  { mid: "MER0000000031184", merchant: "Nilesh", company: "WILDBADGER TECNOLOGY PRIVATE LIMITED", moa: "", incorporation: "15-01-2026", gstNumber: "29AAECW3495Q1ZA" },
  { mid: "MER0000000031185", merchant: "Nilesh", company: "POROZE TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "26-06-2024", gstNumber: "19AAOCP8637P1Z8" },
  { mid: "MER0000000031188", merchant: "Nilesh", company: "EYE-HYVE TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "10-07-2024", gstNumber: "37AAICE2914L1ZF" },
  { mid: "MER0000000031178", merchant: "Nilesh", company: "NEBULUXA SYSTEMS PRIVATE LTD.", moa: "", incorporation: "21-07-2025", gstNumber: "29AAKCN6571E1Z3" },
  { mid: "MER0000000031186", merchant: "Nilesh", company: "RUNNINGWAY TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "08-08-2025", gstNumber: "29AAOCR8999N1ZS" },
  { mid: "MER0000000031204", merchant: "Nilesh", company: "KUICKYTECH SOLUTION PRIVATE LIMITED", moa: "", incorporation: "28-01-2026", gstNumber: "33AAMCK5448B1ZN" },
  { mid: "MER0000000031202", merchant: "Nilesh", company: "NETRATECH PRIVATE LIMITED", moa: "", incorporation: "17-11-2025", gstNumber: "24AALCN0221E1Z1" },
  { mid: "MER0000000031189", merchant: "Nilesh", company: "KARGIL IT SOFTTECH PRIVATE LIMITED", moa: "", incorporation: "24-07-2025", gstNumber: "27AAMCK0431G1ZP" },
  { mid: "MER0000000031203", merchant: "Nilesh", company: "JEWNOX SOLUTIONS PRIVATE LIMITED", moa: "", incorporation: "10-10-2024", gstNumber: "19AAGCJ6042A1ZU" },
  { mid: "MER0000000031177", merchant: "KJ", company: "ADAMBIL ENTERPRISES PRIVATE LIMITED", moa: "DONE", incorporation: "23-10-2025", gstNumber: "07ABECA0147N1ZN" },
  { mid: "", merchant: "Aryan", company: "WHALESTATION ENTERPRISE PRIVATE LIMITED", moa: "", incorporation: "05-08-2024", gstNumber: "24AADCW8963G1ZZ" },
  { mid: "MER0000000031191", merchant: "AJ", company: "NPVS VENTURES PRIVATE LIMITED", moa: "", incorporation: "30-10-2024", gstNumber: "24AAJCN9585B1Z8" },
  { mid: "", merchant: "Aryan", company: "FIDELEX CONSULTANCY PRIVATE LIMITED", moa: "", incorporation: "21-12-2024", gstNumber: "27AAGCF1721J1ZQ" },
  { mid: "MER0000000031220", merchant: "Nilesh", company: "GORA SALES AND DISTRIBUTION PRIVATE LIMITED", moa: "DONE", incorporation: "21-02-2025", gstNumber: "08AALCG7983Q1ZE" },
  { mid: "MER0000000031221", merchant: "Nilesh", company: "BIZOVENT PRIVATE LIMITED", moa: "DONE", incorporation: "10-10-2025", gstNumber: "07AAOCB0274E1ZT" },
  { mid: "MER0000000031222", merchant: "Nilesh", company: "SOFTLYVE TECHNOLOGIES PRIVATE LIMITED", moa: "DONE", incorporation: "04-02-2025", gstNumber: "07ABPCS2027N1ZV" },
  { mid: "MER0000000031223", merchant: "Nilesh", company: "JANTODEC PRIVATE LIMITED", moa: "DONE", incorporation: "20-02-2026", gstNumber: "33AAHCJ2843A1Z2" },
  { mid: "MER0000000031215", merchant: "Nilesh", company: "ORENZA PRIVATE LIMITED", moa: "DONE", incorporation: "05-02-2026", gstNumber: "21AAFCO0484L1ZG" },
  { mid: "MER0000000031214", merchant: "Nilesh", company: "DABLU BABLU TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "04-07-2025", gstNumber: "09AALCD8425G1ZD" },
  { mid: "MER0000000031187", merchant: "Nilesh", company: "SANJAYDEY TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "08-09-2025", gstNumber: "19ABOCS2861H1ZU" },
  { mid: "MER0000000031275", merchant: "Nilesh", company: "AMBB TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "02-09-2024", gstNumber: "19ABBCA3174C1ZZ" },
  { mid: "MER0000000031219", merchant: "Nilesh", company: "XENIAL XPRESS STUDIO PRIVATE LIMITED", moa: "", incorporation: "16-03-2026", gstNumber: "09AAACX6286C1Z5" },
  { mid: "MER0000000031216", merchant: "AJ", company: "SKYNEXIS CONSULTANCY & TECHNOLOGY PVT LTD", moa: "", incorporation: "02-07-2025", gstNumber: "07ABQCS6491K1ZG" },
  { mid: "MER0000000031218", merchant: "Nilesh", company: "QUBIONIX SOLUTIONS PRIVATE LIMITED", moa: "", incorporation: "08-08-2025", gstNumber: "06AABCQ1502H1ZY" },
  { mid: "", merchant: "Nitin", company: "SHREE KHATU SHYAM ENTERPRISES", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nitin", company: "CIRCUTECH SYSTEMS PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "MER0000000031217", merchant: "Nilesh", company: "REVOLUTIONARY ONLINE PRIVATE LIMITED", moa: "", incorporation: "22-10-2025", gstNumber: "19AAPCR1479N1ZF" },
  { mid: "MER0000000031190", merchant: "KJ", company: "VISIONARYVISTA TECH PRIVATE LIMITED", moa: "DONE", incorporation: "15-01-2025", gstNumber: "27AAKCV6794R1ZZ" },
  { mid: "MER0000000031205", merchant: "KJ", company: "DULAARAA PUBLICITY AND MEDIA PRIVATE LIMITED", moa: "", incorporation: "29-08-2024", gstNumber: "09AALCD0872F1ZK" },
  { mid: "", merchant: "KJ", company: "MATILEO ENTERPRISES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "KJ", company: "WENSET INFOCOM PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan", company: "DIGITALFORGE TECHNOLOGIES (OPC) PRIVATE LIMITED", moa: "", incorporation: "01-01-2026", gstNumber: "09AAMCD3121C1Z1" },
  { mid: "MER0000000031272", merchant: "Nilesh", company: "RAGU SALES AND DISTRIBUTION PRIVATE LIMITED", moa: "DONE", incorporation: "27-06-2025", gstNumber: "08AAOCR7806C1Z7" },
  { mid: "MER0000000031267", merchant: "AJ", company: "ADDIRE BEYOND STYLE PRIVATE LIMITED", moa: "", incorporation: "14-11-2017", gstNumber: "09AAQCA2244P1Z3" },
  { mid: "MER0000000031277", merchant: "Shashi", company: "ICHIBAN ECOMMERCE PRIVATE LIMITED", moa: "", incorporation: "27-03-2025", gstNumber: "27AAICI1154N1ZA" },
  { mid: "MER0000000031268", merchant: "AJ", company: "AIWATLAN LOGISTICS PRIVATE LIMITED", moa: "", incorporation: "16-09-2025", gstNumber: "07ABBCA6157D1ZX" },
  { mid: "", merchant: "KJ", company: "LUXELAYER EDGE PRIVATE LIMITED", moa: "", incorporation: "19-01-2026", gstNumber: "" },
  { mid: "", merchant: "Aryan", company: "RAB TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "30-05-2023", gstNumber: "27AAMCR7195N1ZC" },
  { mid: "MER0000000031273", merchant: "Nilesh", company: "MEW BIT TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "27-03-2025", gstNumber: "24AATCM1534F1ZI" },
  { mid: "", merchant: "Nilesh", company: "VENTES SALES AND DISTRIBUTION PRIVATE LIMITED", moa: "", incorporation: "27-06-2025", gstNumber: "08AALCV1956E1Z3" },
  { mid: "MER0000000031271", merchant: "Nilesh", company: "SPARKREVOLUTION TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "09-05-2025", gstNumber: "29ABQCS1967A1Z0" },
  { mid: "MER0000000031280", merchant: "AJ", company: "DYMA INFOTECH PRIVATE LIMITED", moa: "", incorporation: "13-10-2025", gstNumber: "07AAMCD1083H1ZM" },
  { mid: "MER0000000031279", merchant: "AJ", company: "KAVYANSH ENTREPRISES PRIVATE LIMITED", moa: "", incorporation: "08-04-2024", gstNumber: "07AAKCK9401A1ZT" },
  { mid: "", merchant: "Rishabh", company: "UDTEJJ DIGITAL SOLUTIONS PRIVATE LIMITED", moa: "", incorporation: "23-01-2026", gstNumber: "36AAECU0606R1Z0" },
  { mid: "", merchant: "AJ", company: "PHENOX AEROSPACE  INDIA PRIVATE LIMITED", moa: "", incorporation: "10-04-2024", gstNumber: "07AAOCP6508R1ZJ" },
  { mid: "MER0000000031274", merchant: "Nilesh", company: "BARAI AND BISWAS TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "23-02-2025", gstNumber: "19AANCB2729H1ZF" },
  { mid: "", merchant: "KJ", company: "PVKA SOFTWARE PRIVATE LIMITED", moa: "", incorporation: "08-02-2025", gstNumber: "08AAPCP4233N1ZU" },
  { mid: "MER0000000031281", merchant: "AJ", company: "PIONEERS HUB PRIVATE LIMITED", moa: "", incorporation: "23-12-2025", gstNumber: "27AAQCP5030Q1ZQ" },
  { mid: "", merchant: "AJ", company: "GROSMART VENTURES PRIVATE LIMITED", moa: "", incorporation: "09-12-2025", gstNumber: "07AAMCG5648P1ZR" },
  { mid: "", merchant: "AJ", company: "TRUAXIS VENTURES PRIVATE LIMITED", moa: "", incorporation: "15-01-2026", gstNumber: "07AAMCT6963A1Z5" },
  { mid: "", merchant: "Nilesh", company: "INTCOM TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "30-05-2025", gstNumber: "29AAICI2425E1ZP" },
  { mid: "", merchant: "Nilesh", company: "DIPRATA TECHONOLOGY PRIVATE LIMITED", moa: "", incorporation: "25-07-2024", gstNumber: "19AALCD0162E1ZU" },
  { mid: "", merchant: "Nilesh", company: "KSHIRA FIN PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "36AALCK1084E1ZK" },
  { mid: "", merchant: "AJ", company: "BARRINGER PHARMA PRIVATE LIMITED", moa: "", incorporation: "23-04-2018", gstNumber: "07AAHCB7897G1Z5" },
  { mid: "", merchant: "Nilesh", company: "KKBM ENTERPRISES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "VALTARA SALES AND DISTRIBUTION PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "AZUP INNOVATION PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "MONTJAVE VENTURES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "ZIRK TOWN MART PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "DEALDROP TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "AJ", company: "YAKSHAN DIGITECH PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "AJ", company: "PEGYOU TRADEMART PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "AJ", company: "GODAVARY BIZTECH PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "AJ", company: "LOGISQUIRE SOLUTIONS PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "CLICKFUSION TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "AJ", company: "SAVARIYAAN TECHNORAFT PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "WNY IT SERVICES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "VELCYNTRA TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "SHIMSHA TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "AJ", company: "ELANINE TRADE ENTERPRISES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "AJ", company: "PARCENTO SOLUTIONS PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "RMVS SALES AND MARKETING PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "ORIZONTEX TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "VIBEZWEAR TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "ZENFIT TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "AJ", company: "ECOMMXCEL VENTURE PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "RONAM SOLUTION PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan", company: "AKPRAS INFOTECH SYSTEM PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan", company: "FABRICPORT INDIA PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "KHATU TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "CRYSTALBYTE TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "SYNTHETICS E COM SOLUTIONS PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan", company: "AVISHKKAR CATERING PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "PRSNB ENTERPRISES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "RAYTWIN SOLUTION PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "VRINDAVAN DIGITAL MARKETING PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "BOKARO GLOBAL TRADER PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "MANOHIT TESHNI MEDIA SOLUTIONS PRIVATE ILIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan", company: "MEDEDTECH GLOBAL SOLUTIONS PRIVATE LIMITES", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan", company: "SALTION ELECTRONICS ECOM PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "TECHBRIGHT TECHNOLOGIES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "SHARTEJI TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan", company: "DHRUVIKA WORKFORCE PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan", company: "OPULENT METAL WORKS PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "ELITEVERGE TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan", company: "FORGEFLOW IRONWORKS PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "AJ", company: "RUDRA FOOD & BEVERAGE PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "AJ", company: "SKILLION PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "ZERIVOX LOGISTICS PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "VASTRAKRITI RETAIL PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "REFIONTEX COM PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "ADBLOWZ SOLUTION PRIVATE LIMTED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "KLYRA TECHONOLOGIES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "WHITEBOX SPORTS AND ENTERTAINMENT PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Aryan", company: "ULTRIXA ECOMTECH PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "BROWWW TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "AETHERION TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "Nilesh", company: "RELICSPHERE PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
];

const ALL_MERCHANTS = [...new Set(MASTER_DATA.map((d) => d.merchant))].sort();
const REASONS = [
  "Rounding",
  "Missing Txn",
  "Chargeback/Refund",
  "Cycle Mismatch",
  "Bank Charges",
  "TDS Deducted",
  "Gateway Error",
  "Other",
];

function formatINR(n) {
  if (isNaN(n) || n === 0) return "\u20B90.00";
  return (
    "\u20B9" +
    Number(n).toLocaleString("en-IN", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })
  );
}

function formatShort(n) {
  if (n >= 1e7) return "\u20B9" + (n / 1e7).toFixed(2) + " Cr";
  if (n >= 1e5) return "\u20B9" + (n / 1e5).toFixed(2) + " L";
  if (n >= 1e3) return "\u20B9" + (n / 1e3).toFixed(1) + "K";
  return formatINR(n);
}

const CYCLE_META = {
  1: { label: "6 AM → 4 PM", color: "#10b981", start: "6:00 AM", end: "4:00 PM" },
  2: { label: "4 PM → 6 AM", color: "#f59e0b", start: "4:00 PM", end: "6:00 AM" },
};
function getCycle() {
  const h = new Date().getHours();
  const inDay = h >= 6 && h < 16;
  return inDay
    ? { cycle: 1, ...CYCLE_META[1] }
    : { cycle: 2, ...CYCLE_META[2] };
}

function sumR(a) {
  return a.reduce(
    (s, r) => ({
      payin: s.payin + r.payin,
      charge: s.charge + r.charge,
      gst: s.gst + r.gst,
      chargeback: s.chargeback + (r.chargeback || 0),
      deduction: s.deduction + r.deduction,
      settlement: s.settlement + r.settlement,
      count: s.count + 1,
    }),
    { payin: 0, charge: 0, gst: 0, chargeback: 0, deduction: 0, settlement: 0, count: 0 }
  );
}

function LiveClock() {
  const [t, sT] = useState(new Date());
  useEffect(() => {
    const i = setInterval(() => sT(new Date()), 1000);
    return () => clearInterval(i);
  }, []);
  return (
    <span>
      {t.toLocaleTimeString("en-IN", {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        hour12: true,
      })}
    </span>
  );
}

const SI = ({ d }) => (
  <svg
    width="18"
    height="18"
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <path d={d} />
  </svg>
);

const C = {
  bg: "#f0f2f5",
  sidebar: "#1a2e35",
  sideH: "#243d47",
  sideA: "#2d4f5c",
  accent: "#1a7f64",
  accentL: "#d1fae5",
  card: "#ffffff",
  text: "#1e293b",
  muted: "#64748b",
  border: "#e2e8f0",
  green: "#10b981",
  orange: "#f59e0b",
  red: "#ef4444",
  blue: "#3b82f6",
};

export default function App() {
  const [page, setPage] = useState("dashboard");
  const [selM, setSM] = useState("");
  const [selC, setSC] = useState("");
  const [payin, setPay] = useState("");
  const [chargeback, setChargeback] = useState("");
  const [gstRate, setGstRate] = useState(GST_RATE);
  const [entryCycle, setEntryCycle] = useState(null); // null = auto-detect
  const [entryDate, setEntryDate] = useState(""); // "" = today
  const [dlFrom, setDlFrom] = useState("");
  const [dlTo, setDlTo] = useState("");
  // Bank statement upload state
  // Bank settlement supports multiple files (e.g. 3 settlement cycles per day);
  // each row gets a `_sourceFile` tag so we can remove an individual file later.
  const [bankFiles, setBankFiles] = useState([]); // [{ name, rowCount }]
  const [bankRows, setBankRows] = useState([]); // raw rows per company from all bank files
  const [bankError, setBankError] = useState("");
  const [bankDragging, setBankDragging] = useState(false);
  // Source/internal payin upload state (just Name + Sum of Amount per company)
  // Settlement period for the recon (date + time range). Embedded in the downloaded Excel.
  const [reconDate, setReconDate] = useState("");       // ISO yyyy-mm-dd
  const [reconStartTime, setReconStartTime] = useState("");
  const [reconEndTime, setReconEndTime] = useState("");
  // Filter the recon table and the downloaded Excel to a single parent merchant.
  const [reconMerchantFilter, setReconMerchantFilter] = useState("__all__");

  // Persisted recons (localStorage). Each entry captures the inputs needed to restore
  // the full recon view: merchant, date+time, uploaded files, and the parsed rows.
  const RECON_STORAGE_KEY = "settleops_saved_recons_v1";
  const [savedRecons, setSavedRecons] = useState(() => {
    try {
      const raw = localStorage.getItem(RECON_STORAGE_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch {
      return [];
    }
  });
  useEffect(() => {
    try {
      localStorage.setItem(RECON_STORAGE_KEY, JSON.stringify(savedRecons));
    } catch (err) {
      // Quota exceeded or storage unavailable — silently degrade (recons stay in memory only).
      // eslint-disable-next-line no-console
      console.warn("Could not persist saved recons:", err && err.message);
    }
  }, [savedRecons]);
  const [savedReconsFilter, setSavedReconsFilter] = useState("__all__");
  const [sourceFile, setSourceFile] = useState("");
  const [sourceRows, setSourceRows] = useState([]);
  const [sourceError, setSourceError] = useState("");
  const [sourceDragging, setSourceDragging] = useState(false);
  const [mSearch, setMSearch] = useState("");
  const [mFilter, setMFilter] = useState("__all__");
  const [mMoaFilter, setMMoaFilter] = useState("__all__");
  const [rows, setRows] = useState([]);
  const [view, setView] = useState("all");
  const [fM, setFM] = useState("__all__");
  const [fC, setFC] = useState("__all__");
  const [cSearch, setCS] = useState("");
  const [sideOpen, setSO] = useState(true);

  const [recon, setRecon] = useState([]);
  const [rMerchant, setRM] = useState("");
  const [rCompany, setRC] = useState("");
  const [rTxnId, setRTxn] = useState("");
  const [rPayin, setRPay] = useState("");
  const [rClaimed, setRClaimed] = useState("");
  const [rUTR, setRUTR] = useState("");
  const [rCycle, setRCycle] = useState("1");
  const [rDate, setRDate] = useState("");
  const [rFilter, setRFilter] = useState("all");
  const [rFM, setRFM] = useState("__all__");
  const [rReason, setRReason] = useState({});
  const [rNotes, setRNotes] = useState({});
  const [rResolved, setRResolved] = useState({});

  // Upload state
  const [uploadData, setUploadData] = useState([]);
  const [uploadFile, setUploadFile] = useState("");
  const [uploadError, setUploadError] = useState("");
  const [uploadMapping, setUploadMapping] = useState({});
  const [uploadHeaders, setUploadHeaders] = useState([]);
  const [uploadSelected, setUploadSelected] = useState({});
  const [uploadImported, setUploadImported] = useState(0);
  const [isDragging, setIsDragging] = useState(false);

  const compsFor = useMemo(
    () =>
      selM
        ? MASTER_DATA.filter((d) => d.merchant === selM)
            .map((d) => d.company)
            .sort()
        : MASTER_DATA.map((d) => d.company).sort(),
    [selM]
  );
  const filtCL = useMemo(
    () =>
      cSearch
        ? compsFor.filter((c) =>
            c.toLowerCase().includes(cSearch.toLowerCase())
          )
        : compsFor,
    [compsFor, cSearch]
  );
  const rCompsFor = useMemo(
    () =>
      rMerchant
        ? MASTER_DATA.filter((d) => d.merchant === rMerchant)
            .map((d) => d.company)
            .sort()
        : MASTER_DATA.map((d) => d.company).sort(),
    [rMerchant]
  );

  const pn = parseFloat(payin) || 0;
  const cb = parseFloat(chargeback) || 0;
  const ch = (pn * CHARGE_RATE) / 100;
  const gs = (ch * gstRate) / 100;
  const td = ch + gs + cb;
  const stl = pn - td;
  const cyc = getCycle();
  const mCnt = useMemo(() => {
    const m = {};
    MASTER_DATA.forEach((d) => {
      m[d.merchant] = (m[d.merchant] || 0) + 1;
    });
    return m;
  }, []);

  const fRows = useMemo(() => {
    let r = rows;
    if (fM !== "__all__") r = r.filter((x) => x.merchant === fM);
    if (fC !== "__all__") r = r.filter((x) => x.company === fC);
    return r;
  }, [rows, fM, fC]);

  const grouped = useMemo(() => {
    if (view === "company") {
      const m = {};
      fRows.forEach((r) => {
        if (!m[r.company]) m[r.company] = [];
        m[r.company].push(r);
      });
      return Object.entries(m)
        .map(([n, rw]) => ({ name: n, ...sumR(rw) }))
        .sort((a, b) => b.payin - a.payin);
    }
    if (view === "merchant") {
      const m = {};
      fRows.forEach((r) => {
        if (!m[r.merchant]) m[r.merchant] = [];
        m[r.merchant].push(r);
      });
      return Object.entries(m)
        .map(([n, rw]) => ({ name: n, ...sumR(rw) }))
        .sort((a, b) => b.payin - a.payin);
    }
    return null;
  }, [view, fRows]);

  const allT = sumR(rows);

  const addRow = () => {
    if (pn <= 0 || !selM || !selC) return;
    const effectiveCycle = entryCycle != null ? entryCycle : cyc.cycle;
    const effectiveDate = entryDate || new Date().toISOString().slice(0, 10);
    setRows((p) => [
      ...p,
      {
        id: Date.now(),
        merchant: selM,
        company: selC,
        payin: pn,
        charge: ch,
        gst: gs,
        gstRate,
        chargeback: cb,
        deduction: td,
        settlement: stl,
        cycle: effectiveCycle,
        date: effectiveDate,
        time: new Date().toLocaleTimeString("en-IN", {
          hour: "2-digit",
          minute: "2-digit",
          hour12: true,
        }),
      },
    ]);
    setPay("");
    setChargeback("");
  };
  const rmRow = (id) => setRows((p) => p.filter((r) => r.id !== id));
  const clearAll = () => {
    setRows([]);
    setFC("__all__");
    setFM("__all__");
  };
  const canAdd = pn > 0 && selM && selC;

  // ---- Excel cell-style helpers (used by both download functions) ----
  const STYLE_HEADER = {
    font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFFFF" } },
    fill: { patternType: "solid", fgColor: { rgb: "FF305496" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: {
      top:    { style: "thin", color: { rgb: "FF1F3864" } },
      bottom: { style: "thin", color: { rgb: "FF1F3864" } },
      left:   { style: "thin", color: { rgb: "FF1F3864" } },
      right:  { style: "thin", color: { rgb: "FF1F3864" } },
    },
  };
  const STYLE_TOTAL = {
    font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FF1F3864" } },
    fill: { patternType: "solid", fgColor: { rgb: "FFD9E1F2" } },
    alignment: { vertical: "center" },
    border: {
      top:    { style: "medium", color: { rgb: "FF305496" } },
      bottom: { style: "medium", color: { rgb: "FF305496" } },
    },
  };
  const STYLE_BODY_NUM = (numFmt) => ({
    alignment: { horizontal: "right", vertical: "center" },
    ...(numFmt ? { numFmt } : {}),
  });
  // Apply STYLE_HEADER to the row at `rowIdx` (0-based, default 0) for `nCols` columns.
  const styleHeaderRow = (ws, nCols, rowIdx = 0) => {
    for (let c = 0; c < nCols; c++) {
      const ref = XLSX.utils.encode_cell({ r: rowIdx, c });
      if (!ws[ref]) ws[ref] = { t: "s", v: "" };
      ws[ref].s = STYLE_HEADER;
    }
    // A taller header row reads better with the wrap/center.
    if (!ws["!rows"]) ws["!rows"] = [];
    ws["!rows"][rowIdx] = { hpx: 32 };
  };
  // Apply STYLE_TOTAL to a specific row index (0-based) for `nCols` columns.
  const styleTotalRow = (ws, rowIdx, nCols) => {
    for (let c = 0; c < nCols; c++) {
      const ref = XLSX.utils.encode_cell({ r: rowIdx, c });
      if (!ws[ref]) ws[ref] = { t: "s", v: "" };
      ws[ref].s = { ...(ws[ref].s || {}), ...STYLE_TOTAL };
    }
  };

  // Build a worksheet that has an optional metadata banner above the data table.
  // Returns the worksheet AND the row indices needed for styling/numfmt.
  const buildSheetWithMeta = (data, metaInfo) => {
    const showMeta = metaInfo && (metaInfo.date || metaInfo.time);
    if (!showMeta) {
      const ws = XLSX.utils.json_to_sheet(data);
      return { ws, headerRowIdx: 0, dataStartRow: 2, totalRowIdx: data.length };
    }
    // Build banner rows
    const bannerRows = [];
    bannerRows.push(["Settlement Reconciliation Report"]);
    const periodLine = [];
    if (metaInfo.date) periodLine.push("Date:", metaInfo.date);
    if (metaInfo.time) {
      if (periodLine.length) periodLine.push("", "Time:", metaInfo.time);
      else periodLine.push("Time:", metaInfo.time);
    }
    bannerRows.push(periodLine);
    bannerRows.push([]); // blank separator
    const headerRowIdx = bannerRows.length; // 0-based, where the column headers will land
    const ws = XLSX.utils.aoa_to_sheet(bannerRows);
    XLSX.utils.sheet_add_json(ws, data, { origin: `A${headerRowIdx + 1}` });
    // Style the banner title cell (bold + larger font, dark blue)
    if (ws["A1"]) {
      ws["A1"].s = {
        font: { name: "Calibri", sz: 13, bold: true, color: { rgb: "FF1F3864" } },
        alignment: { horizontal: "left", vertical: "center" },
      };
    }
    if (!ws["!rows"]) ws["!rows"] = [];
    ws["!rows"][0] = { hpx: 22 };
    return {
      ws,
      headerRowIdx,
      dataStartRow: headerRowIdx + 2, // 1-based row of the first data row
      totalRowIdx: headerRowIdx + data.length, // 0-based index of the TOTAL row
    };
  };

  const downloadSheet = (rowsToExport, filename) => {
    if (!rowsToExport || rowsToExport.length === 0) return;
    const fmtDate = (iso) => {
      if (!iso) return "";
      const [y, m, d] = iso.split("-");
      return `${d}-${m}-${y}`;
    };
    const fmtFromTs = (ts) => {
      const d = new Date(ts);
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    };
    const getRowDate = (r) => r.date || fmtFromTs(typeof r.id === "number" ? r.id : Date.now());

    let filtered = rowsToExport;
    if (dlFrom || dlTo) {
      filtered = filtered.filter((r) => {
        const d = getRowDate(r);
        if (dlFrom && d < dlFrom) return false;
        if (dlTo && d > dlTo) return false;
        return true;
      });
    }
    if (filtered.length === 0) {
      alert("No entries match the selected date range.");
      return;
    }

    const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
    const data = filtered.map((r) => {
      const iso = getRowDate(r);
      const cycleMeta = CYCLE_META[r.cycle] || CYCLE_META[1];
      const merchantSettle = r.payin - r.charge - r.gst;
      return {
        "Company Name": r.company,
        "From Date": fmtDate(iso),
        "To Date": fmtDate(iso),
        "Payin Start Time": cycleMeta.start,
        "Payin End Time": cycleMeta.end,
        Payin: r2(r.payin),
        "Merchant Fee": r2(r.charge),
        GST: r2(r.gst),
        "Merchant Settle Amount": r2(merchantSettle),
        "Chargeback Received": r2(r.chargeback || 0),
        "Net Merchant Settlement": r2(r.settlement),
      };
    });
    // Use filtered rows for totals
    rowsToExport = filtered;

    const totals = rowsToExport.reduce(
      (s, r) => ({
        payin: s.payin + r.payin,
        charge: s.charge + r.charge,
        gst: s.gst + r.gst,
        merchantSettle: s.merchantSettle + (r.payin - r.charge - r.gst),
        chargeback: s.chargeback + (r.chargeback || 0),
        settlement: s.settlement + r.settlement,
      }),
      { payin: 0, charge: 0, gst: 0, merchantSettle: 0, chargeback: 0, settlement: 0 }
    );
    data.push({
      "Company Name": "TOTAL",
      "From Date": "",
      "To Date": "",
      "Payin Start Time": "",
      "Payin End Time": "",
      Payin: r2(totals.payin),
      "Merchant Fee": r2(totals.charge),
      GST: r2(totals.gst),
      "Merchant Settle Amount": r2(totals.merchantSettle),
      "Chargeback Received": r2(totals.chargeback),
      "Net Merchant Settlement": r2(totals.settlement),
    });

    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [
      { wch: 45 },
      { wch: 12 },
      { wch: 12 },
      { wch: 16 },
      { wch: 16 },
      { wch: 14 },
      { wch: 14 },
      { wch: 12 },
      { wch: 22 },
      { wch: 20 },
      { wch: 24 },
    ];
    // Apply 2-decimal money format to columns F..K (Payin through Net Merchant Settlement)
    const moneyCols = ["F", "G", "H", "I", "J", "K"];
    const numFmt = "#,##0.00";
    for (let row = 2; row <= data.length + 1; row++) {
      for (const col of moneyCols) {
        const ref = col + row;
        const cell = ws[ref];
        if (cell && typeof cell.v === "number") {
          cell.z = numFmt;
          cell.t = "n";
        }
      }
    }
    // Style header + TOTAL row on the settlement report sheet
    const settColCount = Object.keys(data[0] || {}).length;
    styleHeaderRow(ws, settColCount);
    styleTotalRow(ws, data.length, settColCount);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Settlements");
    const ts = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `${filename || "settlements"}_${ts}.xlsx`);
  };

  // Upload logic
  const EXPECTED_FIELDS = [
    { key: "date", label: "Date", patterns: ["date", "txn date", "transaction date", "settlement date"] },
    { key: "txnId", label: "Transaction ID", patterns: ["txn id", "transaction id", "txnid", "trans id", "reference", "ref", "order id", "orderid"] },
    { key: "company", label: "Company", patterns: ["company", "company name", "merchant name", "business name", "entity"] },
    { key: "merchant", label: "Merchant", patterns: ["merchant", "agent", "partner", "assigned to"] },
    { key: "payin", label: "Payin", patterns: ["payin", "pay in", "amount", "txn amount", "transaction amount", "gross amount", "total amount"] },
    { key: "charge", label: "Charge", patterns: ["charge", "charges", "fee", "fees", "commission", "mdr"] },
    { key: "gst", label: "GST", patterns: ["gst", "tax", "gst amount", "service tax"] },
    { key: "deduction", label: "Deduction", patterns: ["deduction", "total deduction", "total charges", "tdr"] },
    { key: "settlement", label: "Settlement", patterns: ["settlement", "settle", "net amount", "net settle", "payout", "credit amount"] },
    { key: "utr", label: "UTR", patterns: ["utr", "utr number", "utr no", "bank ref", "bank reference"] },
    { key: "cycle", label: "Cycle", patterns: ["cycle", "batch", "settlement cycle"] },
  ];

  const fuzzyMatch = (header) => {
    const h = header.toLowerCase().trim();
    for (const field of EXPECTED_FIELDS) {
      for (const pattern of field.patterns) {
        if (h === pattern || h.includes(pattern) || pattern.includes(h)) {
          return field.key;
        }
      }
    }
    return null;
  };

  const autoMapHeaders = (headers) => {
    const mapping = {};
    headers.forEach((h) => {
      const match = fuzzyMatch(h);
      if (match && !Object.values(mapping).includes(match)) {
        mapping[h] = match;
      }
    });
    return mapping;
  };

  const parseFile = useCallback((file) => {
    setUploadError("");
    setUploadData([]);
    setUploadMapping({});
    setUploadHeaders([]);
    setUploadSelected({});
    setUploadImported(0);

    const ext = file.name.split(".").pop().toLowerCase();
    if (!["csv", "xlsx", "xls"].includes(ext)) {
      setUploadError("Unsupported file type. Please upload a CSV or Excel file.");
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        if (json.length === 0) {
          setUploadError("The file appears to be empty.");
          return;
        }

        const headers = Object.keys(json[0]);
        const mapping = autoMapHeaders(headers);
        setUploadHeaders(headers);
        setUploadMapping(mapping);
        setUploadFile(file.name);

        const parsed = json.map((row, idx) => {
          const get = (key) => {
            const col = Object.keys(mapping).find((k) => mapping[k] === key);
            return col ? row[col] : "";
          };
          const payinVal = parseFloat(get("payin")) || 0;
          const chargeVal = parseFloat(get("charge")) || (payinVal * CHARGE_RATE) / 100;
          const gstVal = parseFloat(get("gst")) || (chargeVal * GST_RATE) / 100;
          const deductionVal = parseFloat(get("deduction")) || chargeVal + gstVal;
          const settlementVal = parseFloat(get("settlement")) || payinVal - deductionVal;
          const companyVal = String(get("company") || "").trim();
          const merchantVal = String(get("merchant") || "").trim();

          const masterMatch = MASTER_DATA.find(
            (d) => d.company.toLowerCase() === companyVal.toLowerCase()
          );
          const resolvedMerchant = merchantVal || (masterMatch ? masterMatch.merchant : "");

          const valid = payinVal > 0 && companyVal;

          return {
            _idx: idx,
            _valid: valid,
            _knownCompany: !!masterMatch,
            date: String(get("date") || ""),
            txnId: String(get("txnId") || `UP-${Date.now()}-${idx}`),
            company: companyVal,
            merchant: resolvedMerchant,
            payin: payinVal,
            charge: chargeVal,
            gst: gstVal,
            deduction: deductionVal,
            settlement: settlementVal,
            utr: String(get("utr") || ""),
            cycle: parseInt(get("cycle")) || cyc.cycle,
          };
        });

        setUploadData(parsed);
        const sel = {};
        parsed.forEach((r) => { if (r._valid) sel[r._idx] = true; });
        setUploadSelected(sel);
      } catch (err) {
        setUploadError("Failed to parse file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }, [cyc.cycle]);

  const handleFileDrop = useCallback((e) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer?.files?.[0] || e.target?.files?.[0];
    if (file) parseFile(file);
  }, [parseFile]);

  const handleFileInput = useCallback((e) => {
    const file = e.target.files?.[0];
    if (file) parseFile(file);
  }, [parseFile]);

  const importSelected = () => {
    const toImport = uploadData.filter((r) => uploadSelected[r._idx] && r._valid);
    const newRows = toImport.map((r) => ({
      id: Date.now() + Math.random(),
      merchant: r.merchant || "Unassigned",
      company: r.company,
      payin: r.payin,
      charge: r.charge,
      gst: r.gst,
      deduction: r.deduction,
      settlement: r.settlement,
      cycle: r.cycle,
      time: new Date().toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit", hour12: true }),
    }));
    setRows((p) => [...p, ...newRows]);
    setUploadImported(newRows.length);
    setUploadData([]);
    setUploadFile("");
    setUploadHeaders([]);
    setUploadMapping({});
    setUploadSelected({});
  };

  const clearUpload = () => {
    setUploadData([]);
    setUploadFile("");
    setUploadError("");
    setUploadHeaders([]);
    setUploadMapping({});
    setUploadSelected({});
    setUploadImported(0);
  };

  const uploadSummary = useMemo(() => {
    const selected = uploadData.filter((r) => uploadSelected[r._idx] && r._valid);
    return {
      total: uploadData.length,
      valid: uploadData.filter((r) => r._valid).length,
      invalid: uploadData.filter((r) => !r._valid).length,
      selectedCount: selected.length,
      totalPayin: selected.reduce((s, r) => s + r.payin, 0),
      totalSettlement: selected.reduce((s, r) => s + r.settlement, 0),
    };
  }, [uploadData, uploadSelected]);

  // Bank statement upload (settlement summary per company)
  const BANK_FIELDS = {
    mid: ["merchantsid", "merchant id", "mid", "merchant_id"],
    name: ["name", "company", "company name", "merchant name", "row labels"],
    count: ["count of amount", "count", "txn count", "transaction count", "no of txn"],
    amount: ["sum of amount2", "sum of amount", "payin", "amount", "total amount", "gross", "sum of payin"],
    fee: ["sum of merchant fee", "merchant fee", "fee", "fees", "charge", "charges", "mdr"],
    gst: ["sum of gst", "gst", "tax"],
    settle: ["sum of merchant settle amount", "merchant settle amount", "settle amount", "sum of settle"],
    chargeback: ["chargeback received", "chargeback", "cb", "chargebacks", "refund", "refunds"],
    netSettle: ["net merchant settlement", "net merchant settle", "net merchant settlement amount", "net settle", "net settlement", "net amount"],
  };

  // Normalize a company name: strip "private limited" / "pvt ltd" variants,
  // common typos ("privated", "limte..."), punctuation, and collapse spaces.
  const normCompany = (s) => String(s || "").toLowerCase()
    // Strip "private" + common typos (priavte, privte, privae, prvate), "pvt" variants,
    // "ltd" variants, "limited", "limted", "opc", "llp".
    .replace(/\b(priv\w*|priav\w*|prvt\w*|prvate\w*|pvt\w*|ltd\w*|limited|limte\w*|opc|llp)\.?\b/g, "")
    .replace(/[.,&()]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  // Levenshtein distance — used as a final fuzzy-match fallback for company names
  // that differ by a typo or singular/plural ("Enterprise" vs "Enterprises", "Axiomix" vs "Axiomixx").
  const lev = (a, b) => {
    if (a === b) return 0;
    if (!a.length) return b.length;
    if (!b.length) return a.length;
    const dp = new Array(b.length + 1);
    for (let j = 0; j <= b.length; j++) dp[j] = j;
    for (let i = 1; i <= a.length; i++) {
      let prev = dp[0]; dp[0] = i;
      for (let j = 1; j <= b.length; j++) {
        const tmp = dp[j];
        dp[j] = a[i - 1] === b[j - 1]
          ? prev
          : 1 + Math.min(prev, dp[j], dp[j - 1]);
        prev = tmp;
      }
    }
    return dp[b.length];
  };

  // Find the best master-data company match for `name`, with progressive fallbacks:
  // 1) exact match (case-insensitive)
  // 2) normalized match (strips "private limited"/punctuation/typos)
  // 3) Levenshtein-based fuzzy match (distance ≤ 2 for short names, ≤ ceil(len*0.18) otherwise)
  const findMasterByCompany = (name) => {
    if (!name) return null;
    const lower = name.toLowerCase();
    let m = MASTER_DATA.find((d) => d.company.toLowerCase() === lower);
    if (m) return m;
    const nn = normCompany(name);
    if (!nn) return null;
    m = MASTER_DATA.find((d) => normCompany(d.company) === nn);
    if (m) return m;
    // Fuzzy: closest by Levenshtein within tolerance
    const tolerance = Math.max(2, Math.ceil(nn.length * 0.18));
    let best = null;
    let bestDist = tolerance + 1;
    for (const d of MASTER_DATA) {
      const dn = normCompany(d.company);
      if (!dn) continue;
      // Skip very different lengths early
      if (Math.abs(dn.length - nn.length) > tolerance) continue;
      const dist = lev(nn, dn);
      if (dist < bestDist) { bestDist = dist; best = d; }
    }
    return best;
  };

  // Parse "57,88,735" / "₹ 1,234.50" / 100 → 5788735 / 1234.5 / 100.
  // Strips commas, spaces, currency symbols, and any non-numeric chars before parseFloat.
  const parseAmount = (v) => {
    if (typeof v === "number") return v;
    if (v == null) return 0;
    const cleaned = String(v).replace(/[,\s₹$]/g, "").replace(/[^0-9.\-]/g, "");
    const n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
  };

  // Pick the first sheet that has parseable data. Excel exports often have an empty
  // "Sheet1" before the actual pivot/data sheet ("Company Amounts", etc.).
  const pickDataSheet = (wb) => {
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      const range = ws && ws["!ref"];
      if (!range) continue;
      const json = XLSX.utils.sheet_to_json(ws, { defval: "", header: 1 });
      // Need at least 2 rows (header + 1 data) with some content
      const nonEmpty = json.filter((r) => r.some((c) => c !== "" && c != null));
      if (nonEmpty.length >= 2) return ws;
    }
    return wb.Sheets[wb.SheetNames[0]];
  };

  const matchBankCol = (header) => {
    const h = String(header || "").toLowerCase().trim();
    if (!h) return null;
    let best = null;
    let bestScore = 0;
    for (const [key, patterns] of Object.entries(BANK_FIELDS)) {
      for (const p of patterns) {
        let score = 0;
        if (h === p) score = 1000 + p.length;          // exact match — top priority
        else if (h.includes(p)) score = p.length;       // header contains the pattern
        else if (p.includes(h)) score = h.length * 0.5; // pattern contains the header
        if (score > bestScore) { bestScore = score; best = key; }
      }
    }
    return best;
  };

  // Parse a single bank file and APPEND its rows to bankRows (so multiple files merge).
  // If a file with the same name is already loaded, its rows are replaced.
  const parseBankFile = useCallback((file) => {
    setBankError("");
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["csv", "xlsx", "xls"].includes(ext)) {
      setBankError("Unsupported file type. Please upload a CSV or Excel file.");
      return;
    }
    const reader = new FileReader();
    // Detect the cycle (1/2/3) from the filename — see RECON_CYCLE_PRESETS.
    const detectedCycle = detectBankFileCycle(file.name);
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const ws = pickDataSheet(wb);
        const json = XLSX.utils.sheet_to_json(ws, { defval: "", header: 1 });
        // Find header row: first row that contains a column matching "name" or "mid"
        let headerIdx = -1;
        let mapping = {};
        for (let i = 0; i < Math.min(json.length, 10); i++) {
          const candidate = {};
          (json[i] || []).forEach((h, idx) => {
            const k = matchBankCol(h);
            if (k && !candidate[k]) candidate[k] = idx;
          });
          if (candidate.name != null && (candidate.amount != null || candidate.settle != null || candidate.netSettle != null)) {
            headerIdx = i;
            mapping = candidate;
            break;
          }
        }
        if (headerIdx === -1) {
          setBankError("Could not find a header row with company Name and Settlement/Amount columns.");
          return;
        }
        const out = [];
        for (let r = headerIdx + 1; r < json.length; r++) {
          const row = json[r] || [];
          // Stop at Grand Total — bank files often stack a second pivot below it
          // which would otherwise be re-parsed and double-count every merchant.
          // The label can appear in the merchantsId column (col 0) or the Name column.
          const rowText = row.map((c) => String(c || "").trim()).join(" | ");
          if (/(^|\|\s*)grand\s*total(\s*\||$)/i.test(rowText)) break;
          const name = String(row[mapping.name] || "").trim();
          if (!name) continue;
          if (/^total$/i.test(name)) break;
          const mid = mapping.mid != null ? String(row[mapping.mid] || "").trim() : "";
          const count = parseAmount(row[mapping.count]) || 0;
          // Read whichever columns the file actually has. Missing ones are derived below
          // using the standard rates (0.35% fee + 18% GST on fee).
          let amount = mapping.amount != null ? (parseAmount(row[mapping.amount]) || 0) : 0;
          let fee    = mapping.fee    != null ? (parseAmount(row[mapping.fee])    || 0) : 0;
          let gst    = mapping.gst    != null ? (parseAmount(row[mapping.gst])    || 0) : 0;
          let settle = mapping.settle != null ? (parseAmount(row[mapping.settle]) || 0) : 0;
          const chargeback = mapping.chargeback != null ? (parseAmount(row[mapping.chargeback]) || 0) : 0;
          // Skip empty rows up-front
          if (amount <= 0 && settle <= 0) continue;
          // Back-fill missing values from whatever IS available. Bank files vary:
          //   - "Full" format: has Sum of Amount2 + Fee + GST + Settle → no derivation needed
          //   - "Settle-only" format: only has Sum of Merchant Settle Amount → back-calculate
          //     gross (Payin) from settle, then fee+GST from gross.
          // factor = settle/gross when fees of 0.35% + 18% on fee are applied.
          const factor = 1 - (CHARGE_RATE / 100) - ((CHARGE_RATE / 100) * (GST_RATE / 100));
          if (amount <= 0 && settle > 0) amount = settle / factor; // back-calc gross from net
          if (fee <= 0 && amount > 0)    fee = (amount * CHARGE_RATE) / 100;
          if (gst <= 0 && fee  > 0)      gst = (fee * GST_RATE) / 100;
          if (settle <= 0 && amount > 0) settle = amount - fee - gst;
          const netSettle = mapping.netSettle != null
            ? (parseAmount(row[mapping.netSettle]) || (settle - chargeback))
            : (settle - chargeback);
          // Lookup parent merchant: prefer MID, then exact / normalized / fuzzy company match
          let master = mid ? MASTER_DATA.find((d) => d.mid && d.mid.toLowerCase() === mid.toLowerCase()) : null;
          if (!master) master = findMasterByCompany(name);
          out.push({
            mid,
            name,
            merchant: master ? master.merchant : "Unmatched",
            matched: !!master,
            count,
            amount,
            fee,
            gst,
            settle,
            chargeback,
            netSettle,
            _sourceFile: file.name,
            _cycle: detectedCycle,
          });
        }
        if (out.length === 0) {
          setBankError("No data rows found in the file.");
          return;
        }
        // Replace rows for any previously-loaded file with the same name, then append new rows.
        setBankRows((prev) => [...prev.filter((r) => r._sourceFile !== file.name), ...out]);
        setBankFiles((prev) => [
          ...prev.filter((f) => f.name !== file.name),
          { name: file.name, rowCount: out.length, cycle: detectedCycle },
        ]);
      } catch (err) {
        setBankError("Failed to parse file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleBankDrop = useCallback((e) => {
    e.preventDefault();
    setBankDragging(false);
    const files = Array.from(e.dataTransfer?.files || e.target?.files || []);
    files.forEach((f) => parseBankFile(f));
  }, [parseBankFile]);

  const handleBankInput = useCallback((e) => {
    const files = Array.from(e.target.files || []);
    files.forEach((f) => parseBankFile(f));
    // Reset the input so the same file can be re-selected later
    e.target.value = "";
  }, [parseBankFile]);

  const clearBank = () => {
    setBankRows([]);
    setBankFiles([]);
    setBankError("");
  };

  const removeBankFile = (fileName) => {
    setBankRows((prev) => prev.filter((r) => r._sourceFile !== fileName));
    setBankFiles((prev) => prev.filter((f) => f.name !== fileName));
  };

  // ---- Source / internal payin file parsing (Name + Sum of Amount per company) ----
  const parseSourceFile = useCallback((file) => {
    setSourceError("");
    setSourceRows([]);
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["csv", "xlsx", "xls"].includes(ext)) {
      setSourceError("Unsupported file type. Please upload a CSV or Excel file.");
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const ws = pickDataSheet(wb);
        const json = XLSX.utils.sheet_to_json(ws, { defval: "", header: 1 });
        // Find header row: name + amount
        let headerIdx = -1;
        let mapping = {};
        for (let i = 0; i < Math.min(json.length, 10); i++) {
          const candidate = {};
          (json[i] || []).forEach((h, idx) => {
            const k = matchBankCol(h);
            if (k && candidate[k] == null) candidate[k] = idx;
          });
          if (candidate.name != null && candidate.amount != null) {
            headerIdx = i;
            mapping = candidate;
            break;
          }
        }
        if (headerIdx === -1) {
          setSourceError("Could not find a header row with company Name and Amount columns.");
          return;
        }
        const out = [];
        for (let r = headerIdx + 1; r < json.length; r++) {
          const row = json[r] || [];
          const rowText = row.map((c) => String(c || "").trim()).join(" | ");
          if (/(^|\|\s*)grand\s*total(\s*\||$)/i.test(rowText)) break;
          const name = String(row[mapping.name] || "").trim();
          if (!name) continue;
          if (/^total$/i.test(name)) break;
          const amount = parseAmount(row[mapping.amount]) || 0;
          const count = mapping.count != null ? (parseAmount(row[mapping.count]) || 0) : 0;
          if (amount <= 0) continue;
          // Match company → parent merchant (exact / normalized / fuzzy)
          const master = findMasterByCompany(name);
          out.push({
            name,
            merchant: master ? master.merchant : "Unmatched",
            matched: !!master,
            count,
            amount,
          });
        }
        if (out.length === 0) {
          setSourceError("No data rows found in the file.");
          return;
        }
        setSourceRows(out);
        setSourceFile(file.name);
      } catch (err) {
        setSourceError("Failed to parse file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleSourceDrop = useCallback((e) => {
    e.preventDefault();
    setSourceDragging(false);
    const file = e.dataTransfer?.files?.[0] || e.target?.files?.[0];
    if (file) parseSourceFile(file);
  }, [parseSourceFile]);

  const handleSourceInput = useCallback((e) => {
    const f = e.target.files?.[0];
    if (f) parseSourceFile(f);
  }, [parseSourceFile]);

  const clearSource = () => {
    setSourceRows([]);
    setSourceFile("");
    setSourceError("");
  };

  // ---- Saved Recons (persisted to localStorage) ----
  // Clear all the in-page recon state so the form is ready for a fresh day.
  const clearAllReconInputs = () => {
    setSourceRows([]); setSourceFile(""); setSourceError("");
    setBankRows([]); setBankFiles([]); setBankError("");
    setReconStartTime(""); setReconEndTime("");
  };

  // Snapshot the current recon state (merchant + date/time + uploaded data) into
  // localStorage so the user can reload it later. Returns the saved entry's id.
  const saveCurrentRecon = () => {
    if (sourceRows.length === 0 && bankRows.length === 0) {
      alert("Nothing to save — upload at least one file first.");
      return null;
    }
    const merchantLabel = reconMerchantFilter === "__all__" ? "All merchants" : reconMerchantFilter;
    const dateLabel = reconDate || "(no date)";
    const id = `recon_${Date.now()}`;
    const entry = {
      id,
      merchant: reconMerchantFilter,
      merchantLabel,
      date: reconDate,
      dateLabel,
      startTime: reconStartTime,
      endTime: reconEndTime,
      sourceFile,
      sourceRows: [...sourceRows],
      bankFiles: [...bankFiles],
      bankRows: [...bankRows],
      savedAt: Date.now(),
    };
    // If a recon with the same (merchant, date) already exists, replace it.
    setSavedRecons((prev) => {
      const dupIdx = prev.findIndex((r) => r.merchant === entry.merchant && r.date === entry.date);
      if (dupIdx >= 0) {
        const next = [...prev];
        next[dupIdx] = entry;
        return next;
      }
      return [entry, ...prev];
    });
    return id;
  };

  const loadSavedRecon = (entry) => {
    if (!entry) return;
    setSourceFile(entry.sourceFile || "");
    setSourceRows(entry.sourceRows || []);
    setSourceError("");
    setBankFiles(entry.bankFiles || []);
    setBankRows(entry.bankRows || []);
    setBankError("");
    setReconDate(entry.date || "");
    setReconStartTime(entry.startTime || "");
    setReconEndTime(entry.endTime || "");
    setReconMerchantFilter(entry.merchant || "__all__");
  };

  const deleteSavedRecon = (id) => {
    if (!confirm("Delete this saved recon?")) return;
    setSavedRecons((prev) => prev.filter((r) => r.id !== id));
  };

  // Visible list of saved recons, optionally narrowed by the merchant filter on the panel.
  const visibleSavedRecons = useMemo(() => {
    let list = [...savedRecons].sort((a, b) => (b.savedAt || 0) - (a.savedAt || 0));
    if (savedReconsFilter !== "__all__") list = list.filter((r) => r.merchant === savedReconsFilter);
    return list;
  }, [savedRecons, savedReconsFilter]);

  // Aggregate source rows by parent merchant
  const sourceByMerchant = useMemo(() => {
    if (sourceRows.length === 0) return new Map();
    const groups = new Map();
    sourceRows.forEach((r) => {
      const key = r.merchant || "Unmatched";
      if (!groups.has(key)) groups.set(key, { merchant: key, companies: 0, count: 0, amount: 0, unmatched: 0 });
      const g = groups.get(key);
      g.companies += 1;
      g.count += r.count;
      g.amount += r.amount;
      if (!r.matched) g.unmatched += 1;
    });
    return groups;
  }, [sourceRows]);

  // Aggregate bank rows by parent merchant + recompute using our rates
  const bankByMerchant = useMemo(() => {
    if (bankRows.length === 0) return [];
    const groups = {};
    bankRows.forEach((r) => {
      const key = r.merchant || "Unmatched";
      if (!groups[key]) groups[key] = { merchant: key, companies: 0, count: 0, amount: 0, fee: 0, gst: 0, settle: 0, chargeback: 0, netSettle: 0, unmatched: 0, rows: [] };
      const g = groups[key];
      g.companies += 1;
      g.count += r.count;
      g.amount += r.amount;
      g.fee += r.fee;
      g.gst += r.gst;
      g.settle += r.settle;
      g.chargeback += r.chargeback || 0;
      g.netSettle += r.netSettle || (r.settle - (r.chargeback || 0));
      if (!r.matched) g.unmatched += 1;
      g.rows.push(r);
    });
    return Object.values(groups).map((g) => {
      const ourFee = (g.amount * CHARGE_RATE) / 100;
      const ourGst = (ourFee * GST_RATE) / 100;
      const ourSettle = g.amount - ourFee - ourGst;
      const ourNet = ourSettle - g.chargeback;
      const diff = g.netSettle - ourNet;
      return { ...g, ourFee, ourGst, ourSettle, ourNet, diff };
    }).sort((a, b) => b.amount - a.amount);
  }, [bankRows]);

  const bankTotals = useMemo(() => {
    return bankByMerchant.reduce((s, g) => ({
      companies: s.companies + g.companies,
      count: s.count + g.count,
      amount: s.amount + g.amount,
      fee: s.fee + g.fee,
      gst: s.gst + g.gst,
      settle: s.settle + g.settle,
      chargeback: s.chargeback + g.chargeback,
      netSettle: s.netSettle + g.netSettle,
      ourSettle: s.ourSettle + g.ourSettle,
      ourNet: s.ourNet + g.ourNet,
      diff: s.diff + g.diff,
      unmatched: s.unmatched + g.unmatched,
    }), { companies: 0, count: 0, amount: 0, fee: 0, gst: 0, settle: 0, chargeback: 0, netSettle: 0, ourSettle: 0, ourNet: 0, diff: 0, unmatched: 0 });
  }, [bankByMerchant]);

  // Combined view: bank rows joined with source-payin rows by merchant, with payin difference.
  // Merchants present in only one side still appear in the table (other side shows 0).
  const combinedByMerchant = useMemo(() => {
    const map = new Map();
    bankByMerchant.forEach((g) => {
      map.set(g.merchant, { ...g, sourcePayin: 0, sourceCompanies: 0, sourceCount: 0, sourceUnmatched: 0, hasBank: true });
    });
    sourceByMerchant.forEach((s, key) => {
      const existing = map.get(key);
      if (existing) {
        existing.sourcePayin = s.amount;
        existing.sourceCompanies = s.companies;
        existing.sourceCount = s.count;
        existing.sourceUnmatched = s.unmatched;
      } else {
        // Merchant only in source file
        map.set(key, {
          merchant: key,
          companies: 0, count: 0, amount: 0, fee: 0, gst: 0, settle: 0, chargeback: 0, netSettle: 0,
          ourFee: 0, ourGst: 0, ourSettle: 0, ourNet: 0, diff: 0, unmatched: 0, rows: [],
          sourcePayin: s.amount, sourceCompanies: s.companies, sourceCount: s.count, sourceUnmatched: s.unmatched,
          hasBank: false,
        });
      }
    });
    // Compute payin difference (source − bank)
    map.forEach((g) => { g.payinDiff = (g.sourcePayin || 0) - (g.amount || 0); });
    return [...map.values()].sort((a, b) => Math.max(b.sourcePayin, b.amount) - Math.max(a.sourcePayin, a.amount));
  }, [bankByMerchant, sourceByMerchant]);

  const combinedTotals = useMemo(() => {
    return combinedByMerchant.reduce((s, g) => ({
      sourcePayin: s.sourcePayin + (g.sourcePayin || 0),
      sourceCount: s.sourceCount + (g.sourceCount || 0),
      bankPayin: s.bankPayin + (g.amount || 0),
      payinDiff: s.payinDiff + (g.payinDiff || 0),
    }), { sourcePayin: 0, sourceCount: 0, bankPayin: 0, payinDiff: 0 });
  }, [combinedByMerchant]);

  // Filter view: limit the recon table and the downloaded Excel to a specific merchant.
  const filteredCombinedByMerchant = useMemo(() => {
    if (reconMerchantFilter === "__all__") return combinedByMerchant;
    return combinedByMerchant.filter((g) => g.merchant === reconMerchantFilter);
  }, [combinedByMerchant, reconMerchantFilter]);

  // Totals respect the filter so the TOTAL row reflects only the visible merchant(s).
  const filteredCombinedTotals = useMemo(() => {
    return filteredCombinedByMerchant.reduce((s, g) => ({
      sourcePayin: s.sourcePayin + (g.sourcePayin || 0),
      sourceCount: s.sourceCount + (g.sourceCount || 0),
      bankPayin: s.bankPayin + (g.amount || 0),
      payinDiff: s.payinDiff + (g.payinDiff || 0),
      fee: s.fee + (g.fee || 0),
      gst: s.gst + (g.gst || 0),
      settle: s.settle + (g.settle || 0),
      chargeback: s.chargeback + (g.chargeback || 0),
      netSettle: s.netSettle + (g.netSettle || 0),
      ourSettle: s.ourSettle + (g.ourSettle || 0),
      ourNet: s.ourNet + (g.ourNet || 0),
      diff: s.diff + (g.diff || 0),
    }), { sourcePayin: 0, sourceCount: 0, bankPayin: 0, payinDiff: 0, fee: 0, gst: 0, settle: 0, chargeback: 0, netSettle: 0, ourSettle: 0, ourNet: 0, diff: 0 });
  }, [filteredCombinedByMerchant]);

  const downloadBankRecon = () => {
    if (combinedByMerchant.length === 0) return;
    if (filteredCombinedByMerchant.length === 0) {
      alert(`No data for "${reconMerchantFilter}" — change the merchant filter and try again.`);
      return;
    }
    const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
    const feeLabel = `Our Fee (${CHARGE_RATE}%)`;
    const gstLabel = `Our GST (${GST_RATE}%)`;
    // Respect the in-app merchant filter — if a single merchant is selected, only their
    // row(s) make it into Summary and only their detail sheet is produced.
    const exportMerchants = filteredCombinedByMerchant;
    const exportTotals = filteredCombinedTotals;
    const exportTotalCompanies = exportMerchants.reduce((s, g) => s + (g.companies || 0), 0);
    const exportTotalCount = exportMerchants.reduce((s, g) => s + (g.count || 0), 0);
    const exportTotalUnmatched = exportMerchants.reduce((s, g) => s + (g.unmatched || 0), 0);

    const data = exportMerchants.map((g, i) => ({
      "S.No": i + 1,
      Merchant: g.merchant,
      Companies: g.companies,
      "Txn Count": g.count,
      "Source Payin": r2(g.sourcePayin || 0),
      "Bank Payin": r2(g.amount),
      "Payin Diff (Source − Bank)": r2(g.payinDiff || 0),
      "Bank Fee": r2(g.fee),
      "Bank GST": r2(g.gst),
      "Bank Settlement": r2(g.settle),
      "Chargeback Received": r2(g.chargeback),
      "Net Merchant Settlement": r2(g.netSettle),
      [feeLabel]: r2(g.ourFee),
      [gstLabel]: r2(g.ourGst),
      "Our Settlement": r2(g.ourSettle),
      "Our Net (after CB)": r2(g.ourNet),
      "Difference (Bank Net − Our Net)": r2(g.diff),
      "Unmatched Companies": g.unmatched,
    }));
    data.push({
      "S.No": "",
      Merchant: "TOTAL",
      Companies: exportTotalCompanies,
      "Txn Count": exportTotalCount,
      "Source Payin": r2(exportTotals.sourcePayin),
      "Bank Payin": r2(exportTotals.bankPayin),
      "Payin Diff (Source − Bank)": r2(exportTotals.payinDiff),
      "Bank Fee": r2(exportTotals.fee),
      "Bank GST": r2(exportTotals.gst),
      "Bank Settlement": r2(exportTotals.settle),
      "Chargeback Received": r2(exportTotals.chargeback),
      "Net Merchant Settlement": r2(exportTotals.netSettle),
      [feeLabel]: r2(exportTotals.bankPayin * CHARGE_RATE / 100),
      [gstLabel]: r2((exportTotals.bankPayin * CHARGE_RATE / 100) * GST_RATE / 100),
      "Our Settlement": r2(exportTotals.ourSettle),
      "Our Net (after CB)": r2(exportTotals.ourNet),
      "Difference (Bank Net − Our Net)": r2(exportTotals.diff),
      "Unmatched Companies": exportTotalUnmatched,
    });
    // Build period metadata once; embed it as a banner at the top of every sheet + the filename.
    const periodDate = formatDateLong(reconDate);
    const periodTime = (reconStartTime || reconEndTime) ? `${reconStartTime || "—"} to ${reconEndTime || "—"}` : "";
    const periodMeta = (periodDate || periodTime) ? { date: periodDate, time: periodTime } : null;

    const { ws, headerRowIdx, dataStartRow, totalRowIdx } = buildSheetWithMeta(data, periodMeta);
    ws["!cols"] = [
      { wch: 6 },   // A S.No
      { wch: 14 },  // B Merchant
      { wch: 10 },  // C Companies
      { wch: 10 },  // D Txn Count
      { wch: 16 },  // E Source Payin
      { wch: 16 },  // F Bank Payin
      { wch: 20 },  // G Payin Diff
      { wch: 14 },  // H Bank Fee
      { wch: 12 },  // I Bank GST
      { wch: 16 },  // J Bank Settlement
      { wch: 16 },  // K Chargeback
      { wch: 22 },  // L Net Merchant Settlement
      { wch: 14 },  // M Our Fee
      { wch: 14 },  // N Our GST
      { wch: 16 },  // O Our Settlement
      { wch: 16 },  // P Our Net
      { wch: 24 },  // Q Difference
      { wch: 20 },  // R Unmatched
    ];

    // Apply 2-decimal money format to columns E..Q. Txn Count (D) uses integer format.
    // Data rows start at `dataStartRow` (1-based) and run for data.length rows.
    const moneyCols = ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"];
    const numFmt = "#,##0.00";
    for (let row = dataStartRow; row < dataStartRow + data.length; row++) {
      for (const col of moneyCols) {
        const ref = col + row;
        const cell = ws[ref];
        if (cell && typeof cell.v === "number") {
          cell.z = numFmt;
          cell.t = "n";
        }
      }
      const dRef = "D" + row;
      if (ws[dRef] && typeof ws[dRef].v === "number") {
        ws[dRef].z = "#,##0";
        ws[dRef].t = "n";
      }
    }

    // Style Summary sheet: blue header row + highlighted TOTAL row at the bottom
    const summaryColCount = Object.keys(data[0] || {}).length;
    styleHeaderRow(ws, summaryColCount, headerRowIdx);
    styleTotalRow(ws, totalRowIdx, summaryColCount);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Summary");

    // ----- Cycle Breakdown sheet (per cycle × per company) -----
    // For each cycle that has data, emit one row per company that appeared in that cycle's
    // bank file, plus a "Cycle N Total" subtotal row, and a Grand Total at the end.
    // Columns: Date | Cycle | StartTime | EndTime | Settlement_Date | Company |
    //          Payin | Settlement_Amount | ChargeBack | InternalCharge | Settlement
    // Respects the merchant filter (so when one merchant is selected, only that merchant's
    // companies are listed).
    const cyclesPresent = [1, 2, 3].filter((id) => {
      return bankRows.some((r) => r._cycle === id && (reconMerchantFilter === "__all__" || r.merchant === reconMerchantFilter));
    });
    if (cyclesPresent.length > 0) {
      const dayDate = formatDateLong(reconDate);
      // Per-company day payin (from internal payin file), keyed by normalized company name.
      const sourcePayinByCompany = new Map();
      sourceRows
        .filter((s) => reconMerchantFilter === "__all__" || s.merchant === reconMerchantFilter)
        .forEach((s) => {
          const key = normCompany(s.name);
          sourcePayinByCompany.set(key, (sourcePayinByCompany.get(key) || 0) + (s.amount || 0));
        });
      const filteredSourceTotal = [...sourcePayinByCompany.values()].reduce((s, v) => s + v, 0);

      // Track which companies have already had their Payin printed (so it appears once
      // per company, on the earliest cycle they show up in — mirrors the "merged-cell" look).
      const payinShownFor = new Set();
      const cycleRows = [];

      // Running grand-total accumulators
      let gtSettleAmt = 0, gtChargeback = 0, gtInternal = 0, gtFinal = 0;

      cyclesPresent.forEach((id) => {
        const preset = RECON_CYCLE_PRESETS.find((p) => p.id === id);
        const settlementDate = preset.settleNextDay ? formatDateLong(addDaysToIso(reconDate, 1)) : dayDate;

        // Per-company aggregation within this cycle
        const compMap = new Map();
        bankRows
          .filter((r) => r._cycle === id)
          .filter((r) => reconMerchantFilter === "__all__" || r.merchant === reconMerchantFilter)
          .forEach((r) => {
            const key = normCompany(r.name) || r.name;
            const existing = compMap.get(key);
            if (existing) {
              existing.settle += r.settle || 0;
              existing.chargeback += r.chargeback || 0;
            } else {
              compMap.set(key, { name: r.name, settle: r.settle || 0, chargeback: r.chargeback || 0 });
            }
          });
        if (compMap.size === 0) return;

        const companies = [...compMap.values()].sort((a, b) => b.settle - a.settle);
        let cycSettle = 0, cycCharge = 0, cycInternal = 0, cycFinal = 0, cycPayin = 0;

        companies.forEach((c) => {
          const internal = c.settle * (CHARGE_RATE / 100) * (1 + GST_RATE / 100);
          const finalAmt = c.settle - c.chargeback - internal;
          const compKey = normCompany(c.name);
          // Payin appears once per company (on the first cycle they show up)
          const payinForRow = !payinShownFor.has(compKey) ? (sourcePayinByCompany.get(compKey) || 0) : 0;
          if (!payinShownFor.has(compKey)) payinShownFor.add(compKey);

          cyclePushRow({
            Date: dayDate,
            Cycle: id,
            StartTime: preset.startLabel,
            EndTime: preset.endLabel,
            Settlement_Date: settlementDate,
            Company: c.name,
            Payin: payinForRow ? r2(payinForRow) : "",
            Settlement_Amount: r2(c.settle),
            ChargeBack: r2(c.chargeback),
            InternalCharge: r2(internal),
            Settlement: r2(finalAmt),
            _kind: "company",
          });

          cycSettle += c.settle;
          cycCharge += c.chargeback;
          cycInternal += internal;
          cycFinal += finalAmt;
          cycPayin += payinForRow;
        });

        // Cycle subtotal row
        cyclePushRow({
          Date: "",
          Cycle: "",
          StartTime: "",
          EndTime: "",
          Settlement_Date: "",
          Company: `Cycle ${id} Total`,
          Payin: cycPayin ? r2(cycPayin) : "",
          Settlement_Amount: r2(cycSettle),
          ChargeBack: r2(cycCharge),
          InternalCharge: r2(cycInternal),
          Settlement: r2(cycFinal),
          _kind: "subtotal",
        });

        gtSettleAmt += cycSettle;
        gtChargeback += cycCharge;
        gtInternal += cycInternal;
        gtFinal += cycFinal;
      });

      function cyclePushRow(r) { cycleRows.push(r); }

      // Grand Total row
      cycleRows.push({
        Date: "",
        Cycle: "",
        StartTime: "",
        EndTime: "",
        Settlement_Date: "",
        Company: "GRAND TOTAL",
        Payin: r2(filteredSourceTotal),
        Settlement_Amount: r2(gtSettleAmt),
        ChargeBack: r2(gtChargeback),
        InternalCharge: r2(gtInternal),
        Settlement: r2(gtFinal),
        _kind: "grandtotal",
      });

      // Strip the internal `_kind` field before handing to xlsx, but remember subtotal/grandtotal
      // row indices so we can style them.
      const subtotalRowIdxs = [];
      let grandTotalIdx = -1;
      const cycleData = cycleRows.map((r, i) => {
        if (r._kind === "subtotal") subtotalRowIdxs.push(i);
        if (r._kind === "grandtotal") grandTotalIdx = i;
        const { _kind, ...rest } = r;
        return rest;
      });

      const cycleSheet = buildSheetWithMeta(cycleData, periodMeta);
      const cycleWs = cycleSheet.ws;
      cycleWs["!cols"] = [
        { wch: 13 },  // A Date
        { wch: 7 },   // B Cycle
        { wch: 11 },  // C StartTime
        { wch: 11 },  // D EndTime
        { wch: 15 },  // E Settlement_Date
        { wch: 42 },  // F Company
        { wch: 18 },  // G Payin
        { wch: 22 },  // H Settlement_Amount
        { wch: 14 },  // I ChargeBack
        { wch: 16 },  // J InternalCharge
        { wch: 18 },  // K Settlement
      ];
      // Number-format money columns G..K and integer Cycle in B.
      const cMoney = ["G", "H", "I", "J", "K"];
      for (let row = cycleSheet.dataStartRow; row < cycleSheet.dataStartRow + cycleData.length; row++) {
        for (const col of cMoney) {
          const c = cycleWs[col + row];
          if (c && typeof c.v === "number") { c.z = numFmt; c.t = "n"; }
        }
        const cy = cycleWs["B" + row];
        if (cy && typeof cy.v === "number") { cy.z = "0"; cy.t = "n"; }
      }
      const cycleColCount = Object.keys(cycleData[0]).length;
      styleHeaderRow(cycleWs, cycleColCount, cycleSheet.headerRowIdx);
      // Apply subtotal style (light yellow) to each "Cycle N Total" row and TOTAL style to Grand Total.
      const SUBTOTAL_STYLE = {
        font: { name: "Calibri", sz: 11, bold: true, color: { rgb: "FF7C4A03" } },
        fill: { patternType: "solid", fgColor: { rgb: "FFFEF3C7" } },
        alignment: { vertical: "center" },
      };
      subtotalRowIdxs.forEach((relIdx) => {
        const absRow = cycleSheet.headerRowIdx + 1 + relIdx; // 0-based sheet row
        for (let c = 0; c < cycleColCount; c++) {
          const ref = XLSX.utils.encode_cell({ r: absRow, c });
          if (!cycleWs[ref]) cycleWs[ref] = { t: "s", v: "" };
          cycleWs[ref].s = { ...(cycleWs[ref].s || {}), ...SUBTOTAL_STYLE };
        }
      });
      if (grandTotalIdx >= 0) {
        styleTotalRow(cycleWs, cycleSheet.headerRowIdx + 1 + grandTotalIdx, cycleColCount);
      }
      XLSX.utils.book_append_sheet(wb, cycleWs, "Cycle Breakdown");
    }

    // ----- Per-merchant detail sheets -----
    // normCompany is already defined at the component level for fuzzy company matching.
    const sanitizeSheetName = (name) => {
      let s = String(name || "Unknown").replace(/[\/\\?*\[\]:]/g, " ").trim();
      if (s.length > 31) s = s.slice(0, 31);
      return s || "Unknown";
    };
    const usedNames = new Set(["Summary"]);
    const uniqueName = (base) => {
      let n = sanitizeSheetName(base);
      let i = 2;
      while (usedNames.has(n)) {
        const suffix = ` (${i})`;
        n = sanitizeSheetName(base.slice(0, 31 - suffix.length) + suffix);
        i++;
      }
      usedNames.add(n);
      return n;
    };

    exportMerchants.forEach((g) => {
      // Merge per-company rows from bank + source files under this merchant.
      // The same company may appear in multiple bank settlement files (e.g. 3 cycles
      // per day), so we SUM their numbers rather than overwriting.
      const compMap = new Map();
      bankRows.filter((b) => (b.merchant || "Unmatched") === g.merchant).forEach((b) => {
        const key = (b.mid || normCompany(b.name)) || b.name;
        const existing = compMap.get(key);
        if (existing) {
          existing.bankCount += b.count || 0;
          existing.bankPayin += b.amount || 0;
          existing.bankFee += b.fee || 0;
          existing.bankGst += b.gst || 0;
          existing.bankSettle += b.settle || 0;
          existing.chargeback += b.chargeback || 0;
          existing.netSettle += (b.netSettle != null ? b.netSettle : (b.settle - (b.chargeback || 0)));
          // Keep the first non-empty MID we saw; bank files later in the day may omit it.
          if (!existing.mid && b.mid) existing.mid = b.mid;
        } else {
          compMap.set(key, {
            mid: b.mid || "",
            name: b.name,
            bankCount: b.count || 0,
            bankPayin: b.amount || 0,
            bankFee: b.fee || 0,
            bankGst: b.gst || 0,
            bankSettle: b.settle || 0,
            chargeback: b.chargeback || 0,
            netSettle: (b.netSettle != null ? b.netSettle : (b.settle - (b.chargeback || 0))),
            sourcePayin: 0,
          });
        }
      });
      sourceRows.filter((s) => (s.merchant || "Unmatched") === g.merchant).forEach((s) => {
        const sNorm = normCompany(s.name);
        // Try to match an existing bank row by company name
        let matched = null;
        for (const [, v] of compMap) {
          if (normCompany(v.name) === sNorm) { matched = v; break; }
        }
        if (matched) {
          matched.sourcePayin = s.amount;
          matched.sourceCount = s.count;
        } else {
          compMap.set(s.name, {
            mid: "",
            name: s.name,
            bankCount: 0,
            bankPayin: 0,
            bankFee: 0,
            bankGst: 0,
            bankSettle: 0,
            chargeback: 0,
            netSettle: 0,
            sourcePayin: s.amount,
            sourceCount: s.count,
          });
        }
      });
      if (compMap.size === 0) return;

      const rowsArr = [...compMap.values()].sort((a, b) => (b.sourcePayin || b.bankPayin) - (a.sourcePayin || a.bankPayin));
      const detailData = rowsArr.map((c, i) => ({
        "S.No": i + 1,
        MID: c.mid,
        Company: c.name,
        "Source Payin": r2(c.sourcePayin),
        "Bank Payin": r2(c.bankPayin),
        "Payin Diff": r2((c.sourcePayin || 0) - (c.bankPayin || 0)),
        "Txn Count": c.bankCount || c.sourceCount || 0,
        "Bank Fee": r2(c.bankFee),
        "Bank GST": r2(c.bankGst),
        "Bank Settle": r2(c.bankSettle),
        Chargeback: r2(c.chargeback),
        "Net Settle": r2(c.netSettle),
      }));
      // Total row
      const tot = rowsArr.reduce((s, c) => ({
        sourcePayin: s.sourcePayin + (c.sourcePayin || 0),
        bankPayin: s.bankPayin + (c.bankPayin || 0),
        count: s.count + (c.bankCount || c.sourceCount || 0),
        bankFee: s.bankFee + (c.bankFee || 0),
        bankGst: s.bankGst + (c.bankGst || 0),
        bankSettle: s.bankSettle + (c.bankSettle || 0),
        chargeback: s.chargeback + (c.chargeback || 0),
        netSettle: s.netSettle + (c.netSettle || 0),
      }), { sourcePayin: 0, bankPayin: 0, count: 0, bankFee: 0, bankGst: 0, bankSettle: 0, chargeback: 0, netSettle: 0 });
      detailData.push({
        "S.No": "",
        MID: "",
        Company: "TOTAL",
        "Source Payin": r2(tot.sourcePayin),
        "Bank Payin": r2(tot.bankPayin),
        "Payin Diff": r2(tot.sourcePayin - tot.bankPayin),
        "Txn Count": tot.count,
        "Bank Fee": r2(tot.bankFee),
        "Bank GST": r2(tot.bankGst),
        "Bank Settle": r2(tot.bankSettle),
        Chargeback: r2(tot.chargeback),
        "Net Settle": r2(tot.netSettle),
      });

      // Build with the same Settlement Period banner at the top
      const detail = buildSheetWithMeta(detailData, periodMeta);
      const detailWs = detail.ws;
      detailWs["!cols"] = [
        { wch: 6 }, { wch: 18 }, { wch: 42 }, { wch: 16 }, { wch: 16 }, { wch: 16 },
        { wch: 10 }, { wch: 14 }, { wch: 12 }, { wch: 16 }, { wch: 14 }, { wch: 16 },
      ];
      // Apply 2-decimal format to money columns (D..F, H..L)
      const detailMoney = ["D", "E", "F", "H", "I", "J", "K", "L"];
      for (let row = detail.dataStartRow; row < detail.dataStartRow + detailData.length; row++) {
        for (const col of detailMoney) {
          const cell = detailWs[col + row];
          if (cell && typeof cell.v === "number") { cell.z = numFmt; cell.t = "n"; }
        }
        const cnt = detailWs["G" + row];
        if (cnt && typeof cnt.v === "number") { cnt.z = "#,##0"; cnt.t = "n"; }
      }
      // Style header + TOTAL row on each per-merchant detail sheet
      const detailColCount = Object.keys(detailData[0] || {}).length;
      styleHeaderRow(detailWs, detailColCount, detail.headerRowIdx);
      styleTotalRow(detailWs, detail.totalRowIdx, detailColCount);
      XLSX.utils.book_append_sheet(wb, detailWs, uniqueName(g.merchant));
    });

    // Filename includes the settlement date when set
    const fileDate = reconDate || new Date().toISOString().slice(0, 10);
    const cycleTag = (() => {
      if (!reconStartTime || !reconEndTime) return "";
      const preset = RECON_CYCLE_PRESETS.find(p => p.start === reconStartTime && p.end === reconEndTime);
      return preset ? `_C${preset.id}` : `_${reconStartTime.replace(":", "")}-${reconEndTime.replace(":", "")}`;
    })();
    const merchantTag = reconMerchantFilter !== "__all__"
      ? `_${reconMerchantFilter.replace(/[^a-z0-9-]/gi, "")}`
      : "";
    XLSX.writeFile(wb, `bank_recon_${fileDate}${cycleTag}${merchantTag}.xlsx`);
  };

  const addRecon = () => {
    const rp = parseFloat(rPayin) || 0;
    const rc = parseFloat(rClaimed) || 0;
    if (rp <= 0 || !rMerchant || !rCompany || !rTxnId) return;
    const charge2 = (rp * CHARGE_RATE) / 100;
    const gst2 = (charge2 * GST_RATE) / 100;
    const ourSettle = rp - charge2 - gst2;
    const diff = ourSettle - rc;
    const tolerance = 1;
    const status =
      rc === 0
        ? "Pending"
        : Math.abs(diff) <= tolerance
          ? "Matched"
          : "Mismatch";
    setRecon((p) => [
      ...p,
      {
        id: Date.now(),
        merchant: rMerchant,
        company: rCompany,
        txnId: rTxnId.trim(),
        payin: rp,
        charge: charge2,
        gst: gst2,
        ourSettle,
        claimed: rc,
        diff,
        status,
        cycle: parseInt(rCycle),
        date: rDate,
        utr: rUTR.trim(),
      },
    ]);
    setRTxn("");
    setRPay("");
    setRClaimed("");
    setRUTR("");
    setRDate("");
  };
  const rmRecon = (id) => {
    setRecon((p) => p.filter((r) => r.id !== id));
    const nr = { ...rReason };
    delete nr[id];
    setRReason(nr);
    const nn = { ...rNotes };
    delete nn[id];
    setRNotes(nn);
    const nv = { ...rResolved };
    delete nv[id];
    setRResolved(nv);
  };

  const reconFiltered = useMemo(() => {
    let r = recon;
    if (rFilter === "Resolved") r = recon.filter((x) => rResolved[x.id]);
    else if (rFilter !== "all")
      r = recon.filter((x) => !rResolved[x.id] && x.status === rFilter);
    if (rFM !== "__all__") r = r.filter((x) => x.merchant === rFM);
    return r;
  }, [recon, rFilter, rFM, rResolved]);

  const reconStats = useMemo(() => {
    const matched = recon.filter(
      (r) => r.status === "Matched" && !rResolved[r.id]
    ).length;
    const mismatch = recon.filter(
      (r) => r.status === "Mismatch" && !rResolved[r.id]
    ).length;
    const pending = recon.filter(
      (r) => r.status === "Pending" && !rResolved[r.id]
    ).length;
    const resolved = recon.filter((r) => rResolved[r.id]).length;
    const totalDisc = recon
      .filter((r) => r.status === "Mismatch" && !rResolved[r.id])
      .reduce((s, r) => s + Math.abs(r.diff), 0);
    return {
      matched,
      mismatch,
      pending,
      resolved,
      totalDisc,
      total: recon.length,
    };
  }, [recon, rResolved]);

  const navItems = [
    {
      id: "dashboard",
      label: "Dashboard",
      icon: "M3 12l2-2m0 0l7-7 7 7M5 10v10a1 1 0 001 1h3m10-11l2 2m-2-2v10a1 1 0 01-1 1h-3m-4 0a1 1 0 01-1-1v-4a1 1 0 011-1h2a1 1 0 011 1v4a1 1 0 01-1 1",
    },
    { id: "add", label: "Add Entry", icon: "M12 4v16m8-8H4" },
    {
      id: "reports",
      label: "Reports",
      icon: "M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z",
    },
    {
      id: "recon",
      label: "Reconciliation",
      icon: "M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4",
    },
    {
      id: "merchants",
      label: "Merchants",
      icon: "M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0z",
    },
  ];

  const selS = {
    width: "100%",
    padding: "12px 14px",
    borderRadius: 10,
    border: `1.5px solid ${C.border}`,
    fontSize: 14,
    fontWeight: 500,
    color: C.text,
    background: "#fff",
    cursor: "pointer",
  };
  const inpS = {
    width: "100%",
    padding: "12px 14px",
    borderRadius: 10,
    border: `1.5px solid ${C.border}`,
    fontSize: 14,
    fontWeight: 500,
    color: C.text,
  };
  const labS = {
    fontSize: 11,
    fontWeight: 600,
    color: C.muted,
    display: "block",
    marginBottom: 5,
    textTransform: "uppercase",
    letterSpacing: "0.5px",
  };
  const filtS = {
    padding: "8px 12px",
    borderRadius: 8,
    border: `1px solid ${C.border}`,
    fontSize: 12,
    fontWeight: 500,
    color: C.text,
    cursor: "pointer",
  };

  const statusBadge = (status, resolved) => {
    if (resolved)
      return { bg: "#dbeafe", color: "#1e40af", label: "Resolved" };
    if (status === "Matched")
      return { bg: "#d1fae5", color: "#065f46", label: "Matched" };
    if (status === "Mismatch")
      return { bg: "#fee2e2", color: "#991b1b", label: "Mismatch" };
    return { bg: "#fef3c7", color: "#92400e", label: "Pending" };
  };

  return (
    <div
      style={{
        display: "flex",
        height: "100vh",
        fontFamily:
          "'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif",
        background: C.bg,
        overflow: "hidden",
      }}
    >
      {/* Sidebar */}
      <div
        style={{
          width: sideOpen ? 220 : 64,
          background: C.sidebar,
          display: "flex",
          flexDirection: "column",
          transition: "width .3s",
          flexShrink: 0,
          overflow: "hidden",
        }}
      >
        <div
          style={{
            padding: sideOpen ? "24px 20px 20px" : "24px 14px 20px",
            display: "flex",
            alignItems: "center",
            gap: 10,
            cursor: "pointer",
          }}
          onClick={() => setSO(!sideOpen)}
        >
          <div
            style={{
              width: 34,
              height: 34,
              borderRadius: 10,
              background: "linear-gradient(135deg,#10b981,#059669)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              color: "#fff",
              fontWeight: 800,
              fontSize: 16,
              flexShrink: 0,
            }}
          >
            S
          </div>
          {sideOpen && (
            <span
              style={{
                color: "#fff",
                fontWeight: 700,
                fontSize: 17,
                letterSpacing: "-0.3px",
                whiteSpace: "nowrap",
              }}
            >
              SettleOps
            </span>
          )}
        </div>
        <nav style={{ flex: 1, padding: sideOpen ? "8px 10px" : "8px 8px" }}>
          {navItems.map((item) => {
            const a = page === item.id;
            return (
              <div
                key={item.id}
                onClick={() => setPage(item.id)}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 12,
                  padding: sideOpen ? "11px 14px" : "11px 14px",
                  borderRadius: 10,
                  marginBottom: 4,
                  cursor: "pointer",
                  transition: "all .15s",
                  background: a ? C.sideA : "transparent",
                  color: a ? "#fff" : "#94a3b8",
                  position: "relative",
                }}
                onMouseEnter={(e) => {
                  if (!a) e.currentTarget.style.background = C.sideH;
                }}
                onMouseLeave={(e) => {
                  if (!a) e.currentTarget.style.background = "transparent";
                }}
              >
                <SI d={item.icon} />
                {sideOpen && (
                  <span
                    style={{
                      fontSize: 14,
                      fontWeight: a ? 600 : 500,
                      whiteSpace: "nowrap",
                    }}
                  >
                    {item.label}
                  </span>
                )}
                {item.id === "recon" && reconStats.mismatch > 0 && (
                  <span
                    style={{
                      position: "absolute",
                      right: sideOpen ? 12 : 6,
                      top: sideOpen ? 12 : 6,
                      background: C.red,
                      color: "#fff",
                      fontSize: 9,
                      fontWeight: 700,
                      borderRadius: 10,
                      padding: "1px 6px",
                      minWidth: 18,
                      textAlign: "center",
                    }}
                  >
                    {reconStats.mismatch}
                  </span>
                )}
              </div>
            );
          })}
        </nav>
        <div
          style={{ padding: sideOpen ? "16px 10px 20px" : "16px 8px 20px" }}
        >
          <div
            style={{
              padding: sideOpen ? "12px 14px" : "12px",
              borderRadius: 10,
              background: "rgba(255,255,255,.06)",
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                animation: "pulse 2s infinite",
              }}
            >
              <div
                style={{
                  width: 8,
                  height: 8,
                  borderRadius: "50%",
                  background: cyc.color,
                  boxShadow: `0 0 6px ${cyc.color}`,
                  flexShrink: 0,
                }}
              />
              {sideOpen && (
                <span
                  style={{
                    color: "#94a3b8",
                    fontSize: 11,
                    fontWeight: 600,
                    textTransform: "uppercase",
                    letterSpacing: ".5px",
                  }}
                >
                  Cycle {cyc.cycle}
                </span>
              )}
            </div>
            {sideOpen && (
              <>
                <div
                  style={{
                    color: "#e2e8f0",
                    fontSize: 12,
                    fontWeight: 500,
                    marginTop: 4,
                  }}
                >
                  {cyc.label}
                </div>
                <div
                  style={{
                    color: "#94a3b8",
                    fontSize: 13,
                    fontWeight: 600,
                    marginTop: 4,
                    fontFamily: "'JetBrains Mono',monospace",
                  }}
                >
                  <LiveClock />
                </div>
              </>
            )}
          </div>
        </div>
      </div>

      {/* Main */}
      <div style={{ flex: 1, overflow: "auto", padding: 28 }}>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            marginBottom: 24,
            flexWrap: "wrap",
            gap: 12,
          }}
        >
          <div>
            <h1
              style={{
                fontSize: 22,
                fontWeight: 700,
                color: C.text,
                letterSpacing: "-0.3px",
              }}
            >
              {navItems.find((n) => n.id === page)?.label || "Dashboard"}
            </h1>
            <p style={{ fontSize: 13, color: C.muted, marginTop: 2 }}>
              Charge {CHARGE_RATE}% + GST {page === "add" ? gstRate : GST_RATE}% on charge
            </p>
          </div>
        </div>

        {/* ========== DASHBOARD ========== */}
        {page === "dashboard" && (
          <div className="fade-up">
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))",
                gap: 16,
                marginBottom: 24,
              }}
            >
              {[
                {
                  label: "Total Payin",
                  value: formatShort(allT.payin),
                  sub: `${rows.length} entries`,
                  color: C.accent,
                  bg: "#ecfdf5",
                },
                {
                  label: "Total Charges",
                  value: formatShort(allT.charge),
                  sub: `${CHARGE_RATE}%`,
                  color: C.orange,
                  bg: "#fffbeb",
                },
                {
                  label: "Total GST",
                  value: formatShort(allT.gst),
                  sub: `${GST_RATE}% on charge`,
                  color: C.red,
                  bg: "#fef2f2",
                },
                {
                  label: "Settlement",
                  value: formatShort(allT.settlement),
                  sub:
                    rows.length > 0
                      ? `${((allT.settlement / allT.payin) * 100).toFixed(2)}%`
                      : "\u2014",
                  color: C.green,
                  bg: "#ecfdf5",
                },
                {
                  label: "Recon Issues",
                  value: reconStats.mismatch,
                  sub: `${reconStats.total} checked`,
                  color: reconStats.mismatch > 0 ? C.red : C.green,
                  bg: reconStats.mismatch > 0 ? "#fef2f2" : "#ecfdf5",
                },
              ].map((s, i) => (
                <div
                  key={i}
                  style={{
                    background: C.card,
                    borderRadius: 14,
                    padding: 20,
                    border: `1px solid ${C.border}`,
                    position: "relative",
                    overflow: "hidden",
                  }}
                >
                  <div
                    style={{
                      position: "absolute",
                      top: 0,
                      right: 0,
                      width: 80,
                      height: 80,
                      borderRadius: "0 0 0 80px",
                      background: s.bg,
                      opacity: 0.6,
                    }}
                  />
                  <div
                    style={{
                      fontSize: 11,
                      fontWeight: 600,
                      color: C.muted,
                      textTransform: "uppercase",
                      letterSpacing: ".5px",
                      marginBottom: 8,
                    }}
                  >
                    {s.label}
                  </div>
                  <div
                    style={{
                      fontSize: 22,
                      fontWeight: 700,
                      color: s.color,
                      fontFamily: "'JetBrains Mono',monospace",
                    }}
                  >
                    {s.value}
                  </div>
                  <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>
                    {s.sub}
                  </div>
                </div>
              ))}
            </div>
            {rows.length > 0 ? (
              <div
                style={{
                  background: C.card,
                  borderRadius: 14,
                  border: `1px solid ${C.border}`,
                  overflow: "hidden",
                }}
              >
                <div
                  style={{
                    padding: "16px 20px",
                    borderBottom: `1px solid ${C.border}`,
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                  }}
                >
                  <span
                    style={{ fontSize: 15, fontWeight: 700, color: C.text }}
                  >
                    Recent Entries
                  </span>
                  <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                    <button
                      onClick={() => downloadSheet(rows, "settlements")}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        padding: "6px 14px",
                        borderRadius: 8,
                        border: "none",
                        background: C.accent,
                        color: "#fff",
                        fontSize: 12,
                        fontWeight: 600,
                        cursor: "pointer",
                      }}
                    >
                      <SI d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1M12 12v8m0 0l-4-4m4 4l4-4M12 4v4" />
                      Download Sheet
                    </button>
                    <span
                      style={{
                        fontSize: 12,
                        color: C.accent,
                        fontWeight: 600,
                        cursor: "pointer",
                      }}
                      onClick={() => setPage("reports")}
                    >
                      View All &rarr;
                    </span>
                  </div>
                </div>
                <div style={{ overflowX: "auto" }}>
                  <table
                    style={{
                      width: "100%",
                      borderCollapse: "collapse",
                      fontSize: 13,
                    }}
                  >
                    <thead>
                      <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                        {[
                          "Merchant",
                          "Company",
                          "Payin",
                          "Settlement",
                          "Cycle",
                        ].map((h) => (
                          <th
                            key={h}
                            style={{
                              padding: "12px 16px",
                              textAlign:
                                h === "Merchant" || h === "Company"
                                  ? "left"
                                  : "right",
                              fontSize: 10,
                              fontWeight: 600,
                              color: C.muted,
                              textTransform: "uppercase",
                            }}
                          >
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {rows
                        .slice(-5)
                        .reverse()
                        .map((r) => (
                          <tr
                            key={r.id}
                            style={{
                              borderBottom: `1px solid ${C.border}`,
                            }}
                          >
                            <td
                              style={{
                                padding: "12px 16px",
                                fontWeight: 600,
                                color: C.accent,
                              }}
                            >
                              {r.merchant}
                            </td>
                            <td
                              style={{
                                padding: "12px 16px",
                                color: C.text,
                                maxWidth: 200,
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {r.company}
                            </td>
                            <td
                              style={{
                                padding: "12px 16px",
                                textAlign: "right",
                                fontFamily: "'JetBrains Mono',monospace",
                                fontWeight: 600,
                              }}
                            >
                              {formatINR(r.payin)}
                            </td>
                            <td
                              style={{
                                padding: "12px 16px",
                                textAlign: "right",
                                fontFamily: "'JetBrains Mono',monospace",
                                fontWeight: 700,
                                color: C.green,
                              }}
                            >
                              {formatINR(r.settlement)}
                            </td>
                            <td
                              style={{
                                padding: "12px 16px",
                                textAlign: "right",
                              }}
                            >
                              <span
                                style={{
                                  padding: "3px 10px",
                                  borderRadius: 6,
                                  fontSize: 11,
                                  fontWeight: 600,
                                  background:
                                    r.cycle === 2 ? "#fef3c7" : "#d1fae5",
                                  color:
                                    r.cycle === 2 ? "#92400e" : "#065f46",
                                }}
                              >
                                C{r.cycle}
                              </span>
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <div
                style={{
                  background: C.card,
                  borderRadius: 14,
                  border: `1px solid ${C.border}`,
                  padding: "48px 24px",
                  textAlign: "center",
                }}
              >
                <div style={{ fontSize: 40, marginBottom: 12 }}>
                  {"\uD83D\uDCCA"}
                </div>
                <div
                  style={{
                    fontSize: 16,
                    fontWeight: 600,
                    color: C.text,
                    marginBottom: 6,
                  }}
                >
                  No entries yet
                </div>
                <button
                  onClick={() => setPage("add")}
                  style={{
                    background: C.accent,
                    color: "#fff",
                    border: "none",
                    borderRadius: 10,
                    padding: "10px 24px",
                    fontSize: 14,
                    fontWeight: 600,
                    cursor: "pointer",
                  }}
                >
                  + Add Entry
                </button>
              </div>
            )}
          </div>
        )}

        {/* ========== ADD ENTRY ========== */}
        {page === "add" && (
          <div className="fade-up" style={{ maxWidth: 640 }}>
            <div
              style={{
                background: C.card,
                borderRadius: 14,
                border: `1px solid ${C.border}`,
                padding: 28,
              }}
            >
              <div
                style={{
                  fontSize: 16,
                  fontWeight: 700,
                  color: C.text,
                  marginBottom: 20,
                }}
              >
                New Settlement Entry
              </div>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "1fr 1fr",
                  gap: 16,
                  marginBottom: 16,
                }}
              >
                <div>
                  <label style={labS}>Merchant</label>
                  <select
                    value={selM}
                    onChange={(e) => {
                      setSM(e.target.value);
                      setSC("");
                      setCS("");
                    }}
                    style={selS}
                  >
                    <option value="">&mdash; Select &mdash;</option>
                    {ALL_MERCHANTS.map((m) => (
                      <option key={m} value={m}>
                        {m} ({mCnt[m]})
                      </option>
                    ))}
                  </select>
                </div>
                <div style={{ position: "relative" }}>
                  <label style={labS}>Company</label>
                  <input
                    value={selC || cSearch}
                    onChange={(e) => {
                      setCS(e.target.value);
                      setSC("");
                    }}
                    placeholder={
                      selM
                        ? `Search ${compsFor.length}...`
                        : "Select merchant"
                    }
                    style={inpS}
                    list="co-l"
                  />
                  <datalist id="co-l">
                    {filtCL.map((c) => (
                      <option key={c} value={c} />
                    ))}
                  </datalist>
                  {cSearch && !selC && filtCL.length > 0 && (
                    <div
                      style={{
                        position: "absolute",
                        top: "100%",
                        left: 0,
                        right: 0,
                        marginTop: 4,
                        background: "#fff",
                        border: `1px solid ${C.border}`,
                        borderRadius: 10,
                        maxHeight: 180,
                        overflowY: "auto",
                        zIndex: 10,
                        boxShadow: "0 8px 24px rgba(0,0,0,.12)",
                      }}
                    >
                      {filtCL.slice(0, 8).map((c) => (
                        <div
                          key={c}
                          onClick={() => {
                            setSC(c);
                            setCS("");
                          }}
                          style={{
                            padding: "10px 14px",
                            fontSize: 13,
                            cursor: "pointer",
                            borderBottom: `1px solid ${C.border}`,
                          }}
                          onMouseEnter={(e) =>
                            (e.currentTarget.style.background = "#f0fdf4")
                          }
                          onMouseLeave={(e) =>
                            (e.currentTarget.style.background = "#fff")
                          }
                        >
                          {c}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
              {selC && (
                <div
                  style={{
                    marginBottom: 16,
                    padding: "10px 16px",
                    background: "#ecfdf5",
                    borderRadius: 10,
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    border: "1px solid #a7f3d0",
                  }}
                >
                  <div style={{ fontSize: 13 }}>
                    <span style={{ fontWeight: 700, color: C.accent }}>
                      {selM}
                    </span>
                    <span style={{ color: C.muted, margin: "0 8px" }}>
                      &rarr;
                    </span>
                    <span style={{ color: C.text, fontWeight: 500 }}>
                      {selC}
                    </span>
                  </div>
                  <button
                    onClick={() => {
                      setSC("");
                      setCS("");
                    }}
                    style={{
                      background: "none",
                      border: "none",
                      color: C.muted,
                      cursor: "pointer",
                      fontSize: 18,
                    }}
                  >
                    &times;
                  </button>
                </div>
              )}
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "1.2fr 1fr 1fr",
                  gap: 14,
                  marginBottom: 16,
                }}
              >
                <div>
                  <label style={labS}>Payin Amount (₹)</label>
                  <input
                    type="number"
                    value={payin}
                    onChange={(e) => setPay(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && addRow()}
                    placeholder="e.g. 10000"
                    style={{
                      ...inpS,
                      fontSize: 18,
                      fontFamily: "'JetBrains Mono',monospace",
                      fontWeight: 600,
                    }}
                  />
                </div>
                <div>
                  <label style={labS}>
                    Chargeback (₹){" "}
                    <span style={{ color: C.muted, fontWeight: 500, textTransform: "none", letterSpacing: 0 }}>
                      optional
                    </span>
                  </label>
                  <input
                    type="number"
                    value={chargeback}
                    onChange={(e) => setChargeback(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && addRow()}
                    placeholder="0"
                    style={{
                      ...inpS,
                      fontSize: 18,
                      fontFamily: "'JetBrains Mono',monospace",
                      fontWeight: 600,
                      color: cb > 0 ? C.red : C.text,
                    }}
                  />
                </div>
                <div>
                  <label style={labS}>GST Rate</label>
                  <div style={{ display: "flex", gap: 6 }}>
                    {GST_OPTIONS.map((opt) => (
                      <button
                        key={opt}
                        type="button"
                        onClick={() => setGstRate(opt)}
                        style={{
                          flex: 1,
                          padding: "12px 4px",
                          borderRadius: 10,
                          border: `1.5px solid ${gstRate === opt ? C.accent : C.border}`,
                          background: gstRate === opt ? C.accent : "#fff",
                          color: gstRate === opt ? "#fff" : C.muted,
                          fontSize: 13,
                          fontWeight: 700,
                          cursor: "pointer",
                          fontFamily: "'JetBrains Mono',monospace",
                          transition: "all .15s",
                        }}
                      >
                        {opt}%
                      </button>
                    ))}
                  </div>
                </div>
              </div>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "1fr 1fr",
                  gap: 14,
                  marginBottom: 16,
                }}
              >
                <div>
                  <label style={labS}>
                    Entry Date{" "}
                    <span style={{ color: C.muted, fontWeight: 500, textTransform: "none", letterSpacing: 0 }}>
                      (defaults to today)
                    </span>
                  </label>
                  <input
                    type="date"
                    value={entryDate}
                    onChange={(e) => setEntryDate(e.target.value)}
                    style={inpS}
                  />
                </div>
                <div>
                  <label style={labS}>
                    Cycle{" "}
                    <span style={{ color: C.muted, fontWeight: 500, textTransform: "none", letterSpacing: 0 }}>
                      (auto: C{cyc.cycle})
                    </span>
                  </label>
                  <div style={{ display: "flex", gap: 6 }}>
                    {[
                      { c: null, label: "Auto", sub: `C${cyc.cycle}` },
                      { c: 1, label: "C1", sub: "6AM–4PM" },
                      { c: 2, label: "C2", sub: "4PM–6AM" },
                    ].map((opt) => {
                      const active = entryCycle === opt.c;
                      return (
                        <button
                          key={opt.label}
                          type="button"
                          onClick={() => setEntryCycle(opt.c)}
                          style={{
                            flex: 1,
                            padding: "8px 4px",
                            borderRadius: 10,
                            border: `1.5px solid ${active ? C.accent : C.border}`,
                            background: active ? C.accent : "#fff",
                            color: active ? "#fff" : C.muted,
                            cursor: "pointer",
                            transition: "all .15s",
                            display: "flex",
                            flexDirection: "column",
                            alignItems: "center",
                            gap: 2,
                          }}
                        >
                          <span style={{ fontSize: 12, fontWeight: 700 }}>{opt.label}</span>
                          <span style={{ fontSize: 9, fontWeight: 500, opacity: 0.85 }}>{opt.sub}</span>
                        </button>
                      );
                    })}
                  </div>
                </div>
              </div>
              <div style={{ marginBottom: 20 }}>
                <button
                  onClick={addRow}
                  disabled={!canAdd}
                  style={{
                    background: canAdd ? C.accent : "#e2e8f0",
                    color: canAdd ? "#fff" : "#94a3b8",
                    border: "none",
                    borderRadius: 10,
                    padding: "14px 28px",
                    fontSize: 15,
                    fontWeight: 700,
                    cursor: canAdd ? "pointer" : "not-allowed",
                    whiteSpace: "nowrap",
                    width: "100%",
                  }}
                >
                  + Add Entry
                </button>
              </div>
              {pn > 0 && (
                <div
                  style={{
                    display: "grid",
                    gridTemplateColumns: "repeat(5,1fr)",
                    gap: 10,
                  }}
                >
                  {[
                    { l: "Charge", v: ch, c: C.orange },
                    { l: `GST (${gstRate}%)`, v: gs, c: C.red },
                    { l: "Chargeback", v: cb, c: "#dc2626" },
                    { l: "Deduction", v: td, c: "#f97316" },
                    { l: "Final Settlement", v: stl, c: C.green },
                  ].map((x) => (
                    <div
                      key={x.l}
                      style={{
                        background: C.bg,
                        borderRadius: 10,
                        padding: 12,
                        borderLeft: `3px solid ${x.c}`,
                      }}
                    >
                      <div
                        style={{
                          fontSize: 9,
                          color: C.muted,
                          textTransform: "uppercase",
                          fontWeight: 600,
                          marginBottom: 4,
                        }}
                      >
                        {x.l}
                      </div>
                      <div
                        style={{
                          fontSize: 14,
                          fontFamily: "'JetBrains Mono',monospace",
                          fontWeight: 700,
                          color: x.c,
                        }}
                      >
                        {formatINR(x.v)}
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {/* ========== REPORTS ========== */}
        {page === "reports" && (
          <div className="fade-up">
            <div
              style={{
                background: C.card,
                borderRadius: 14,
                border: `1px solid ${C.border}`,
                padding: "14px 20px",
                marginBottom: 20,
                display: "flex",
                flexWrap: "wrap",
                alignItems: "center",
                gap: 12,
              }}
            >
              <div
                style={{
                  display: "flex",
                  gap: 4,
                  background: C.bg,
                  borderRadius: 8,
                  padding: 3,
                }}
              >
                {[
                  { k: "all", l: "All" },
                  { k: "company", l: "Company" },
                  { k: "merchant", l: "Merchant" },
                ].map((t) => (
                  <button
                    key={t.k}
                    onClick={() => setView(t.k)}
                    style={{
                      background: view === t.k ? C.accent : "transparent",
                      color: view === t.k ? "#fff" : C.muted,
                      border: "none",
                      borderRadius: 6,
                      padding: "8px 16px",
                      fontSize: 12,
                      fontWeight: 600,
                      cursor: "pointer",
                    }}
                  >
                    {t.l}
                  </button>
                ))}
              </div>
              <div
                style={{
                  display: "flex",
                  gap: 10,
                  flexWrap: "wrap",
                  flex: 1,
                  justifyContent: "flex-end",
                }}
              >
                <select
                  value={fM}
                  onChange={(e) => setFM(e.target.value)}
                  style={filtS}
                >
                  <option value="__all__">All Merchants</option>
                  {[...new Set(rows.map((r) => r.merchant))]
                    .sort()
                    .map((m) => (
                      <option key={m} value={m}>
                        {m}
                      </option>
                    ))}
                </select>
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 6,
                    padding: "4px 10px",
                    borderRadius: 8,
                    border: `1px solid ${C.border}`,
                    background: "#fff",
                  }}
                  title="Filter downloads by date range"
                >
                  <span style={{ fontSize: 10, fontWeight: 600, color: C.muted, textTransform: "uppercase" }}>
                    From
                  </span>
                  <input
                    type="date"
                    value={dlFrom}
                    onChange={(e) => setDlFrom(e.target.value)}
                    style={{ border: "none", fontSize: 12, color: C.text, padding: 0, outline: "none" }}
                  />
                  <span style={{ fontSize: 10, fontWeight: 600, color: C.muted, textTransform: "uppercase" }}>
                    To
                  </span>
                  <input
                    type="date"
                    value={dlTo}
                    onChange={(e) => setDlTo(e.target.value)}
                    style={{ border: "none", fontSize: 12, color: C.text, padding: 0, outline: "none" }}
                  />
                  {(dlFrom || dlTo) && (
                    <button
                      onClick={() => { setDlFrom(""); setDlTo(""); }}
                      style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 14, padding: 0 }}
                      title="Clear date range"
                    >
                      &times;
                    </button>
                  )}
                </div>
                {rows.length > 0 && (
                  <>
                    <button
                      onClick={() => downloadSheet(fRows, "settlement_report")}
                      style={{
                        ...filtS,
                        border: `1px solid ${C.accent}`,
                        background: C.accent,
                        color: "#fff",
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                      }}
                    >
                      <SI d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1M12 12v8m0 0l-4-4m4 4l4-4M12 4v4" />
                      Download Sheet
                    </button>
                    <button
                      onClick={clearAll}
                      style={{
                        ...filtS,
                        border: "1px solid #fca5a5",
                        background: "#fef2f2",
                        color: C.red,
                      }}
                    >
                      Clear All
                    </button>
                  </>
                )}
              </div>
            </div>
            {grouped && grouped.length > 0 && (
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns:
                    "repeat(auto-fill,minmax(260px,1fr))",
                  gap: 14,
                  marginBottom: 20,
                }}
              >
                {grouped.map((g, i) => (
                  <div
                    key={i}
                    className="fade-up"
                    style={{
                      background: C.card,
                      border: `1px solid ${C.border}`,
                      borderRadius: 14,
                      padding: 20,
                      borderTop: `3px solid ${view === "company" ? C.accent : C.green}`,
                    }}
                  >
                    <div
                      style={{
                        fontSize: 14,
                        fontWeight: 700,
                        color: C.text,
                        marginBottom: 10,
                        lineHeight: 1.3,
                      }}
                    >
                      {g.name}
                    </div>
                    <div
                      style={{
                        display: "grid",
                        gridTemplateColumns: "1fr 1fr",
                        gap: 8,
                      }}
                    >
                      {[
                        { l: "Entries", v: g.count, raw: 1, c: C.muted },
                        { l: "Total Payin", v: g.payin, c: C.text },
                        { l: "Deductions", v: g.deduction, c: C.orange },
                        { l: "Settlement", v: g.settlement, c: C.green },
                      ].map((x) => (
                        <div key={x.l}>
                          <div
                            style={{
                              fontSize: 9,
                              color: C.muted,
                              textTransform: "uppercase",
                              fontWeight: 600,
                            }}
                          >
                            {x.l}
                          </div>
                          <div
                            style={{
                              fontSize: 13,
                              fontFamily: "'JetBrains Mono',monospace",
                              fontWeight: 700,
                              color: x.c,
                            }}
                          >
                            {x.raw ? x.v : formatINR(x.v)}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            )}
            {fRows.length > 0 && view === "all" && (
              <div
                style={{
                  background: C.card,
                  borderRadius: 14,
                  border: `1px solid ${C.border}`,
                  overflow: "hidden",
                }}
              >
                <div
                  style={{
                    padding: "14px 20px",
                    borderBottom: `1px solid ${C.border}`,
                    fontSize: 14,
                    fontWeight: 700,
                    color: C.text,
                  }}
                >
                  All Entries ({fRows.length})
                </div>
                <div style={{ overflowX: "auto" }}>
                  <table
                    style={{
                      width: "100%",
                      borderCollapse: "collapse",
                      fontSize: 13,
                    }}
                  >
                    <thead>
                      <tr style={{ borderBottom: `1px solid ${C.border}` }}>
                        {[
                          "#",
                          "Merchant",
                          "Company",
                          "Payin",
                          "Deduction",
                          "Settlement",
                          "C",
                          "",
                        ].map((h, i) => (
                          <th
                            key={i}
                            style={{
                              padding: "11px 14px",
                              textAlign: i <= 2 ? "left" : "right",
                              fontSize: 10,
                              fontWeight: 700,
                              color: C.muted,
                              textTransform: "uppercase",
                            }}
                          >
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {fRows.map((r, i) => (
                        <tr
                          key={r.id}
                          style={{
                            borderBottom: `1px solid ${C.border}`,
                          }}
                          onMouseEnter={(e) =>
                            (e.currentTarget.style.background = "#f8fafc")
                          }
                          onMouseLeave={(e) =>
                            (e.currentTarget.style.background = "#fff")
                          }
                        >
                          <td
                            style={{
                              padding: "10px 14px",
                              color: C.muted,
                              fontSize: 11,
                            }}
                          >
                            {i + 1}
                          </td>
                          <td
                            style={{
                              padding: "10px 14px",
                              fontWeight: 600,
                              color: C.accent,
                            }}
                          >
                            {r.merchant}
                          </td>
                          <td
                            style={{
                              padding: "10px 14px",
                              color: C.text,
                              maxWidth: 200,
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                              whiteSpace: "nowrap",
                            }}
                            title={r.company}
                          >
                            {r.company}
                          </td>
                          <td
                            style={{
                              padding: "10px 14px",
                              textAlign: "right",
                              fontFamily: "'JetBrains Mono',monospace",
                              fontWeight: 600,
                            }}
                          >
                            {formatINR(r.payin)}
                          </td>
                          <td
                            style={{
                              padding: "10px 14px",
                              textAlign: "right",
                              fontFamily: "'JetBrains Mono',monospace",
                              color: C.orange,
                            }}
                          >
                            {formatINR(r.deduction)}
                          </td>
                          <td
                            style={{
                              padding: "10px 14px",
                              textAlign: "right",
                              fontFamily: "'JetBrains Mono',monospace",
                              fontWeight: 700,
                              color: C.green,
                            }}
                          >
                            {formatINR(r.settlement)}
                          </td>
                          <td
                            style={{
                              padding: "10px 14px",
                              textAlign: "center",
                            }}
                          >
                            <span
                              style={{
                                padding: "2px 8px",
                                borderRadius: 5,
                                fontSize: 11,
                                fontWeight: 600,
                                background:
                                  r.cycle === 2 ? "#fef3c7" : "#d1fae5",
                                color:
                                  r.cycle === 2 ? "#92400e" : "#065f46",
                              }}
                            >
                              C{r.cycle}
                            </span>
                          </td>
                          <td style={{ padding: "10px 8px" }}>
                            <button
                              onClick={() => rmRow(r.id)}
                              style={{
                                background: "none",
                                border: "none",
                                color: "#cbd5e1",
                                cursor: "pointer",
                                fontSize: 16,
                              }}
                            >
                              &times;
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
            {rows.length === 0 && (
              <div
                style={{
                  background: C.card,
                  borderRadius: 14,
                  border: `1px solid ${C.border}`,
                  padding: "48px 24px",
                  textAlign: "center",
                  color: C.muted,
                }}
              >
                No entries.{" "}
                <span
                  style={{
                    color: C.accent,
                    cursor: "pointer",
                    fontWeight: 600,
                  }}
                  onClick={() => setPage("add")}
                >
                  Add one &rarr;
                </span>
              </div>
            )}
          </div>
        )}

        {/* ========== RECONCILIATION ========== */}
        {page === "recon" && (
          <div className="fade-up">
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))",
                gap: 14,
                marginBottom: 24,
              }}
            >
              {[
                {
                  l: "Total Checked",
                  v: reconStats.total,
                  c: C.blue,
                  bg: "#eff6ff",
                },
                {
                  l: "Matched",
                  v: reconStats.matched,
                  c: C.green,
                  bg: "#ecfdf5",
                },
                {
                  l: "Mismatch",
                  v: reconStats.mismatch,
                  c: C.red,
                  bg: "#fef2f2",
                },
                {
                  l: "Pending",
                  v: reconStats.pending,
                  c: C.orange,
                  bg: "#fffbeb",
                },
                {
                  l: "Resolved",
                  v: reconStats.resolved,
                  c: C.blue,
                  bg: "#eff6ff",
                },
                {
                  l: "Total Discrepancy",
                  v: formatINR(reconStats.totalDisc),
                  c: C.red,
                  bg: "#fef2f2",
                  raw: 1,
                },
              ].map((s, i) => (
                <div
                  key={i}
                  style={{
                    background: C.card,
                    borderRadius: 12,
                    padding: "16px 18px",
                    border: `1px solid ${C.border}`,
                    borderTop: `3px solid ${s.c}`,
                  }}
                >
                  <div
                    style={{
                      fontSize: 10,
                      fontWeight: 600,
                      color: C.muted,
                      textTransform: "uppercase",
                      letterSpacing: ".5px",
                      marginBottom: 6,
                    }}
                  >
                    {s.l}
                  </div>
                  <div
                    style={{
                      fontSize: s.raw ? 16 : 24,
                      fontWeight: 700,
                      color: s.c,
                      fontFamily: "'JetBrains Mono',monospace",
                    }}
                  >
                    {s.v}
                  </div>
                </div>
              ))}
            </div>

            {/* Saved Recons (persisted to localStorage) — load past days back any time. */}
            {savedRecons.length > 0 && (
              <div
                style={{
                  background: C.card,
                  borderRadius: 14,
                  border: `1px solid ${C.border}`,
                  padding: "18px 20px",
                  marginBottom: 16,
                }}
              >
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 700, color: C.text }}>
                      💾 Saved Recons
                    </div>
                    <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>
                      Click any row to reload it — overrides the current files & date.
                    </div>
                  </div>
                  <select
                    value={savedReconsFilter}
                    onChange={(e) => setSavedReconsFilter(e.target.value)}
                    style={{ padding: "6px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 12, fontWeight: 600, background: "#fff", cursor: "pointer" }}
                  >
                    <option value="__all__">All merchants ({savedRecons.length})</option>
                    {[...new Set(savedRecons.map((r) => r.merchant))].sort().map((m) => (
                      <option key={m} value={m}>
                        {m === "__all__" ? "All merchants" : m} ({savedRecons.filter((r) => r.merchant === m).length})
                      </option>
                    ))}
                  </select>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6, maxHeight: 220, overflowY: "auto" }}>
                  {visibleSavedRecons.length === 0 && (
                    <div style={{ fontSize: 12, color: C.muted, padding: "12px 0", textAlign: "center" }}>
                      No saved recons match this filter.
                    </div>
                  )}
                  {visibleSavedRecons.map((entry) => {
                    const totalSource = (entry.sourceRows || []).reduce((s, r) => s + (r.amount || 0), 0);
                    const totalBank = (entry.bankRows || []).reduce((s, r) => s + (r.settle || 0), 0);
                    const dateText = entry.date ? formatDateLong(entry.date) : "(no date)";
                    const cycles = [...new Set((entry.bankFiles || []).map((f) => f.cycle).filter(Boolean))].sort();
                    return (
                      <div
                        key={entry.id}
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                          padding: "10px 12px",
                          borderRadius: 8,
                          border: `1px solid ${C.border}`,
                          background: "#fafbfc",
                          fontSize: 12,
                          gap: 8,
                        }}
                      >
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                            <span style={{ fontWeight: 700, color: C.accent, fontSize: 13 }}>
                              {entry.merchantLabel || "All merchants"}
                            </span>
                            <span style={{ color: C.muted }}>·</span>
                            <span style={{ fontWeight: 600, color: C.text }}>{dateText}</span>
                            {cycles.length > 0 && (
                              <span style={{ marginLeft: 4, display: "flex", gap: 3 }}>
                                {cycles.map((c) => (
                                  <span key={c} style={{ padding: "1px 6px", borderRadius: 4, background: "#065f46", color: "#fff", fontSize: 9, fontWeight: 700 }}>
                                    C{c}
                                  </span>
                                ))}
                              </span>
                            )}
                          </div>
                          <div style={{ fontSize: 10, color: C.muted, marginTop: 2 }}>
                            Source: {formatINR(totalSource)} · Bank settle: {formatINR(totalBank)} · {(entry.bankFiles || []).length} bank file(s)
                            {entry.startTime && entry.endTime && ` · ${entry.startTime}–${entry.endTime}`}
                          </div>
                        </div>
                        <div style={{ display: "flex", gap: 6 }}>
                          <button
                            onClick={() => loadSavedRecon(entry)}
                            style={{ padding: "6px 12px", borderRadius: 6, border: "none", background: C.accent, color: "#fff", fontSize: 11, fontWeight: 700, cursor: "pointer" }}
                          >
                            Load
                          </button>
                          <button
                            onClick={() => deleteSavedRecon(entry.id)}
                            style={{ padding: "6px 10px", borderRadius: 6, border: `1px solid ${C.border}`, background: "#fff", color: C.red, fontSize: 11, fontWeight: 700, cursor: "pointer" }}
                            title="Delete this saved recon"
                          >
                            ×
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* Reconciliation Uploads */}
            <div
              style={{
                background: C.card,
                borderRadius: 14,
                border: `1px solid ${C.border}`,
                padding: 24,
                marginBottom: 20,
              }}
            >
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14, flexWrap: "wrap", gap: 8 }}>
                <div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: C.text }}>
                    Reconciliation Uploads
                  </div>
                  <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>
                    Upload internal payin and bank settlement files — auto-aggregates by merchant
                  </div>
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                  {/* Merchant filter — primary picker; limits both the in-app table and the
                      downloaded Excel. Shown even before any files are uploaded so the user
                      can pick the merchant first and then upload that merchant's data. */}
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{ fontSize: 10, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: ".5px" }}>
                      Merchant
                    </span>
                    <select
                      value={reconMerchantFilter}
                      onChange={(e) => setReconMerchantFilter(e.target.value)}
                      style={{
                        padding: "7px 10px",
                        borderRadius: 8,
                        border: `1px solid ${reconMerchantFilter === "__all__" ? C.border : C.accent}`,
                        background: reconMerchantFilter === "__all__" ? "#fff" : "#ecfdf5",
                        color: C.text,
                        fontSize: 12,
                        fontWeight: 600,
                        cursor: "pointer",
                        minWidth: 180,
                      }}
                    >
                      <option value="__all__">
                        {combinedByMerchant.length > 0
                          ? `All merchants (${combinedByMerchant.length})`
                          : "All merchants"}
                      </option>
                      {combinedByMerchant.length > 0
                        ? combinedByMerchant.map((g) => (
                            <option key={g.merchant} value={g.merchant}>{g.merchant}</option>
                          ))
                        : ALL_MERCHANTS.map((m) => (
                            <option key={m} value={m}>{m}</option>
                          ))}
                    </select>
                  </div>
                  {(bankRows.length > 0 || sourceRows.length > 0) && (
                    <>
                      <button
                        onClick={() => {
                          const id = saveCurrentRecon();
                          if (id) alert("Recon saved — you can load it back from the Saved Recons panel.");
                        }}
                        style={{ padding: "8px 14px", borderRadius: 8, border: `1px solid ${C.accent}`, background: "#fff", color: C.accent, fontSize: 12, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}
                        title="Save current files + date so you can reload this recon later"
                      >
                        💾 Save Recon
                      </button>
                      <button
                        onClick={() => {
                          if (confirm("Clear all uploaded files and reset the period? (Saved recons are NOT affected.)")) {
                            clearAllReconInputs();
                          }
                        }}
                        style={{ padding: "8px 14px", borderRadius: 8, border: `1px solid ${C.border}`, background: "#fff", color: C.muted, fontSize: 12, fontWeight: 600, cursor: "pointer" }}
                        title="Start a fresh recon (e.g. for the next day) — saved recons aren't deleted"
                      >
                        New day
                      </button>
                      <button
                        onClick={downloadBankRecon}
                        style={{ padding: "8px 14px", borderRadius: 8, border: "none", background: C.accent, color: "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}
                      >
                        <SI d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1M12 12v8m0 0l-4-4m4 4l4-4M12 4v4" />
                        {reconMerchantFilter === "__all__" ? "Download Recon" : `Download ${reconMerchantFilter}`}
                      </button>
                    </>
                  )}
                </div>
              </div>

              {/* Settlement Period: which date + time range this recon is for.
                  The values get embedded in the downloaded Excel and the filename. */}
              <div
                style={{
                  marginBottom: 14,
                  padding: "12px 14px",
                  borderRadius: 10,
                  background: "#f8fafc",
                  border: `1px solid ${C.border}`,
                }}
              >
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, flexWrap: "wrap", gap: 6 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: ".5px" }}>
                    Settlement Period
                  </div>
                  {(reconDate || reconStartTime || reconEndTime) && (
                    <button
                      onClick={() => { setReconDate(""); setReconStartTime(""); setReconEndTime(""); }}
                      style={{ background: "none", border: "none", color: C.muted, fontSize: 11, fontWeight: 600, cursor: "pointer" }}
                    >
                      Clear ×
                    </button>
                  )}
                </div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center" }}>
                  <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                    <label style={{ fontSize: 10, fontWeight: 600, color: C.muted }}>Date</label>
                    <input
                      type="date"
                      value={reconDate}
                      onChange={(e) => setReconDate(e.target.value)}
                      style={{ padding: "6px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 12, color: C.text, background: "#fff" }}
                    />
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                    <label style={{ fontSize: 10, fontWeight: 600, color: C.muted }}>From Time</label>
                    <input
                      type="time"
                      value={reconStartTime}
                      onChange={(e) => setReconStartTime(e.target.value)}
                      style={{ padding: "6px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 12, color: C.text, background: "#fff", fontFamily: "'JetBrains Mono',monospace" }}
                    />
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                    <label style={{ fontSize: 10, fontWeight: 600, color: C.muted }}>To Time</label>
                    <input
                      type="time"
                      value={reconEndTime}
                      onChange={(e) => setReconEndTime(e.target.value)}
                      style={{ padding: "6px 10px", borderRadius: 8, border: `1px solid ${C.border}`, fontSize: 12, color: C.text, background: "#fff", fontFamily: "'JetBrains Mono',monospace" }}
                    />
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                    <label style={{ fontSize: 10, fontWeight: 600, color: C.muted }}>Quick fill (cycle)</label>
                    <div style={{ display: "flex", gap: 4 }}>
                      {RECON_CYCLE_PRESETS.map((p) => {
                        const active = reconStartTime === p.start && reconEndTime === p.end;
                        return (
                          <button
                            key={p.id}
                            onClick={() => {
                              setReconStartTime(p.start);
                              setReconEndTime(p.end);
                              if (!reconDate) setReconDate(new Date().toISOString().slice(0, 10));
                            }}
                            title={p.range}
                            style={{
                              padding: "6px 10px",
                              borderRadius: 8,
                              border: `1px solid ${active ? C.accent : C.border}`,
                              background: active ? C.accent : "#fff",
                              color: active ? "#fff" : C.text,
                              fontSize: 11,
                              fontWeight: 600,
                              cursor: "pointer",
                              display: "flex",
                              flexDirection: "column",
                              alignItems: "center",
                              gap: 1,
                            }}
                          >
                            <span>{p.label}</span>
                            <span style={{ fontSize: 9, fontWeight: 500, opacity: 0.85 }}>{p.range}</span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                </div>
                {(reconDate || (reconStartTime && reconEndTime)) && (
                  <div style={{ marginTop: 8, fontSize: 11, color: "#1e40af", fontWeight: 600 }}>
                    📅 {formatDateLong(reconDate) || "—"}
                    {(reconStartTime || reconEndTime) && (
                      <span> · ⏱ {reconStartTime || "—"} to {reconEndTime || "—"}</span>
                    )}
                  </div>
                )}
              </div>

              {/* Two side-by-side drop zones */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
                {/* Source / internal payin */}
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: ".5px" }}>
                      1. Internal Payin
                    </div>
                    {sourceRows.length > 0 && (
                      <button onClick={clearSource} style={{ background: "none", border: "none", color: C.muted, fontSize: 11, fontWeight: 600, cursor: "pointer" }}>
                        Clear ×
                      </button>
                    )}
                  </div>
                  {sourceRows.length === 0 ? (
                    <div
                      onDragOver={(e) => { e.preventDefault(); setSourceDragging(true); }}
                      onDragLeave={() => setSourceDragging(false)}
                      onDrop={handleSourceDrop}
                      onClick={() => document.getElementById("source-file-input")?.click()}
                      style={{
                        background: sourceDragging ? "#eff6ff" : C.bg,
                        border: sourceDragging ? `2.5px dashed ${C.blue}` : `2px dashed ${C.border}`,
                        borderRadius: 12,
                        padding: "24px 14px",
                        textAlign: "center",
                        cursor: "pointer",
                        transition: "all .2s",
                      }}
                    >
                      <input id="source-file-input" type="file" accept=".csv,.xlsx,.xls" onChange={handleSourceInput} style={{ display: "none" }} />
                      <div style={{ fontSize: 24, marginBottom: 6 }}>{sourceDragging ? "📥" : "📊"}</div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: C.text, marginBottom: 2 }}>
                        {sourceDragging ? "Drop here" : "Drag & drop internal payin Excel"}
                      </div>
                      <div style={{ fontSize: 10, color: C.muted }}>
                        Detects: Row Labels / Name + Sum of Amount
                      </div>
                    </div>
                  ) : (
                    <div style={{ padding: "12px 14px", background: "#eff6ff", borderRadius: 10, border: "1px solid #bfdbfe", fontSize: 11, color: "#1e40af" }}>
                      <div><strong>📊 {sourceFile}</strong></div>
                      <div style={{ marginTop: 4 }}>
                        {sourceRows.length} companies · {sourceByMerchant.size} merchants · Total{" "}
                        <strong>{formatINR(sourceRows.reduce((s, r) => s + r.amount, 0))}</strong>
                      </div>
                    </div>
                  )}
                  {sourceError && (
                    <div style={{ marginTop: 8, padding: "8px 12px", borderRadius: 6, background: "#fef2f2", border: "1px solid #fca5a5", color: "#991b1b", fontSize: 11, fontWeight: 600 }}>
                      {sourceError}
                    </div>
                  )}
                </div>

                {/* Bank settlement — supports multiple files (e.g. 3 settlement cycles per day) */}
                <div>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: ".5px" }}>
                      2. Bank Settlement{bankFiles.length > 1 && ` (${bankFiles.length} files)`}
                    </div>
                    {bankFiles.length > 0 && (
                      <button onClick={clearBank} style={{ background: "none", border: "none", color: C.muted, fontSize: 11, fontWeight: 600, cursor: "pointer" }}>
                        Clear all ×
                      </button>
                    )}
                  </div>
                  <div
                    onDragOver={(e) => { e.preventDefault(); setBankDragging(true); }}
                    onDragLeave={() => setBankDragging(false)}
                    onDrop={handleBankDrop}
                    onClick={() => document.getElementById("bank-file-input")?.click()}
                    style={{
                      background: bankDragging ? "#ecfdf5" : C.bg,
                      border: bankDragging ? `2.5px dashed ${C.accent}` : `2px dashed ${C.border}`,
                      borderRadius: 12,
                      padding: bankFiles.length > 0 ? "14px 14px" : "24px 14px",
                      textAlign: "center",
                      cursor: "pointer",
                      transition: "all .2s",
                    }}
                  >
                    <input id="bank-file-input" type="file" accept=".csv,.xlsx,.xls" multiple onChange={handleBankInput} style={{ display: "none" }} />
                    <div style={{ fontSize: bankFiles.length > 0 ? 18 : 24, marginBottom: 4 }}>
                      {bankDragging ? "📥" : "🏦"}
                    </div>
                    <div style={{ fontSize: 12, fontWeight: 700, color: C.text, marginBottom: 2 }}>
                      {bankDragging
                        ? "Drop file(s) here"
                        : bankFiles.length > 0
                          ? "+ Add another bank settlement file"
                          : "Drag & drop bank settlement Excel(s) — multiple allowed"}
                    </div>
                    {bankFiles.length === 0 && (
                      <div style={{ fontSize: 10, color: C.muted }}>
                        Detects: MID · Name · Amount · Fee · GST · Settle · Chargeback
                      </div>
                    )}
                  </div>
                  {bankFiles.length > 0 && (
                    <div style={{ marginTop: 8, display: "flex", flexDirection: "column", gap: 6 }}>
                      {bankFiles.map((f) => {
                        const fileRows = bankRows.filter((r) => r._sourceFile === f.name);
                        const fileTotal = fileRows.reduce((s, r) => s + (r.amount || 0), 0);
                        return (
                          <div
                            key={f.name}
                            style={{
                              padding: "8px 12px",
                              background: "#ecfdf5",
                              borderRadius: 8,
                              border: "1px solid #a7f3d0",
                              fontSize: 11,
                              color: "#065f46",
                              display: "flex",
                              justifyContent: "space-between",
                              alignItems: "center",
                              gap: 8,
                            }}
                          >
                            <div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "flex", alignItems: "center", gap: 6 }} title={f.name}>
                              {f.cycle ? (
                                <span style={{ padding: "2px 7px", borderRadius: 5, background: "#065f46", color: "#fff", fontSize: 10, fontWeight: 700, letterSpacing: ".3px" }}>
                                  C{f.cycle}
                                </span>
                              ) : (
                                <span style={{ padding: "2px 7px", borderRadius: 5, background: "#fef3c7", color: "#92400e", fontSize: 10, fontWeight: 700 }} title="No cycle detected — name your file 1/2/3 or 'Cycle 1' etc.">
                                  C?
                                </span>
                              )}
                              <strong>{f.name}</strong>
                              <span style={{ opacity: 0.8 }}>
                                · {f.rowCount} cos · {formatINR(fileTotal)}
                              </span>
                            </div>
                            <button
                              onClick={(e) => { e.stopPropagation(); removeBankFile(f.name); }}
                              style={{ background: "none", border: "none", color: "#065f46", fontSize: 14, fontWeight: 700, cursor: "pointer", padding: 0, lineHeight: 1 }}
                              title="Remove this file"
                            >
                              ×
                            </button>
                          </div>
                        );
                      })}
                      <div style={{ fontSize: 11, color: C.muted, fontWeight: 600, paddingLeft: 4 }}>
                        Total: {bankRows.length} companies · {bankByMerchant.length} merchants
                        {bankTotals.unmatched > 0 && (
                          <span style={{ color: "#991b1b", marginLeft: 6 }}>· {bankTotals.unmatched} unmatched</span>
                        )}
                      </div>
                    </div>
                  )}
                  {bankError && (
                    <div style={{ marginTop: 8, padding: "8px 12px", borderRadius: 6, background: "#fef2f2", border: "1px solid #fca5a5", color: "#991b1b", fontSize: 11, fontWeight: 600 }}>
                      {bankError}
                    </div>
                  )}
                </div>
              </div>

              {(bankRows.length > 0 || sourceRows.length > 0) && (
                <>

                  {/* Aggregated summary table */}
                  <div style={{ marginTop: 14, overflowX: "auto", border: `1px solid ${C.border}`, borderRadius: 10 }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                      <thead>
                        <tr style={{ borderBottom: `2px solid ${C.border}`, background: "#f8fafc" }}>
                          {(() => {
                            const headers = ["Merchant"];
                            if (sourceRows.length > 0) headers.push("Source Payin");
                            if (bankRows.length > 0) headers.push("Bank Payin");
                            if (sourceRows.length > 0 && bankRows.length > 0) headers.push("Payin Diff");
                            if (bankRows.length > 0) headers.push("Bank Fee", "Bank GST", "Bank Settle", "Chargeback", "Net Settle", "Our Settle", "Difference");
                            return headers.map((h, i) => (
                              <th key={i} style={{ padding: "10px 10px", textAlign: i <= 0 ? "left" : "right", fontSize: 9, fontWeight: 700, color: C.muted, textTransform: "uppercase", letterSpacing: ".5px", whiteSpace: "nowrap" }}>
                                {h}
                              </th>
                            ));
                          })()}
                        </tr>
                      </thead>
                      <tbody>
                        {filteredCombinedByMerchant.map((g) => (
                          <tr key={g.merchant} style={{ borderBottom: `1px solid ${C.border}` }} onMouseEnter={(e) => (e.currentTarget.style.background = "#f8fafc")} onMouseLeave={(e) => (e.currentTarget.style.background = "#fff")}>
                            <td style={{ padding: "10px 10px", fontWeight: 700, color: g.merchant === "Unmatched" ? C.red : C.accent }}>
                              {g.merchant}
                              {g.unmatched > 0 && g.merchant !== "Unmatched" && (
                                <span style={{ marginLeft: 6, padding: "1px 6px", background: "#fef3c7", color: "#92400e", borderRadius: 4, fontSize: 9, fontWeight: 700 }} title="unmatched bank rows">
                                  {g.unmatched} ub
                                </span>
                              )}
                            </td>
                            {sourceRows.length > 0 && (
                              <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 600, color: "#7c3aed" }}>
                                {g.sourcePayin > 0 ? formatINR(g.sourcePayin) : "—"}
                              </td>
                            )}
                            {bankRows.length > 0 && (
                              <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 600 }}>
                                {g.amount > 0 ? formatINR(g.amount) : "—"}
                              </td>
                            )}
                            {sourceRows.length > 0 && bankRows.length > 0 && (
                              <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: Math.abs(g.payinDiff) <= 1 ? C.green : C.red }}>
                                {g.payinDiff > 0 ? "+" : ""}{formatINR(g.payinDiff)}
                              </td>
                            )}
                            {bankRows.length > 0 && (
                              <>
                                <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", color: C.orange }}>{formatINR(g.fee)}</td>
                                <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", color: C.red }}>{formatINR(g.gst)}</td>
                                <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: C.blue }}>{formatINR(g.settle)}</td>
                                <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", color: g.chargeback > 0 ? "#dc2626" : C.muted }}>
                                  {g.chargeback > 0 ? "−" + formatINR(g.chargeback) : "—"}
                                </td>
                                <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: "#1e40af" }}>{formatINR(g.netSettle)}</td>
                                <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: C.green }}>{formatINR(g.ourNet)}</td>
                                <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: Math.abs(g.diff) <= 1 ? C.green : C.red }}>
                                  {g.diff > 0 ? "+" : ""}{formatINR(g.diff)}
                                </td>
                              </>
                            )}
                          </tr>
                        ))}
                        <tr style={{ background: "#f1f5f9", borderTop: `2px solid ${C.border}` }}>
                          <td style={{ padding: "10px 10px", fontWeight: 700, color: C.text }}>TOTAL</td>
                          {sourceRows.length > 0 && (
                            <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: "#7c3aed" }}>{formatINR(filteredCombinedTotals.sourcePayin)}</td>
                          )}
                          {bankRows.length > 0 && (
                            <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700 }}>{formatINR(filteredCombinedTotals.bankPayin)}</td>
                          )}
                          {sourceRows.length > 0 && bankRows.length > 0 && (
                            <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: Math.abs(filteredCombinedTotals.payinDiff) <= 1 ? C.green : C.red }}>
                              {filteredCombinedTotals.payinDiff > 0 ? "+" : ""}{formatINR(filteredCombinedTotals.payinDiff)}
                            </td>
                          )}
                          {bankRows.length > 0 && (
                            <>
                              <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: C.orange }}>{formatINR(filteredCombinedTotals.fee)}</td>
                              <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: C.red }}>{formatINR(filteredCombinedTotals.gst)}</td>
                              <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: C.blue }}>{formatINR(filteredCombinedTotals.settle)}</td>
                              <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: filteredCombinedTotals.chargeback > 0 ? "#dc2626" : C.muted }}>
                                {filteredCombinedTotals.chargeback > 0 ? "−" + formatINR(filteredCombinedTotals.chargeback) : "—"}
                              </td>
                              <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: "#1e40af" }}>{formatINR(filteredCombinedTotals.netSettle)}</td>
                              <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: C.green }}>{formatINR(filteredCombinedTotals.ourNet)}</td>
                              <td style={{ padding: "10px 10px", textAlign: "right", fontFamily: "'JetBrains Mono',monospace", fontWeight: 700, color: Math.abs(filteredCombinedTotals.diff) <= 1 ? C.green : C.red }}>
                                {filteredCombinedTotals.diff > 0 ? "+" : ""}{formatINR(filteredCombinedTotals.diff)}
                              </td>
                            </>
                          )}
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>

            <div
              style={{
                background: C.card,
                borderRadius: 14,
                border: `1px solid ${C.border}`,
                padding: 24,
                marginBottom: 20,
              }}
            >
              <div
                style={{
                  fontSize: 15,
                  fontWeight: 700,
                  color: C.text,
                  marginBottom: 16,
                }}
              >
                Verify Transaction
              </div>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "1fr 1fr 1fr",
                  gap: 14,
                  marginBottom: 14,
                }}
              >
                <div>
                  <label style={labS}>Merchant</label>
                  <select
                    value={rMerchant}
                    onChange={(e) => {
                      setRM(e.target.value);
                      setRC("");
                    }}
                    style={selS}
                  >
                    <option value="">&mdash; Select &mdash;</option>
                    {ALL_MERCHANTS.map((m) => (
                      <option key={m} value={m}>
                        {m}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label style={labS}>Company</label>
                  <select
                    value={rCompany}
                    onChange={(e) => setRC(e.target.value)}
                    style={selS}
                  >
                    <option value="">&mdash; Select &mdash;</option>
                    {rCompsFor.map((c) => (
                      <option key={c} value={c}>
                        {c}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label style={labS}>Transaction ID</label>
                  <input
                    value={rTxnId}
                    onChange={(e) => setRTxn(e.target.value)}
                    placeholder="TXN-001"
                    style={inpS}
                  />
                </div>
              </div>
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr",
                  gap: 14,
                  marginBottom: 14,
                }}
              >
                <div>
                  <label style={labS}>Payin (₹)</label>
                  <input
                    type="number"
                    value={rPayin}
                    onChange={(e) => setRPay(e.target.value)}
                    placeholder="10000"
                    style={{
                      ...inpS,
                      fontFamily: "'JetBrains Mono',monospace",
                    }}
                  />
                </div>
                <div>
                  <label style={labS}>Amount to be Settled in Bank (₹)</label>
                  <input
                    type="number"
                    value={rClaimed}
                    onChange={(e) => setRClaimed(e.target.value)}
                    placeholder="9950"
                    style={{
                      ...inpS,
                      fontFamily: "'JetBrains Mono',monospace",
                    }}
                  />
                </div>
                <div>
                  <label style={labS}>UTR Number</label>
                  <input
                    value={rUTR}
                    onChange={(e) => setRUTR(e.target.value)}
                    placeholder="UTR123456"
                    style={inpS}
                  />
                </div>
                <div>
                  <label style={labS}>Date</label>
                  <input
                    type="date"
                    value={rDate}
                    onChange={(e) => setRDate(e.target.value)}
                    style={inpS}
                  />
                </div>
                <div>
                  <label style={labS}>Cycle</label>
                  <select
                    value={rCycle}
                    onChange={(e) => setRCycle(e.target.value)}
                    style={selS}
                  >
                    <option value="1">Cycle 1</option>
                    <option value="2">Cycle 2</option>
                  </select>
                </div>
              </div>

              {(parseFloat(rPayin) || 0) > 0 &&
                (() => {
                  const rp = parseFloat(rPayin);
                  const rc2 = parseFloat(rClaimed) || 0;
                  const ch2 = (rp * CHARGE_RATE) / 100;
                  const gs2 = (ch2 * GST_RATE) / 100;
                  const os = rp - ch2 - gs2;
                  const df = os - rc2;
                  return (
                    <div
                      style={{
                        background: C.bg,
                        borderRadius: 10,
                        padding: "14px 18px",
                        marginBottom: 14,
                        display: "flex",
                        flexWrap: "wrap",
                        gap: 20,
                        alignItems: "center",
                      }}
                    >
                      <div>
                        <div
                          style={{
                            fontSize: 9,
                            color: C.muted,
                            textTransform: "uppercase",
                            fontWeight: 600,
                          }}
                        >
                          Our Settlement
                        </div>
                        <div
                          style={{
                            fontSize: 16,
                            fontFamily: "'JetBrains Mono',monospace",
                            fontWeight: 700,
                            color: C.green,
                          }}
                        >
                          {formatINR(os)}
                        </div>
                      </div>
                      {rc2 > 0 && (
                        <>
                          <div>
                            <div
                              style={{
                                fontSize: 9,
                                color: C.muted,
                                textTransform: "uppercase",
                                fontWeight: 600,
                              }}
                            >
                              Bank Settlement
                            </div>
                            <div
                              style={{
                                fontSize: 16,
                                fontFamily: "'JetBrains Mono',monospace",
                                fontWeight: 700,
                                color: C.blue,
                              }}
                            >
                              {formatINR(rc2)}
                            </div>
                          </div>
                          <div>
                            <div
                              style={{
                                fontSize: 9,
                                color: C.muted,
                                textTransform: "uppercase",
                                fontWeight: 600,
                              }}
                            >
                              Difference
                            </div>
                            <div
                              style={{
                                fontSize: 16,
                                fontFamily: "'JetBrains Mono',monospace",
                                fontWeight: 700,
                                color:
                                  Math.abs(df) <= 1 ? C.green : C.red,
                              }}
                            >
                              {df > 0 ? "+" : ""}
                              {formatINR(df)}
                            </div>
                          </div>
                          <div
                            style={{
                              padding: "6px 14px",
                              borderRadius: 8,
                              fontSize: 12,
                              fontWeight: 700,
                              ...(Math.abs(df) <= 1
                                ? { background: "#d1fae5", color: "#065f46" }
                                : { background: "#fee2e2", color: "#991b1b" }),
                            }}
                          >
                            {Math.abs(df) <= 1 ? "MATCH" : "MISMATCH"}
                          </div>
                        </>
                      )}
                    </div>
                  );
                })()}

              <button
                onClick={addRecon}
                disabled={
                  !(
                    parseFloat(rPayin) > 0 &&
                    rMerchant &&
                    rCompany &&
                    rTxnId
                  )
                }
                style={{
                  background:
                    parseFloat(rPayin) > 0 &&
                    rMerchant &&
                    rCompany &&
                    rTxnId
                      ? C.accent
                      : "#e2e8f0",
                  color:
                    parseFloat(rPayin) > 0 &&
                    rMerchant &&
                    rCompany &&
                    rTxnId
                      ? "#fff"
                      : "#94a3b8",
                  border: "none",
                  borderRadius: 10,
                  padding: "12px 28px",
                  fontSize: 14,
                  fontWeight: 700,
                  cursor:
                    parseFloat(rPayin) > 0 &&
                    rMerchant &&
                    rCompany &&
                    rTxnId
                      ? "pointer"
                      : "not-allowed",
                }}
              >
                + Add to Reconciliation
              </button>
            </div>

            {recon.length > 0 && (
              <div
                style={{
                  background: C.card,
                  borderRadius: 14,
                  border: `1px solid ${C.border}`,
                  padding: "12px 20px",
                  marginBottom: 16,
                  display: "flex",
                  flexWrap: "wrap",
                  alignItems: "center",
                  gap: 10,
                }}
              >
                <div
                  style={{
                    display: "flex",
                    gap: 4,
                    background: C.bg,
                    borderRadius: 8,
                    padding: 3,
                  }}
                >
                  {[
                    { k: "all", l: "All" },
                    { k: "Matched", l: "Matched" },
                    { k: "Mismatch", l: "Mismatch" },
                    { k: "Pending", l: "Pending" },
                    { k: "Resolved", l: "Resolved" },
                  ].map((t) => (
                    <button
                      key={t.k}
                      onClick={() => setRFilter(t.k)}
                      style={{
                        background:
                          rFilter === t.k
                            ? t.k === "Mismatch"
                              ? C.red
                              : t.k === "Matched"
                                ? C.green
                                : t.k === "Resolved"
                                  ? C.blue
                                  : C.accent
                            : "transparent",
                        color: rFilter === t.k ? "#fff" : C.muted,
                        border: "none",
                        borderRadius: 6,
                        padding: "6px 14px",
                        fontSize: 11,
                        fontWeight: 600,
                        cursor: "pointer",
                      }}
                    >
                      {t.l}
                    </button>
                  ))}
                </div>
                <select
                  value={rFM}
                  onChange={(e) => setRFM(e.target.value)}
                  style={filtS}
                >
                  <option value="__all__">All Merchants</option>
                  {[...new Set(recon.map((r) => r.merchant))]
                    .sort()
                    .map((m) => (
                      <option key={m} value={m}>
                        {m}
                      </option>
                    ))}
                </select>
                <button
                  onClick={() => {
                    setRecon([]);
                    setRReason({});
                    setRNotes({});
                    setRResolved({});
                  }}
                  style={{
                    ...filtS,
                    border: "1px solid #fca5a5",
                    background: "#fef2f2",
                    color: C.red,
                    marginLeft: "auto",
                  }}
                >
                  Clear All
                </button>
              </div>
            )}

            {reconFiltered.length > 0 && (
              <div
                style={{
                  background: C.card,
                  borderRadius: 14,
                  border: `1px solid ${C.border}`,
                  overflow: "hidden",
                }}
              >
                <div
                  style={{
                    padding: "14px 20px",
                    borderBottom: `1px solid ${C.border}`,
                    fontSize: 14,
                    fontWeight: 700,
                    color: C.text,
                  }}
                >
                  Reconciliation Entries ({reconFiltered.length})
                </div>
                <div style={{ overflowX: "auto" }}>
                  <table
                    style={{
                      width: "100%",
                      borderCollapse: "collapse",
                      fontSize: 12,
                    }}
                  >
                    <thead>
                      <tr
                        style={{ borderBottom: `2px solid ${C.border}` }}
                      >
                        {[
                          "Txn ID",
                          "Merchant",
                          "Company",
                          "Payin",
                          "Our Settle",
                          "Bank Settle",
                          "Difference",
                          "Status",
                          "Reason",
                          "Action",
                        ].map((h, i) => (
                          <th
                            key={i}
                            style={{
                              padding: "11px 12px",
                              textAlign: i <= 2 ? "left" : "right",
                              fontSize: 9,
                              fontWeight: 700,
                              color: C.muted,
                              textTransform: "uppercase",
                              letterSpacing: ".5px",
                              whiteSpace: "nowrap",
                            }}
                          >
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {reconFiltered.map((r) => {
                        const badge = statusBadge(
                          r.status,
                          rResolved[r.id]
                        );
                        return (
                          <tr
                            key={r.id}
                            style={{
                              borderBottom: `1px solid ${C.border}`,
                            }}
                            onMouseEnter={(e) =>
                              (e.currentTarget.style.background =
                                "#f8fafc")
                            }
                            onMouseLeave={(e) =>
                              (e.currentTarget.style.background = "#fff")
                            }
                          >
                            <td
                              style={{
                                padding: "10px 12px",
                                fontFamily: "'JetBrains Mono',monospace",
                                fontWeight: 600,
                                color: C.text,
                                fontSize: 12,
                              }}
                            >
                              {r.txnId}
                            </td>
                            <td
                              style={{
                                padding: "10px 12px",
                                fontWeight: 600,
                                color: C.accent,
                              }}
                            >
                              {r.merchant}
                            </td>
                            <td
                              style={{
                                padding: "10px 12px",
                                color: C.text,
                                maxWidth: 160,
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                whiteSpace: "nowrap",
                              }}
                              title={r.company}
                            >
                              {r.company}
                            </td>
                            <td
                              style={{
                                padding: "10px 12px",
                                textAlign: "right",
                                fontFamily: "'JetBrains Mono',monospace",
                              }}
                            >
                              {formatINR(r.payin)}
                            </td>
                            <td
                              style={{
                                padding: "10px 12px",
                                textAlign: "right",
                                fontFamily: "'JetBrains Mono',monospace",
                                fontWeight: 600,
                                color: C.green,
                              }}
                            >
                              {formatINR(r.ourSettle)}
                            </td>
                            <td
                              style={{
                                padding: "10px 12px",
                                textAlign: "right",
                                fontFamily: "'JetBrains Mono',monospace",
                                fontWeight: 600,
                                color:
                                  r.claimed > 0 ? C.blue : C.muted,
                              }}
                            >
                              {r.claimed > 0
                                ? formatINR(r.claimed)
                                : "\u2014"}
                            </td>
                            <td
                              style={{
                                padding: "10px 12px",
                                textAlign: "right",
                                fontFamily: "'JetBrains Mono',monospace",
                                fontWeight: 700,
                                color:
                                  r.status === "Mismatch"
                                    ? C.red
                                    : r.status === "Matched"
                                      ? C.green
                                      : C.orange,
                              }}
                            >
                              {r.status === "Mismatch"
                                ? (r.diff > 0 ? "+" : "") +
                                  formatINR(r.diff)
                                : r.status === "Matched"
                                  ? "\u20B90.00"
                                  : "\u2014"}
                            </td>
                            <td
                              style={{
                                padding: "10px 12px",
                                textAlign: "right",
                              }}
                            >
                              <span
                                style={{
                                  padding: "3px 10px",
                                  borderRadius: 6,
                                  fontSize: 10,
                                  fontWeight: 700,
                                  background: badge.bg,
                                  color: badge.color,
                                }}
                              >
                                {badge.label}
                              </span>
                            </td>
                            <td
                              style={{
                                padding: "10px 12px",
                                textAlign: "right",
                              }}
                            >
                              {r.status === "Mismatch" &&
                                !rResolved[r.id] && (
                                  <select
                                    value={rReason[r.id] || ""}
                                    onChange={(e) =>
                                      setRReason((p) => ({
                                        ...p,
                                        [r.id]: e.target.value,
                                      }))
                                    }
                                    style={{
                                      padding: "4px 8px",
                                      borderRadius: 6,
                                      border: `1px solid ${C.border}`,
                                      fontSize: 11,
                                      color: C.text,
                                      cursor: "pointer",
                                      maxWidth: 120,
                                    }}
                                  >
                                    <option value="">Reason...</option>
                                    {REASONS.map((re) => (
                                      <option key={re} value={re}>
                                        {re}
                                      </option>
                                    ))}
                                  </select>
                                )}
                            </td>
                            <td
                              style={{
                                padding: "10px 8px",
                                textAlign: "center",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {!rResolved[r.id] &&
                                r.status === "Mismatch" &&
                                rReason[r.id] && (
                                  <button
                                    onClick={() =>
                                      setRResolved((p) => ({
                                        ...p,
                                        [r.id]: true,
                                      }))
                                    }
                                    style={{
                                      background: "#dbeafe",
                                      color: "#1e40af",
                                      border: "none",
                                      borderRadius: 6,
                                      padding: "4px 10px",
                                      fontSize: 10,
                                      fontWeight: 700,
                                      cursor: "pointer",
                                      marginRight: 4,
                                    }}
                                  >
                                    Resolve
                                  </button>
                                )}
                              <button
                                onClick={() => rmRecon(r.id)}
                                style={{
                                  background: "none",
                                  border: "none",
                                  color: "#cbd5e1",
                                  cursor: "pointer",
                                  fontSize: 14,
                                }}
                              >
                                &times;
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {recon.length === 0 && (
              <div
                style={{
                  background: C.card,
                  borderRadius: 14,
                  border: `1px solid ${C.border}`,
                  padding: "48px 24px",
                  textAlign: "center",
                }}
              >
                <div style={{ fontSize: 40, marginBottom: 12 }}>
                  {"\uD83D\uDD0D"}
                </div>
                <div
                  style={{
                    fontSize: 16,
                    fontWeight: 600,
                    color: C.text,
                    marginBottom: 6,
                  }}
                >
                  No reconciliation entries
                </div>
                <div style={{ fontSize: 13, color: C.muted }}>
                  Add transactions above to verify settlement amounts
                  against merchant claims
                </div>
              </div>
            )}
          </div>
        )}

        {/* ========== MERCHANTS ========== */}
        {page === "merchants" && (
          <div className="fade-up">
            {/* Top bar: search + summary */}
            <div
              style={{
                background: C.card,
                borderRadius: 14,
                border: `1px solid ${C.border}`,
                padding: "14px 20px",
                marginBottom: 16,
                display: "flex",
                flexWrap: "wrap",
                alignItems: "center",
                gap: 12,
              }}
            >
              <input
                value={mSearch}
                onChange={(e) => setMSearch(e.target.value)}
                placeholder="Search by company, MID, or GST number..."
                style={{
                  flex: 1,
                  minWidth: 240,
                  padding: "10px 14px",
                  borderRadius: 10,
                  border: `1.5px solid ${C.border}`,
                  fontSize: 13,
                  color: C.text,
                }}
              />
              <select
                value={mFilter}
                onChange={(e) => setMFilter(e.target.value)}
                style={filtS}
              >
                <option value="__all__">All Merchants ({MASTER_DATA.length})</option>
                {ALL_MERCHANTS.map((m) => (
                  <option key={m} value={m}>
                    {m} ({MASTER_DATA.filter((d) => d.merchant === m).length})
                  </option>
                ))}
              </select>
            </div>

            {/* Merchant cards */}
            <div
              style={{
                display: "grid",
                gridTemplateColumns:
                  "repeat(auto-fill,minmax(380px,1fr))",
                gap: 16,
              }}
            >
              {ALL_MERCHANTS.filter((m) =>
                mFilter === "__all__" ? true : m === mFilter
              ).map((m) => {
                const comps = MASTER_DATA.filter(
                  (d) => d.merchant === m
                ).filter((c) => {
                  if (!mSearch) return true;
                  const q = mSearch.toLowerCase();
                  return (
                    c.company.toLowerCase().includes(q) ||
                    (c.mid || "").toLowerCase().includes(q) ||
                    (c.gstNumber || "").toLowerCase().includes(q)
                  );
                });
                if (comps.length === 0) return null;
                const entries = rows.filter((r) => r.merchant === m);
                const mT = sumR(entries);
                return (
                  <div
                    key={m}
                    style={{
                      background: C.card,
                      borderRadius: 14,
                      border: `1px solid ${C.border}`,
                      overflow: "hidden",
                    }}
                  >
                    <div
                      style={{
                        padding: "14px 18px",
                        borderBottom: `1px solid ${C.border}`,
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        background: "linear-gradient(135deg,#f8fafc,#f0fdf4)",
                      }}
                    >
                      <div>
                        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                          <div
                            style={{
                              width: 32,
                              height: 32,
                              borderRadius: 8,
                              background: C.accent,
                              color: "#fff",
                              display: "flex",
                              alignItems: "center",
                              justifyContent: "center",
                              fontWeight: 700,
                              fontSize: 13,
                            }}
                          >
                            {m.substring(0, 2).toUpperCase()}
                          </div>
                          <div>
                            <div
                              style={{
                                fontSize: 15,
                                fontWeight: 700,
                                color: C.text,
                              }}
                            >
                              {m}
                            </div>
                            <div style={{ fontSize: 11, color: C.muted }}>
                              {comps.length} companies
                            </div>
                          </div>
                        </div>
                      </div>
                      {entries.length > 0 && (
                        <div style={{ textAlign: "right" }}>
                          <div style={{ fontSize: 10, color: C.muted, textTransform: "uppercase", fontWeight: 600 }}>
                            Settled
                          </div>
                          <div
                            style={{
                              fontSize: 15,
                              fontWeight: 700,
                              color: C.green,
                              fontFamily: "'JetBrains Mono',monospace",
                            }}
                          >
                            {formatShort(mT.settlement)}
                          </div>
                        </div>
                      )}
                    </div>
                    <div
                      style={{
                        maxHeight: 340,
                        overflowY: "auto",
                      }}
                    >
                      {comps.map((c, i) => (
                        <div
                          key={i}
                          style={{
                            padding: "10px 16px",
                            borderBottom: `1px solid ${C.border}`,
                          }}
                          onMouseEnter={(e) => (e.currentTarget.style.background = "#f8fafc")}
                          onMouseLeave={(e) => (e.currentTarget.style.background = "#fff")}
                        >
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div
                                style={{
                                  fontSize: 12,
                                  fontWeight: 600,
                                  color: C.text,
                                  overflow: "hidden",
                                  textOverflow: "ellipsis",
                                  whiteSpace: "nowrap",
                                }}
                                title={c.company}
                              >
                                {c.company}
                              </div>
                              <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 4, alignItems: "center" }}>
                                {c.mid && (
                                  <span
                                    style={{
                                      fontSize: 10,
                                      fontFamily: "'JetBrains Mono',monospace",
                                      color: C.muted,
                                      background: C.bg,
                                      padding: "2px 6px",
                                      borderRadius: 4,
                                    }}
                                    title="Merchant ID"
                                  >
                                    {c.mid}
                                  </span>
                                )}
                                {c.gstNumber && (
                                  <span
                                    style={{
                                      fontSize: 10,
                                      fontFamily: "'JetBrains Mono',monospace",
                                      color: C.blue,
                                    }}
                                    title="GST Number"
                                  >
                                    {c.gstNumber}
                                  </span>
                                )}
                                {c.incorporation && (
                                  <span style={{ fontSize: 10, color: C.muted }} title="Incorporation date">
                                    {"\uD83D\uDCC5 "}
                                    {c.incorporation}
                                  </span>
                                )}
                              </div>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
