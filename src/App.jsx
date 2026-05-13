import { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";

const CHARGE_RATE = 0.35;
const GST_RATE = 28;
const GST_OPTIONS = [5, 12, 18, 28];

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
  { mid: "MER0000000030988", merchant: "Aryan", company: "SIREPLES INDIA PRIVATE LIMITED", moa: "DONE", incorporation: "27-02-2025", gstNumber: "08ABPCS3115B1ZJ" },
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
  { mid: "MER0000000031064", merchant: "Shashi", company: "Felunor Supplier Private Limited", moa: "DONE", incorporation: "17-09-2025", gstNumber: "07AAGCF6262L1ZA" },
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
  { mid: "", merchant: "Nilesh", company: "ONESTOP SHOPPING STATION PVT LTD", moa: "", incorporation: "02-12-2023", gstNumber: "24AAECO2407J1ZO" },
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
  { mid: "MER0000000031153", merchant: "AK", company: "JOVLERA TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "06-12-2025", gstNumber: "07AAHCJ1595M1Z1" },
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
  { mid: "MER0000000031190", merchant: "KJ", company: "VISINORYVISTA TECH PRIAVTE LIMITED", moa: "DONE", incorporation: "15-01-2025", gstNumber: "27AAKCV6794R1ZZ" },
  { mid: "MER0000000031205", merchant: "KJ", company: "DULAARAA PUBLICITY AND MEDIA PRIVATE LIMITED", moa: "", incorporation: "29-08-2024", gstNumber: "09AALCD0872F1ZK" },
  { mid: "", merchant: "KJ", company: "MATILEO ENTERPRISES PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "" },
  { mid: "", merchant: "KJ", company: "WENSET INFOCOM PRIAVTE LIMITED", moa: "", incorporation: "", gstNumber: "" },
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
  { mid: "", merchant: "AJ", company: "PHENOX AEROSPACE INDIA PRIVATE LIMITED", moa: "", incorporation: "10-04-2024", gstNumber: "07AAOCP6508R1ZJ" },
  { mid: "MER0000000031274", merchant: "Nilesh", company: "BARAI AND BISWAS TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "23-02-2025", gstNumber: "19AANCB2729H1ZF" },
  { mid: "", merchant: "KJ", company: "PVKA SOFTWARE PRIVATE LIMITED", moa: "", incorporation: "08-02-2025", gstNumber: "08AAPCP4233N1ZU" },
  { mid: "MER0000000031281", merchant: "AJ", company: "PIONEERS HUB PRIVATE LIMITED", moa: "", incorporation: "23-12-2025", gstNumber: "27AAQCP5030Q1ZQ" },
  { mid: "", merchant: "AJ", company: "GROSMART VENTURES PRIVATE LIMITED", moa: "", incorporation: "09-12-2025", gstNumber: "07AAMCG5648P1ZR" },
  { mid: "", merchant: "AJ", company: "TRUAXIS VENTURES PRIVATE LIMITED", moa: "", incorporation: "15-01-2026", gstNumber: "07AAMCT6963A1Z5" },
  { mid: "", merchant: "Nilesh", company: "INTCOM TECHNOLOGY PRIVATE LIMITED", moa: "", incorporation: "30-05-2025", gstNumber: "29AAICI2425E1ZP" },
  { mid: "", merchant: "Nilesh", company: "DIPRATA TECHONOLOGY PRIVATE LIMITED", moa: "", incorporation: "25-07-2024", gstNumber: "19AALCD0162E1ZU" },
  { mid: "", merchant: "Nilesh", company: "KSHIRA FIN PRIVATE LIMITED", moa: "", incorporation: "", gstNumber: "36AALCK1084E1ZK" },
  { mid: "", merchant: "AJ", company: "BARRINGER PHARMA PRIVATE LIMITED", moa: "", incorporation: "23-04-2018", gstNumber: "07AAHCB7897G1Z5" },
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
        Payin: r.payin,
        "Merchant Fee": r.charge,
        GST: r.gst,
        "Merchant Settle Amount": merchantSettle,
        "Chargeback Received": r.chargeback || 0,
        "Net Merchant Settlement": r.settlement,
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
      Payin: totals.payin,
      "Merchant Fee": totals.charge,
      GST: totals.gst,
      "Merchant Settle Amount": totals.merchantSettle,
      "Chargeback Received": totals.chargeback,
      "Net Merchant Settlement": totals.settlement,
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
