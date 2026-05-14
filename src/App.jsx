import { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import { Card, CardHeader, CardTitle, CardDescription, CardContent } from "@/components/ui/card";
import { Label } from "@/components/ui/label";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Separator } from "@/components/ui/separator";
import { Badge } from "@/components/ui/badge";
import { Tabs, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { cn } from "@/lib/utils";
import {
  Plus,
  X,
  Calendar as CalendarIcon,
  Building2,
  IndianRupee,
  Sun,
  Moon,
  Receipt,
  Download,
  Trash2,
  Upload,
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle,
  Search,
  Users,
  ArrowRight,
  Menu,
  Zap,
  Home,
  PlusCircle,
  BarChart3,
  ClipboardCheck,
  Building,
  Sparkles,
  ChevronsLeft,
  ChevronsRight,
  Activity as ActivityIcon,
} from "lucide-react";
import Automation from "@/Automation.jsx";
import Login from "@/Login.jsx";
import { api as backendApi, getStoredAuth } from "@/lib/api.js";
import { LogOut } from "lucide-react";

const CHARGE_RATE = 0.41;
const GST_RATE = 0;
const GST_OPTIONS = [0, 5, 12, 18, 28];

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
  bg: "#f6f5fb",
  sidebar: "#1d1535",
  sideH: "#2a1f4a",
  sideA: "#3a2a64",
  accent: "#7c3aed",
  accentL: "#ede9fe",
  card: "#ffffff",
  text: "#15131e",
  muted: "#6b6480",
  border: "#e8e4ee",
  green: "#10b981",
  orange: "#f59e0b",
  red: "#ef4444",
  blue: "#3b82f6",
};

export default function App() {
  // ----- Auth gate -----
  const [auth, setAuth] = useState(() => getStoredAuth());

  useEffect(() => {
    function onExpired() {
      setAuth(null);
    }
    window.addEventListener("settleops:auth-expired", onExpired);
    return () => window.removeEventListener("settleops:auth-expired", onExpired);
  }, []);

  if (!auth?.access_token) {
    return <Login onSuccess={(r) => setAuth(r)} />;
  }

  return <AuthenticatedApp user={auth.user} onLogout={() => { backendApi.logout(); setAuth(null); }} />;
}

function AuthenticatedApp({ user, onLogout }) {
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
  const [bankFile, setBankFile] = useState("");
  const [bankRows, setBankRows] = useState([]); // raw rows per company from bank file
  const [bankError, setBankError] = useState("");
  const [bankDragging, setBankDragging] = useState(false);
  // Source/internal payin upload state (just Name + Sum of Amount per company)
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
  const [mobileNav, setMobileNav] = useState(false);
  const [backendUp, setBackendUp] = useState(null); // null=checking, true/false=known
  const [serverStats, setServerStats] = useState(null); // backend /api/dashboard/stats snapshot

  // Probe backend every 30s so the header pill + nav badge update live
  useEffect(() => {
    let cancelled = false;
    async function probe() {
      try {
        await backendApi.health();
        if (!cancelled) setBackendUp(true);
      } catch {
        if (!cancelled) setBackendUp(false);
      }
    }
    probe();
    const id = setInterval(probe, 30_000);
    return () => {
      cancelled = true;
      clearInterval(id);
    };
  }, []);

  // When on the Dashboard, fetch the backend snapshot (if backend is up)
  useEffect(() => {
    if (!backendUp) {
      setServerStats(null);
      return;
    }
    let cancelled = false;
    async function load() {
      try {
        const stats = await backendApi.dashboard();
        if (!cancelled) setServerStats(stats);
      } catch {
        if (!cancelled) setServerStats(null);
      }
    }
    load();
  }, [backendUp, page]);

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
    .replace(/\b(priv\w*|pvt\w*|ltd\w*|limited|limte\w*|opc|llp)\.?\b/g, "")
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

  const parseBankFile = useCallback((file) => {
    setBankError("");
    setBankRows([]);
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["csv", "xlsx", "xls"].includes(ext)) {
      setBankError("Unsupported file type. Please upload a CSV or Excel file.");
      return;
    }
    const reader = new FileReader();
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
          const amount = parseAmount(row[mapping.amount]) || 0;
          const fee = parseAmount(row[mapping.fee]) || 0;
          const gst = parseAmount(row[mapping.gst]) || 0;
          const settle = parseAmount(row[mapping.settle]) || (amount - fee - gst);
          const chargeback = mapping.chargeback != null ? (parseAmount(row[mapping.chargeback]) || 0) : 0;
          const netSettle = mapping.netSettle != null
            ? (parseAmount(row[mapping.netSettle]) || (settle - chargeback))
            : (settle - chargeback);
          if (amount <= 0 && settle <= 0) continue;
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
          });
        }
        if (out.length === 0) {
          setBankError("No data rows found in the file.");
          return;
        }
        setBankRows(out);
        setBankFile(file.name);
      } catch (err) {
        setBankError("Failed to parse file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleBankDrop = useCallback((e) => {
    e.preventDefault();
    setBankDragging(false);
    const file = e.dataTransfer?.files?.[0] || e.target?.files?.[0];
    if (file) parseBankFile(file);
  }, [parseBankFile]);

  const handleBankInput = useCallback((e) => {
    const f = e.target.files?.[0];
    if (f) parseBankFile(f);
  }, [parseBankFile]);

  const clearBank = () => {
    setBankRows([]);
    setBankFile("");
    setBankError("");
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

  const downloadBankRecon = () => {
    if (combinedByMerchant.length === 0) return;
    const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
    const feeLabel = `Our Fee (${CHARGE_RATE}%)`;
    const gstLabel = `Our GST (${GST_RATE}%)`;
    const data = combinedByMerchant.map((g, i) => ({
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
      Companies: bankTotals.companies,
      "Txn Count": bankTotals.count,
      "Source Payin": r2(combinedTotals.sourcePayin),
      "Bank Payin": r2(combinedTotals.bankPayin),
      "Payin Diff (Source − Bank)": r2(combinedTotals.payinDiff),
      "Bank Fee": r2(bankTotals.fee),
      "Bank GST": r2(bankTotals.gst),
      "Bank Settlement": r2(bankTotals.settle),
      "Chargeback Received": r2(bankTotals.chargeback),
      "Net Merchant Settlement": r2(bankTotals.netSettle),
      [feeLabel]: r2(bankTotals.amount * CHARGE_RATE / 100),
      [gstLabel]: r2((bankTotals.amount * CHARGE_RATE / 100) * GST_RATE / 100),
      "Our Settlement": r2(bankTotals.ourSettle),
      "Our Net (after CB)": r2(bankTotals.ourNet),
      "Difference (Bank Net − Our Net)": r2(bankTotals.diff),
      "Unmatched Companies": bankTotals.unmatched,
    });
    const ws = XLSX.utils.json_to_sheet(data);
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
    const moneyCols = ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q"];
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
      const dRef = "D" + row;
      if (ws[dRef] && typeof ws[dRef].v === "number") {
        ws[dRef].z = "#,##0";
        ws[dRef].t = "n";
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Summary");

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

    combinedByMerchant.forEach((g) => {
      // Merge per-company rows from bank + source files under this merchant
      const compMap = new Map();
      bankRows.filter((b) => (b.merchant || "Unmatched") === g.merchant).forEach((b) => {
        const key = (b.mid || normCompany(b.name)) || b.name;
        compMap.set(key, {
          mid: b.mid || "",
          name: b.name,
          bankCount: b.count,
          bankPayin: b.amount,
          bankFee: b.fee,
          bankGst: b.gst,
          bankSettle: b.settle,
          chargeback: b.chargeback || 0,
          netSettle: b.netSettle || (b.settle - (b.chargeback || 0)),
          sourcePayin: 0,
        });
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

      const detailWs = XLSX.utils.json_to_sheet(detailData);
      detailWs["!cols"] = [
        { wch: 6 }, { wch: 18 }, { wch: 42 }, { wch: 16 }, { wch: 16 }, { wch: 16 },
        { wch: 10 }, { wch: 14 }, { wch: 12 }, { wch: 16 }, { wch: 14 }, { wch: 16 },
      ];
      // Apply 2-decimal format to money columns (D..F, H..L)
      const detailMoney = ["D", "E", "F", "H", "I", "J", "K", "L"];
      for (let row = 2; row <= detailData.length + 1; row++) {
        for (const col of detailMoney) {
          const cell = detailWs[col + row];
          if (cell && typeof cell.v === "number") { cell.z = numFmt; cell.t = "n"; }
        }
        const cnt = detailWs["G" + row];
        if (cnt && typeof cnt.v === "number") { cnt.z = "#,##0"; cnt.t = "n"; }
      }
      XLSX.utils.book_append_sheet(wb, detailWs, uniqueName(g.merchant));
    });

    XLSX.writeFile(wb, `bank_recon_${new Date().toISOString().slice(0, 10)}.xlsx`);
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

  const navSections = [
    {
      label: "Workspace",
      items: [
        { id: "dashboard", label: "Dashboard", Icon: Home },
        { id: "add", label: "Add Entry", Icon: PlusCircle },
        { id: "reports", label: "Reports", Icon: BarChart3 },
      ],
    },
    {
      label: "Reconcile",
      items: [
        { id: "recon", label: "Reconciliation", Icon: ClipboardCheck },
        { id: "merchants", label: "Merchants", Icon: Building },
      ],
    },
    {
      label: "Tools",
      items: [
        { id: "automation", label: "Automation", Icon: Sparkles },
      ],
    },
  ];
  const navItems = navSections.flatMap((s) => s.items); // legacy: page-title lookup still uses this

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

  const expanded = sideOpen || mobileNav;

  return (
    <div
      style={{
        display: "flex",
        height: "100vh",
        fontFamily:
          "'Saviom','Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif",
        background: C.bg,
        overflow: "hidden",
        letterSpacing: "-0.01em",
      }}
    >
      {/* Mobile backdrop */}
      {mobileNav && (
        <div
          onClick={() => setMobileNav(false)}
          className="fixed inset-0 z-40 bg-black/55 backdrop-blur-[2px] md:hidden"
        />
      )}

      {/* Sidebar */}
      <aside
        className={cn(
          "fixed inset-y-0 left-0 z-50 flex flex-col overflow-hidden border-r border-white/[0.06] md:static md:translate-x-0",
          mobileNav ? "translate-x-0" : "-translate-x-full md:translate-x-0"
        )}
        style={{
          width: expanded ? 244 : 72,
          background:
            "radial-gradient(120% 80% at 0% 0%, rgba(167,139,250,0.10), transparent 60%)," +
            "radial-gradient(80% 60% at 100% 100%, rgba(124,58,237,0.08), transparent 70%)," +
            "linear-gradient(180deg, #15102a 0%, #0d0a1d 100%)",
          transition: "width .25s ease, transform .25s ease",
        }}
      >
        {/* Brand */}
        <div
          className={cn(
            "relative flex items-center px-4 pt-5 pb-4",
            expanded ? "justify-between" : "justify-center"
          )}
        >
          <div
            className={cn(
              "flex min-w-0 items-center gap-3 cursor-pointer",
              expanded && "flex-1"
            )}
            onClick={() => {
              if (mobileNav) setMobileNav(false);
              else if (!sideOpen) setSO(true);
            }}
            title={!expanded ? "Expand" : undefined}
          >
            <div
              className="flex h-9 w-9 flex-shrink-0 items-center justify-center rounded-lg text-white font-extrabold relative overflow-hidden"
              style={{
                background: "linear-gradient(135deg,#a78bfa 0%,#8b5cf6 45%,#6d28d9 100%)",
                boxShadow:
                  "0 6px 16px rgba(124,58,237,.45), inset 0 1px 0 rgba(255,255,255,.35), inset 0 -1px 0 rgba(0,0,0,.2)",
                fontSize: 15,
                letterSpacing: "-0.5px",
              }}
            >
              <span className="relative z-10">S</span>
              <div
                className="pointer-events-none absolute inset-0"
                style={{
                  background:
                    "radial-gradient(60% 50% at 30% 20%, rgba(255,255,255,.4), transparent 70%)",
                }}
              />
            </div>
            {expanded && (
              <div className="flex min-w-0 flex-col leading-tight">
                <span className="truncate text-[15px] font-bold text-white tracking-tight">
                  SettleOps
                </span>
                <span className="mt-0.5 flex items-center gap-1.5 text-[9.5px] font-semibold uppercase tracking-[.14em] text-white/40">
                  <span className="h-1 w-1 rounded-full bg-emerald-400 shadow-[0_0_4px_rgba(52,211,153,.8)]" />
                  PRO
                </span>
              </div>
            )}
          </div>
        </div>

        {/* Always-visible collapse/expand pill on the right edge — never gets clipped */}
        <button
          type="button"
          onClick={() => setSO(!sideOpen)}
          title={expanded ? "Collapse sidebar" : "Expand sidebar"}
          className={cn(
            "absolute z-10 hidden h-6 w-6 items-center justify-center rounded-full border border-white/10 bg-[#1a1334] text-white/60 shadow-md transition-all hover:bg-violet-500/30 hover:text-white md:flex",
            "top-7"
          )}
          style={{ right: -12 }}
        >
          {expanded ? <ChevronsLeft className="h-3 w-3" /> : <ChevronsRight className="h-3 w-3" />}
        </button>

        {/* Divider */}
        <div className="mx-4 h-px bg-gradient-to-r from-transparent via-white/[0.06] to-transparent" />

        {/* Sectioned nav */}
        <nav className="flex-1 overflow-y-auto px-3 py-3">
          {navSections.map((section, sIdx) => (
            <div key={section.label} className={cn(sIdx > 0 && "mt-4")}>
              {expanded && (
                <div className="mb-1.5 px-3 text-[9.5px] font-semibold uppercase tracking-[.14em] text-white/35">
                  {section.label}
                </div>
              )}
              {section.items.map((item) => {
                const a = page === item.id;
                const Icon = item.Icon;
                const showBadge = item.id === "recon" && reconStats.mismatch > 0;
                return (
                  <button
                    key={item.id}
                    type="button"
                    onClick={() => {
                      setPage(item.id);
                      setMobileNav(false);
                    }}
                    title={!expanded ? item.label : undefined}
                    className={cn(
                      "group relative mb-0.5 flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-left transition-all duration-150",
                      a
                        ? "text-white"
                        : "text-white/60 hover:bg-white/[0.05] hover:text-white"
                    )}
                    style={
                      a
                        ? {
                            background:
                              "linear-gradient(90deg, rgba(167,139,250,0.22) 0%, rgba(124,58,237,0.10) 100%)",
                            boxShadow:
                              "inset 0 0 0 1px rgba(167,139,250,0.28), 0 1px 0 rgba(0,0,0,0.2)",
                          }
                        : undefined
                    }
                  >
                    {a && (
                      <span
                        className="absolute -left-3 top-1/2 h-5 w-[3px] -translate-y-1/2 rounded-r"
                        style={{
                          background: "linear-gradient(180deg,#c4b5fd,#7c3aed)",
                          boxShadow: "0 0 8px rgba(167,139,250,.7)",
                        }}
                      />
                    )}
                    <Icon
                      className={cn(
                        "h-[18px] w-[18px] flex-shrink-0 transition-colors",
                        a ? "text-violet-200" : "text-white/55 group-hover:text-white"
                      )}
                      strokeWidth={a ? 2.2 : 1.8}
                    />
                    {expanded && (
                      <span
                        className={cn(
                          "flex-1 truncate text-[13.5px] tracking-tight",
                          a ? "font-semibold" : "font-medium"
                        )}
                      >
                        {item.label}
                      </span>
                    )}
                    {showBadge && expanded && (
                      <span
                        className="ml-auto rounded-md px-1.5 py-0.5 text-[10px] font-bold tabular-nums text-white"
                        style={{
                          background: "linear-gradient(135deg,#f43f5e,#e11d48)",
                          boxShadow: "0 2px 6px rgba(244,63,94,.45)",
                        }}
                      >
                        {reconStats.mismatch}
                      </span>
                    )}
                    {showBadge && !expanded && (
                      <span
                        className="absolute right-1.5 top-1.5 h-2 w-2 rounded-full"
                        style={{
                          background: "#f43f5e",
                          boxShadow: "0 0 6px rgba(244,63,94,.7)",
                        }}
                      />
                    )}
                  </button>
                );
              })}
            </div>
          ))}
        </nav>

        {/* Cycle status card */}
        <div className="px-3 pt-2">
          {expanded ? (
            <div
              className="relative overflow-hidden rounded-lg p-3"
              style={{
                background:
                  "linear-gradient(135deg, rgba(255,255,255,0.04) 0%, rgba(167,139,250,0.06) 100%)",
                boxShadow:
                  "inset 0 0 0 1px rgba(255,255,255,0.06), 0 4px 12px rgba(0,0,0,0.25)",
              }}
            >
              <div
                className="pointer-events-none absolute -right-4 -top-4 h-14 w-14 rounded-full blur-2xl"
                style={{ background: `radial-gradient(circle, ${cyc.color}45, transparent 70%)` }}
              />
              <div className="relative">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-1.5">
                    <span
                      className="inline-block h-1.5 w-1.5 flex-shrink-0 rounded-full"
                      style={{
                        background: cyc.color,
                        boxShadow: `0 0 6px ${cyc.color}, 0 0 2px ${cyc.color}`,
                        animation: "pulse 2s infinite",
                      }}
                    />
                    <span className="text-[9px] font-bold uppercase tracking-[.18em] text-white/55">
                      Active cycle
                    </span>
                  </div>
                  <span
                    className="rounded px-1.5 py-0.5 font-mono text-[10px] font-bold tabular-nums"
                    style={{ background: `${cyc.color}26`, color: cyc.color }}
                  >
                    C{cyc.cycle}
                  </span>
                </div>
                <div className="mt-2 text-[11px] font-medium tracking-tight text-white/65">
                  {cyc.label}
                </div>
                <div
                  className="mt-1 font-mono text-[14px] font-bold leading-none text-white tabular-nums"
                  style={{ letterSpacing: "0.5px" }}
                >
                  <LiveClock />
                </div>
              </div>
            </div>
          ) : (
            <div
              className="relative flex flex-col items-center gap-1.5 rounded-lg px-1 py-3"
              style={{
                background:
                  "linear-gradient(180deg, rgba(255,255,255,0.04) 0%, rgba(167,139,250,0.04) 100%)",
                boxShadow: "inset 0 0 0 1px rgba(255,255,255,0.06)",
              }}
              title={`Cycle ${cyc.cycle} · ${cyc.label}`}
            >
              <span
                className="h-2 w-2 rounded-full"
                style={{
                  background: cyc.color,
                  boxShadow: `0 0 8px ${cyc.color}, 0 0 2px ${cyc.color}`,
                  animation: "pulse 2s infinite",
                }}
              />
              <span
                className="rounded-md px-1.5 py-0.5 font-mono text-[10px] font-bold tabular-nums"
                style={{ background: `${cyc.color}26`, color: cyc.color }}
              >
                C{cyc.cycle}
              </span>
            </div>
          )}
        </div>

        {/* Account row — at the very bottom now */}
        <div className="mx-3 my-3 h-px bg-gradient-to-r from-transparent via-white/[0.08] to-transparent" />
        <div className="px-3 pb-4">
          {expanded ? (
            <div
              className="group flex items-center gap-2.5 rounded-lg px-2.5 py-2 transition-colors hover:bg-white/[0.05]"
              style={{
                background: "linear-gradient(135deg,rgba(255,255,255,0.02),rgba(167,139,250,0.04))",
                boxShadow: "inset 0 0 0 1px rgba(255,255,255,0.04)",
              }}
            >
              <div
                className="relative flex h-8 w-8 flex-shrink-0 items-center justify-center overflow-hidden rounded-md text-[12px] font-bold uppercase text-white"
                style={{
                  background: "linear-gradient(135deg,#a78bfa 0%,#7c3aed 100%)",
                  boxShadow:
                    "inset 0 1px 0 rgba(255,255,255,.3), 0 2px 6px rgba(124,58,237,.35)",
                }}
                title={user?.username}
              >
                <span className="relative z-10">{(user?.username || "?").slice(0, 1)}</span>
                {/* Online indicator */}
                <span
                  className="absolute -bottom-0.5 -right-0.5 h-2.5 w-2.5 rounded-full border-2 border-[#15102a] bg-emerald-400"
                  style={{ boxShadow: "0 0 4px rgba(52,211,153,.7)" }}
                />
              </div>
              <div className="min-w-0 flex-1 leading-tight">
                <div className="truncate text-[12.5px] font-semibold text-white">
                  {user?.username}
                </div>
                <div className="flex items-center gap-1 text-[9.5px] font-medium text-emerald-300/80">
                  <span className="h-1 w-1 rounded-full bg-emerald-400" />
                  Signed in
                </div>
              </div>
              <button
                type="button"
                onClick={onLogout}
                title="Sign out"
                aria-label="Sign out"
                className="flex h-7 w-7 flex-shrink-0 items-center justify-center rounded-md text-white/45 transition-all hover:bg-red-500/15 hover:text-red-300"
              >
                <LogOut className="h-3.5 w-3.5" />
              </button>
            </div>
          ) : (
            /* Collapsed: avatar with online dot on top, logout below */
            <div className="flex flex-col items-center gap-1.5">
              <div
                className="relative flex h-8 w-8 items-center justify-center overflow-hidden rounded-md text-[12px] font-bold uppercase text-white"
                style={{
                  background: "linear-gradient(135deg,#a78bfa 0%,#7c3aed 100%)",
                  boxShadow:
                    "inset 0 1px 0 rgba(255,255,255,.3), 0 2px 6px rgba(124,58,237,.35)",
                }}
                title={user?.username}
              >
                <span className="relative z-10">{(user?.username || "?").slice(0, 1)}</span>
                <span
                  className="absolute -bottom-0.5 -right-0.5 h-2 w-2 rounded-full border border-[#15102a] bg-emerald-400"
                  style={{ boxShadow: "0 0 3px rgba(52,211,153,.7)" }}
                />
              </div>
              <button
                type="button"
                onClick={onLogout}
                title={`Sign out ${user?.username || ""}`}
                aria-label="Sign out"
                className="flex h-6 w-6 items-center justify-center rounded-md text-white/45 transition-all hover:bg-red-500/15 hover:text-red-300"
              >
                <LogOut className="h-3 w-3" />
              </button>
            </div>
          )}
        </div>
      </aside>

      {/* Main */}
      <div className="flex-1 overflow-auto px-4 py-5 md:px-9 md:py-8">
        <div className="mb-6 flex flex-wrap items-start justify-between gap-3 border-b border-border pb-4 md:mb-7 md:items-end md:pb-[18px]">
          <div className="flex min-w-0 items-start gap-3">
            <button
              type="button"
              onClick={() => setMobileNav(true)}
              className="mt-0.5 flex h-9 w-9 flex-shrink-0 items-center justify-center rounded-md border border-border bg-card text-foreground shadow-xs transition-colors hover:bg-muted md:hidden"
              aria-label="Open menu"
            >
              <Menu className="h-5 w-5" />
            </button>
            <div className="min-w-0">
              <div className="mb-1.5 text-[10.5px] font-semibold uppercase tracking-[.1em] text-muted-foreground md:text-[11px]">
                SettleOps · {navItems.find((n) => n.id === page)?.label || "Dashboard"}
              </div>
              <h1
                className="text-2xl font-bold leading-[1.1] tracking-tight text-foreground md:text-[28px]"
                style={{ letterSpacing: "-0.8px" }}
              >
                {navItems.find((n) => n.id === page)?.label || "Dashboard"}
              </h1>
              <p className="mt-1.5 text-xs text-muted-foreground md:text-[13px]">
                Charge {CHARGE_RATE}% + GST {page === "add" ? gstRate : GST_RATE}% applied on charge
              </p>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              onClick={() => setPage("automation")}
              title={
                backendUp === true
                  ? "Backend connected — click to open Automation"
                  : backendUp === false
                  ? "Backend offline — click for setup"
                  : "Checking backend…"
              }
              className={cn(
                "flex items-center gap-1.5 rounded-full border px-2.5 py-1 text-[10px] font-semibold tracking-wide transition-colors md:text-[11px]",
                backendUp === true && "border-emerald-200 bg-emerald-50 text-emerald-700 hover:bg-emerald-100",
                backendUp === false && "border-red-200 bg-red-50 text-red-700 hover:bg-red-100",
                backendUp === null && "border-border bg-muted/60 text-muted-foreground"
              )}
            >
              <span
                className={cn(
                  "inline-block h-1.5 w-1.5 rounded-full",
                  backendUp === true && "bg-emerald-500 shadow-[0_0_6px_rgba(16,185,129,.7)]",
                  backendUp === false && "bg-red-500",
                  backendUp === null && "bg-slate-400 animate-pulse"
                )}
              />
              API
            </button>
            <div
              className="flex items-center gap-2 rounded-full border px-3 py-1.5 md:px-3.5 md:py-2"
              style={{ background: C.accentL, borderColor: `${C.accent}22` }}
            >
              <span
                className="h-1.5 w-1.5 rounded-full"
                style={{ background: cyc.color, boxShadow: `0 0 6px ${cyc.color}` }}
              />
              <span
                className="text-[10.5px] font-semibold tracking-wide md:text-[11.5px]"
                style={{ color: C.accent }}
              >
                <span className="hidden sm:inline">Cycle </span>C{cyc.cycle} · {cyc.label}
              </span>
            </div>
          </div>
        </div>

        {/* ========== DASHBOARD ========== */}
        {page === "dashboard" && (
          <div className="fade-up space-y-6">
            <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))" }}>
              {[
                {
                  label: "Total Payin",
                  value: formatShort(allT.payin),
                  sub: `${rows.length} entries`,
                  tone: "violet",
                },
                {
                  label: "Total Charges",
                  value: formatShort(allT.charge),
                  sub: `${CHARGE_RATE}%`,
                  tone: "amber",
                },
                {
                  label: "Total GST",
                  value: formatShort(allT.gst),
                  sub: `${GST_RATE}% on charge`,
                  tone: "rose",
                },
                {
                  label: "Settlement",
                  value: formatShort(allT.settlement),
                  sub:
                    rows.length > 0
                      ? `${((allT.settlement / allT.payin) * 100).toFixed(2)}% of payin`
                      : "\u2014",
                  tone: "emerald",
                },
                {
                  label: "Recon Issues",
                  value: reconStats.mismatch,
                  sub: `${reconStats.total} checked`,
                  tone: reconStats.mismatch > 0 ? "red" : "emerald",
                },
              ].map((s, i) => {
                const toneMap = {
                  violet: { value: "text-violet-700", bg: "from-violet-200/60 to-violet-100/0", ring: "ring-violet-200" },
                  amber: { value: "text-amber-700", bg: "from-amber-200/60 to-amber-100/0", ring: "ring-amber-200" },
                  rose: { value: "text-rose-700", bg: "from-rose-200/60 to-rose-100/0", ring: "ring-rose-200" },
                  emerald: { value: "text-emerald-700", bg: "from-emerald-200/60 to-emerald-100/0", ring: "ring-emerald-200" },
                  red: { value: "text-red-700", bg: "from-red-200/60 to-red-100/0", ring: "ring-red-200" },
                }[s.tone];
                return (
                  <Card key={i} className="relative overflow-hidden border-border/70 shadow-xs">
                    <div
                      className={cn(
                        "pointer-events-none absolute -right-10 -top-10 h-32 w-32 rounded-full bg-gradient-to-br blur-2xl",
                        toneMap.bg
                      )}
                    />
                    <CardContent className="relative px-5 py-5">
                      <div className="mb-2 text-[10.5px] font-semibold uppercase tracking-[.08em] text-muted-foreground">
                        {s.label}
                      </div>
                      <div className={cn("font-mono text-[26px] font-bold leading-none tabular-nums", toneMap.value)}>
                        {s.value}
                      </div>
                      <div className="mt-2 text-xs text-muted-foreground">
                        {s.sub}
                      </div>
                    </CardContent>
                  </Card>
                );
              })}
            </div>

            {/* Secondary insight cards */}
            {(() => {
              const todayStr = new Date().toISOString().slice(0, 10);
              const todayRows = rows.filter((r) => (r.date || "").slice(0, 10) === todayStr);
              const todayPayin = todayRows.reduce((s, r) => s + (r.payin || 0), 0);
              const c1 = rows.filter((r) => r.cycle === 1);
              const c2 = rows.filter((r) => r.cycle === 2);
              const c1Sum = c1.reduce((s, r) => s + (r.payin || 0), 0);
              const c2Sum = c2.reduce((s, r) => s + (r.payin || 0), 0);
              const totalCnt = c1.length + c2.length || 1;
              const c1Pct = (c1.length / totalCnt) * 100;
              const merchantCount = serverStats?.coverage?.merchant_count ?? new Set(MASTER_DATA.map((d) => d.merchant)).size;
              const companyCount = serverStats?.coverage?.company_count ?? MASTER_DATA.length;
              const reconTotal = reconStats.total || 1;
              return (
                <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
                  {/* Today */}
                  {(() => {
                    const todaySettle = todayRows.reduce((s, r) => s + (r.settlement || 0), 0);
                    const c1Today = todayRows.filter((r) => r.cycle === 1).length;
                    const c2Today = todayRows.filter((r) => r.cycle === 2).length;
                    // 7-day rolling
                    const sevenAgo = new Date();
                    sevenAgo.setDate(sevenAgo.getDate() - 7);
                    const week = rows.filter((r) => new Date(r.date || 0) >= sevenAgo);
                    const weekAvg = week.length > 0 ? week.length / 7 : 0;
                    // Top company today
                    const byCo = {};
                    todayRows.forEach((r) => { byCo[r.company] = (byCo[r.company] || 0) + (r.payin || 0); });
                    const topCo = Object.entries(byCo).sort((a, b) => b[1] - a[1])[0];
                    return (
                      <Card className="relative overflow-hidden border-border/70 shadow-xs">
                        <div className="pointer-events-none absolute -right-12 -top-12 h-32 w-32 rounded-full bg-gradient-to-br from-blue-200/60 to-violet-100/0 blur-2xl" />
                        <CardContent className="relative space-y-3 px-5 py-5">
                          <div className="flex items-center justify-between">
                            <div className="flex items-center gap-2">
                              <div className="flex h-6 w-6 items-center justify-center rounded-md bg-blue-100 text-blue-700">
                                <CalendarIcon className="h-3.5 w-3.5" />
                              </div>
                              <div className="text-[10.5px] font-semibold uppercase tracking-[.08em] text-muted-foreground">
                                Today
                              </div>
                            </div>
                            <span className="font-mono text-[10px] text-muted-foreground">{todayStr}</span>
                          </div>

                          <div className="flex items-baseline gap-1.5">
                            <span className="text-[28px] font-bold leading-none tracking-tight proportional-nums text-foreground">
                              {todayRows.length}
                            </span>
                            <span className="text-[11px] font-medium text-muted-foreground">entries today</span>
                          </div>

                          <div className="grid grid-cols-2 gap-2">
                            <div className="rounded-lg bg-violet-50/70 px-3 py-2 ring-1 ring-violet-200/40">
                              <div className="text-[9.5px] font-semibold uppercase tracking-wider text-muted-foreground">Payin</div>
                              <div className="font-mono text-sm font-bold tabular-nums text-violet-700">{formatShort(todayPayin)}</div>
                            </div>
                            <div className="rounded-lg bg-emerald-50/70 px-3 py-2 ring-1 ring-emerald-200/40">
                              <div className="text-[9.5px] font-semibold uppercase tracking-wider text-muted-foreground">Settlement</div>
                              <div className="font-mono text-sm font-bold tabular-nums text-emerald-700">{formatShort(todaySettle)}</div>
                            </div>
                          </div>

                          <div className="space-y-1 pt-0.5 text-[10.5px]">
                            <div className="flex items-center justify-between text-muted-foreground">
                              <span>7-day avg</span>
                              <span className="font-semibold tabular-nums text-foreground">{weekAvg.toFixed(1)} / day</span>
                            </div>
                            <div className="flex items-center justify-between text-muted-foreground">
                              <span>Cycle split</span>
                              <span className="flex items-center gap-1 font-semibold">
                                <span className="rounded bg-violet-100 px-1.5 py-0.5 text-[10px] text-violet-700 tabular-nums">C1 {c1Today}</span>
                                <span className="rounded bg-amber-100 px-1.5 py-0.5 text-[10px] text-amber-700 tabular-nums">C2 {c2Today}</span>
                              </span>
                            </div>
                            {topCo && (
                              <div className="flex items-center justify-between gap-2 text-muted-foreground">
                                <span>Top company</span>
                                <span className="truncate text-right font-semibold text-foreground" title={topCo[0]}>
                                  {topCo[0]}
                                </span>
                              </div>
                            )}
                          </div>
                        </CardContent>
                      </Card>
                    );
                  })()}

                  {/* By Cycle */}
                  {(() => {
                    const totalEntries = c1.length + c2.length;
                    const c1PctNum = totalEntries > 0 ? (c1.length / totalEntries) * 100 : 0;
                    const c2PctNum = totalEntries > 0 ? (c2.length / totalEntries) * 100 : 0;
                    const totalPayin = c1Sum + c2Sum;
                    return (
                      <Card className="relative overflow-hidden border-border/70 shadow-xs">
                        <div className="pointer-events-none absolute -right-12 -top-12 h-32 w-32 rounded-full bg-gradient-to-br from-amber-200/60 to-violet-100/0 blur-2xl" />
                        <CardContent className="relative space-y-3 px-5 py-5">
                          <div className="flex items-center justify-between">
                            <div className="flex items-center gap-2">
                              <div className="flex h-6 w-6 items-center justify-center rounded-md bg-amber-100 text-amber-700">
                                <Sun className="h-3.5 w-3.5" />
                              </div>
                              <div className="text-[10.5px] font-semibold uppercase tracking-[.08em] text-muted-foreground">
                                By Cycle
                              </div>
                            </div>
                            <span className="font-mono text-[10px] text-muted-foreground tabular-nums">
                              {totalEntries} total
                            </span>
                          </div>

                          {/* Stacked split bar */}
                          <div className="flex h-2 w-full overflow-hidden rounded-full bg-muted">
                            <div
                              className="h-full bg-gradient-to-r from-violet-400 to-violet-600 transition-all"
                              style={{ width: `${c1PctNum}%` }}
                              title={`C1: ${c1PctNum.toFixed(0)}%`}
                            />
                            <div
                              className="h-full bg-gradient-to-r from-amber-400 to-amber-500 transition-all"
                              style={{ width: `${c2PctNum}%` }}
                              title={`C2: ${c2PctNum.toFixed(0)}%`}
                            />
                          </div>

                          <div className="grid grid-cols-2 gap-2">
                            <div className="rounded-lg bg-violet-50/70 px-3 py-2 ring-1 ring-violet-200/40">
                              <div className="flex items-center justify-between">
                                <div className="flex items-center gap-1 text-violet-700">
                                  <Sun className="h-3 w-3" />
                                  <span className="text-[10.5px] font-bold">C1</span>
                                </div>
                                <span className="text-[10px] font-mono tabular-nums text-violet-600">{c1PctNum.toFixed(0)}%</span>
                              </div>
                              <div className="mt-0.5 text-[15px] font-bold leading-none proportional-nums text-violet-700">
                                {c1.length}
                              </div>
                              <div className="mt-0.5 font-mono text-[10px] tabular-nums text-muted-foreground">{formatShort(c1Sum)}</div>
                            </div>
                            <div className="rounded-lg bg-amber-50/70 px-3 py-2 ring-1 ring-amber-200/40">
                              <div className="flex items-center justify-between">
                                <div className="flex items-center gap-1 text-amber-700">
                                  <Moon className="h-3 w-3" />
                                  <span className="text-[10.5px] font-bold">C2</span>
                                </div>
                                <span className="text-[10px] font-mono tabular-nums text-amber-600">{c2PctNum.toFixed(0)}%</span>
                              </div>
                              <div className="mt-0.5 text-[15px] font-bold leading-none proportional-nums text-amber-700">
                                {c2.length}
                              </div>
                              <div className="mt-0.5 font-mono text-[10px] tabular-nums text-muted-foreground">{formatShort(c2Sum)}</div>
                            </div>
                          </div>

                          <div className="pt-0.5 text-[10.5px] text-muted-foreground">
                            <div className="flex items-center justify-between">
                              <span>Active now</span>
                              <span className="inline-flex items-center gap-1 font-semibold text-foreground">
                                <span
                                  className="h-1.5 w-1.5 rounded-full"
                                  style={{ background: cyc.color, boxShadow: `0 0 4px ${cyc.color}` }}
                                />
                                C{cyc.cycle} · {cyc.label}
                              </span>
                            </div>
                            <div className="mt-0.5 flex items-center justify-between">
                              <span>Total payin</span>
                              <span className="font-mono font-semibold tabular-nums text-foreground">{formatShort(totalPayin)}</span>
                            </div>
                          </div>
                        </CardContent>
                      </Card>
                    );
                  })()}

                  {/* Reconciliation Status */}
                  <Card className="relative overflow-hidden border-border/70 shadow-xs">
                    <div className="pointer-events-none absolute -right-12 -top-12 h-32 w-32 rounded-full bg-gradient-to-br from-emerald-200/60 to-rose-100/0 blur-2xl" />
                    <CardContent className="relative space-y-3 px-5 py-5">
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <div className="flex h-6 w-6 items-center justify-center rounded-md bg-emerald-100 text-emerald-700">
                            <CheckCircle2 className="h-3.5 w-3.5" />
                          </div>
                          <div className="text-[10.5px] font-semibold uppercase tracking-[.08em] text-muted-foreground">
                            Reconciliation
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={() => setPage("recon")}
                          className="inline-flex items-center gap-0.5 text-[10px] font-semibold text-primary hover:underline"
                        >
                          View <ArrowRight className="h-2.5 w-2.5" />
                        </button>
                      </div>

                      <div className="flex items-baseline gap-1.5">
                        <span className="text-[28px] font-bold leading-none tracking-tight proportional-nums text-foreground">
                          {reconStats.total}
                        </span>
                        <span className="text-[11px] font-medium text-muted-foreground">checks</span>
                      </div>

                      <div className="flex h-2 w-full overflow-hidden rounded-full bg-muted">
                        <div className="h-full bg-emerald-500 transition-all" style={{ width: `${(reconStats.matched / reconTotal) * 100}%` }} />
                        <div className="h-full bg-red-500 transition-all" style={{ width: `${(reconStats.mismatch / reconTotal) * 100}%` }} />
                        <div className="h-full bg-amber-500 transition-all" style={{ width: `${(reconStats.pending / reconTotal) * 100}%` }} />
                      </div>

                      <div className="space-y-1 text-[11px]">
                        <div className="flex items-center justify-between">
                          <span className="flex items-center gap-1.5 text-emerald-700">
                            <span className="h-1.5 w-1.5 rounded-full bg-emerald-500" /> Matched
                          </span>
                          <span className="font-bold tabular-nums">{reconStats.matched}</span>
                        </div>
                        <div className="flex items-center justify-between">
                          <span className="flex items-center gap-1.5 text-red-700">
                            <span className="h-1.5 w-1.5 rounded-full bg-red-500" /> Mismatch
                          </span>
                          <span className="font-bold tabular-nums">{reconStats.mismatch}</span>
                        </div>
                        <div className="flex items-center justify-between">
                          <span className="flex items-center gap-1.5 text-amber-700">
                            <span className="h-1.5 w-1.5 rounded-full bg-amber-500" /> Pending
                          </span>
                          <span className="font-bold tabular-nums">{reconStats.pending}</span>
                        </div>
                      </div>

                      <div className="border-t border-border/60 pt-2 text-[10.5px]">
                        <div className="flex items-center justify-between text-muted-foreground">
                          <span>Discrepancy</span>
                          <span className={cn(
                            "font-mono font-semibold tabular-nums",
                            reconStats.totalDisc > 0 ? "text-red-700" : "text-emerald-700"
                          )}>
                            {formatShort(reconStats.totalDisc || 0)}
                          </span>
                        </div>
                        {reconStats.resolved > 0 && (
                          <div className="mt-0.5 flex items-center justify-between text-muted-foreground">
                            <span>Resolved</span>
                            <span className="font-semibold tabular-nums text-blue-700">{reconStats.resolved}</span>
                          </div>
                        )}
                      </div>
                    </CardContent>
                  </Card>

                  {/* Coverage */}
                  <Card className="relative overflow-hidden border-border/70 shadow-xs">
                    <div className="pointer-events-none absolute -right-12 -top-12 h-32 w-32 rounded-full bg-gradient-to-br from-violet-200/60 to-blue-100/0 blur-2xl" />
                    <CardContent className="relative space-y-3 px-5 py-5">
                      {/* Header */}
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <div className="flex h-6 w-6 items-center justify-center rounded-md bg-primary/10 text-primary">
                            <Users className="h-3.5 w-3.5" />
                          </div>
                          <div className="text-[10.5px] font-semibold uppercase tracking-[.08em] text-muted-foreground">
                            Coverage
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={() => setPage("merchants")}
                          className="inline-flex items-center gap-0.5 text-[10px] font-semibold text-primary hover:underline"
                        >
                          View <ArrowRight className="h-2.5 w-2.5" />
                        </button>
                      </div>

                      {/* Headline split */}
                      <div className="grid grid-cols-2 gap-2">
                        <div className="rounded-lg bg-violet-50/70 px-3 py-2.5 ring-1 ring-violet-200/50">
                          <div className="text-[22px] font-bold leading-none tracking-tight proportional-nums text-violet-700">
                            {merchantCount}
                          </div>
                          <div className="mt-1 text-[9.5px] font-semibold uppercase tracking-wider text-muted-foreground">
                            Merchants
                          </div>
                        </div>
                        <div className="rounded-lg bg-blue-50/70 px-3 py-2.5 ring-1 ring-blue-200/50">
                          <div className="text-[22px] font-bold leading-none tracking-tight proportional-nums text-blue-700">
                            {companyCount}
                          </div>
                          <div className="mt-1 text-[9.5px] font-semibold uppercase tracking-wider text-muted-foreground">
                            Companies
                          </div>
                        </div>
                      </div>

                      {/* Top 3 + insight */}
                      {(() => {
                        const byMerchant = {};
                        for (const d of MASTER_DATA) {
                          byMerchant[d.merchant] = (byMerchant[d.merchant] || 0) + 1;
                        }
                        const ranked = Object.entries(byMerchant).sort((a, b) => b[1] - a[1]);
                        const avg = merchantCount > 0 ? Math.round(companyCount / merchantCount) : 0;
                        const top3 = ranked.slice(0, 3);
                        const max = top3[0]?.[1] || 1;
                        return (
                          <div className="space-y-1.5 pt-0.5">
                            <div className="flex items-center justify-between text-[10px] text-muted-foreground">
                              <span>
                                Avg <span className="font-semibold tabular-nums text-foreground">{avg}</span> companies / merchant
                              </span>
                              <span>
                                Top: <span className="font-semibold text-foreground">{top3[0]?.[0] || "—"}</span>
                              </span>
                            </div>
                            <div className="space-y-1">
                              {top3.map(([name, count]) => (
                                <div key={name} className="flex items-center gap-2 text-[10.5px]">
                                  <span className="w-14 truncate font-semibold text-foreground" title={name}>
                                    {name}
                                  </span>
                                  <div className="relative h-1.5 flex-1 overflow-hidden rounded-full bg-muted">
                                    <div
                                      className="absolute inset-y-0 left-0 rounded-full bg-gradient-to-r from-violet-400 to-violet-600"
                                      style={{ width: `${(count / max) * 100}%` }}
                                    />
                                  </div>
                                  <span className="w-7 text-right text-[10px] font-semibold tabular-nums text-muted-foreground">
                                    {count}
                                  </span>
                                </div>
                              ))}
                            </div>
                          </div>
                        );
                      })()}
                    </CardContent>
                  </Card>
                </div>
              );
            })()}

            {/* Top Merchants + Recent Recon */}
            {(rows.length > 0 || recon.length > 0) && (() => {
              const topMerchants = (() => {
                const m = {};
                rows.forEach((r) => {
                  if (!m[r.merchant]) m[r.merchant] = { merchant: r.merchant, payin: 0, settlement: 0, count: 0 };
                  m[r.merchant].payin += r.payin || 0;
                  m[r.merchant].settlement += r.settlement || 0;
                  m[r.merchant].count += 1;
                });
                return Object.values(m).sort((a, b) => b.settlement - a.settlement).slice(0, 5);
              })();
              const maxSettlement = topMerchants[0]?.settlement || 1;
              const recentIssues = recon
                .filter((r) => r.status === "Mismatch" && !rResolved[r.id])
                .slice(-5)
                .reverse();
              return (
                <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
                  {/* Top Merchants */}
                  <Card className="border-border/70 shadow-xs">
                    <CardHeader className="border-b border-border/60 pb-3">
                      <div className="flex items-center justify-between">
                        <div>
                          <CardTitle className="text-sm font-semibold tracking-tight">Top Merchants</CardTitle>
                          <CardDescription className="text-xs">By settlement</CardDescription>
                        </div>
                        <Button size="sm" variant="ghost" onClick={() => setPage("reports")} className="text-xs">
                          All <ArrowRight className="ml-1 h-3.5 w-3.5" />
                        </Button>
                      </div>
                    </CardHeader>
                    <CardContent className="px-5 py-4">
                      {topMerchants.length === 0 ? (
                        <div className="py-6 text-center text-xs text-muted-foreground">No data yet</div>
                      ) : (
                        <div className="space-y-3">
                          {topMerchants.map((m, idx) => (
                            <div key={m.merchant} className="space-y-1.5">
                              <div className="flex items-center justify-between text-xs">
                                <div className="flex items-center gap-2 min-w-0">
                                  <span className="flex h-5 w-5 flex-shrink-0 items-center justify-center rounded bg-primary/10 font-mono text-[10px] font-bold text-primary">
                                    {idx + 1}
                                  </span>
                                  <span className="truncate font-semibold text-foreground">{m.merchant}</span>
                                  <span className="text-[10px] text-muted-foreground">· {m.count}</span>
                                </div>
                                <span className="ml-2 font-mono text-xs font-bold tabular-nums text-emerald-700">
                                  {formatShort(m.settlement)}
                                </span>
                              </div>
                              <div className="h-1 w-full overflow-hidden rounded-full bg-muted">
                                <div
                                  className="h-full rounded-full bg-gradient-to-r from-violet-400 to-violet-600"
                                  style={{ width: `${(m.settlement / maxSettlement) * 100}%` }}
                                />
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </CardContent>
                  </Card>

                  {/* Recent Recon Issues */}
                  <Card className="border-border/70 shadow-xs">
                    <CardHeader className="border-b border-border/60 pb-3">
                      <div className="flex items-center justify-between">
                        <div>
                          <CardTitle className="text-sm font-semibold tracking-tight">Open Recon Issues</CardTitle>
                          <CardDescription className="text-xs">Latest unresolved mismatches</CardDescription>
                        </div>
                        <Button size="sm" variant="ghost" onClick={() => setPage("recon")} className="text-xs">
                          All <ArrowRight className="ml-1 h-3.5 w-3.5" />
                        </Button>
                      </div>
                    </CardHeader>
                    <CardContent className="px-5 py-4">
                      {recentIssues.length === 0 ? (
                        <div className="flex flex-col items-center justify-center py-6 text-center">
                          <CheckCircle2 className="mb-1.5 h-5 w-5 text-emerald-600" />
                          <div className="text-xs font-semibold text-foreground">All clear</div>
                          <div className="text-[11px] text-muted-foreground">No open issues</div>
                        </div>
                      ) : (
                        <div className="space-y-2">
                          {recentIssues.map((r) => (
                            <div
                              key={r.id}
                              className="flex items-center justify-between gap-3 rounded-md border border-border/60 bg-muted/30 px-3 py-2"
                            >
                              <div className="min-w-0 flex-1">
                                <div className="flex items-center gap-2">
                                  <Badge variant="secondary" className="bg-red-100 px-1.5 py-0 font-mono text-[9.5px] text-red-800 hover:bg-red-100">
                                    {r.txnId}
                                  </Badge>
                                  <span className="truncate text-[11px] font-semibold text-foreground">{r.merchant}</span>
                                </div>
                                <div className="mt-0.5 truncate text-[10px] text-muted-foreground">{r.company}</div>
                              </div>
                              <div className="text-right">
                                <div className="font-mono text-[11px] font-bold tabular-nums text-red-700">
                                  {r.diff > 0 ? "+" : ""}{formatINR(r.diff)}
                                </div>
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </CardContent>
                  </Card>
                </div>
              );
            })()}

            {rows.length > 0 ? (
              <Card className="overflow-hidden border-border/70 shadow-xs">
                <div className="flex flex-wrap items-center justify-between gap-3 border-b border-border/70 px-5 py-4">
                  <div>
                    <div className="text-[15px] font-semibold tracking-tight text-foreground">
                      Recent Entries
                    </div>
                    <div className="text-xs text-muted-foreground">Last 5 settlements</div>
                  </div>
                  <div className="flex items-center gap-2">
                    <Button
                      size="sm"
                      variant="outline"
                      onClick={() => setPage("reports")}
                    >
                      View all <ArrowRight className="ml-1 h-3.5 w-3.5" />
                    </Button>
                    <Button
                      size="sm"
                      onClick={() => downloadSheet(rows, "settlements")}
                    >
                      <SI d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1M12 12v8m0 0l-4-4m4 4l4-4M12 4v4" />
                      Download
                    </Button>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-muted/40">
                      <tr>
                        {["Merchant", "Company", "Payin", "Settlement", "Cycle"].map((h) => (
                          <th
                            key={h}
                            className={cn(
                              "px-5 py-2.5 text-[10px] font-semibold uppercase tracking-wider text-muted-foreground",
                              h === "Merchant" || h === "Company" ? "text-left" : "text-right"
                            )}
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
                            className="border-t border-border/60 transition-colors hover:bg-muted/30"
                          >
                            <td className="px-5 py-3 font-semibold text-primary">{r.merchant}</td>
                            <td className="max-w-[240px] overflow-hidden text-ellipsis whitespace-nowrap px-5 py-3 text-foreground">
                              {r.company}
                            </td>
                            <td className="px-5 py-3 text-right font-mono font-semibold tabular-nums text-foreground">
                              {formatINR(r.payin)}
                            </td>
                            <td className="px-5 py-3 text-right font-mono font-bold tabular-nums text-emerald-700">
                              {formatINR(r.settlement)}
                            </td>
                            <td className="px-5 py-3 text-right">
                              <Badge
                                variant="secondary"
                                className={cn(
                                  "font-mono text-[10.5px]",
                                  r.cycle === 2
                                    ? "bg-amber-100 text-amber-800 hover:bg-amber-100"
                                    : "bg-violet-100 text-violet-800 hover:bg-violet-100"
                                )}
                              >
                                C{r.cycle}
                              </Badge>
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              </Card>
            ) : (
              <Card className="border-dashed border-border/80 bg-background/40 shadow-none">
                <CardContent className="flex flex-col items-center justify-center px-6 py-16 text-center">
                  <div className="mb-4 flex h-14 w-14 items-center justify-center rounded-2xl bg-primary/10 text-primary">
                    <Receipt className="h-7 w-7" />
                  </div>
                  <div className="text-base font-semibold text-foreground">
                    No entries yet
                  </div>
                  <div className="mb-5 mt-1 max-w-sm text-sm text-muted-foreground">
                    Create your first settlement entry to see KPIs, recent activity, and reconciliation status here.
                  </div>
                  <Button onClick={() => setPage("add")}>
                    <Plus className="mr-1 h-4 w-4" />
                    Add Entry
                  </Button>
                </CardContent>
              </Card>
            )}
          </div>
        )}

        {/* ========== ADD ENTRY ========== */}
        {page === "add" && (
          <div className="fade-up">
            <Card className="border-border shadow-sm">
              <CardHeader className="border-b border-border/60 pb-5">
                <div className="flex items-center gap-3">
                  <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-primary/10 text-primary">
                    <Receipt className="h-5 w-5" />
                  </div>
                  <div>
                    <CardTitle className="text-base font-semibold text-foreground">
                      New Settlement Entry
                    </CardTitle>
                    <CardDescription className="text-xs text-muted-foreground">
                      Charge {(CHARGE_RATE).toFixed(2)}% on payin · GST applied on charge
                    </CardDescription>
                  </div>
                </div>
              </CardHeader>

              <CardContent className="space-y-6 pt-6">
                {/* Row 1: Merchant + Company */}
                <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                  <div className="space-y-2">
                    <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                      Merchant
                    </Label>
                    <Select
                      value={selM || undefined}
                      onValueChange={(v) => {
                        setSM(v);
                        setSC("");
                        setCS("");
                      }}
                    >
                      <SelectTrigger className="h-11 w-full">
                        <SelectValue placeholder="— Select merchant —" />
                      </SelectTrigger>
                      <SelectContent>
                        {ALL_MERCHANTS.map((m) => (
                          <SelectItem key={m} value={m}>
                            <span className="font-medium">{m}</span>
                            <span className="ml-2 text-xs text-muted-foreground">({mCnt[m]})</span>
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>

                  <div className="relative space-y-2">
                    <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                      Company
                    </Label>
                    <div className="relative">
                      <Building2 className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground/70" />
                      <Input
                        value={selC || cSearch}
                        onChange={(e) => {
                          setCS(e.target.value);
                          setSC("");
                        }}
                        placeholder={
                          selM ? `Search ${compsFor.length} companies…` : "Select merchant first"
                        }
                        disabled={!selM}
                        className="h-11 pl-9"
                        list="co-l"
                      />
                    </div>
                    <datalist id="co-l">
                      {filtCL.map((c) => (
                        <option key={c} value={c} />
                      ))}
                    </datalist>
                    {cSearch && !selC && filtCL.length > 0 && (
                      <div className="absolute left-0 right-0 top-full z-20 mt-1 max-h-48 overflow-y-auto rounded-md border border-border bg-popover shadow-lg">
                        {filtCL.slice(0, 8).map((c) => (
                          <button
                            key={c}
                            type="button"
                            onClick={() => {
                              setSC(c);
                              setCS("");
                            }}
                            className="block w-full border-b border-border/60 px-3 py-2.5 text-left text-xs text-popover-foreground last:border-b-0 hover:bg-accent hover:text-accent-foreground"
                          >
                            {c}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                </div>

                {/* Selected pair chip */}
                {selC && (
                  <div className="flex items-center justify-between rounded-lg border border-primary/30 bg-primary/5 px-4 py-2.5">
                    <div className="flex items-center gap-2 text-sm">
                      <Badge variant="secondary" className="bg-primary/15 text-primary hover:bg-primary/15">
                        {selM}
                      </Badge>
                      <span className="text-muted-foreground">→</span>
                      <span className="font-medium text-foreground">{selC}</span>
                    </div>
                    <Button
                      type="button"
                      variant="ghost"
                      size="icon"
                      onClick={() => {
                        setSC("");
                        setCS("");
                      }}
                      className="h-7 w-7 text-muted-foreground hover:bg-primary/10 hover:text-foreground"
                    >
                      <X className="h-4 w-4" />
                    </Button>
                  </div>
                )}

                {/* Row 2: Payin + Chargeback + GST */}
                <div className="grid grid-cols-1 gap-4 md:grid-cols-[1.2fr_1fr_1.2fr]">
                  <div className="space-y-2">
                    <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                      Payin Amount (₹)
                    </Label>
                    <div className="relative">
                      <IndianRupee className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground/70" />
                      <Input
                        type="number"
                        value={payin}
                        onChange={(e) => setPay(e.target.value)}
                        onKeyDown={(e) => e.key === "Enter" && addRow()}
                        placeholder="10,000"
                        className="h-11 pl-9 font-mono text-base font-semibold tabular-nums"
                      />
                    </div>
                  </div>

                  <div className="space-y-2">
                    <Label className="flex items-baseline justify-between text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                      <span>Chargeback (₹)</span>
                      <span className="text-[10px] font-normal normal-case tracking-normal text-muted-foreground/70">optional</span>
                    </Label>
                    <Input
                      type="number"
                      value={chargeback}
                      onChange={(e) => setChargeback(e.target.value)}
                      onKeyDown={(e) => e.key === "Enter" && addRow()}
                      placeholder="0"
                      className={cn(
                        "h-11 font-mono text-base font-semibold tabular-nums",
                        cb > 0 && "text-destructive"
                      )}
                    />
                  </div>

                  <div className="space-y-2">
                    <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                      GST Rate
                    </Label>
                    <div className="inline-flex h-11 w-full items-center rounded-md bg-muted p-1">
                      {GST_OPTIONS.map((opt) => {
                        const active = gstRate === opt;
                        return (
                          <button
                            key={opt}
                            type="button"
                            onClick={() => setGstRate(opt)}
                            className={cn(
                              "flex-1 rounded-sm font-mono text-xs font-bold transition-all",
                              "h-full",
                              active
                                ? "bg-background text-primary shadow-sm"
                                : "text-muted-foreground hover:text-foreground"
                            )}
                          >
                            {opt}%
                          </button>
                        );
                      })}
                    </div>
                  </div>
                </div>

                {/* Row 3: Date + Cycle */}
                <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                  <div className="space-y-2">
                    <Label className="flex items-baseline justify-between text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                      <span>Entry Date</span>
                      <span className="text-[10px] font-normal normal-case tracking-normal text-muted-foreground/70">defaults to today</span>
                    </Label>
                    <div className="relative">
                      <CalendarIcon className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground/70" />
                      <Input
                        type="date"
                        value={entryDate}
                        onChange={(e) => setEntryDate(e.target.value)}
                        className="h-11 pl-9"
                      />
                    </div>
                  </div>

                  <div className="space-y-2">
                    <Label className="flex items-baseline justify-between text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                      <span>Cycle</span>
                      <span className="text-[10px] font-normal normal-case tracking-normal text-muted-foreground/70">auto: C{cyc.cycle}</span>
                    </Label>
                    <div className="grid h-11 grid-cols-3 gap-1.5">
                      {[
                        { c: null, label: "Auto", sub: `C${cyc.cycle}`, icon: null },
                        { c: 1, label: "C1", sub: "6AM–4PM", icon: Sun },
                        { c: 2, label: "C2", sub: "4PM–6AM", icon: Moon },
                      ].map((opt) => {
                        const active = entryCycle === opt.c;
                        const Icon = opt.icon;
                        return (
                          <button
                            key={opt.label}
                            type="button"
                            onClick={() => setEntryCycle(opt.c)}
                            className={cn(
                              "flex flex-col items-center justify-center gap-0.5 rounded-md border text-xs font-semibold transition-all",
                              active
                                ? "border-primary bg-primary text-primary-foreground shadow-sm"
                                : "border-border bg-background text-muted-foreground hover:border-primary/40 hover:bg-accent hover:text-foreground"
                            )}
                          >
                            <span className="flex items-center gap-1 leading-none">
                              {Icon && <Icon className="h-3 w-3" />}
                              {opt.label}
                            </span>
                            <span className={cn("text-[9px] font-medium leading-none", active ? "opacity-90" : "opacity-70")}>{opt.sub}</span>
                          </button>
                        );
                      })}
                    </div>
                  </div>
                </div>

                {/* Submit */}
                <Button
                  onClick={addRow}
                  disabled={!canAdd}
                  size="lg"
                  className="h-12 w-full text-sm font-semibold tracking-wide"
                >
                  <Plus className="mr-1 h-4 w-4" />
                  Add Entry
                </Button>

                {/* Preview totals */}
                {pn > 0 && (
                  <>
                    <Separator />
                    <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-5">
                      {[
                        { l: "Charge", v: ch, tone: "amber" },
                        { l: `GST (${gstRate}%)`, v: gs, tone: "rose" },
                        { l: "Chargeback", v: cb, tone: "red" },
                        { l: "Deduction", v: td, tone: "orange" },
                        { l: "Final Settlement", v: stl, tone: "primary" },
                      ].map((x) => {
                        const toneMap = {
                          amber: { bar: "bg-amber-500", text: "text-amber-600" },
                          rose: { bar: "bg-rose-500", text: "text-rose-600" },
                          red: { bar: "bg-red-600", text: "text-red-700" },
                          orange: { bar: "bg-orange-500", text: "text-orange-600" },
                          primary: { bar: "bg-primary", text: "text-primary" },
                        }[x.tone];
                        return (
                          <div
                            key={x.l}
                            className="relative overflow-hidden rounded-lg border border-border bg-muted/40 p-3"
                          >
                            <div className={cn("absolute left-0 top-0 h-full w-1", toneMap.bar)} />
                            <div className="pl-1.5">
                              <div className="mb-1 text-[9px] font-semibold uppercase tracking-wider text-muted-foreground">
                                {x.l}
                              </div>
                              <div className={cn("font-mono text-sm font-bold tabular-nums", toneMap.text)}>
                                {formatINR(x.v)}
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </>
                )}
              </CardContent>
            </Card>
          </div>
        )}

        {/* ========== REPORTS ========== */}
        {page === "reports" && (
          <div className="fade-up space-y-5">
            <Card className="border-border/70 shadow-xs">
              <CardContent className="flex flex-wrap items-center gap-3 px-5 py-4">
                <div className="-mx-1 max-w-full overflow-x-auto px-1">
                  <Tabs value={view} onValueChange={setView}>
                    <TabsList>
                      <TabsTrigger value="all">All</TabsTrigger>
                      <TabsTrigger value="company">By Company</TabsTrigger>
                      <TabsTrigger value="merchant">By Merchant</TabsTrigger>
                    </TabsList>
                  </Tabs>
                </div>
                <div className="ml-auto flex flex-wrap items-center gap-2">
                  <Select
                    value={fM}
                    onValueChange={setFM}
                  >
                    <SelectTrigger className="h-9 w-[180px]">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__all__">All Merchants</SelectItem>
                      {[...new Set(rows.map((r) => r.merchant))]
                        .sort()
                        .map((m) => (
                          <SelectItem key={m} value={m}>{m}</SelectItem>
                        ))}
                    </SelectContent>
                  </Select>

                  <div
                    className="flex h-9 min-w-0 items-center gap-2 rounded-md border border-border bg-background px-2.5"
                    title="Filter downloads by date range"
                  >
                    <CalendarIcon className="h-3.5 w-3.5 flex-shrink-0 text-muted-foreground" />
                    <input
                      type="date"
                      value={dlFrom}
                      onChange={(e) => setDlFrom(e.target.value)}
                      className="w-[110px] min-w-0 border-0 bg-transparent text-xs outline-none"
                    />
                    <span className="text-[10px] font-semibold uppercase text-muted-foreground">to</span>
                    <input
                      type="date"
                      value={dlTo}
                      onChange={(e) => setDlTo(e.target.value)}
                      className="w-[110px] min-w-0 border-0 bg-transparent text-xs outline-none"
                    />
                    {(dlFrom || dlTo) && (
                      <button
                        type="button"
                        onClick={() => { setDlFrom(""); setDlTo(""); }}
                        className="text-muted-foreground hover:text-foreground"
                        title="Clear date range"
                      >
                        <X className="h-3.5 w-3.5" />
                      </button>
                    )}
                  </div>

                  {rows.length > 0 && (
                    <>
                      <Button
                        size="sm"
                        onClick={() => downloadSheet(fRows, "settlement_report")}
                      >
                        <Download className="mr-1 h-3.5 w-3.5" />
                        Download
                      </Button>
                      <Button
                        size="sm"
                        variant="outline"
                        onClick={clearAll}
                        className="border-destructive/30 text-destructive hover:bg-destructive/5 hover:text-destructive"
                      >
                        <Trash2 className="mr-1 h-3.5 w-3.5" />
                        Clear all
                      </Button>
                    </>
                  )}
                </div>
              </CardContent>
            </Card>

            {grouped && grouped.length > 0 && (
              <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fill,minmax(280px,1fr))" }}>
                {grouped.map((g, i) => (
                  <Card
                    key={i}
                    className="fade-up overflow-hidden border-border/70 shadow-xs"
                  >
                    <div className={cn("h-1 w-full", view === "company" ? "bg-primary" : "bg-emerald-500")} />
                    <CardContent className="px-5 py-4">
                      <div className="mb-3 text-sm font-semibold leading-snug tracking-tight text-foreground">
                        {g.name}
                      </div>
                      <div className="grid grid-cols-2 gap-3">
                        {[
                          { l: "Entries", v: g.count, raw: 1, c: "text-muted-foreground" },
                          { l: "Total Payin", v: g.payin, c: "text-foreground" },
                          { l: "Deductions", v: g.deduction, c: "text-amber-600" },
                          { l: "Settlement", v: g.settlement, c: "text-emerald-700" },
                        ].map((x) => (
                          <div key={x.l}>
                            <div className="text-[9.5px] font-semibold uppercase tracking-wider text-muted-foreground">
                              {x.l}
                            </div>
                            <div className={cn("font-mono text-[13px] font-bold tabular-nums", x.c)}>
                              {x.raw ? x.v : formatINR(x.v)}
                            </div>
                          </div>
                        ))}
                      </div>
                    </CardContent>
                  </Card>
                ))}
              </div>
            )}

            {fRows.length > 0 && view === "all" && (
              <Card className="overflow-hidden border-border/70 shadow-xs">
                <div className="flex items-center justify-between border-b border-border/70 px-5 py-3.5">
                  <div>
                    <div className="text-sm font-semibold tracking-tight text-foreground">
                      All Entries
                    </div>
                    <div className="text-xs text-muted-foreground">{fRows.length} rows</div>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-muted/40">
                      <tr>
                        {["#", "Merchant", "Company", "Payin", "Deduction", "Settlement", "Cycle", ""].map((h, i) => (
                          <th
                            key={i}
                            className={cn(
                              "px-4 py-2.5 text-[10px] font-semibold uppercase tracking-wider text-muted-foreground",
                              i <= 2 ? "text-left" : "text-right"
                            )}
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
                          className="border-t border-border/60 transition-colors hover:bg-muted/30"
                        >
                          <td className="px-4 py-2.5 text-[11px] text-muted-foreground">{i + 1}</td>
                          <td className="px-4 py-2.5 font-semibold text-primary">{r.merchant}</td>
                          <td
                            className="max-w-[220px] overflow-hidden text-ellipsis whitespace-nowrap px-4 py-2.5 text-foreground"
                            title={r.company}
                          >
                            {r.company}
                          </td>
                          <td className="px-4 py-2.5 text-right font-mono font-semibold tabular-nums text-foreground">
                            {formatINR(r.payin)}
                          </td>
                          <td className="px-4 py-2.5 text-right font-mono tabular-nums text-amber-600">
                            {formatINR(r.deduction)}
                          </td>
                          <td className="px-4 py-2.5 text-right font-mono font-bold tabular-nums text-emerald-700">
                            {formatINR(r.settlement)}
                          </td>
                          <td className="px-4 py-2.5 text-right">
                            <Badge
                              variant="secondary"
                              className={cn(
                                "font-mono text-[10.5px]",
                                r.cycle === 2
                                  ? "bg-amber-100 text-amber-800 hover:bg-amber-100"
                                  : "bg-violet-100 text-violet-800 hover:bg-violet-100"
                              )}
                            >
                              C{r.cycle}
                            </Badge>
                          </td>
                          <td className="px-2 py-2.5 text-right">
                            <Button
                              variant="ghost"
                              size="icon"
                              onClick={() => rmRow(r.id)}
                              className="h-7 w-7 text-muted-foreground/60 hover:bg-destructive/10 hover:text-destructive"
                            >
                              <X className="h-4 w-4" />
                            </Button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Card>
            )}

            {rows.length === 0 && (
              <Card className="border-dashed border-border/80 bg-background/40 shadow-none">
                <CardContent className="flex flex-col items-center justify-center px-6 py-14 text-center">
                  <div className="mb-3 flex h-12 w-12 items-center justify-center rounded-2xl bg-primary/10 text-primary">
                    <FileSpreadsheet className="h-6 w-6" />
                  </div>
                  <div className="text-base font-semibold text-foreground">No entries yet</div>
                  <div className="mb-4 mt-1 text-sm text-muted-foreground">
                    Add your first settlement to populate this report.
                  </div>
                  <Button onClick={() => setPage("add")}>
                    <Plus className="mr-1 h-4 w-4" />
                    Add Entry
                  </Button>
                </CardContent>
              </Card>
            )}
          </div>
        )}

        {/* ========== RECONCILIATION ========== */}
        {page === "recon" && (
          <div className="fade-up space-y-5">
            <div className="grid gap-3" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))" }}>
              {[
                { l: "Total Checked", v: reconStats.total, tone: "violet" },
                { l: "Matched", v: reconStats.matched, tone: "emerald" },
                { l: "Mismatch", v: reconStats.mismatch, tone: "red" },
                { l: "Pending", v: reconStats.pending, tone: "amber" },
                { l: "Resolved", v: reconStats.resolved, tone: "blue" },
                { l: "Total Discrepancy", v: formatINR(reconStats.totalDisc), tone: "red", raw: 1 },
              ].map((s, i) => {
                const toneMap = {
                  violet: { value: "text-violet-700", bg: "from-violet-200/60 to-violet-100/0" },
                  emerald: { value: "text-emerald-700", bg: "from-emerald-200/60 to-emerald-100/0" },
                  red: { value: "text-red-700", bg: "from-red-200/60 to-red-100/0" },
                  amber: { value: "text-amber-700", bg: "from-amber-200/60 to-amber-100/0" },
                  blue: { value: "text-blue-700", bg: "from-blue-200/60 to-blue-100/0" },
                }[s.tone];
                return (
                  <Card key={i} className="relative overflow-hidden border-border/70 shadow-xs">
                    <div className={cn("pointer-events-none absolute -right-8 -top-8 h-24 w-24 rounded-full bg-gradient-to-br blur-2xl", toneMap.bg)} />
                    <CardContent className="relative px-5 py-4">
                      <div className="mb-1.5 text-[10px] font-semibold uppercase tracking-[.08em] text-muted-foreground">
                        {s.l}
                      </div>
                      <div className={cn("font-mono font-bold tabular-nums leading-none", s.raw ? "text-[17px]" : "text-[24px]", toneMap.value)}>
                        {s.v}
                      </div>
                    </CardContent>
                  </Card>
                );
              })}
            </div>

            {/* Reconciliation Uploads */}
            <Card className="border-border/70 shadow-xs">
              <CardHeader className="flex flex-row flex-wrap items-start justify-between gap-3 border-b border-border/60 pb-4">
                <div>
                  <CardTitle className="text-[15px] font-semibold tracking-tight text-foreground">
                    Reconciliation Uploads
                  </CardTitle>
                  <CardDescription className="text-xs text-muted-foreground">
                    Upload internal payin and bank settlement files — auto-aggregates by merchant
                  </CardDescription>
                </div>
                {(bankRows.length > 0 || sourceRows.length > 0) && (
                  <Button size="sm" onClick={downloadBankRecon}>
                    <Download className="mr-1 h-3.5 w-3.5" />
                    Download Recon
                  </Button>
                )}
              </CardHeader>
              <CardContent className="space-y-4 pt-5">
                {/* Two side-by-side drop zones */}
                <div className="grid gap-3 md:grid-cols-2">
                  {/* Source / internal payin */}
                  <div>
                    <div className="mb-2 flex items-center justify-between">
                      <div className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                        1. Internal Payin
                      </div>
                      {sourceRows.length > 0 && (
                        <button
                          type="button"
                          onClick={clearSource}
                          className="text-[11px] font-semibold text-muted-foreground hover:text-foreground"
                        >
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
                        className={cn(
                          "cursor-pointer rounded-xl border-2 border-dashed px-4 py-7 text-center transition-all",
                          sourceDragging
                            ? "border-blue-400 bg-blue-50/70"
                            : "border-border bg-muted/30 hover:border-primary/40 hover:bg-muted/50"
                        )}
                      >
                        <input id="source-file-input" type="file" accept=".csv,.xlsx,.xls" onChange={handleSourceInput} className="hidden" />
                        <div className={cn("mx-auto mb-3 flex h-11 w-11 items-center justify-center rounded-xl", sourceDragging ? "bg-blue-100 text-blue-700" : "bg-primary/10 text-primary")}>
                          {sourceDragging ? <Upload className="h-5 w-5" /> : <FileSpreadsheet className="h-5 w-5" />}
                        </div>
                        <div className="text-xs font-semibold text-foreground">
                          {sourceDragging ? "Drop here" : "Drag & drop internal payin Excel"}
                        </div>
                        <div className="mt-1 text-[10px] text-muted-foreground">
                          Detects: Row Labels / Name + Sum of Amount
                        </div>
                      </div>
                    ) : (
                      <div className="rounded-lg border border-blue-200 bg-blue-50/70 px-3.5 py-3 text-[11px] text-blue-900">
                        <div className="flex items-center gap-1.5 font-semibold">
                          <FileSpreadsheet className="h-3.5 w-3.5" /> {sourceFile}
                        </div>
                        <div className="mt-1.5">
                          {sourceRows.length} companies · {sourceByMerchant.size} merchants · Total{" "}
                          <strong>{formatINR(sourceRows.reduce((s, r) => s + r.amount, 0))}</strong>
                        </div>
                      </div>
                    )}
                    {sourceError && (
                      <div className="mt-2 flex items-start gap-2 rounded-md border border-destructive/30 bg-destructive/5 px-3 py-2 text-[11px] font-semibold text-destructive">
                        <AlertCircle className="mt-px h-3.5 w-3.5 flex-shrink-0" />
                        {sourceError}
                      </div>
                    )}
                  </div>

                  {/* Bank settlement */}
                  <div>
                    <div className="mb-2 flex items-center justify-between">
                      <div className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">
                        2. Bank Settlement
                      </div>
                      {bankRows.length > 0 && (
                        <button
                          type="button"
                          onClick={clearBank}
                          className="text-[11px] font-semibold text-muted-foreground hover:text-foreground"
                        >
                          Clear ×
                        </button>
                      )}
                    </div>
                    {bankRows.length === 0 ? (
                      <div
                        onDragOver={(e) => { e.preventDefault(); setBankDragging(true); }}
                        onDragLeave={() => setBankDragging(false)}
                        onDrop={handleBankDrop}
                        onClick={() => document.getElementById("bank-file-input")?.click()}
                        className={cn(
                          "cursor-pointer rounded-xl border-2 border-dashed px-4 py-7 text-center transition-all",
                          bankDragging
                            ? "border-emerald-400 bg-emerald-50/70"
                            : "border-border bg-muted/30 hover:border-primary/40 hover:bg-muted/50"
                        )}
                      >
                        <input id="bank-file-input" type="file" accept=".csv,.xlsx,.xls" onChange={handleBankInput} className="hidden" />
                        <div className={cn("mx-auto mb-3 flex h-11 w-11 items-center justify-center rounded-xl", bankDragging ? "bg-emerald-100 text-emerald-700" : "bg-emerald-50 text-emerald-700")}>
                          {bankDragging ? <Upload className="h-5 w-5" /> : <FileSpreadsheet className="h-5 w-5" />}
                        </div>
                        <div className="text-xs font-semibold text-foreground">
                          {bankDragging ? "Drop here" : "Drag & drop bank settlement Excel"}
                        </div>
                        <div className="mt-1 text-[10px] text-muted-foreground">
                          Detects: MID · Name · Amount · Fee · GST · Settle · Chargeback
                        </div>
                      </div>
                    ) : (
                      <div className="rounded-lg border border-emerald-200 bg-emerald-50/70 px-3.5 py-3 text-[11px] text-emerald-900">
                        <div className="flex items-center gap-1.5 font-semibold">
                          <FileSpreadsheet className="h-3.5 w-3.5" /> {bankFile}
                        </div>
                        <div className="mt-1.5">
                          {bankRows.length} companies · {bankByMerchant.length} merchants
                          {bankTotals.unmatched > 0 && (
                            <span className="ml-1.5 text-red-700">· {bankTotals.unmatched} unmatched</span>
                          )}
                        </div>
                      </div>
                    )}
                    {bankError && (
                      <div className="mt-2 flex items-start gap-2 rounded-md border border-destructive/30 bg-destructive/5 px-3 py-2 text-[11px] font-semibold text-destructive">
                        <AlertCircle className="mt-px h-3.5 w-3.5 flex-shrink-0" />
                        {bankError}
                      </div>
                    )}
                  </div>
                </div>

                {(bankRows.length > 0 || sourceRows.length > 0) && (
                  <div className="overflow-x-auto rounded-lg border border-border">
                    <table className="w-full text-xs">
                      <thead className="bg-muted/40">
                        <tr>
                          {(() => {
                            const headers = ["Merchant"];
                            if (sourceRows.length > 0) headers.push("Source Payin");
                            if (bankRows.length > 0) headers.push("Bank Payin");
                            if (sourceRows.length > 0 && bankRows.length > 0) headers.push("Payin Diff");
                            if (bankRows.length > 0) headers.push("Bank Fee", "Bank GST", "Bank Settle", "Chargeback", "Net Settle", "Our Settle", "Difference");
                            return headers.map((h, i) => (
                              <th
                                key={i}
                                className={cn(
                                  "whitespace-nowrap px-3 py-2.5 text-[9.5px] font-semibold uppercase tracking-wider text-muted-foreground",
                                  i <= 0 ? "text-left" : "text-right"
                                )}
                              >
                                {h}
                              </th>
                            ));
                          })()}
                        </tr>
                      </thead>
                      <tbody>
                        {combinedByMerchant.map((g) => (
                          <tr key={g.merchant} className="border-t border-border/60 transition-colors hover:bg-muted/30">
                            <td className={cn("px-3 py-2.5 font-bold", g.merchant === "Unmatched" ? "text-red-700" : "text-primary")}>
                              {g.merchant}
                              {g.unmatched > 0 && g.merchant !== "Unmatched" && (
                                <Badge
                                  variant="secondary"
                                  className="ml-2 bg-amber-100 px-1.5 py-0 text-[9px] font-bold text-amber-800 hover:bg-amber-100"
                                  title="unmatched bank rows"
                                >
                                  {g.unmatched} ub
                                </Badge>
                              )}
                            </td>
                            {sourceRows.length > 0 && (
                              <td className="px-3 py-2.5 text-right font-mono font-semibold tabular-nums text-violet-700">
                                {g.sourcePayin > 0 ? formatINR(g.sourcePayin) : "—"}
                              </td>
                            )}
                            {bankRows.length > 0 && (
                              <td className="px-3 py-2.5 text-right font-mono font-semibold tabular-nums text-foreground">
                                {g.amount > 0 ? formatINR(g.amount) : "—"}
                              </td>
                            )}
                            {sourceRows.length > 0 && bankRows.length > 0 && (
                              <td className={cn("px-3 py-2.5 text-right font-mono font-bold tabular-nums", Math.abs(g.payinDiff) <= 1 ? "text-emerald-700" : "text-red-700")}>
                                {g.payinDiff > 0 ? "+" : ""}{formatINR(g.payinDiff)}
                              </td>
                            )}
                            {bankRows.length > 0 && (
                              <>
                                <td className="px-3 py-2.5 text-right font-mono tabular-nums text-amber-600">{formatINR(g.fee)}</td>
                                <td className="px-3 py-2.5 text-right font-mono tabular-nums text-rose-600">{formatINR(g.gst)}</td>
                                <td className="px-3 py-2.5 text-right font-mono font-bold tabular-nums text-blue-700">{formatINR(g.settle)}</td>
                                <td className={cn("px-3 py-2.5 text-right font-mono tabular-nums", g.chargeback > 0 ? "text-red-700" : "text-muted-foreground")}>
                                  {g.chargeback > 0 ? "−" + formatINR(g.chargeback) : "—"}
                                </td>
                                <td className="px-3 py-2.5 text-right font-mono font-bold tabular-nums text-blue-900">{formatINR(g.netSettle)}</td>
                                <td className="px-3 py-2.5 text-right font-mono font-bold tabular-nums text-emerald-700">{formatINR(g.ourNet)}</td>
                                <td className={cn("px-3 py-2.5 text-right font-mono font-bold tabular-nums", Math.abs(g.diff) <= 1 ? "text-emerald-700" : "text-red-700")}>
                                  {g.diff > 0 ? "+" : ""}{formatINR(g.diff)}
                                </td>
                              </>
                            )}
                          </tr>
                        ))}
                        <tr className="border-t-2 border-border bg-muted/60">
                          <td className="px-3 py-2.5 font-bold text-foreground">TOTAL</td>
                          {sourceRows.length > 0 && (
                            <td className="px-3 py-2.5 text-right font-mono font-bold tabular-nums text-violet-700">{formatINR(combinedTotals.sourcePayin)}</td>
                          )}
                          {bankRows.length > 0 && (
                            <td className="px-3 py-2.5 text-right font-mono font-bold tabular-nums text-foreground">{formatINR(combinedTotals.bankPayin)}</td>
                          )}
                          {sourceRows.length > 0 && bankRows.length > 0 && (
                            <td className={cn("px-3 py-2.5 text-right font-mono font-bold tabular-nums", Math.abs(combinedTotals.payinDiff) <= 1 ? "text-emerald-700" : "text-red-700")}>
                              {combinedTotals.payinDiff > 0 ? "+" : ""}{formatINR(combinedTotals.payinDiff)}
                            </td>
                          )}
                          {bankRows.length > 0 && (
                            <>
                              <td className="px-3 py-2.5 text-right font-mono font-bold tabular-nums text-amber-600">{formatINR(bankTotals.fee)}</td>
                              <td className="px-3 py-2.5 text-right font-mono font-bold tabular-nums text-rose-600">{formatINR(bankTotals.gst)}</td>
                              <td className="px-3 py-2.5 text-right font-mono font-bold tabular-nums text-blue-700">{formatINR(bankTotals.settle)}</td>
                              <td className={cn("px-3 py-2.5 text-right font-mono font-bold tabular-nums", bankTotals.chargeback > 0 ? "text-red-700" : "text-muted-foreground")}>
                                {bankTotals.chargeback > 0 ? "−" + formatINR(bankTotals.chargeback) : "—"}
                              </td>
                              <td className="px-3 py-2.5 text-right font-mono font-bold tabular-nums text-blue-900">{formatINR(bankTotals.netSettle)}</td>
                              <td className="px-3 py-2.5 text-right font-mono font-bold tabular-nums text-emerald-700">{formatINR(bankTotals.ourNet)}</td>
                              <td className={cn("px-3 py-2.5 text-right font-mono font-bold tabular-nums", Math.abs(bankTotals.diff) <= 1 ? "text-emerald-700" : "text-red-700")}>
                                {bankTotals.diff > 0 ? "+" : ""}{formatINR(bankTotals.diff)}
                              </td>
                            </>
                          )}
                        </tr>
                      </tbody>
                    </table>
                  </div>
                )}
              </CardContent>
            </Card>

            <Card className="border-border/70 shadow-xs">
              <CardHeader className="border-b border-border/60 pb-4">
                <div className="flex items-center gap-3">
                  <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-blue-100 text-blue-700">
                    <CheckCircle2 className="h-5 w-5" />
                  </div>
                  <div>
                    <CardTitle className="text-[15px] font-semibold tracking-tight text-foreground">
                      Verify Transaction
                    </CardTitle>
                    <CardDescription className="text-xs text-muted-foreground">
                      Cross-check a single transaction against bank settlement
                    </CardDescription>
                  </div>
                </div>
              </CardHeader>
              <CardContent className="space-y-4 pt-5">
                <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 md:grid-cols-3">
                  <div className="space-y-2">
                    <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">Merchant</Label>
                    <Select
                      value={rMerchant || undefined}
                      onValueChange={(v) => { setRM(v); setRC(""); }}
                    >
                      <SelectTrigger className="h-10 w-full">
                        <SelectValue placeholder="— Select —" />
                      </SelectTrigger>
                      <SelectContent>
                        {ALL_MERCHANTS.map((m) => (
                          <SelectItem key={m} value={m}>{m}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="space-y-2">
                    <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">Company</Label>
                    <Select
                      value={rCompany || undefined}
                      onValueChange={setRC}
                      disabled={!rMerchant}
                    >
                      <SelectTrigger className="h-10 w-full">
                        <SelectValue placeholder={rMerchant ? "— Select —" : "Select merchant first"} />
                      </SelectTrigger>
                      <SelectContent>
                        {rCompsFor.map((c) => (
                          <SelectItem key={c} value={c}>{c}</SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </div>
                  <div className="space-y-2">
                    <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">Transaction ID</Label>
                    <Input
                      value={rTxnId}
                      onChange={(e) => setRTxn(e.target.value)}
                      placeholder="TXN-001"
                      className="h-10 font-mono"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-5">
                  <div className="space-y-2">
                    <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">Payin (₹)</Label>
                    <div className="relative">
                      <IndianRupee className="pointer-events-none absolute left-3 top-1/2 h-3.5 w-3.5 -translate-y-1/2 text-muted-foreground/70" />
                      <Input
                        type="number"
                        value={rPayin}
                        onChange={(e) => setRPay(e.target.value)}
                        placeholder="10000"
                        className="h-10 pl-8 font-mono font-semibold tabular-nums"
                      />
                    </div>
                  </div>
                  <div className="space-y-2">
                    <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">Bank Settle (₹)</Label>
                    <div className="relative">
                      <IndianRupee className="pointer-events-none absolute left-3 top-1/2 h-3.5 w-3.5 -translate-y-1/2 text-muted-foreground/70" />
                      <Input
                        type="number"
                        value={rClaimed}
                        onChange={(e) => setRClaimed(e.target.value)}
                        placeholder="9950"
                        className="h-10 pl-8 font-mono font-semibold tabular-nums"
                      />
                    </div>
                  </div>
                  <div className="space-y-2">
                    <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">UTR Number</Label>
                    <Input
                      value={rUTR}
                      onChange={(e) => setRUTR(e.target.value)}
                      placeholder="UTR123456"
                      className="h-10 font-mono"
                    />
                  </div>
                  <div className="space-y-2">
                    <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">Date</Label>
                    <div className="relative">
                      <CalendarIcon className="pointer-events-none absolute left-3 top-1/2 h-3.5 w-3.5 -translate-y-1/2 text-muted-foreground/70" />
                      <Input
                        type="date"
                        value={rDate}
                        onChange={(e) => setRDate(e.target.value)}
                        className="h-10 pl-8"
                      />
                    </div>
                  </div>
                  <div className="space-y-2">
                    <Label className="text-[11px] font-semibold uppercase tracking-wider text-muted-foreground">Cycle</Label>
                    <Select value={rCycle} onValueChange={setRCycle}>
                      <SelectTrigger className="h-10 w-full">
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="1">Cycle 1</SelectItem>
                        <SelectItem value="2">Cycle 2</SelectItem>
                      </SelectContent>
                    </Select>
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
                    const isMatch = Math.abs(df) <= 1;
                    return (
                      <div className="flex flex-wrap items-center gap-5 rounded-lg border border-border bg-muted/40 px-5 py-4">
                        <div>
                          <div className="text-[9.5px] font-semibold uppercase tracking-wider text-muted-foreground">Our Settlement</div>
                          <div className="font-mono text-base font-bold tabular-nums text-emerald-700">{formatINR(os)}</div>
                        </div>
                        {rc2 > 0 && (
                          <>
                            <Separator orientation="vertical" className="h-9" />
                            <div>
                              <div className="text-[9.5px] font-semibold uppercase tracking-wider text-muted-foreground">Bank Settlement</div>
                              <div className="font-mono text-base font-bold tabular-nums text-blue-700">{formatINR(rc2)}</div>
                            </div>
                            <Separator orientation="vertical" className="h-9" />
                            <div>
                              <div className="text-[9.5px] font-semibold uppercase tracking-wider text-muted-foreground">Difference</div>
                              <div className={cn("font-mono text-base font-bold tabular-nums", isMatch ? "text-emerald-700" : "text-red-700")}>
                                {df > 0 ? "+" : ""}{formatINR(df)}
                              </div>
                            </div>
                            <Badge
                              className={cn(
                                "ml-auto font-bold tracking-wide",
                                isMatch
                                  ? "bg-emerald-100 text-emerald-800 hover:bg-emerald-100"
                                  : "bg-red-100 text-red-800 hover:bg-red-100"
                              )}
                            >
                              {isMatch ? <CheckCircle2 className="mr-1 h-3.5 w-3.5" /> : <AlertCircle className="mr-1 h-3.5 w-3.5" />}
                              {isMatch ? "MATCH" : "MISMATCH"}
                            </Badge>
                          </>
                        )}
                      </div>
                    );
                  })()}

                <Button
                  size="lg"
                  onClick={addRecon}
                  disabled={!(parseFloat(rPayin) > 0 && rMerchant && rCompany && rTxnId)}
                  className="h-11 px-7"
                >
                  <Plus className="mr-1 h-4 w-4" />
                  Add to Reconciliation
                </Button>
              </CardContent>
            </Card>

            {recon.length > 0 && (
              <Card className="border-border/70 shadow-xs">
                <CardContent className="flex flex-wrap items-center gap-2.5 px-5 py-3">
                  <div className="-mx-1 max-w-full overflow-x-auto px-1">
                    <Tabs value={rFilter} onValueChange={setRFilter}>
                      <TabsList>
                        <TabsTrigger value="all">All</TabsTrigger>
                        <TabsTrigger value="Matched">Matched</TabsTrigger>
                        <TabsTrigger value="Mismatch">Mismatch</TabsTrigger>
                        <TabsTrigger value="Pending">Pending</TabsTrigger>
                        <TabsTrigger value="Resolved">Resolved</TabsTrigger>
                      </TabsList>
                    </Tabs>
                  </div>
                  <Select value={rFM} onValueChange={setRFM}>
                    <SelectTrigger className="h-9 w-[180px]">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__all__">All Merchants</SelectItem>
                      {[...new Set(recon.map((r) => r.merchant))]
                        .sort()
                        .map((m) => (
                          <SelectItem key={m} value={m}>{m}</SelectItem>
                        ))}
                    </SelectContent>
                  </Select>
                  <Button
                    size="sm"
                    variant="outline"
                    onClick={() => {
                      setRecon([]);
                      setRReason({});
                      setRNotes({});
                      setRResolved({});
                    }}
                    className="ml-auto border-destructive/30 text-destructive hover:bg-destructive/5 hover:text-destructive"
                  >
                    <Trash2 className="mr-1 h-3.5 w-3.5" />
                    Clear all
                  </Button>
                </CardContent>
              </Card>
            )}

            {reconFiltered.length > 0 && (
              <Card className="overflow-hidden border-border/70 shadow-xs">
                <div className="flex items-center justify-between border-b border-border/70 px-5 py-3.5">
                  <div>
                    <div className="text-sm font-semibold tracking-tight text-foreground">
                      Reconciliation Entries
                    </div>
                    <div className="text-xs text-muted-foreground">{reconFiltered.length} rows</div>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead className="bg-muted/40">
                      <tr>
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
                            className={cn(
                              "whitespace-nowrap px-3 py-2.5 text-[9.5px] font-semibold uppercase tracking-wider text-muted-foreground",
                              i <= 2 ? "text-left" : "text-right"
                            )}
                          >
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {reconFiltered.map((r) => {
                        const badge = statusBadge(r.status, rResolved[r.id]);
                        return (
                          <tr
                            key={r.id}
                            className="border-t border-border/60 transition-colors hover:bg-muted/30"
                          >
                            <td className="px-3 py-2.5 font-mono text-xs font-semibold text-foreground">{r.txnId}</td>
                            <td className="px-3 py-2.5 font-semibold text-primary">{r.merchant}</td>
                            <td
                              className="max-w-[160px] overflow-hidden text-ellipsis whitespace-nowrap px-3 py-2.5 text-foreground"
                              title={r.company}
                            >
                              {r.company}
                            </td>
                            <td className="px-3 py-2.5 text-right font-mono tabular-nums text-foreground">{formatINR(r.payin)}</td>
                            <td className="px-3 py-2.5 text-right font-mono font-semibold tabular-nums text-emerald-700">{formatINR(r.ourSettle)}</td>
                            <td className={cn("px-3 py-2.5 text-right font-mono font-semibold tabular-nums", r.claimed > 0 ? "text-blue-700" : "text-muted-foreground")}>
                              {r.claimed > 0 ? formatINR(r.claimed) : "\u2014"}
                            </td>
                            <td className={cn(
                              "px-3 py-2.5 text-right font-mono font-bold tabular-nums",
                              r.status === "Mismatch" ? "text-red-700" : r.status === "Matched" ? "text-emerald-700" : "text-amber-600"
                            )}>
                              {r.status === "Mismatch"
                                ? (r.diff > 0 ? "+" : "") + formatINR(r.diff)
                                : r.status === "Matched"
                                  ? "\u20B90.00"
                                  : "\u2014"}
                            </td>
                            <td className="px-3 py-2.5 text-right">
                              <span
                                className="inline-flex items-center rounded-md px-2 py-0.5 text-[10px] font-bold"
                                style={{ background: badge.bg, color: badge.color }}
                              >
                                {badge.label}
                              </span>
                            </td>
                            <td className="px-3 py-2.5 text-right">
                              {r.status === "Mismatch" && !rResolved[r.id] && (
                                <Select
                                  value={rReason[r.id] || undefined}
                                  onValueChange={(v) => setRReason((p) => ({ ...p, [r.id]: v }))}
                                >
                                  <SelectTrigger className="h-7 w-[130px] text-[11px]">
                                    <SelectValue placeholder="Reason..." />
                                  </SelectTrigger>
                                  <SelectContent>
                                    {REASONS.map((re) => (
                                      <SelectItem key={re} value={re}>{re}</SelectItem>
                                    ))}
                                  </SelectContent>
                                </Select>
                              )}
                            </td>
                            <td className="whitespace-nowrap px-2 py-2.5 text-center">
                              {!rResolved[r.id] && r.status === "Mismatch" && rReason[r.id] && (
                                <Button
                                  size="sm"
                                  variant="outline"
                                  onClick={() => setRResolved((p) => ({ ...p, [r.id]: true }))}
                                  className="mr-1 h-7 border-blue-300 px-2.5 text-[10px] font-bold text-blue-700 hover:bg-blue-50"
                                >
                                  Resolve
                                </Button>
                              )}
                              <Button
                                variant="ghost"
                                size="icon"
                                onClick={() => rmRecon(r.id)}
                                className="h-7 w-7 text-muted-foreground/60 hover:bg-destructive/10 hover:text-destructive"
                              >
                                <X className="h-3.5 w-3.5" />
                              </Button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </Card>
            )}

            {recon.length === 0 && (
              <Card className="border-dashed border-border/80 bg-background/40 shadow-none">
                <CardContent className="flex flex-col items-center justify-center px-6 py-14 text-center">
                  <div className="mb-3 flex h-12 w-12 items-center justify-center rounded-2xl bg-primary/10 text-primary">
                    <Search className="h-6 w-6" />
                  </div>
                  <div className="text-base font-semibold text-foreground">No reconciliation entries</div>
                  <div className="mt-1 max-w-sm text-sm text-muted-foreground">
                    Add transactions above to verify settlement amounts against merchant claims.
                  </div>
                </CardContent>
              </Card>
            )}
          </div>
        )}

        {/* ========== MERCHANTS ========== */}
        {page === "merchants" && (
          <div className="fade-up space-y-5">
            <Card className="border-border/70 shadow-xs">
              <CardContent className="flex flex-wrap items-center gap-3 px-5 py-4">
                <div className="relative min-w-[260px] flex-1">
                  <Search className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-muted-foreground/70" />
                  <Input
                    value={mSearch}
                    onChange={(e) => setMSearch(e.target.value)}
                    placeholder="Search by company, MID, or GST number..."
                    className="h-10 pl-9"
                  />
                </div>
                <Select value={mFilter} onValueChange={setMFilter}>
                  <SelectTrigger className="h-10 w-[220px]">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__all__">All Merchants ({MASTER_DATA.length})</SelectItem>
                    {ALL_MERCHANTS.map((m) => (
                      <SelectItem key={m} value={m}>
                        {m} ({MASTER_DATA.filter((d) => d.merchant === m).length})
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </CardContent>
            </Card>

            <div className="grid gap-4" style={{ gridTemplateColumns: "repeat(auto-fill,minmax(min(100%,400px),1fr))" }}>
              {ALL_MERCHANTS.filter((m) =>
                mFilter === "__all__" ? true : m === mFilter
              ).map((m) => {
                const comps = MASTER_DATA.filter((d) => d.merchant === m).filter((c) => {
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
                  <Card key={m} className="overflow-hidden border-border/70 shadow-xs">
                    <div className="flex items-center justify-between gap-3 border-b border-border/60 bg-gradient-to-br from-primary/8 to-primary/0 px-5 py-3.5">
                      <div className="flex items-center gap-3">
                        <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-gradient-to-br from-violet-500 to-violet-700 text-xs font-bold text-white shadow-sm">
                          {m.substring(0, 2).toUpperCase()}
                        </div>
                        <div>
                          <div className="text-sm font-semibold tracking-tight text-foreground">
                            {m}
                          </div>
                          <div className="flex items-center gap-1.5 text-[11px] text-muted-foreground">
                            <Users className="h-3 w-3" /> {comps.length} {comps.length === 1 ? "company" : "companies"}
                          </div>
                        </div>
                      </div>
                      {entries.length > 0 && (
                        <div className="text-right">
                          <div className="text-[9.5px] font-semibold uppercase tracking-wider text-muted-foreground">
                            Settled
                          </div>
                          <div className="font-mono text-sm font-bold tabular-nums text-emerald-700">
                            {formatShort(mT.settlement)}
                          </div>
                        </div>
                      )}
                    </div>
                    <div className="max-h-[340px] overflow-y-auto">
                      {comps.map((c, i) => (
                        <div
                          key={i}
                          className="border-b border-border/50 px-4 py-2.5 transition-colors last:border-b-0 hover:bg-muted/40"
                        >
                          <div
                            className="overflow-hidden text-ellipsis whitespace-nowrap text-xs font-semibold text-foreground"
                            title={c.company}
                          >
                            {c.company}
                          </div>
                          <div className="mt-1 flex flex-wrap items-center gap-1.5">
                            {c.mid && (
                              <span
                                className="rounded bg-muted px-1.5 py-0.5 font-mono text-[10px] text-muted-foreground"
                                title="Merchant ID"
                              >
                                {c.mid}
                              </span>
                            )}
                            {c.gstNumber && (
                              <span
                                className="rounded bg-blue-50 px-1.5 py-0.5 font-mono text-[10px] text-blue-700"
                                title="GST Number"
                              >
                                {c.gstNumber}
                              </span>
                            )}
                            {c.incorporation && (
                              <span
                                className="flex items-center gap-1 text-[10px] text-muted-foreground"
                                title="Incorporation date"
                              >
                                <CalendarIcon className="h-2.5 w-2.5" /> {c.incorporation}
                              </span>
                            )}
                          </div>
                        </div>
                      ))}
                    </div>
                  </Card>
                );
              })}
            </div>
          </div>
        )}

        {/* ========== AUTOMATION ========== */}
        {page === "automation" && <Automation />}
      </div>
    </div>
  );
}
