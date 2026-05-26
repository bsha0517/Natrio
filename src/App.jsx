import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { db, isConnected } from "./supabase.js";

/* ── Palette ── */
const C = {
  bg:"#F3F4F6",surface:"#FFFFFF",border:"#E5E7EB",
  accent:"#2563EB",accentSoft:"#DBEAFE",
  green:"#16A34A",greenSoft:"#DCFCE7",
  amber:"#D97706",amberSoft:"#FEF3C7",
  red:"#DC2626",redSoft:"#FEE2E2",
  purple:"#7C3AED",purpleSoft:"#EDE9FE",
  teal:"#0D9488",tealSoft:"#CCFBF1",
  text:"#111827",muted:"#6B7280",subtle:"#9CA3AF",
};

/* ── Helpers ── */
const fmt = n => "PKR " + new Intl.NumberFormat("en-PK",{maximumFractionDigits:0}).format(n||0);
const today = () => new Date().toISOString().slice(0,10);
const uid = () => Date.now() + Math.floor(Math.random()*9999);
const EXP_CATS = ["Rent","Utilities","Transport","Salaries","Marketing","Maintenance","Other"];
const ORDER_STATUSES = ["Invoiced","Paid","Delivered","Due","Pending","Cancelled"];
const DEL_STATUSES = ["Dispatched","En Route","Delivered","Failed"];
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

/* ── Mobile hook ── */
function useIsMobile() {
  const [mobile, setMobile] = useState(window.innerWidth < 768);
  useEffect(() => {
    const h = () => setMobile(window.innerWidth < 768);
    window.addEventListener("resize", h);
    return () => window.removeEventListener("resize", h);
  }, []);
  return mobile;
}

/* ── Persistent audit log (localStorage) ── */
const AUDIT_KEY = "natrio_audit_log";
function loadAudit() {
  try { return JSON.parse(localStorage.getItem(AUDIT_KEY)||"[]"); } catch { return []; }
}
function saveAudit(log) {
  try { localStorage.setItem(AUDIT_KEY, JSON.stringify(log.slice(0,1000))); } catch {}
}

/* ── Excel helpers ── */
function exportToExcel(rows, filename) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Data");
  XLSX.writeFile(wb, filename);
}
function importFromExcel(file, cb) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, {type:"binary"});
      const ws = wb.Sheets[wb.SheetNames[0]];
      cb(XLSX.utils.sheet_to_json(ws));
    } catch { alert("Could not parse file. Make sure it's a valid Excel (.xlsx) file."); }
  };
  reader.readAsBinaryString(file);
}

/* ── Seed data ── */
const seedInventory = [];
const seedCustomers = [];
const seedSales = [];
const seedSuppliers = [];
const seedPOs = [];
const seedExpenses = [];
const seedDeliveries = [];
const seedUsers = [];
const DEFAULT_SETTINGS = {
  companyName:"Natrio Organics",tagline:"Distribution Suite",
  address1:"",address2:"",city:"",country:"Pakistan",
  phone:"",email:"",website:"",
  taxLabel:"NTN",taxNumber:"",taxLabel2:"STRN",taxNumber2:"",
  bankName:"",bankAccount:"",bankIBAN:"",
  invoiceNotes:"Thank you for your business.",logoDataUrl:"",
};
const ROLE_TABS = {
  admin:["dashboard","pos","orders","quotations","inventory","purchase","crm","delivery","expenses","reports","analytics","ai","users","audit","settings"],
  sales:["dashboard","pos","orders","quotations","crm","delivery","ai"],
  viewer:["dashboard","orders","reports","analytics"],
};
const ALL_NAV = [
  {id:"dashboard",label:"Dashboard",icon:"▦"},
  {id:"pos",label:"Sales / POS",icon:"⊕"},
  {id:"orders",label:"Orders",icon:"≡"},
  {id:"quotations",label:"Quotations",icon:"📋"},
  {id:"inventory",label:"Inventory",icon:"⊞"},
  {id:"purchase",label:"Purchase Orders",icon:"🛒"},
  {id:"crm",label:"Customers",icon:"◎"},
  {id:"delivery",label:"Deliveries",icon:"🚚"},
  {id:"expenses",label:"Expenses",icon:"💸"},
  {id:"reports",label:"Reports",icon:"◈"},
  {id:"analytics",label:"Analytics",icon:"📊"},
  {id:"ai",label:"AI Assistant",icon:"✦"},
  {id:"users",label:"Users",icon:"◉"},
  {id:"audit",label:"Audit Log",icon:"🔍"},
  {id:"settings",label:"Settings",icon:"⚙"},
];

/* ── Shared UI ── */
const Badge = ({label}) => {
  const map={
    Gold:["#D97706","#FEF3C7"],Platinum:["#7C3AED","#EDE9FE"],Silver:["#6B7280","#F3F4F6"],Bronze:["#92400E","#FEF3C7"],
    Paid:["#16A34A","#DCFCE7"],Invoiced:["#2563EB","#DBEAFE"],Delivered:["#0D9488","#CCFBF1"],
    Due:["#DC2626","#FEE2E2"],Pending:["#D97706","#FEF3C7"],Cancelled:["#6B7280","#F3F4F6"],
    Received:["#16A34A","#DCFCE7"],"En Route":["#2563EB","#DBEAFE"],Dispatched:["#D97706","#FEF3C7"],Failed:["#DC2626","#FEE2E2"],
    Low:["#DC2626","#FEE2E2"],OK:["#16A34A","#DCFCE7"],
    Admin:["#7C3AED","#EDE9FE"],Sales:["#16A34A","#DCFCE7"],Viewer:["#D97706","#FEF3C7"],
    create:["#16A34A","#DCFCE7"],update:["#2563EB","#DBEAFE"],delete:["#DC2626","#FEE2E2"],login:["#7C3AED","#EDE9FE"],
    Draft:["#6B7280","#F3F4F6"],Sent:["#2563EB","#DBEAFE"],Converted:["#16A34A","#DCFCE7"],
    Rent:["#7C3AED","#EDE9FE"],Utilities:["#0D9488","#CCFBF1"],Transport:["#D97706","#FEF3C7"],
    Salaries:["#DC2626","#FEE2E2"],Marketing:["#2563EB","#DBEAFE"],Maintenance:["#92400E","#FEF3C7"],Other:["#6B7280","#F3F4F6"],
  };
  const [fg,bg]=map[label]||["#6B7280","#F3F4F6"];
  return <span style={{background:bg,color:fg,border:`1px solid ${fg}22`,borderRadius:4,padding:"2px 6px",fontSize:10,fontWeight:700,letterSpacing:.5,textTransform:"uppercase",whiteSpace:"nowrap"}}>{label}</span>;
};

const StatCard = ({label,value,sub,color,icon}) => (
  <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:10,padding:"13px 14px",flex:1,boxShadow:"0 1px 3px #0000000a",minWidth:0}}>
    <div style={{display:"flex",justifyContent:"space-between"}}>
      <div style={{fontSize:10,color:C.muted,textTransform:"uppercase",letterSpacing:1.5,fontWeight:700}}>{label}</div>
      {icon&&<span style={{fontSize:14}}>{icon}</span>}
    </div>
    <div style={{fontSize:18,fontWeight:800,color:color||C.text,fontFamily:"'DM Mono',monospace",letterSpacing:-1,marginTop:5,wordBreak:"break-all"}}>{value}</div>
    {sub&&<div style={{fontSize:10,color:C.muted,marginTop:2}}>{sub}</div>}
  </div>
);

/* Fixed Input — controlled without re-mount bug */
const Input = ({style,...p}) => (
  <input style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:7,padding:"8px 11px",color:C.text,fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",fontFamily:"inherit",...style}} {...p}/>
);
const Textarea = ({style,...p}) => (
  <textarea style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:7,padding:"8px 11px",color:C.text,fontSize:13,outline:"none",width:"100%",boxSizing:"border-box",fontFamily:"inherit",resize:"vertical",...style}} {...p}/>
);
const Sel = ({value,onChange,children,style}) => (
  <select value={value} onChange={onChange} style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:7,padding:"8px 11px",color:C.text,fontSize:13,fontFamily:"inherit",outline:"none",...style}}>{children}</select>
);
const Btn = ({children,onClick,variant="primary",style,disabled,title}) => {
  const s={
    primary:{background:C.accent,color:"#fff",border:"none"},
    ghost:{background:"transparent",color:C.muted,border:`1px solid ${C.border}`},
    danger:{background:C.redSoft,color:C.red,border:`1px solid ${C.red}33`},
    success:{background:C.greenSoft,color:C.green,border:`1px solid ${C.green}33`},
    purple:{background:C.purpleSoft,color:C.purple,border:`1px solid ${C.purple}33`},
    teal:{background:C.tealSoft,color:C.teal,border:`1px solid ${C.teal}33`},
  };
  return <button title={title} onClick={onClick} disabled={disabled} style={{...s[variant],borderRadius:7,padding:"7px 12px",fontSize:12,fontWeight:600,cursor:disabled?"not-allowed":"pointer",fontFamily:"inherit",opacity:disabled?.5:1,whiteSpace:"nowrap",...style}}>{children}</button>;
};
const Card = ({children,style,title,action}) => (
  <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:10,padding:15,boxShadow:"0 1px 3px #0000000a",...style}}>
    {title&&<div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:11}}>
      <div style={{fontSize:10,fontWeight:800,color:C.muted,textTransform:"uppercase",letterSpacing:1.5}}>{title}</div>{action}</div>}
    {children}
  </div>
);
const Th = ({children}) => <th style={{textAlign:"left",padding:"7px 9px",fontSize:10,color:C.muted,textTransform:"uppercase",letterSpacing:1,fontWeight:700,background:C.bg,whiteSpace:"nowrap"}}>{children}</th>;
const Td = ({children,style}) => <td style={{padding:"8px 9px",fontSize:12,color:C.text,borderBottom:`1px solid ${C.border}`,...style}}>{children}</td>;

/* Tooltip wrapper */
const Tip = ({label,children}) => {
  const [show,setShow]=useState(false);
  return (
    <div style={{position:"relative",display:"inline-flex"}} onMouseEnter={()=>setShow(true)} onMouseLeave={()=>setShow(false)}>
      {children}
      {show&&<div style={{position:"absolute",left:"110%",top:"50%",transform:"translateY(-50%)",background:"#1f2937",color:"#fff",padding:"4px 9px",borderRadius:6,fontSize:11,fontWeight:600,whiteSpace:"nowrap",pointerEvents:"none",zIndex:999}}>{label}</div>}
    </div>
  );
};

/* Import/Export bar — reusable across tabs */
function ImportExportBar({exportFn,importFn,label="data"}) {
  const ref = useRef();
  return (
    <div style={{display:"flex",gap:7,alignItems:"center"}}>
      <Btn onClick={exportFn} variant="ghost" style={{fontSize:11,padding:"5px 10px"}} title={`Export ${label} to Excel`}>⬇ Excel</Btn>
      {importFn&&<>
        <Btn onClick={()=>ref.current.click()} variant="ghost" style={{fontSize:11,padding:"5px 10px"}} title={`Import ${label} from Excel`}>⬆ Import</Btn>
        <input ref={ref} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>{if(e.target.files[0])importFn(e.target.files[0]);e.target.value="";}}/>
      </>}
    </div>
  );
}

/* ── Print helpers ── */
function printDoc(order, customers, settings={}, isQuote=false) {
  const s={...DEFAULT_SETTINGS,...settings};
  const cust=customers.find(c=>c.id===order.customerId)||{};
  const logo=s.logoDataUrl?`<img src="${s.logoDataUrl}" style="height:58px;max-width:155px;object-fit:contain;display:block;margin-bottom:4px"/>`:
    `<h1 style="font-size:19px;color:#2d5a27;font-weight:900;margin:0 0 3px">${s.companyName}</h1>`;
  const from=[s.address1,s.address2,s.city,s.country,s.phone,s.email,s.taxNumber?`${s.taxLabel}: ${s.taxNumber}`:"",s.taxNumber2?`${s.taxLabel2}: ${s.taxNumber2}`:""].filter(Boolean);
  const bank=[s.bankName,s.bankAccount?`Acc: ${s.bankAccount}`:"",s.bankIBAN?`IBAN: ${s.bankIBAN}`:""].filter(Boolean);
  const paid=(order.payments||[]).reduce((a,p)=>a+p.amount,0);
  const due=order.total-paid;
  const title=isQuote?`QUOTATION ${order.id}`:`INVOICE #${order.id}`;
  const w=window.open("","_blank");
  w.document.write(`<!DOCTYPE html><html><head><title>${title}</title>
<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Segoe UI',sans-serif;color:#111;padding:32px;max-width:760px;margin:auto}
.h{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:24px;padding-bottom:14px;border-bottom:3px solid #2d5a27}
.m{text-align:right}.m h2{font-size:19px;font-weight:900;color:#2d5a27}.m p{color:#555;font-size:11px;margin-top:3px}
.p2{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:18px}
.p h4{font-size:9px;color:#888;text-transform:uppercase;letter-spacing:2px;margin-bottom:5px;font-weight:700;padding-bottom:2px;border-bottom:1px solid #e5e7eb}
.p p{font-size:11px;color:#374151;margin-bottom:2px;line-height:1.4}.nm{font-size:13px;font-weight:700;color:#111;margin-bottom:3px}
table{width:100%;border-collapse:collapse;margin-bottom:16px}
thead tr{background:#2d5a27;color:#fff}thead th{padding:8px 10px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.8px}
tbody tr{border-bottom:1px solid #e5e7eb}tbody tr:nth-child(even){background:#f9fafb}
tbody td{padding:7px 10px;font-size:11px;color:#374151}.num{font-family:monospace;text-align:right}
.bot{display:grid;grid-template-columns:1fr auto;gap:14px;margin-bottom:18px;align-items:start}
.bk h4{font-size:9px;color:#888;text-transform:uppercase;letter-spacing:2px;margin-bottom:4px;font-weight:700}.bk p{font-size:11px;color:#374151;margin-bottom:2px}
.tot{min-width:210px}.tr{display:flex;justify-content:space-between;padding:5px 0;font-size:11px;border-bottom:1px solid #e5e7eb;color:#4b5563}
.tr.g{border-top:2px solid #2d5a27;border-bottom:none;margin-top:3px;padding-top:8px;font-size:15px;font-weight:800;color:#111}
.tr.due{color:#dc2626;font-weight:700}.sb{display:inline-block;padding:2px 7px;border-radius:20px;font-size:9px;font-weight:700;text-transform:uppercase}
.ft{text-align:center;color:#9ca3af;font-size:10px;padding-top:12px;border-top:1px solid #e5e7eb;line-height:1.7}
@media print{body{padding:12px}}</style></head><body>
<div class="h"><div>${logo}<p style="color:#777;font-size:9px">${s.tagline}</p></div>
<div class="m"><h2>${title}</h2><p>Date: <strong>${order.date}</strong></p>
${!isQuote?`<p>Status: <span class="sb" style="background:${order.status==="Paid"?"#dcfce7":order.status==="Due"?"#fee2e2":"#dbeafe"};color:${order.status==="Paid"?"#166534":order.status==="Due"?"#991b1b":"#1e40af"}">${order.status}</span></p>`:""}
${order.validUntil?`<p>Valid Until: <strong>${order.validUntil}</strong></p>`:""}</div></div>
<div class="p2"><div class="p"><h4>From</h4><p class="nm">${s.companyName}</p>${from.map(l=>`<p>${l}</p>`).join("")}</div>
<div class="p"><h4>${isQuote?"Prepared For":"Bill To"}</h4><p class="nm">${order.customer}</p>${cust.email?`<p>${cust.email}</p>`:""}${cust.phone?`<p>${cust.phone}</p>`:""}${cust.address?`<p>${cust.address}</p>`:""}</div></div>
<table><thead><tr><th>#</th><th>Product</th><th>SKU</th><th style="text-align:right">Qty</th><th style="text-align:right">Unit Price</th><th style="text-align:right">Amount</th></tr></thead>
<tbody>${order.items.map((it,i)=>`<tr><td>${i+1}</td><td>${it.name}</td><td style="font-family:monospace;color:#6b7280;font-size:10px">${it.sku||""}</td><td class="num">${it.qty}</td><td class="num">PKR ${(it.price||0).toLocaleString()}</td><td class="num">PKR ${((it.qty||0)*(it.price||0)).toLocaleString()}</td></tr>`).join("")}</tbody></table>
<div class="bot"><div class="bk">${bank.length?`<h4>Payment Details</h4>${bank.map(l=>`<p>${l}</p>`).join("")}`:""}</div>
<div class="tot"><div class="tr"><span>Subtotal</span><span>PKR ${order.total.toLocaleString()}</span></div>
${!isQuote&&paid>0?`<div class="tr"><span>Paid</span><span>PKR ${paid.toLocaleString()}</span></div>`:""}
${!isQuote&&due>0?`<div class="tr due"><span>Balance Due</span><span>PKR ${due.toLocaleString()}</span></div>`:""}
<div class="tr g"><span>Total</span><span>PKR ${order.total.toLocaleString()}</span></div></div></div>
${order.note?`<p style="font-size:10px;color:#555;margin-bottom:12px">Note: ${order.note}</p>`:""}
<div class="ft"><p>${s.invoiceNotes}</p>${s.companyName?`<p style="margin-top:3px;font-weight:600">${s.companyName}</p>`:""}</div>
<script>window.onload=()=>{window.print();}<\/script></body></html>`);
  w.document.close();
}
function printPackingSlip(order){
  const w=window.open("","_blank");
  w.document.write(`<!DOCTYPE html><html><head><title>Packing Slip</title>
<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Segoe UI',sans-serif;padding:26px;max-width:560px;margin:auto}
h1{font-size:16px;font-weight:900;margin-bottom:3px}p{font-size:11px;color:#555;margin-bottom:10px}
table{width:100%;border-collapse:collapse}th{background:#111;color:#fff;padding:7px 9px;text-align:left;font-size:9px;text-transform:uppercase}
td{padding:7px 9px;border-bottom:1px solid #e5e7eb;font-size:11px}.ck{width:18px;height:18px;border:2px solid #ccc;border-radius:2px;display:inline-block}
@media print{body{padding:10px}}</style></head><body>
<h1>PACKING SLIP</h1><p>Order #${order.id} · ${order.date} · <strong>${order.customer}</strong></p>
<table><thead><tr><th>✓</th><th>SKU</th><th>Product</th><th style="text-align:right">Qty</th></tr></thead>
<tbody>${order.items.map(it=>`<tr><td><span class="ck"></span></td><td style="font-family:monospace;color:#555;font-size:10px">${it.sku||""}</td><td>${it.name}</td><td style="text-align:right;font-weight:700">${it.qty}</td></tr>`).join("")}</tbody></table>
<p style="margin-top:14px;font-size:9px;color:#888">Packed by: ___________________ &nbsp;&nbsp; Date: ___________________</p>
<script>window.onload=()=>{window.print();}<\/script></body></html>`);
  w.document.close();
}
function printReportPDF(title, rows, columns) {
  const w=window.open("","_blank");
  w.document.write(`<!DOCTYPE html><html><head><title>${title}</title>
<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Segoe UI',sans-serif;padding:28px;color:#111}
h1{font-size:18px;font-weight:900;margin-bottom:4px;color:#2d5a27}p{font-size:11px;color:#555;margin-bottom:14px}
table{width:100%;border-collapse:collapse}
thead tr{background:#2d5a27;color:#fff}thead th{padding:8px 10px;text-align:left;font-size:10px;text-transform:uppercase}
tbody tr{border-bottom:1px solid #e5e7eb}tbody tr:nth-child(even){background:#f9fafb}
tbody td{padding:7px 10px;font-size:11px}
.footer{margin-top:16px;font-size:10px;color:#9ca3af;text-align:center}
@media print{body{padding:14px}}</style></head><body>
<h1>${title}</h1><p>Generated: ${new Date().toLocaleString()} &nbsp;·&nbsp; Natrio Organics</p>
<table><thead><tr>${columns.map(c=>`<th>${c}</th>`).join("")}</tr></thead>
<tbody>${rows.map(r=>`<tr>${r.map(v=>`<td>${v}</td>`).join("")}</tr>`).join("")}</tbody></table>
<div class="footer">Natrio Organics — Distribution Suite</div>
<script>window.onload=()=>{window.print();}<\/script></body></html>`);
  w.document.close();
}

/* ── Login ── */
function LoginScreen({users,onLogin}){
  const [u,setU]=useState("");const [p,setP]=useState("");const [err,setErr]=useState("");const [shake,setShake]=useState(false);
  const go=()=>{const f=users.find(x=>x.username===u.trim().toLowerCase()&&x.password===p);if(f)onLogin(f);else{setErr("Invalid username or password.");setShake(true);setTimeout(()=>setShake(false),600);}};
  return(
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:C.bg,fontFamily:"'DM Sans','Segoe UI',sans-serif",padding:16}}>
      <div style={{position:"absolute",inset:0,backgroundImage:`linear-gradient(${C.border} 1px,transparent 1px),linear-gradient(90deg,${C.border} 1px,transparent 1px)`,backgroundSize:"40px 40px",opacity:.7}}/>
      <div style={{position:"relative",width:"100%",maxWidth:380,background:C.surface,border:`1px solid ${C.border}`,borderRadius:18,padding:"28px 24px",transform:shake?"translateX(-6px)":"none",transition:shake?"transform .1s":"transform .2s",boxShadow:"0 8px 32px #00000014"}}>
        <div style={{textAlign:"center",marginBottom:20}}>
          <img src="/logo.jpeg" alt="Natrio" style={{width:88,height:88,objectFit:"contain",marginBottom:8,borderRadius:10}}/>
          <div style={{fontSize:11,color:C.muted}}>Sign in to your account</div>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:9}}>
          <div><div style={{fontSize:10,color:C.muted,marginBottom:4,fontWeight:700,textTransform:"uppercase",letterSpacing:1}}>Username</div><Input value={u} onChange={e=>setU(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()} placeholder="Enter username"/></div>
          <div><div style={{fontSize:10,color:C.muted,marginBottom:4,fontWeight:700,textTransform:"uppercase",letterSpacing:1}}>Password</div><Input type="password" value={p} onChange={e=>setP(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()} placeholder="Enter password"/></div>
          {err&&<div style={{background:C.redSoft,border:`1px solid ${C.red}44`,borderRadius:7,padding:"7px 11px",fontSize:12,color:C.red}}>{err}</div>}
          <Btn onClick={go} style={{width:"100%",padding:11,fontSize:13,marginTop:2}}>Sign In</Btn>
        </div>
        <div style={{marginTop:16,padding:"10px 12px",background:C.bg,border:`1px solid ${C.border}`,borderRadius:8}}>
          <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:5}}>Demo accounts</div>
          {[["admin","admin123","Admin"],["sales","sales123","Sales"],["viewer","viewer123","Viewer"]].map(([un,pw,r])=>(
            <div key={un} onClick={()=>{setU(un);setP(pw);setErr("");}} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"5px 0",cursor:"pointer",borderBottom:`1px solid ${C.border}`}}>
              <span style={{fontSize:11,color:C.text,fontFamily:"'DM Mono',monospace"}}>{un} / {pw}</span><Badge label={r}/>
            </div>
          ))}
          <div style={{fontSize:9,color:C.muted,marginTop:5}}>Tap any row to auto-fill</div>
        </div>
      </div>
    </div>
  );
}

/* ── AI ── */
function AIAssistant({inventory,customers,sales,expenses}){
  const [key]=useState(()=>localStorage.getItem("natrio_api_key")||"");
  const [saved,setSaved]=useState(key);const [kInput,setKInput]=useState("");
  const [msgs,setMsgs]=useState([{role:"assistant",text:"Hi! I'm your Natrio AI. Ask me about inventory, sales, customers, expenses, or any business insights."}]);
  const [inp,setInp]=useState("");const [load,setLoad]=useState(false);const bot=useRef();
  useEffect(()=>{bot.current?.scrollIntoView({behavior:"smooth"});},[msgs]);
  if(!saved)return(
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100%",gap:12,padding:24}}>
      <div style={{fontSize:24}}>✦</div><div style={{fontSize:16,fontWeight:800}}>AI Assistant Setup</div>
      <div style={{fontSize:12,color:C.muted,textAlign:"center",maxWidth:340}}>Enter your Anthropic API key. Get one free at <a href="https://console.anthropic.com" target="_blank" rel="noreferrer" style={{color:C.accent}}>console.anthropic.com</a></div>
      <div style={{width:"100%",maxWidth:340,display:"flex",flexDirection:"column",gap:8}}>
        <Input placeholder="sk-ant-api03-..." type="password" value={kInput} onChange={e=>setKInput(e.target.value)}/>
        <Btn onClick={()=>{if(kInput.startsWith("sk-")){localStorage.setItem("natrio_api_key",kInput);setSaved(kInput);}}} style={{width:"100%",padding:10}}>Save & Enable AI</Btn>
      </div>
    </div>
  );
  const sys=`You are a wholesale/distribution AI for Natrio Organics. Currency PKR.
INV:${inventory.map(i=>`${i.sku}|${i.name}|Qty:${i.qty}(min:${i.reorder})|PKR${i.price}`).join(";")}
CUST:${customers.map(c=>`${c.name}|${c.tier}|Bal:PKR${c.balance}`).join(";")}
SALES:${sales.map(s=>`#${s.id}|${s.customer}|PKR${s.total}|${s.status}`).join(";")}
EXP total:PKR${expenses.reduce((a,e)=>a+e.amount,0)}
Rev:PKR${sales.reduce((a,s)=>a+s.total,0)}|LowStock:${inventory.filter(i=>i.qty<=i.reorder).map(i=>i.name).join(",")||"None"}
Be concise, use PKR, give actionable insights.`;
  const send=async()=>{
    if(!inp.trim()||load)return;
    const m=inp.trim();setInp("");setMsgs(x=>[...x,{role:"user",text:m}]);setLoad(true);
    try{
      const r=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json","x-api-key":saved,"anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,system:sys,messages:[...msgs.slice(1).map(x=>({role:x.role==="assistant"?"assistant":"user",content:x.text})),{role:"user",content:m}]})});
      const d=await r.json();if(d.error)throw new Error(d.error.message);
      setMsgs(x=>[...x,{role:"assistant",text:d.content?.[0]?.text||"Sorry, try again."}]);
    }catch(e){setMsgs(x=>[...x,{role:"assistant",text:`Error: ${e.message}`}]);}
    setLoad(false);
  };
  return(
    <div style={{display:"flex",flexDirection:"column",height:"100%"}}>
      <div style={{flex:1,overflowY:"auto",padding:14,display:"flex",flexDirection:"column",gap:9}}>
        {msgs.map((m,i)=>(
          <div key={i} style={{display:"flex",justifyContent:m.role==="user"?"flex-end":"flex-start"}}>
            <div style={{maxWidth:"85%",padding:"9px 13px",borderRadius:m.role==="user"?"13px 13px 4px 13px":"13px 13px 13px 4px",background:m.role==="user"?C.accent:C.surface,color:m.role==="user"?"#fff":C.text,fontSize:13,lineHeight:1.5,border:m.role==="assistant"?`1px solid ${C.border}`:"none",whiteSpace:"pre-wrap"}}>{m.text}</div>
          </div>
        ))}
        {load&&<div style={{display:"flex"}}><div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:"13px 13px 13px 4px",padding:"9px 13px",color:C.muted,fontSize:12}}>Thinking…</div></div>}
        <div ref={bot}/>
      </div>
      <div style={{padding:"10px 14px",borderTop:`1px solid ${C.border}`,display:"flex",gap:8}}>
        <Input value={inp} onChange={e=>setInp(e.target.value)} onKeyDown={e=>e.key==="Enter"&&send()} placeholder="Ask about your business…" style={{flex:1}}/>
        <Btn onClick={send} disabled={load||!inp.trim()}>Send</Btn>
      </div>
    </div>
  );
}

/* ── Dashboard ── */
function Dashboard({inventory,customers,sales,expenses,deliveries}){
  const isMobile=useIsMobile();
  const [period,setPeriod]=useState("all");
  const now=new Date();
  const filterSales=s=>{
    const d=new Date(s.date);
    if(period==="month") return d.getFullYear()===now.getFullYear()&&d.getMonth()===now.getMonth();
    if(period==="quarter"){const q=Math.floor(now.getMonth()/3);return d.getFullYear()===now.getFullYear()&&Math.floor(d.getMonth()/3)===q;}
    if(period==="year") return d.getFullYear()===now.getFullYear();
    return true;
  };
  const filtered=sales.filter(filterSales);
  const rev=filtered.reduce((a,s)=>a+s.total,0);
  const cogs=filtered.reduce((a,s)=>a+s.items.reduce((b,it)=>{const inv=inventory.find(i=>i.sku===it.sku);return b+(inv?it.qty*inv.cost:0);},0),0);
  const outs=filtered.filter(s=>["Due","Invoiced","Pending"].includes(s.status)).reduce((a,s)=>a+s.total,0);
  const low=inventory.filter(i=>i.qty<=i.reorder);
  const over=filtered.filter(s=>s.status==="Due");
  const transit=deliveries.filter(d=>["Dispatched","En Route"].includes(d.status));
  const periodLabel={all:"All Time",month:"This Month",quarter:"This Quarter",year:"This Year"}[period];
  return(
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <div style={{fontSize:12,color:C.muted,fontWeight:600}}>{periodLabel}</div>
        <div style={{display:"flex",gap:6}}>
          {["all","month","quarter","year"].map(p=>(
            <button key={p} onClick={()=>setPeriod(p)} style={{background:period===p?C.accent:"transparent",color:period===p?"#fff":C.muted,border:`1px solid ${period===p?C.accent:C.border}`,borderRadius:6,padding:"5px 10px",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",textTransform:"capitalize"}}>{p==="all"?"All":p.charAt(0).toUpperCase()+p.slice(1)}</button>
          ))}
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr 1fr":"repeat(3,1fr) repeat(3,1fr)",gap:9}}>
        <StatCard label="Revenue" value={fmt(rev)} sub={`${filtered.length} orders`} color={C.green} icon="💰"/>
        <StatCard label="Gross Profit" value={fmt(rev-cogs)} sub={`${rev?Math.round(((rev-cogs)/rev)*100):0}% margin`} color={C.accent} icon="📈"/>
        <StatCard label="Outstanding" value={fmt(outs)} sub="unpaid" color={C.amber} icon="⏳"/>
        <StatCard label="Low Stock" value={low.length} color={low.length?C.red:C.green} icon="📦"/>
        <StatCard label="In Transit" value={transit.length} color={C.teal} icon="🚚"/>
        <StatCard label="Overdue" value={over.length} color={over.length?C.red:C.green} icon="🔴"/>
      </div>
      <div style={{display:"flex",flexDirection:isMobile?"column":"row",gap:12}}>
        <Card style={{flex:2}} title="Recent Orders">
          {isMobile?(
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {filtered.slice(0,5).map(s=>(
                <div key={s.id} style={{background:C.bg,borderRadius:8,padding:"10px 12px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                    <div><div style={{fontSize:12,fontWeight:700}}>#{s.id} — {s.customer}</div><div style={{fontSize:10,color:C.muted,marginTop:2}}>{s.date}</div></div>
                    <Badge label={s.status}/>
                  </div>
                  <div style={{fontSize:13,fontWeight:800,color:C.accent,fontFamily:"'DM Mono',monospace",marginTop:5}}>{fmt(s.total)}</div>
                </div>
              ))}
            </div>
          ):(
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr><Th>Order</Th><Th>Date</Th><Th>Customer</Th><Th>Total</Th><Th>Status</Th></tr></thead>
              <tbody>{filtered.slice(0,6).map(s=>(
                <tr key={s.id}><Td style={{color:C.accent,fontFamily:"'DM Mono',monospace",fontWeight:700}}>#{s.id}</Td><Td style={{color:C.muted}}>{s.date}</Td><Td style={{fontWeight:600}}>{s.customer}</Td><Td style={{fontFamily:"'DM Mono',monospace"}}>{fmt(s.total)}</Td><Td><Badge label={s.status}/></Td></tr>
              ))}</tbody>
            </table>
          )}
        </Card>
        <div style={{flex:1,display:"flex",flexDirection:"column",gap:10}}>
          <Card title="⚠ Low Stock">
            {low.length===0?<div style={{color:C.muted,fontSize:12}}>All items OK ✓</div>:low.slice(0,4).map(i=>(
              <div key={i.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${C.border}`}}>
                <div><div style={{fontSize:12,fontWeight:600}}>{i.name}</div><div style={{fontSize:9,color:C.muted}}>{i.sku}</div></div>
                <div style={{textAlign:"right"}}><div style={{fontSize:13,fontWeight:800,color:C.red,fontFamily:"'DM Mono',monospace"}}>{i.qty}</div><div style={{fontSize:9,color:C.muted}}>min {i.reorder}</div></div>
              </div>
            ))}
          </Card>
          <Card title="🔴 Overdue">
            {over.length===0?<div style={{color:C.muted,fontSize:12}}>No overdue ✓</div>:over.slice(0,3).map(s=>(
              <div key={s.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${C.border}`}}>
                <div><div style={{fontSize:12,fontWeight:600}}>{s.customer}</div><div style={{fontSize:9,color:C.muted}}>#{s.id}</div></div>
                <div style={{fontSize:12,fontWeight:800,color:C.red,fontFamily:"'DM Mono',monospace"}}>{fmt(s.total)}</div>
              </div>
            ))}
          </Card>
        </div>
      </div>
    </div>
  );
}

/* ── POS ── */
function POS({inventory,setInventory,customers,sales,setSales,currentUser,addAudit}){
  const isMobile=useIsMobile();
  const [cart,setCart]=useState([]);const [custId,setCustId]=useState("");const [search,setSearch]=useState("");const [saved,setSaved]=useState(false);const [disc,setDisc]=useState(0);const [note,setNote]=useState("");const [showCart,setShowCart]=useState(false);
  const filtered=inventory.filter(i=>i.name.toLowerCase().includes(search.toLowerCase())||i.sku.toLowerCase().includes(search.toLowerCase())||(i.barcode||"").includes(search));
  const sub=cart.reduce((a,c)=>a+c.qty*c.price,0);const da=Math.round(sub*(disc/100));const total=sub-da;
  const add=item=>setCart(c=>{const ex=c.find(x=>x.sku===item.sku);if(ex)return c.map(x=>x.sku===item.sku?{...x,qty:x.qty+1}:x);return[...c,{...item,qty:1}];});
  const upd=(sku,raw)=>{const q=parseInt(raw,10);if(isNaN(q)||q<=0){setCart(c=>c.filter(x=>x.sku!==sku));return;}setCart(c=>c.map(x=>x.sku===sku?{...x,qty:q}:x));};
  const place=()=>{
    if(!cart.length||!custId)return;
    const cust=customers.find(c=>c.id===parseInt(custId));
    const order={id:uid(),date:today(),customer:cust.name,customerId:cust.id,items:cart.map(c=>({sku:c.sku,name:c.name,qty:c.qty,price:c.price})),subtotal:sub,discount:disc,discAmt:da,total,note,status:"Invoiced",payments:[]};
    setSales(s=>[order,...s]);setInventory(inv=>inv.map(i=>{const ci=cart.find(c=>c.sku===i.sku);return ci?{...i,qty:i.qty-ci.qty}:i;}));
    addAudit("create","Order",`#${order.id} for ${cust.name} — ${fmt(total)}`,currentUser.name);
    setCart([]);setCustId("");setDisc(0);setNote("");setSaved(true);setShowCart(false);setTimeout(()=>setSaved(false),3000);
  };
  const CartPanel=()=>(
    <Card style={{display:"flex",flexDirection:"column",gap:9,flex:isMobile?undefined:1}} title="Order Cart">
      <Sel value={custId} onChange={e=>setCustId(e.target.value)}><option value="">— Select Customer —</option>{customers.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</Sel>
      <div style={{flex:1,overflowY:"auto",maxHeight:isMobile?"40vh":"auto"}}>
        {cart.length===0?<div style={{color:C.muted,fontSize:12,textAlign:"center",marginTop:16}}>Tap products to add</div>:cart.map(item=>(
          <div key={item.sku} style={{borderBottom:`1px solid ${C.border}`,padding:"8px 0"}}>
            <div style={{fontSize:12,fontWeight:700,marginBottom:4}}>{item.name}</div>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <div style={{display:"flex",alignItems:"center",gap:4}}>
                <Btn onClick={()=>upd(item.sku,item.qty-1)} variant="ghost" style={{padding:"2px 7px",fontSize:13}}>−</Btn>
                <input type="number" min="1" value={item.qty} onChange={e=>upd(item.sku,e.target.value)} style={{width:42,textAlign:"center",background:C.bg,border:`1px solid ${C.border}`,borderRadius:5,padding:"3px 2px",color:C.text,fontSize:12,fontFamily:"'DM Mono',monospace",outline:"none"}}/>
                <Btn onClick={()=>upd(item.sku,item.qty+1)} variant="ghost" style={{padding:"2px 7px",fontSize:13}}>+</Btn>
              </div>
              <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:1}}>
                <span style={{color:C.accent,fontFamily:"'DM Mono',monospace",fontWeight:700,fontSize:11}}>{fmt(item.qty*item.price)}</span>
                <button onClick={()=>setCart(c=>c.filter(x=>x.sku!==item.sku))} style={{background:"none",border:"none",color:C.red,fontSize:9,cursor:"pointer",padding:0}}>Remove</button>
              </div>
            </div>
          </div>
        ))}
      </div>
      <div style={{display:"flex",alignItems:"center",gap:6}}>
        <span style={{fontSize:10,color:C.muted,whiteSpace:"nowrap"}}>Discount %</span>
        <input type="number" min="0" max="100" value={disc} onChange={e=>setDisc(Math.max(0,Math.min(100,+e.target.value)))} style={{width:48,background:C.bg,border:`1px solid ${C.border}`,borderRadius:5,padding:"4px 6px",color:C.text,fontSize:12,outline:"none"}}/>
      </div>
      <Input value={note} onChange={e=>setNote(e.target.value)} placeholder="Order note (optional)"/>
      <div style={{borderTop:`1px solid ${C.border}`,paddingTop:9}}>
        {disc>0&&<div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:C.muted,marginBottom:3}}><span>Subtotal</span><span>{fmt(sub)}</span></div>}
        {disc>0&&<div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:C.green,marginBottom:3}}><span>Discount ({disc}%)</span><span>-{fmt(da)}</span></div>}
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:9}}><span style={{color:C.muted,fontWeight:700,fontSize:12}}>Total</span><span style={{fontSize:17,fontWeight:800,fontFamily:"'DM Mono',monospace"}}>{fmt(total)}</span></div>
        <Btn onClick={place} disabled={!cart.length||!custId} style={{width:"100%",padding:"10px",fontSize:12}}>{saved?"✓ Order Placed!":"Place Order"}</Btn>
      </div>
    </Card>
  );
  return(
    <div style={{display:"flex",flexDirection:"column",gap:10,height:isMobile?"auto":"calc(100vh - 130px)"}}>
      <Input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search by name, SKU or barcode…"/>
      {isMobile&&cart.length>0&&(
        <button onClick={()=>setShowCart(v=>!v)} style={{background:C.accent,color:"#fff",border:"none",borderRadius:8,padding:"10px",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>🛒 Cart ({cart.length} items) — {fmt(total)}</button>
      )}
      {isMobile&&showCart&&<CartPanel/>}
      <div style={{display:isMobile?"none":"flex",gap:12,flex:1,overflow:"hidden"}}>
        <div style={{flex:2,display:"flex",flexDirection:"column",gap:9}}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:8,overflowY:"auto",flex:1}}>
            {filtered.map(item=>(
              <div key={item.id} onClick={()=>add(item)} style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,padding:11,cursor:"pointer",transition:"border-color .15s",boxShadow:"0 1px 3px #0000000a"}} onMouseEnter={e=>e.currentTarget.style.borderColor=C.accent} onMouseLeave={e=>e.currentTarget.style.borderColor=C.border}>
                <div style={{fontSize:9,color:C.muted,fontFamily:"'DM Mono',monospace",marginBottom:2}}>{item.sku}</div>
                <div style={{fontSize:12,fontWeight:700,marginBottom:5,lineHeight:1.3}}>{item.name}</div>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <span style={{fontSize:12,fontWeight:800,color:C.accent,fontFamily:"'DM Mono',monospace"}}>{fmt(item.price)}</span>
                  <Badge label={item.qty<=item.reorder?"Low":"OK"}/>
                </div>
                <div style={{fontSize:9,color:C.muted,marginTop:3}}>Stock: {item.qty}</div>
              </div>
            ))}
          </div>
        </div>
        <CartPanel/>
      </div>
      {isMobile&&(
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:8}}>
          {filtered.map(item=>(
            <div key={item.id} onClick={()=>add(item)} style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,padding:11,cursor:"pointer",boxShadow:"0 1px 3px #0000000a"}}>
              <div style={{fontSize:9,color:C.muted,fontFamily:"'DM Mono',monospace",marginBottom:2}}>{item.sku}</div>
              <div style={{fontSize:12,fontWeight:700,marginBottom:5,lineHeight:1.3}}>{item.name}</div>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span style={{fontSize:12,fontWeight:800,color:C.accent,fontFamily:"'DM Mono',monospace"}}>{fmt(item.price)}</span>
                <Badge label={item.qty<=item.reorder?"Low":"OK"}/>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/* ── Orders ── */
function Orders({sales,setSales,customers,settings,currentUser,addAudit,canEdit}){
  const isMobile=useIsMobile();
  const [fC,setFC]=useState("all");const [fS,setFS]=useState("all");const [exp,setExp]=useState(null);const [pay,setPay]=useState(null);const [pf,setPf]=useState({amount:"",method:"Bank Transfer",note:""});const [editId,setEditId]=useState(null);const [ef,setEf]=useState({});
  const fil=sales.filter(s=>(fC==="all"||s.customerId===parseInt(fC))&&(fS==="all"||s.status===fS));
  const upSt=(id,st)=>{setSales(s=>s.map(o=>o.id===id?{...o,status:st}:o));addAudit("update","Order",`#${id} → ${st}`,currentUser.name);};
  const del=id=>{if(!window.confirm("Delete order?"))return;setSales(s=>s.filter(o=>o.id!==id));addAudit("delete","Order",`#${id}`,currentUser.name);};
  const addPay=()=>{if(!pay||!pf.amount)return;setSales(s=>s.map(o=>{if(o.id!==pay.id)return o;const payments=[...(o.payments||[]),{amount:+pf.amount,date:today(),method:pf.method,note:pf.note}];const pd=payments.reduce((a,p)=>a+p.amount,0);return{...o,payments,status:pd>=o.total?"Paid":o.status};}));addAudit("update","Payment",`PKR${pf.amount} on #${pay.id}`,currentUser.name);setPay(null);setPf({amount:"",method:"Bank Transfer",note:""});};
  const startEdit=o=>{setEditId(o.id);setEf({date:o.date,customer:o.customer,note:o.note||"",status:o.status});};
  const saveEdit=()=>{setSales(s=>s.map(o=>o.id===editId?{...o,...ef}:o));addAudit("update","Order",`#${editId} edited`,currentUser.name);setEditId(null);};
  const tot={rev:fil.reduce((a,s)=>a+s.total,0),paid:fil.filter(s=>s.status==="Paid").reduce((a,s)=>a+s.total,0),due:fil.filter(s=>["Due","Invoiced","Pending"].includes(s.status)).reduce((a,s)=>a+s.total,0)};
  const doExport=()=>exportToExcel(fil.map(o=>({ID:o.id,Date:o.date,Customer:o.customer,Total:o.total,Paid:(o.payments||[]).reduce((a,p)=>a+p.amount,0),Status:o.status,Note:o.note||""})),"natrio_orders.xlsx");
  return(
    <div style={{display:"flex",flexDirection:"column",gap:11}}>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}><StatCard label="Orders" value={fil.length}/><StatCard label="Value" value={fmt(tot.rev)}/><StatCard label="Collected" value={fmt(tot.paid)} color={C.green}/><StatCard label="Outstanding" value={fmt(tot.due)} color={C.amber}/></div>
      <Card>
        <div style={{display:"flex",gap:7,alignItems:"center",flexWrap:"wrap"}}>
          <Sel value={fC} onChange={e=>setFC(e.target.value)} style={{flex:1,minWidth:120}}><option value="all">All Customers</option>{customers.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</Sel>
          <Sel value={fS} onChange={e=>setFS(e.target.value)} style={{flex:1,minWidth:100}}><option value="all">All Statuses</option>{ORDER_STATUSES.map(s=><option key={s}>{s}</option>)}</Sel>
          <Btn onClick={()=>{setFC("all");setFS("all");}} variant="ghost" style={{fontSize:10,padding:"5px 9px"}}>Clear</Btn>
          <ImportExportBar exportFn={doExport} label="orders"/>
        </div>
      </Card>
      {pay&&<div style={{position:"fixed",inset:0,background:"#00000066",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:16}} onClick={e=>e.target===e.currentTarget&&setPay(null)}>
        <Card style={{width:"100%",maxWidth:360}} title={`Record Payment — #${pay.id}`}>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            <div style={{fontSize:12,color:C.muted}}>Total: <strong>{fmt(pay.total)}</strong> · Paid: <strong>{fmt((pay.payments||[]).reduce((a,p)=>a+p.amount,0))}</strong></div>
            <Input type="number" placeholder="Amount (PKR)" value={pf.amount} onChange={e=>setPf(f=>({...f,amount:e.target.value}))}/>
            <Sel value={pf.method} onChange={e=>setPf(f=>({...f,method:e.target.value}))}><option>Bank Transfer</option><option>Cash</option><option>Cheque</option><option>Online</option></Sel>
            <Input placeholder="Note (optional)" value={pf.note} onChange={e=>setPf(f=>({...f,note:e.target.value}))}/>
            <div style={{display:"flex",gap:7}}><Btn onClick={addPay} variant="success" style={{flex:1}}>Save Payment</Btn><Btn onClick={()=>setPay(null)} variant="ghost" style={{flex:1}}>Cancel</Btn></div>
          </div>
        </Card>
      </div>}
      {isMobile?(
        <div style={{display:"flex",flexDirection:"column",gap:9}}>
          {fil.length===0&&<div style={{textAlign:"center",color:C.muted,padding:24,fontSize:13}}>No orders match filters.</div>}
          {fil.map(order=>{const pd=(order.payments||[]).reduce((a,p)=>a+p.amount,0);return(
            <Card key={order.id}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
                <div><div style={{fontSize:13,fontWeight:800,color:C.accent,fontFamily:"'DM Mono',monospace"}}>#{order.id}</div><div style={{fontSize:11,color:C.muted,marginTop:2}}>{order.date} · {order.customer}</div></div>
                <Badge label={order.status}/>
              </div>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                <div><div style={{fontSize:10,color:C.muted}}>Total</div><div style={{fontSize:14,fontWeight:800,fontFamily:"'DM Mono',monospace"}}>{fmt(order.total)}</div></div>
                <div style={{textAlign:"right"}}><div style={{fontSize:10,color:C.muted}}>Paid</div><div style={{fontSize:14,fontWeight:800,color:pd>=order.total?C.green:C.amber,fontFamily:"'DM Mono',monospace"}}>{fmt(pd)}</div></div>
              </div>
              {canEdit&&<div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                <Sel value={order.status} onChange={e=>upSt(order.id,e.target.value)} style={{fontSize:10,padding:"4px 6px",flex:1}}>{ORDER_STATUSES.map(s=><option key={s}>{s}</option>)}</Sel>
                <Btn onClick={()=>setPay(order)} variant="success" style={{fontSize:10,padding:"4px 8px"}}>💳 Pay</Btn>
                <Btn onClick={()=>printDoc(order,customers,settings)} variant="purple" style={{fontSize:10,padding:"4px 8px"}}>🖨</Btn>
                <Btn onClick={()=>del(order.id)} variant="danger" style={{fontSize:10,padding:"4px 8px"}}>✕</Btn>
              </div>}
            </Card>
          );})}
        </div>
      ):(
        <Card><table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr><Th>Order</Th><Th>Date</Th><Th>Customer</Th><Th>Total</Th><Th>Paid</Th><Th>Status</Th><Th>Actions</Th></tr></thead>
          <tbody>
            {fil.length===0&&<tr><td colSpan={7} style={{padding:"24px",textAlign:"center",color:C.muted,fontSize:12}}>No orders match filters.</td></tr>}
            {fil.map(order=>{const pd=(order.payments||[]).reduce((a,p)=>a+p.amount,0);return(<>
              {editId===order.id?(
                <tr key={`${order.id}-edit`} style={{background:C.accentSoft}}>
                  <Td style={{color:C.accent,fontFamily:"'DM Mono',monospace",fontWeight:700}}>#{order.id}</Td>
                  <Td><Input value={ef.date} onChange={e=>setEf(f=>({...f,date:e.target.value}))} type="date" style={{fontSize:11,padding:"3px 6px"}}/></Td>
                  <Td><Input value={ef.customer} onChange={e=>setEf(f=>({...f,customer:e.target.value}))} style={{fontSize:11,padding:"3px 6px"}}/></Td>
                  <Td style={{fontFamily:"'DM Mono',monospace"}}>{fmt(order.total)}</Td>
                  <Td><Input value={ef.note} onChange={e=>setEf(f=>({...f,note:e.target.value}))} placeholder="Note" style={{fontSize:11,padding:"3px 6px"}}/></Td>
                  <Td><Sel value={ef.status} onChange={e=>setEf(f=>({...f,status:e.target.value}))} style={{fontSize:11,padding:"3px 6px"}}>{ORDER_STATUSES.map(s=><option key={s}>{s}</option>)}</Sel></Td>
                  <Td><div style={{display:"flex",gap:4}}><Btn onClick={saveEdit} variant="success" style={{fontSize:10,padding:"3px 7px"}}>Save</Btn><Btn onClick={()=>setEditId(null)} variant="ghost" style={{fontSize:10,padding:"3px 7px"}}>✕</Btn></div></Td>
                </tr>
              ):(
                <>
                <tr key={order.id} style={{borderBottom:`1px solid ${C.border}`,cursor:"pointer",background:exp===order.id?C.bg:"transparent"}} onClick={()=>setExp(exp===order.id?null:order.id)}>
                  <Td style={{color:C.accent,fontFamily:"'DM Mono',monospace",fontWeight:700}}>#{order.id}</Td>
                  <Td style={{color:C.muted}}>{order.date}</Td><Td style={{fontWeight:600}}>{order.customer}</Td>
                  <Td style={{fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fmt(order.total)}</Td>
                  <Td style={{fontFamily:"'DM Mono',monospace",color:pd>=order.total?C.green:C.amber}}>{fmt(pd)}</Td>
                  <Td><Badge label={order.status}/></Td>
                  <Td onClick={e=>e.stopPropagation()}><div style={{display:"flex",gap:3,flexWrap:"wrap"}}>
                    {canEdit&&<Sel value={order.status} onChange={e=>upSt(order.id,e.target.value)} style={{fontSize:10,padding:"3px 5px",width:"auto"}}>{ORDER_STATUSES.map(s=><option key={s}>{s}</option>)}</Sel>}
                    {canEdit&&<Tip label="Record Payment"><Btn onClick={()=>setPay(order)} variant="success" style={{fontSize:10,padding:"3px 7px"}}>💳</Btn></Tip>}
                    {canEdit&&<Tip label="Edit Order"><Btn onClick={()=>startEdit(order)} variant="ghost" style={{fontSize:10,padding:"3px 7px"}}>✎</Btn></Tip>}
                    <Tip label="Print Invoice"><Btn onClick={()=>printDoc(order,customers,settings)} variant="purple" style={{fontSize:10,padding:"3px 7px"}}>🖨</Btn></Tip>
                    <Tip label="Packing Slip"><Btn onClick={()=>printPackingSlip(order)} variant="teal" style={{fontSize:10,padding:"3px 7px"}}>📋</Btn></Tip>
                    {canEdit&&<Tip label="Delete"><Btn onClick={()=>del(order.id)} variant="danger" style={{fontSize:10,padding:"3px 7px"}}>✕</Btn></Tip>}
                  </div></Td>
                </tr>
                {exp===order.id&&(
                  <tr key={`${order.id}-d`}><td colSpan={7} style={{background:C.bg,padding:"0 12px 12px 28px",borderBottom:`1px solid ${C.border}`}}>
                    <div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:1,margin:"9px 0 6px"}}>Items</div>
                    <table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr><Th>SKU</Th><Th>Product</Th><Th>Qty</Th><Th>Price</Th><Th>Total</Th></tr></thead>
                    <tbody>{order.items.map((it,i)=><tr key={i}><Td style={{fontFamily:"'DM Mono',monospace",color:C.muted,fontSize:10}}>{it.sku||"—"}</Td><Td style={{fontWeight:600}}>{it.name}</Td><Td>{it.qty}</Td><Td>{fmt(it.price)}</Td><Td style={{color:C.accent,fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fmt(it.qty*it.price)}</Td></tr>)}</tbody></table>
                    {(order.payments||[]).length>0&&<><div style={{fontSize:9,color:C.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:1,margin:"9px 0 6px"}}>Payments</div>
                    <table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr><Th>Date</Th><Th>Amount</Th><Th>Method</Th><Th>Note</Th></tr></thead>
                    <tbody>{order.payments.map((p,i)=><tr key={i}><Td>{p.date}</Td><Td style={{fontFamily:"'DM Mono',monospace",fontWeight:700,color:C.green}}>{fmt(p.amount)}</Td><Td>{p.method}</Td><Td style={{color:C.muted}}>{p.note||"—"}</Td></tr>)}</tbody></table></>}
                    {order.note&&<div style={{marginTop:7,fontSize:11,color:C.muted}}>Note: {order.note}</div>}
                  </td></tr>
                )}
                </>
              )}
            </>);})}
          </tbody>
        </table></Card>
      )}
    </div>
  );
}

/* ── Quotations ── */
function Quotations({inventory,customers,setSales,setInventory,settings,currentUser,addAudit}){
  const isMobile=useIsMobile();
  const [quotes,setQuotes]=useState([{id:"QT-001",date:"2026-05-22",customer:"Acme Industries",customerId:1,items:[{sku:"WHL-001",name:"Industrial Bearing Set",qty:50,price:7800}],total:390000,status:"Sent",validUntil:"2026-06-22",note:"Bulk order"}]);
  const [add,setAdd]=useState(false);const [form,setForm]=useState({customerId:"",validUntil:"",note:""});const [cq,setCq]=useState([]);
  const addItem=item=>setCq(c=>{const ex=c.find(x=>x.sku===item.sku);if(ex)return c.map(x=>x.sku===item.sku?{...x,qty:x.qty+1}:x);return[...c,{...item,qty:1}];});
  const save=()=>{const cust=customers.find(c=>c.id===parseInt(form.customerId));if(!cust||!cq.length)return;const tot=cq.reduce((a,c)=>a+c.qty*c.price,0);const q={id:`QT-${String(quotes.length+1).padStart(3,"0")}`,date:today(),customer:cust.name,customerId:cust.id,items:cq.map(c=>({sku:c.sku,name:c.name,qty:c.qty,price:c.price})),total:tot,status:"Draft",validUntil:form.validUntil,note:form.note};setQuotes(qs=>[q,...qs]);addAudit("create","Quotation",`${q.id} for ${cust.name}`,currentUser.name);setAdd(false);setCq([]);setForm({customerId:"",validUntil:"",note:""});};
  const convert=q=>{if(!window.confirm("Convert to order?"))return;const order={id:uid(),date:today(),customer:q.customer,customerId:q.customerId,items:q.items,total:q.total,status:"Invoiced",payments:[],note:`From ${q.id}`};setSales(s=>[order,...s]);setInventory(inv=>inv.map(i=>{const ci=q.items.find(x=>x.sku===i.sku);return ci?{...i,qty:i.qty-ci.qty}:i;}));setQuotes(qs=>qs.map(x=>x.id===q.id?{...x,status:"Converted"}:x));addAudit("create","Order",`From ${q.id}`,currentUser.name);};
  const doExport=()=>exportToExcel(quotes.map(q=>({ID:q.id,Date:q.date,Customer:q.customer,Total:q.total,Status:q.status,ValidUntil:q.validUntil||"",Note:q.note||""})),"natrio_quotes.xlsx");
  return(
    <div style={{display:"flex",flexDirection:"column",gap:11}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <ImportExportBar exportFn={doExport} label="quotations"/>
        <Btn onClick={()=>setAdd(true)}>+ New Quotation</Btn>
      </div>
      {add&&<Card title="New Quotation" style={{border:`1px solid ${C.accent}`}}>
        <div style={{display:"flex",gap:8,marginBottom:9,flexWrap:"wrap"}}>
          <Sel value={form.customerId} onChange={e=>setForm(f=>({...f,customerId:e.target.value}))} style={{flex:1,minWidth:140}}><option value="">Select Customer</option>{customers.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</Sel>
          <Input type="date" value={form.validUntil} onChange={e=>setForm(f=>({...f,validUntil:e.target.value}))} style={{flex:1,minWidth:130}}/>
          <Input value={form.note} onChange={e=>setForm(f=>({...f,note:e.target.value}))} placeholder="Note" style={{flex:2,minWidth:160}}/>
        </div>
        <div style={{display:"grid",gridTemplateColumns:`repeat(auto-fill,minmax(${isMobile?"130px":"140px"},1fr))`,gap:6,marginBottom:9}}>
          {inventory.map(item=><div key={item.id} onClick={()=>addItem(item)} style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:6,padding:8,cursor:"pointer",fontSize:11}}><div style={{fontWeight:700,marginBottom:2}}>{item.name}</div><div style={{color:C.accent,fontFamily:"'DM Mono',monospace",fontSize:10}}>{fmt(item.price)}</div></div>)}
        </div>
        {cq.length>0&&<div style={{marginBottom:9}}>{cq.map(it=><div key={it.sku} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"5px 0",borderBottom:`1px solid ${C.border}`}}><span style={{fontSize:12}}>{it.name}</span><div style={{display:"flex",gap:6,alignItems:"center"}}><input type="number" min="1" value={it.qty} onChange={e=>setCq(c=>c.map(x=>x.sku===it.sku?{...x,qty:+e.target.value||1}:x))} style={{width:42,background:C.bg,border:`1px solid ${C.border}`,borderRadius:4,padding:"3px 5px",outline:"none",fontSize:11}}/><span style={{fontFamily:"'DM Mono',monospace",fontSize:10,color:C.accent}}>{fmt(it.qty*it.price)}</span><button onClick={()=>setCq(c=>c.filter(x=>x.sku!==it.sku))} style={{background:"none",border:"none",color:C.red,cursor:"pointer",fontSize:12}}>✕</button></div></div>)}</div>}
        <div style={{display:"flex",gap:7}}><Btn onClick={save} variant="success" disabled={!form.customerId||!cq.length}>Save</Btn><Btn onClick={()=>{setAdd(false);setCq([]);}} variant="ghost">Cancel</Btn></div>
      </Card>}
      {isMobile?(
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {quotes.map(q=>(
            <Card key={q.id}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                <div><div style={{fontSize:13,fontWeight:800,color:C.purple,fontFamily:"'DM Mono',monospace"}}>{q.id}</div><div style={{fontSize:11,color:C.muted}}>{q.date} · {q.customer}</div></div>
                <Badge label={q.status}/>
              </div>
              <div style={{fontSize:14,fontWeight:800,fontFamily:"'DM Mono',monospace",marginBottom:8}}>{fmt(q.total)}</div>
              <div style={{display:"flex",gap:5}}>
                {q.status!=="Converted"&&<Btn onClick={()=>convert(q)} variant="success" style={{fontSize:10,padding:"4px 8px",flex:1}}>→ Order</Btn>}
                <Btn onClick={()=>printDoc(q,customers,settings,true)} variant="ghost" style={{fontSize:10,padding:"4px 8px"}}>🖨 Print</Btn>
              </div>
            </Card>
          ))}
        </div>
      ):(
        <Card><table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr><Th>ID</Th><Th>Date</Th><Th>Customer</Th><Th>Total</Th><Th>Valid Until</Th><Th>Status</Th><Th>Actions</Th></tr></thead>
          <tbody>{quotes.map(q=>(
            <tr key={q.id} style={{borderBottom:`1px solid ${C.border}`}}>
              <Td style={{color:C.purple,fontFamily:"'DM Mono',monospace",fontWeight:700}}>{q.id}</Td>
              <Td style={{color:C.muted}}>{q.date}</Td><Td style={{fontWeight:600}}>{q.customer}</Td>
              <Td style={{fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fmt(q.total)}</Td>
              <Td style={{color:C.muted}}>{q.validUntil||"—"}</Td><Td><Badge label={q.status}/></Td>
              <Td><div style={{display:"flex",gap:5}}>
                {q.status!=="Converted"&&<Btn onClick={()=>convert(q)} variant="success" style={{fontSize:10,padding:"3px 7px"}}>→ Order</Btn>}
                <Btn onClick={()=>printDoc(q,customers,settings,true)} variant="ghost" style={{fontSize:10,padding:"3px 7px"}}>🖨 Print</Btn>
              </div></Td>
            </tr>
          ))}</tbody>
        </table></Card>
      )}
    </div>
  );
}

/* ── Inventory ── */
function Inventory({inventory,setInventory,currentUser,addAudit}){
  const isMobile=useIsMobile();
  const [editing,setEditing]=useState(null);const [form,setForm]=useState({});const [adding,setAdding]=useState(false);const [adj,setAdj]=useState(null);const [af,setAf]=useState({qty:"",reason:"",type:"add"});
  const startEdit=item=>{setEditing(item.id);setForm({...item});};
  const save=()=>{setInventory(inv=>inv.map(i=>i.id===editing?{...form,qty:+form.qty,reorder:+form.reorder,cost:+form.cost,price:+form.price}:i));addAudit("update","Inventory",`Updated ${form.name}`,currentUser.name);setEditing(null);};
  const addNew=()=>{const item={...form,id:uid(),qty:+form.qty||0,reorder:+form.reorder||0,cost:+form.cost||0,price:+form.price||0};setInventory(inv=>[...inv,item]);addAudit("create","Inventory",`Added ${form.name}`,currentUser.name);setAdding(false);setForm({});};
  const applyAdj=()=>{if(!adj)return;const d=af.type==="add"?+af.qty:-+af.qty;setInventory(inv=>inv.map(i=>i.id===adj.id?{...i,qty:Math.max(0,i.qty+d)}:i));addAudit("update","StockAdj",`${af.type==="add"?"+":"-"}${af.qty} × ${adj.name} (${af.reason})`,currentUser.name);setAdj(null);setAf({qty:"",reason:"",type:"add"});};
  const doExport=()=>exportToExcel(inventory.map(i=>({SKU:i.sku,Name:i.name,Category:i.category,Qty:i.qty,Reorder:i.reorder,Cost:i.cost,Price:i.price,Supplier:i.supplier,Barcode:i.barcode||""})),"natrio_inventory.xlsx");
  const doImport=file=>importFromExcel(file,rows=>{const mapped=rows.map(r=>({id:uid(),sku:r.SKU||"",name:r.Name||"",category:r.Category||"",qty:+r.Qty||0,reorder:+r.Reorder||0,cost:+r.Cost||0,price:+r.Price||0,supplier:r.Supplier||"",barcode:r.Barcode||""}));setInventory(inv=>[...inv,...mapped]);addAudit("create","Inventory",`Imported ${mapped.length} products`,currentUser.name);});
  return(
    <div style={{display:"flex",flexDirection:"column",gap:11}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <ImportExportBar exportFn={doExport} importFn={doImport} label="inventory"/>
        <Btn onClick={()=>{setAdding(true);setForm({});}}>+ Add Product</Btn>
      </div>
      {adj&&<div style={{position:"fixed",inset:0,background:"#00000066",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:16}} onClick={e=>e.target===e.currentTarget&&setAdj(null)}>
        <Card style={{width:"100%",maxWidth:320}} title={`Stock Adjustment — ${adj.name}`}>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            <div style={{fontSize:12,color:C.muted}}>Current qty: <strong>{adj.qty}</strong></div>
            <div style={{display:"flex",gap:6}}><Btn onClick={()=>setAf(f=>({...f,type:"add"}))} variant={af.type==="add"?"success":"ghost"} style={{flex:1}}>+ Add</Btn><Btn onClick={()=>setAf(f=>({...f,type:"remove"}))} variant={af.type==="remove"?"danger":"ghost"} style={{flex:1}}>− Remove</Btn></div>
            <Input type="number" min="1" placeholder="Quantity" value={af.qty} onChange={e=>setAf(f=>({...f,qty:e.target.value}))}/>
            <Sel value={af.reason} onChange={e=>setAf(f=>({...f,reason:e.target.value}))}><option value="">Select Reason</option><option>Physical Count</option><option>Damaged Goods</option><option>Customer Return</option><option>Supplier Return</option><option>Internal Use</option><option>Other</option></Sel>
            <div style={{display:"flex",gap:7}}><Btn onClick={applyAdj} style={{flex:1}} disabled={!af.qty||!af.reason}>Apply</Btn><Btn onClick={()=>setAdj(null)} variant="ghost" style={{flex:1}}>Cancel</Btn></div>
          </div>
        </Card>
      </div>}
      {adding&&<Card title="New Product" style={{border:`1px solid ${C.accent}`}}>
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr 1fr":"1fr 1fr 1fr",gap:8,marginBottom:8}}>
          {["sku","name","category","supplier","barcode","qty","reorder","cost","price"].map(f=><Input key={f} placeholder={f.charAt(0).toUpperCase()+f.slice(1)} value={form[f]||""} onChange={e=>setForm(x=>({...x,[f]:e.target.value}))}/>)}
        </div>
        <div style={{display:"flex",gap:7}}><Btn onClick={addNew} variant="success">Save</Btn><Btn onClick={()=>setAdding(false)} variant="ghost">Cancel</Btn></div>
      </Card>}
      {isMobile?(
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {inventory.map(item=>(
            <Card key={item.id}>
              {editing===item.id?(
                <div style={{display:"flex",flexDirection:"column",gap:7}}>
                  {["sku","name","category","supplier","qty","reorder","cost","price"].map(f=><Input key={f} placeholder={f.charAt(0).toUpperCase()+f.slice(1)} value={form[f]||""} onChange={e=>setForm(x=>({...x,[f]:e.target.value}))}/>)}
                  <div style={{display:"flex",gap:6}}><Btn onClick={save} variant="success" style={{flex:1}}>Save</Btn><Btn onClick={()=>setEditing(null)} variant="ghost" style={{flex:1}}>Cancel</Btn></div>
                </div>
              ):(
                <>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                  <div><div style={{fontSize:13,fontWeight:700}}>{item.name}</div><div style={{fontSize:10,color:C.muted}}>{item.sku} · {item.category}</div></div>
                  <Badge label={item.qty<=item.reorder?"Low":"OK"}/>
                </div>
                <div style={{display:"flex",gap:10,marginBottom:8}}>
                  <div><div style={{fontSize:9,color:C.muted}}>Stock</div><div style={{fontSize:14,fontWeight:800,color:item.qty<=item.reorder?C.red:C.text,fontFamily:"'DM Mono',monospace"}}>{item.qty}</div></div>
                  <div><div style={{fontSize:9,color:C.muted}}>Price</div><div style={{fontSize:13,fontFamily:"'DM Mono',monospace"}}>{fmt(item.price)}</div></div>
                  <div><div style={{fontSize:9,color:C.muted}}>Margin</div><div style={{fontSize:13,color:C.green,fontFamily:"'DM Mono',monospace"}}>{item.price?Math.round(((item.price-item.cost)/item.price)*100):0}%</div></div>
                </div>
                <div style={{display:"flex",gap:5}}><Btn onClick={()=>startEdit(item)} variant="ghost" style={{fontSize:10,padding:"4px 8px",flex:1}}>✎ Edit</Btn><Btn onClick={()=>setAdj(item)} variant="teal" style={{fontSize:10,padding:"4px 8px",flex:1}}>Adjust</Btn><Btn onClick={()=>{if(window.confirm(`Delete "${item.name}"?`)){setInventory(inv=>inv.filter(i=>i.id!==item.id));addAudit("delete","Inventory",item.name,currentUser.name);}}} variant="danger" style={{fontSize:10,padding:"4px 8px"}}>✕</Btn></div>
                </>
              )}
            </Card>
          ))}
        </div>
      ):(
        <Card><div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr><Th>SKU</Th><Th>Product</Th><Th>Category</Th><Th>Stock</Th><Th>Reorder</Th><Th>Cost</Th><Th>Price</Th><Th>Margin</Th><Th>Status</Th><Th>Actions</Th></tr></thead>
            <tbody>{inventory.map(item=>(
              <tr key={item.id} style={{borderBottom:`1px solid ${C.border}`}}>
                {editing===item.id?(<><td colSpan={9} style={{padding:"8px 9px"}}><div style={{display:"grid",gridTemplateColumns:"repeat(8,1fr)",gap:4}}>{["sku","name","category","qty","reorder","cost","price","supplier"].map(f=><Input key={f} value={form[f]||""} onChange={e=>setForm(x=>({...x,[f]:e.target.value}))} style={{fontSize:10,padding:"3px 6px"}}/>)}</div></td>
                <td style={{padding:"8px 9px",whiteSpace:"nowrap"}}><Btn onClick={save} variant="success" style={{fontSize:10,padding:"3px 7px",marginRight:3}}>Save</Btn><Btn onClick={()=>setEditing(null)} variant="ghost" style={{fontSize:10,padding:"3px 7px"}}>✕</Btn></td></>
                ):(<>
                <Td style={{fontFamily:"'DM Mono',monospace",color:C.muted,fontSize:9}}>{item.sku}</Td>
                <Td><div style={{fontWeight:700,fontSize:12}}>{item.name}</div><div style={{fontSize:9,color:C.muted}}>{item.supplier}</div></Td>
                <Td style={{color:C.muted,fontSize:11}}>{item.category}</Td>
                <Td style={{fontWeight:800,color:item.qty<=item.reorder?C.red:C.text,fontFamily:"'DM Mono',monospace"}}>{item.qty}</Td>
                <Td style={{color:C.muted,fontFamily:"'DM Mono',monospace"}}>{item.reorder}</Td>
                <Td style={{color:C.muted,fontFamily:"'DM Mono',monospace",fontSize:11}}>{fmt(item.cost)}</Td>
                <Td style={{fontFamily:"'DM Mono',monospace",fontSize:11}}>{fmt(item.price)}</Td>
                <Td style={{color:C.green,fontFamily:"'DM Mono',monospace"}}>{item.price?Math.round(((item.price-item.cost)/item.price)*100):0}%</Td>
                <Td><Badge label={item.qty<=item.reorder?"Low":"OK"}/></Td>
                <Td><div style={{display:"flex",gap:3}}>
                  <Tip label="Edit"><Btn onClick={()=>startEdit(item)} variant="ghost" style={{fontSize:10,padding:"3px 6px"}}>✎</Btn></Tip>
                  <Tip label="Adjust Stock"><Btn onClick={()=>setAdj(item)} variant="teal" style={{fontSize:10,padding:"3px 6px"}}>±</Btn></Tip>
                  <Tip label="Delete"><Btn onClick={()=>{if(window.confirm(`Delete "${item.name}"?`)){setInventory(inv=>inv.filter(i=>i.id!==item.id));addAudit("delete","Inventory",item.name,currentUser.name);}}} variant="danger" style={{fontSize:10,padding:"3px 6px"}}>✕</Btn></Tip>
                </div></Td>
                </>)}
              </tr>
            ))}</tbody>
          </table>
        </div></Card>
      )}
    </div>
  );
}

/* ── Purchase Orders ── */
function PurchaseOrders({pos,setPOs,inventory,setInventory,suppliers,currentUser,addAudit}){
  const [add,setAdd]=useState(false);const [form,setForm]=useState({supplierId:"",expectedDate:""});const [cpo,setCpo]=useState([]);
  const addItem=item=>setCpo(c=>{const ex=c.find(x=>x.sku===item.sku);if(ex)return c.map(x=>x.sku===item.sku?{...x,qty:x.qty+1}:x);return[...c,{...item,qty:1,cost:item.cost}];});
  const save=()=>{const s=suppliers.find(x=>x.id===parseInt(form.supplierId));if(!s||!cpo.length)return;const tot=cpo.reduce((a,c)=>a+c.qty*c.cost,0);const po={id:`PO-${String(pos.length+1).padStart(3,"0")}`,date:today(),supplierId:s.id,supplierName:s.name,items:cpo.map(c=>({sku:c.sku,name:c.name,qty:c.qty,cost:c.cost})),total:tot,status:"Pending",expectedDate:form.expectedDate};setPOs(p=>[po,...p]);addAudit("create","PO",`${po.id} from ${s.name}`,currentUser.name);setAdd(false);setCpo([]);setForm({supplierId:"",expectedDate:""});};
  const recv=po=>{if(!window.confirm("Mark received & update stock?"))return;setInventory(inv=>inv.map(i=>{const ci=po.items.find(x=>x.sku===i.sku);return ci?{...i,qty:i.qty+ci.qty}:i;}));setPOs(p=>p.map(x=>x.id===po.id?{...x,status:"Received"}:x));addAudit("update","PO",`${po.id} received`,currentUser.name);};
  const delPO=po=>{if(!window.confirm(`Delete ${po.id}?`))return;setPOs(p=>p.filter(x=>x.id!==po.id));addAudit("delete","PO",po.id,currentUser.name);};
  const doExport=()=>exportToExcel(pos.map(p=>({ID:p.id,Date:p.date,Supplier:p.supplierName,Total:p.total,Status:p.status,Expected:p.expectedDate||""})),"natrio_purchase_orders.xlsx");
  return(
    <div style={{display:"flex",flexDirection:"column",gap:11}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <ImportExportBar exportFn={doExport} label="purchase orders"/>
        <Btn onClick={()=>setAdd(true)}>+ New Purchase Order</Btn>
      </div>
      {add&&<Card title="New PO" style={{border:`1px solid ${C.accent}`}}>
        <div style={{display:"flex",gap:8,marginBottom:9,flexWrap:"wrap"}}><Sel value={form.supplierId} onChange={e=>setForm(f=>({...f,supplierId:e.target.value}))} style={{flex:1,minWidth:140}}><option value="">Select Supplier</option>{suppliers.map(s=><option key={s.id} value={s.id}>{s.name}</option>)}</Sel><Input type="date" value={form.expectedDate} onChange={e=>setForm(f=>({...f,expectedDate:e.target.value}))} style={{flex:1,minWidth:130}}/></div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))",gap:6,marginBottom:9}}>{inventory.map(item=><div key={item.id} onClick={()=>addItem(item)} style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:6,padding:8,cursor:"pointer",fontSize:11}}><div style={{fontWeight:700,marginBottom:2}}>{item.name}</div><div style={{color:C.muted,fontFamily:"'DM Mono',monospace",fontSize:10}}>{fmt(item.cost)}</div><div style={{fontSize:9,color:C.muted}}>Stock: {item.qty}</div></div>)}</div>
        {cpo.length>0&&<div style={{marginBottom:9}}>{cpo.map(it=><div key={it.sku} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"4px 0",borderBottom:`1px solid ${C.border}`}}><span style={{fontSize:12}}>{it.name}</span><div style={{display:"flex",gap:6,alignItems:"center"}}><input type="number" min="1" value={it.qty} onChange={e=>setCpo(c=>c.map(x=>x.sku===it.sku?{...x,qty:+e.target.value||1}:x))} style={{width:40,background:C.bg,border:`1px solid ${C.border}`,borderRadius:4,padding:"2px 4px",outline:"none",fontSize:11}}/><span style={{fontFamily:"'DM Mono',monospace",fontSize:10,color:C.muted}}>{fmt(it.qty*it.cost)}</span><button onClick={()=>setCpo(c=>c.filter(x=>x.sku!==it.sku))} style={{background:"none",border:"none",color:C.red,cursor:"pointer",fontSize:12}}>✕</button></div></div>)}</div>}
        <div style={{display:"flex",gap:7}}><Btn onClick={save} variant="success" disabled={!form.supplierId||!cpo.length}>Create PO</Btn><Btn onClick={()=>{setAdd(false);setCpo([]);}} variant="ghost">Cancel</Btn></div>
      </Card>}
      <Card><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
        <thead><tr><Th>PO #</Th><Th>Date</Th><Th>Supplier</Th><Th>Items</Th><Th>Total Cost</Th><Th>Expected</Th><Th>Status</Th><Th>Actions</Th></tr></thead>
        <tbody>{pos.map(po=>(
          <tr key={po.id} style={{borderBottom:`1px solid ${C.border}`}}>
            <Td style={{color:C.teal,fontFamily:"'DM Mono',monospace",fontWeight:700}}>{po.id}</Td>
            <Td style={{color:C.muted}}>{po.date}</Td><Td style={{fontWeight:600}}>{po.supplierName}</Td>
            <Td style={{color:C.muted}}>{po.items.length}</Td>
            <Td style={{fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fmt(po.total)}</Td>
            <Td style={{color:C.muted}}>{po.expectedDate||"—"}</Td>
            <Td><Badge label={po.status}/></Td>
            <Td><div style={{display:"flex",gap:4}}>
              {po.status==="Pending"&&<Btn onClick={()=>recv(po)} variant="success" style={{fontSize:10,padding:"3px 7px"}}>✓ Received</Btn>}
              <Tip label="Delete PO"><Btn onClick={()=>delPO(po)} variant="danger" style={{fontSize:10,padding:"3px 7px"}}>✕</Btn></Tip>
            </div></Td>
          </tr>
        ))}</tbody>
      </table></div></Card>
    </div>
  );
}

/* ── CRM ── */
function CRM({customers,setCustomers,sales,currentUser,addAudit}){
  const isMobile=useIsMobile();
  const [sel,setSel]=useState(null);
  const [adding,setAdding]=useState(false);
  const [form,setForm]=useState({name:"",email:"",phone:"",address:"",credit:"",tier:"Bronze"});
  const cust=sel?customers.find(c=>c.id===sel):null;
  const cs=cust?sales.filter(s=>s.customerId===cust.id):[];
  const doExport=()=>exportToExcel(customers.map(c=>({Name:c.name,Email:c.email,Phone:c.phone,Tier:c.tier,CreditLimit:c.credit,Balance:c.balance,Orders:c.orders,Since:c.since,Address:c.address||""})),"natrio_customers.xlsx");
  const doImport=file=>importFromExcel(file,rows=>{const mapped=rows.map(r=>({id:uid(),name:r.Name||"",email:r.Email||"",phone:r.Phone||"",tier:r.Tier||"Bronze",credit:+r.CreditLimit||0,balance:+r.Balance||0,orders:+r.Orders||0,since:r.Since||today(),address:r.Address||""}));setCustomers(c=>[...c,...mapped]);addAudit("create","Customer",`Imported ${mapped.length} customers`,currentUser.name);});
  const saveCustomer=()=>{
    if(!form.name.trim())return;
    const nc={id:uid(),name:form.name,email:form.email,phone:form.phone,credit:+form.credit||0,balance:0,tier:form.tier,since:today(),orders:0,address:form.address};
    setCustomers(c=>[...c,nc]);
    addAudit("create","Customer",nc.name,currentUser.name);
    setAdding(false);
    setForm({name:"",email:"",phone:"",address:"",credit:"",tier:"Bronze"});
    setSel(nc.id);
  };
  const del=id=>{
    if(!window.confirm("Delete customer?"))return;
    setCustomers(c=>c.filter(x=>x.id!==id));
    addAudit("delete","Customer",customers.find(c=>c.id===id)?.name||"",currentUser.name);
    if(sel===id)setSel(null);
  };

  /* Customer list panel — plain JSX, no inner component */
  const listPanel=(
    <div style={{flex:1,overflowY:"auto",display:"flex",flexDirection:"column",gap:8,minWidth:0}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,flexWrap:"wrap"}}>
        <ImportExportBar exportFn={doExport} importFn={doImport} label="customers"/>
        <Btn onClick={()=>{setAdding(true);setForm({name:"",email:"",phone:"",address:"",credit:"",tier:"Bronze"});setSel(null);}}>+ Add Customer</Btn>
      </div>
      {adding&&(
        <Card style={{border:`1px solid ${C.accent}`}}>
          <div style={{display:"flex",flexDirection:"column",gap:7}}>
            <Input
              placeholder="Company Name *"
              value={form.name}
              onChange={e=>setForm(f=>({...f,name:e.target.value}))}
            />
            <Input
              placeholder="Email"
              value={form.email}
              onChange={e=>setForm(f=>({...f,email:e.target.value}))}
            />
            <Input
              placeholder="Phone"
              value={form.phone}
              onChange={e=>setForm(f=>({...f,phone:e.target.value}))}
            />
            <Input
              placeholder="Address"
              value={form.address}
              onChange={e=>setForm(f=>({...f,address:e.target.value}))}
            />
            <Input
              placeholder="Credit Limit (PKR)"
              type="number"
              value={form.credit}
              onChange={e=>setForm(f=>({...f,credit:e.target.value}))}
            />
            <Sel value={form.tier} onChange={e=>setForm(f=>({...f,tier:e.target.value}))}>
              {["Bronze","Silver","Gold","Platinum"].map(t=><option key={t}>{t}</option>)}
            </Sel>
            <div style={{display:"flex",gap:7}}>
              <Btn onClick={saveCustomer} variant="success" style={{flex:1}}>Save</Btn>
              <Btn onClick={()=>setAdding(false)} variant="ghost" style={{flex:1}}>Cancel</Btn>
            </div>
          </div>
        </Card>
      )}
      {customers.map(c=>(
        <div key={c.id} onClick={()=>{setSel(c.id);setAdding(false);}}
          style={{background:sel===c.id?C.accentSoft:C.surface,border:`1px solid ${sel===c.id?C.accent:C.border}`,borderRadius:8,padding:11,cursor:"pointer",transition:"all .12s",boxShadow:"0 1px 3px #0000000a"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}><div style={{fontSize:13,fontWeight:700}}>{c.name}</div><Badge label={c.tier}/></div>
          <div style={{fontSize:10,color:C.muted,marginTop:3}}>{c.email}</div>
          <div style={{display:"flex",gap:11,marginTop:6}}>
            <div><div style={{fontSize:9,color:C.muted}}>Balance</div><div style={{fontSize:12,color:c.balance>c.credit*.8?C.red:C.text,fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fmt(c.balance)}</div></div>
            <div><div style={{fontSize:9,color:C.muted}}>Orders</div><div style={{fontSize:12,fontFamily:"'DM Mono',monospace",fontWeight:700}}>{c.orders}</div></div>
          </div>
        </div>
      ))}
    </div>
  );

  /* Customer detail panel — plain JSX */
  const detailPanel=(
    <div style={{flex:2,background:C.surface,border:`1px solid ${C.border}`,borderRadius:10,padding:18,overflowY:"auto",boxShadow:"0 1px 3px #0000000a"}}>
      {!cust
        ?<div style={{color:C.muted,fontSize:13,textAlign:"center",marginTop:52}}>Select a customer</div>
        :<>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
            <div><div style={{fontSize:17,fontWeight:800}}>{cust.name}</div><div style={{fontSize:11,color:C.muted,marginTop:2}}>Since {cust.since}</div></div>
            <div style={{display:"flex",gap:6}}><Badge label={cust.tier}/><Btn onClick={()=>del(cust.id)} variant="danger" style={{fontSize:10,padding:"3px 9px"}}>Delete</Btn></div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:9,marginBottom:12}}>
            <StatCard label="Revenue" value={fmt(cs.reduce((a,s)=>a+s.total,0))} color={C.green}/>
            <StatCard label="Collected" value={fmt(cs.filter(s=>s.status==="Paid").reduce((a,s)=>a+s.total,0))} color={C.accent}/>
            <StatCard label="Overdue" value={fmt(cs.filter(s=>s.status==="Due").reduce((a,s)=>a+s.total,0))} color={C.red}/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
            {[["Email",cust.email],["Phone",cust.phone],["Address",cust.address],["Credit Limit",fmt(cust.credit)],["Balance",fmt(cust.balance)],["Credit Used",`${cust.credit?Math.round((cust.balance/cust.credit)*100):0}%`]].map(([l,v])=>(
              <div key={l} style={{background:C.bg,borderRadius:6,padding:"8px 10px"}}><div style={{fontSize:9,color:C.muted,marginBottom:2,textTransform:"uppercase",letterSpacing:1}}>{l}</div><div style={{fontSize:12,fontWeight:600}}>{v||"—"}</div></div>
            ))}
          </div>
          <div style={{borderTop:`1px solid ${C.border}`,paddingTop:12}}>
            <div style={{fontSize:9,fontWeight:800,color:C.muted,textTransform:"uppercase",letterSpacing:1.5,marginBottom:8}}>Order History</div>
            {cs.length===0?<div style={{color:C.muted,fontSize:12}}>No orders yet.</div>:cs.map(s=>(
              <div key={s.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${C.border}`}}>
                <div><span style={{color:C.accent,fontFamily:"'DM Mono',monospace",fontSize:11}}>#{s.id}</span><span style={{color:C.muted,fontSize:11,marginLeft:7}}>{s.date}</span></div>
                <div style={{display:"flex",alignItems:"center",gap:7}}><span style={{fontFamily:"'DM Mono',monospace",fontWeight:700,fontSize:12}}>{fmt(s.total)}</span><Badge label={s.status}/></div>
              </div>
            ))}
          </div>
        </>
      }
    </div>
  );

  if(isMobile){
    return(
      <div style={{display:"flex",flexDirection:"column",gap:11}}>
        {!sel
          ?listPanel
          :<>
            <Btn onClick={()=>setSel(null)} variant="ghost" style={{alignSelf:"flex-start",fontSize:11}}>← Back to list</Btn>
            {detailPanel}
          </>
        }
      </div>
    );
  }
  return(
    <div style={{display:"flex",gap:12,height:"calc(100vh - 140px)"}}>
      {listPanel}
      {detailPanel}
    </div>
  );
}

/* ── Deliveries ── */
function DeliveryTracker({deliveries,setDeliveries,sales,currentUser,addAudit}){
  const isMobile=useIsMobile();
  const [add,setAdd]=useState(false);const [form,setForm]=useState({orderId:"",driver:"",vehicle:"",date:"",notes:"",area:""});const [editId,setEditId]=useState(null);const [ef,setEf]=useState({});
  const unassigned=sales.filter(s=>["Invoiced","Due"].includes(s.status)&&!deliveries.find(d=>d.orderId===s.id));
  const create=()=>{const o=sales.find(s=>s.id===parseInt(form.orderId));if(!o)return;const d={id:`DEL-${String(deliveries.length+1).padStart(3,"0")}`,orderId:o.id,customer:o.customer,driver:form.driver,vehicle:form.vehicle,date:form.date,status:"Dispatched",notes:form.notes,area:form.area};setDeliveries(x=>[d,...x]);addAudit("create","Delivery",`${d.id} → ${o.customer}`,currentUser.name);setAdd(false);setForm({orderId:"",driver:"",vehicle:"",date:"",notes:"",area:""});};
  const upd=(id,st)=>{setDeliveries(d=>d.map(x=>x.id===id?{...x,status:st}:x));addAudit("update","Delivery",`${id} → ${st}`,currentUser.name);};
  const startEdit=d=>{setEditId(d.id);setEf({driver:d.driver,vehicle:d.vehicle,date:d.date,notes:d.notes||"",area:d.area||""});};
  const saveEdit=()=>{setDeliveries(d=>d.map(x=>x.id===editId?{...x,...ef}:x));addAudit("update","Delivery",`${editId} edited`,currentUser.name);setEditId(null);};
  const delDel=d=>{if(!window.confirm(`Delete ${d.id}?`))return;setDeliveries(x=>x.filter(v=>v.id!==d.id));addAudit("delete","Delivery",d.id,currentUser.name);};
  const doExport=()=>exportToExcel(deliveries.map(d=>({ID:d.id,OrderID:d.orderId,Customer:d.customer,Driver:d.driver,Vehicle:d.vehicle,Date:d.date,Status:d.status,Area:d.area||"",Notes:d.notes||""})),"natrio_deliveries.xlsx");
  return(
    <div style={{display:"flex",flexDirection:"column",gap:11}}>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}><StatCard label="Total" value={deliveries.length} icon="🚚"/><StatCard label="In Transit" value={deliveries.filter(d=>["Dispatched","En Route"].includes(d.status)).length} color={C.amber}/><StatCard label="Delivered" value={deliveries.filter(d=>d.status==="Delivered").length} color={C.green}/><StatCard label="Unassigned" value={unassigned.length} color={unassigned.length?C.red:C.text}/></div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <ImportExportBar exportFn={doExport} label="deliveries"/>
        <Btn onClick={()=>setAdd(true)}>+ Schedule Delivery</Btn>
      </div>
      {add&&<Card title="Schedule Delivery" style={{border:`1px solid ${C.accent}`}}>
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:8,marginBottom:8}}>
          <Sel value={form.orderId} onChange={e=>setForm(f=>({...f,orderId:e.target.value}))}><option value="">Select Order</option>{unassigned.map(s=><option key={s.id} value={s.id}>#{s.id} — {s.customer}</option>)}</Sel>
          <Input placeholder="Driver Name" value={form.driver} onChange={e=>setForm(f=>({...f,driver:e.target.value}))}/>
          <Input placeholder="Vehicle / Plate" value={form.vehicle} onChange={e=>setForm(f=>({...f,vehicle:e.target.value}))}/>
          <Input type="date" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))}/>
          <Input placeholder="Delivery Area" value={form.area} onChange={e=>setForm(f=>({...f,area:e.target.value}))}/>
          <Input placeholder="Notes" value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))}/>
        </div>
        <div style={{display:"flex",gap:7}}><Btn onClick={create} variant="success" disabled={!form.orderId||!form.driver}>Create</Btn><Btn onClick={()=>setAdd(false)} variant="ghost">Cancel</Btn></div>
      </Card>}
      {isMobile?(
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {deliveries.map(d=>(
            <Card key={d.id}>
              {editId===d.id?(
                <div style={{display:"flex",flexDirection:"column",gap:7}}>
                  <Input placeholder="Driver" value={ef.driver} onChange={e=>setEf(f=>({...f,driver:e.target.value}))}/>
                  <Input placeholder="Vehicle" value={ef.vehicle} onChange={e=>setEf(f=>({...f,vehicle:e.target.value}))}/>
                  <Input type="date" value={ef.date} onChange={e=>setEf(f=>({...f,date:e.target.value}))}/>
                  <Input placeholder="Area" value={ef.area} onChange={e=>setEf(f=>({...f,area:e.target.value}))}/>
                  <Input placeholder="Notes" value={ef.notes} onChange={e=>setEf(f=>({...f,notes:e.target.value}))}/>
                  <div style={{display:"flex",gap:6}}><Btn onClick={saveEdit} variant="success" style={{flex:1}}>Save</Btn><Btn onClick={()=>setEditId(null)} variant="ghost" style={{flex:1}}>Cancel</Btn></div>
                </div>
              ):(
                <>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                  <div><div style={{fontSize:12,fontWeight:700,color:C.teal}}>{d.id}</div><div style={{fontSize:11,color:C.muted}}>Order #{d.orderId} · {d.customer}</div></div>
                  <Badge label={d.status}/>
                </div>
                <div style={{fontSize:12,marginBottom:4}}>🚗 {d.driver} · {d.vehicle}</div>
                <div style={{fontSize:11,color:C.muted,marginBottom:8}}>{d.date}{d.area?` · ${d.area}`:""}</div>
                <div style={{display:"flex",gap:6,alignItems:"center"}}>
                  <Sel value={d.status} onChange={e=>upd(d.id,e.target.value)} style={{fontSize:11,padding:"4px 7px",flex:1}}>{DEL_STATUSES.map(s=><option key={s}>{s}</option>)}</Sel>
                  <Btn onClick={()=>startEdit(d)} variant="ghost" style={{fontSize:10,padding:"4px 8px"}}>✎ Edit</Btn>
                  <Btn onClick={()=>delDel(d)} variant="danger" style={{fontSize:10,padding:"4px 8px"}}>✕</Btn>
                </div>
                </>
              )}
            </Card>
          ))}
        </div>
      ):(
        <Card><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr><Th>DEL #</Th><Th>Order</Th><Th>Customer</Th><Th>Driver</Th><Th>Vehicle</Th><Th>Area</Th><Th>Date</Th><Th>Status</Th><Th>Actions</Th></tr></thead>
          <tbody>{deliveries.map(d=>(
            <tr key={d.id} style={{borderBottom:`1px solid ${C.border}`}}>
              {editId===d.id?(
                <><td colSpan={8} style={{padding:"8px 9px"}}><div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:5}}><Input value={ef.driver} onChange={e=>setEf(f=>({...f,driver:e.target.value}))} style={{fontSize:11,padding:"3px 6px"}} placeholder="Driver"/><Input value={ef.vehicle} onChange={e=>setEf(f=>({...f,vehicle:e.target.value}))} style={{fontSize:11,padding:"3px 6px"}} placeholder="Vehicle"/><Input type="date" value={ef.date} onChange={e=>setEf(f=>({...f,date:e.target.value}))} style={{fontSize:11,padding:"3px 6px"}}/><Input value={ef.area} onChange={e=>setEf(f=>({...f,area:e.target.value}))} style={{fontSize:11,padding:"3px 6px"}} placeholder="Area"/><Input value={ef.notes} onChange={e=>setEf(f=>({...f,notes:e.target.value}))} style={{fontSize:11,padding:"3px 6px"}} placeholder="Notes"/></div></td>
                <td style={{padding:"8px 9px",whiteSpace:"nowrap"}}><Btn onClick={saveEdit} variant="success" style={{fontSize:10,padding:"3px 7px",marginRight:3}}>Save</Btn><Btn onClick={()=>setEditId(null)} variant="ghost" style={{fontSize:10,padding:"3px 7px"}}>✕</Btn></td></>
              ):(
                <>
                <Td style={{color:C.teal,fontFamily:"'DM Mono',monospace",fontWeight:700}}>{d.id}</Td>
                <Td style={{color:C.accent,fontFamily:"'DM Mono',monospace"}}>#{d.orderId}</Td>
                <Td style={{fontWeight:600}}>{d.customer}</Td><Td>{d.driver}</Td><Td style={{color:C.muted}}>{d.vehicle}</Td><Td style={{color:C.muted}}>{d.area||"—"}</Td><Td style={{color:C.muted}}>{d.date}</Td>
                <Td><Badge label={d.status}/></Td>
                <Td><div style={{display:"flex",gap:4}}>
                  <Sel value={d.status} onChange={e=>upd(d.id,e.target.value)} style={{fontSize:10,padding:"3px 5px",width:"auto"}}>{DEL_STATUSES.map(s=><option key={s}>{s}</option>)}</Sel>
                  <Tip label="Edit"><Btn onClick={()=>startEdit(d)} variant="ghost" style={{fontSize:10,padding:"3px 6px"}}>✎</Btn></Tip>
                  <Tip label="Delete"><Btn onClick={()=>delDel(d)} variant="danger" style={{fontSize:10,padding:"3px 6px"}}>✕</Btn></Tip>
                </div></Td>
                </>
              )}
            </tr>
          ))}</tbody>
        </table></div></Card>
      )}
    </div>
  );
}

/* ── Expenses ── */
function Expenses({expenses,setExpenses,currentUser,addAudit}){
  const isMobile=useIsMobile();
  const [add,setAdd]=useState(false);const [form,setForm]=useState({date:today(),category:"Rent",description:"",amount:""});
  const [viewMode,setViewMode]=useState("all");const [selYear,setSelYear]=useState(new Date().getFullYear());const [selMonth,setSelMonth]=useState(new Date().getMonth());
  const years=[...new Set(expenses.map(e=>new Date(e.date).getFullYear()))].sort((a,b)=>b-a);
  const filtered=viewMode==="all"?expenses:viewMode==="year"?expenses.filter(e=>new Date(e.date).getFullYear()===selYear):expenses.filter(e=>new Date(e.date).getFullYear()===selYear&&new Date(e.date).getMonth()===selMonth);
  const total=filtered.reduce((a,e)=>a+e.amount,0);
  const byCat=EXP_CATS.map(cat=>({cat,total:filtered.filter(e=>e.category===cat).reduce((a,e)=>a+e.amount,0)})).filter(x=>x.total>0).sort((a,b)=>b.total-a.total);
  const max=Math.max(...byCat.map(x=>x.total),1);
  const byMonth=MONTHS.map((m,i)=>({m,total:expenses.filter(e=>new Date(e.date).getFullYear()===selYear&&new Date(e.date).getMonth()===i).reduce((a,e)=>a+e.amount,0)}));
  const create=()=>{if(!form.description||!form.amount)return;const e={id:uid(),date:form.date,category:form.category,description:form.description,amount:+form.amount};setExpenses(x=>[e,...x]);addAudit("create","Expense",`${form.category}: ${fmt(+form.amount)}`,currentUser.name);setAdd(false);setForm({date:today(),category:"Rent",description:"",amount:""});};
  const del=id=>{if(!window.confirm("Delete?"))return;setExpenses(x=>x.filter(e=>e.id!==id));};
  const doExport=()=>exportToExcel(filtered.map(e=>({Date:e.date,Category:e.category,Description:e.description,Amount:e.amount})),"natrio_expenses.xlsx");
  const doImport=file=>importFromExcel(file,rows=>{const mapped=rows.map(r=>({id:uid(),date:r.Date||today(),category:r.Category||"Other",description:r.Description||"",amount:+r.Amount||0}));setExpenses(x=>[...x,...mapped]);addAudit("create","Expense",`Imported ${mapped.length} expenses`,currentUser.name);});
  return(
    <div style={{display:"flex",flexDirection:"column",gap:11}}>
      <div style={{display:"flex",gap:10,flexWrap:"wrap"}}><StatCard label="Total (Filtered)" value={fmt(total)} color={C.red} icon="💸"/><StatCard label="Categories" value={byCat.length}/><StatCard label="Entries" value={filtered.length}/></div>
      <div style={{display:"flex",gap:7,alignItems:"center",flexWrap:"wrap"}}>
        {["all","month","year"].map(v=><button key={v} onClick={()=>setViewMode(v)} style={{background:viewMode===v?C.accent:"transparent",color:viewMode===v?"#fff":C.muted,border:`1px solid ${viewMode===v?C.accent:C.border}`,borderRadius:6,padding:"5px 11px",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",textTransform:"capitalize"}}>{v}</button>)}
        {viewMode!=="all"&&<Sel value={selYear} onChange={e=>setSelYear(+e.target.value)} style={{fontSize:11,padding:"5px 8px"}}>{(years.length?years:[new Date().getFullYear()]).map(y=><option key={y}>{y}</option>)}</Sel>}
        {viewMode==="month"&&<Sel value={selMonth} onChange={e=>setSelMonth(+e.target.value)} style={{fontSize:11,padding:"5px 8px"}}>{MONTHS.map((m,i)=><option key={i} value={i}>{m}</option>)}</Sel>}
        <div style={{marginLeft:"auto",display:"flex",gap:6,alignItems:"center"}}>
          <ImportExportBar exportFn={doExport} importFn={doImport} label="expenses"/>
          <Btn onClick={()=>setAdd(true)}>+ Add</Btn>
        </div>
      </div>
      {viewMode==="year"&&<Card title={`Monthly Breakdown — ${selYear}`}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(90px,1fr))",gap:8}}>
          {byMonth.map(({m,total:t})=>(
            <div key={m} style={{background:C.bg,borderRadius:7,padding:"8px 10px",textAlign:"center",cursor:"pointer",border:`1px solid ${t>0?C.accent+"44":C.border}`}} onClick={()=>{setViewMode("month");setSelMonth(MONTHS.indexOf(m));}}>
              <div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4}}>{m}</div>
              <div style={{fontSize:12,fontWeight:800,color:t>0?C.red:C.muted,fontFamily:"'DM Mono',monospace"}}>{t>0?fmt(t):"—"}</div>
            </div>
          ))}
        </div>
      </Card>}
      {add&&<Card style={{border:`1px solid ${C.accent}`}}>
        <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr":"1fr 1fr",gap:8,marginBottom:8}}>
          <Input type="date" value={form.date} onChange={e=>setForm(f=>({...f,date:e.target.value}))}/>
          <Sel value={form.category} onChange={e=>setForm(f=>({...f,category:e.target.value}))}>{EXP_CATS.map(c=><option key={c}>{c}</option>)}</Sel>
          <Input placeholder="Description" value={form.description} onChange={e=>setForm(f=>({...f,description:e.target.value}))} style={{gridColumn:isMobile?"1":"1/-1"}}/>
          <Input type="number" placeholder="Amount (PKR)" value={form.amount} onChange={e=>setForm(f=>({...f,amount:e.target.value}))}/>
        </div>
        <div style={{display:"flex",gap:7}}><Btn onClick={create} variant="success">Save</Btn><Btn onClick={()=>setAdd(false)} variant="ghost">Cancel</Btn></div>
      </Card>}
      <div style={{display:"flex",flexDirection:isMobile?"column":"row",gap:12}}>
        <div style={{flex:2}}>
          {isMobile?(
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {filtered.map(e=>(
                <Card key={e.id}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                    <div><div style={{fontSize:12,fontWeight:700}}>{e.description}</div><div style={{fontSize:10,color:C.muted,marginTop:2}}>{e.date}</div><div style={{marginTop:4}}><Badge label={e.category}/></div></div>
                    <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6}}>
                      <div style={{fontSize:14,fontWeight:800,color:C.red,fontFamily:"'DM Mono',monospace"}}>{fmt(e.amount)}</div>
                      <Btn onClick={()=>del(e.id)} variant="danger" style={{fontSize:10,padding:"3px 7px"}}>✕</Btn>
                    </div>
                  </div>
                </Card>
              ))}
            </div>
          ):(
            <Card><table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr><Th>Date</Th><Th>Category</Th><Th>Description</Th><Th>Amount</Th><Th></Th></tr></thead>
              <tbody>{filtered.map(e=>(
                <tr key={e.id} style={{borderBottom:`1px solid ${C.border}`}}>
                  <Td style={{color:C.muted}}>{e.date}</Td><Td><Badge label={e.category}/></Td>
                  <Td>{e.description}</Td>
                  <Td style={{fontFamily:"'DM Mono',monospace",fontWeight:700,color:C.red}}>{fmt(e.amount)}</Td>
                  <Td><Btn onClick={()=>del(e.id)} variant="danger" style={{fontSize:10,padding:"2px 6px"}}>✕</Btn></Td>
                </tr>
              ))}</tbody>
            </table></Card>
          )}
        </div>
        <Card style={{width:isMobile?"auto":"200px"}} title="By Category">
          {byCat.map(({cat,total:t})=>(
            <div key={cat} style={{marginBottom:9}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3,fontSize:11}}><span style={{fontWeight:600}}>{cat}</span><span style={{color:C.red,fontFamily:"'DM Mono',monospace",fontWeight:700,fontSize:10}}>{fmt(t)}</span></div>
              <div style={{height:4,background:C.bg,borderRadius:3}}><div style={{height:"100%",width:`${(t/max)*100}%`,background:C.red,borderRadius:3}}/></div>
            </div>
          ))}
        </Card>
      </div>
    </div>
  );
}

/* ── Reports ── */
function Reports({inventory,customers,sales,expenses}){
  const isMobile=useIsMobile();
  const rev=sales.reduce((a,s)=>a+s.total,0);
  const exp=expenses.reduce((a,e)=>a+e.amount,0);
  const cogs=sales.reduce((a,s)=>a+s.items.reduce((b,it)=>{const inv=inventory.find(i=>i.sku===it.sku);return b+(inv?it.qty*inv.cost:0);},0),0);
  const gross=rev-cogs;const net=gross-exp;
  const ti={};sales.forEach(s=>s.items.forEach(it=>{ti[it.name]=(ti[it.name]||0)+it.qty*it.price;}));
  const ts=Object.entries(ti).sort((a,b)=>b[1]-a[1]);
  const tc=customers.map(c=>({...c,rev:sales.filter(s=>s.customerId===c.id).reduce((a,s)=>a+s.total,0)})).sort((a,b)=>b.rev-a.rev);
  const mr=Math.max(...ts.map(x=>x[1]),1);const mc=Math.max(...tc.map(x=>x.rev),1);
  const aging=customers.map(c=>({...c,due:sales.filter(s=>s.customerId===c.id&&s.status==="Due").reduce((a,s)=>a+s.total,0)})).filter(c=>c.due>0).sort((a,b)=>b.due-a.due);
  const exportPnL=()=>exportToExcel([{Metric:"Revenue",Amount:rev},{Metric:"COGS",Amount:cogs},{Metric:"Gross Profit",Amount:gross},{Metric:"Expenses",Amount:exp},{Metric:"Net Profit",Amount:net}],"natrio_pnl.xlsx");
  const exportCustomers=()=>exportToExcel(tc.map(c=>({Customer:c.name,Tier:c.tier,Revenue:c.rev,Orders:sales.filter(s=>s.customerId===c.id).length})),"natrio_customer_report.xlsx");
  const exportInventory=()=>exportToExcel(inventory.map(i=>({SKU:i.sku,Product:i.name,Qty:i.qty,CostValue:i.qty*i.cost,SellValue:i.qty*i.price,Profit:i.qty*(i.price-i.cost)})),"natrio_inventory_value.xlsx");
  const printPnL=()=>printReportPDF("Profit & Loss Statement",[[["Revenue",fmt(rev)],["COGS",fmt(cogs)],["Gross Profit",fmt(gross)],["Expenses",fmt(exp)],["Net Profit",fmt(net)]].map(r=>r)],["Metric","Amount"]);
  const printAging=()=>printReportPDF("Customer Aging Report",aging.map(c=>[c.name,c.tier,fmt(c.due)]),["Customer","Tier","Overdue Amount"]);
  return(
    <div style={{display:"flex",flexDirection:"column",gap:13}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <div style={{fontSize:12,fontWeight:700,color:C.muted}}>Financial Summary</div>
        <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
          <Btn onClick={exportPnL} variant="ghost" style={{fontSize:11,padding:"5px 10px"}}>⬇ P&L Excel</Btn>
          <Btn onClick={printPnL} variant="ghost" style={{fontSize:11,padding:"5px 10px"}}>🖨 P&L PDF</Btn>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr 1fr":"repeat(5,1fr)",gap:9}}>
        <StatCard label="Revenue" value={fmt(rev)} color={C.green} icon="💰"/>
        <StatCard label="COGS" value={fmt(cogs)} color={C.amber} icon="📦"/>
        <StatCard label="Gross Profit" value={fmt(gross)} sub={`${rev?Math.round((gross/rev)*100):0}% margin`} color={gross>=0?C.accent:C.red} icon="📈"/>
        <StatCard label="Expenses" value={fmt(exp)} color={C.red} icon="💸"/>
        <StatCard label="Net Profit" value={fmt(net)} color={net>=0?C.green:C.red} icon="🏦"/>
      </div>
      <div style={{display:"flex",flexDirection:isMobile?"column":"row",gap:11}}>
        <Card style={{flex:1}} title="Top Products" action={<Btn onClick={()=>exportToExcel(ts.map(([n,r])=>({Product:n,Revenue:r})),"top_products.xlsx")} variant="ghost" style={{fontSize:10,padding:"3px 7px"}}>⬇</Btn>}>
          {ts.map(([name,r])=>(<div key={name} style={{marginBottom:9}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:3,fontSize:11}}><span style={{fontWeight:600}}>{name}</span><span style={{color:C.accent,fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fmt(r)}</span></div><div style={{height:4,background:C.bg,borderRadius:3}}><div style={{height:"100%",width:`${(r/mr)*100}%`,background:C.accent,borderRadius:3}}/></div></div>))}
        </Card>
        <Card style={{flex:1}} title="Top Customers" action={<Btn onClick={exportCustomers} variant="ghost" style={{fontSize:10,padding:"3px 7px"}}>⬇</Btn>}>
          {tc.map(c=>(<div key={c.id} style={{marginBottom:9}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:3,fontSize:11}}><div style={{display:"flex",gap:5,alignItems:"center"}}><span style={{fontWeight:600}}>{c.name}</span><Badge label={c.tier}/></div><span style={{color:C.green,fontFamily:"'DM Mono',monospace",fontWeight:700}}>{fmt(c.rev)}</span></div><div style={{height:4,background:C.bg,borderRadius:3}}><div style={{height:"100%",width:`${(c.rev/mc)*100}%`,background:C.green,borderRadius:3}}/></div></div>))}
        </Card>
        <Card style={{flex:1}} title="Customer Aging" action={<Btn onClick={printAging} variant="ghost" style={{fontSize:10,padding:"3px 7px"}}>🖨</Btn>}>
          {aging.length===0?<div style={{color:C.muted,fontSize:12}}>No overdue ✓</div>:aging.map(c=>(
            <div key={c.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${C.border}`}}>
              <div><div style={{fontSize:12,fontWeight:700}}>{c.name}</div><Badge label={c.tier}/></div>
              <div style={{fontSize:13,fontWeight:800,color:C.red,fontFamily:"'DM Mono',monospace"}}>{fmt(c.due)}</div>
            </div>
          ))}
        </Card>
      </div>
      <Card title="Inventory Value & Turnover" action={<Btn onClick={exportInventory} variant="ghost" style={{fontSize:10,padding:"3px 7px"}}>⬇ Excel</Btn>}>
        {isMobile?(
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {inventory.map(i=>{const sold=sales.reduce((a,s)=>a+(s.items.find(x=>x.sku===i.sku)?.qty||0),0);const t=i.qty+sold>0?Math.round((sold/(i.qty+sold))*100):0;return(
              <div key={i.id} style={{background:C.bg,borderRadius:7,padding:"10px 12px"}}>
                <div style={{fontSize:12,fontWeight:700,marginBottom:5}}>{i.name}</div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:4,fontSize:11}}>
                  <div><span style={{color:C.muted}}>Cost Val: </span><span style={{fontFamily:"'DM Mono',monospace"}}>{fmt(i.qty*i.cost)}</span></div>
                  <div><span style={{color:C.muted}}>Sell Val: </span><span style={{color:C.green,fontFamily:"'DM Mono',monospace"}}>{fmt(i.qty*i.price)}</span></div>
                  <div><span style={{color:C.muted}}>Profit: </span><span style={{color:C.accent,fontFamily:"'DM Mono',monospace"}}>{fmt(i.qty*(i.price-i.cost))}</span></div>
                  <div><span style={{color:C.muted}}>Turnover: </span><span style={{color:t>50?C.green:C.amber,fontFamily:"'DM Mono',monospace"}}>{t}%</span></div>
                </div>
              </div>
            );})}
          </div>
        ):(
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr><Th>Product</Th><Th>Qty</Th><Th>Cost Value</Th><Th>Sell Value</Th><Th>Potential Profit</Th><Th>Turnover</Th></tr></thead>
            <tbody>{inventory.map(i=>{const sold=sales.reduce((a,s)=>a+(s.items.find(x=>x.sku===i.sku)?.qty||0),0);const t=i.qty+sold>0?Math.round((sold/(i.qty+sold))*100):0;return(
              <tr key={i.id} style={{borderBottom:`1px solid ${C.border}`}}>
                <Td style={{fontWeight:600}}>{i.name}</Td><Td style={{fontFamily:"'DM Mono',monospace"}}>{i.qty}</Td>
                <Td style={{fontFamily:"'DM Mono',monospace"}}>{fmt(i.qty*i.cost)}</Td>
                <Td style={{fontFamily:"'DM Mono',monospace",color:C.green}}>{fmt(i.qty*i.price)}</Td>
                <Td style={{fontFamily:"'DM Mono',monospace",color:C.accent}}>{fmt(i.qty*(i.price-i.cost))}</Td>
                <Td><div style={{display:"flex",alignItems:"center",gap:5}}><div style={{width:48,height:4,background:C.bg,borderRadius:3}}><div style={{height:"100%",width:`${t}%`,background:t>50?C.green:C.amber,borderRadius:3}}/></div><span style={{fontSize:10,color:C.muted}}>{t}%</span></div></Td>
              </tr>
            );})}
            </tbody>
          </table>
        )}
      </Card>
    </div>
  );
}

/* ── Analytics ── */
function Analytics({sales,expenses,inventory}){
  const isMobile=useIsMobile();
  const last30=sales.filter(s=>(new Date()-new Date(s.date))/864e5<=30);
  const bySt=ORDER_STATUSES.map(st=>({st,count:sales.filter(s=>s.status===st).length,value:sales.filter(s=>s.status===st).reduce((a,s)=>a+s.total,0)})).filter(x=>x.count>0);
  const expT=expenses.reduce((a,e)=>a+e.amount,1);
  const byCat=EXP_CATS.map(cat=>({cat,total:expenses.filter(e=>e.category===cat).reduce((a,e)=>a+e.amount,0)})).filter(x=>x.total>0);
  const bs=Object.entries(sales.reduce((m,s)=>{s.items.forEach(it=>{m[it.name]=(m[it.name]||0)+it.qty;});return m;},{})).sort((a,b)=>b[1]-a[1]).slice(0,6);
  const rev=sales.reduce((a,s)=>a+s.total,0);
  const convR=sales.filter(s=>s.status==="Paid").length/(sales.length||1)*100;
  const doExport=()=>exportToExcel([...bySt.map(x=>({Type:"OrderStatus",Label:x.st,Count:x.count,Value:x.value})),...bs.map(([n,q])=>({Type:"BestSeller",Label:n,Count:q,Value:0})),...byCat.map(x=>({Type:"Expense",Label:x.cat,Count:0,Value:x.total}))],"natrio_analytics.xlsx");
  return(
    <div style={{display:"flex",flexDirection:"column",gap:13}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <div style={{fontSize:12,fontWeight:700,color:C.muted}}>Business Analytics</div>
        <ImportExportBar exportFn={doExport} label="analytics"/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:isMobile?"1fr 1fr":"repeat(4,1fr)",gap:9}}>
        <StatCard label="Avg Order" value={fmt(rev/(sales.length||1))} icon="📊"/>
        <StatCard label="Payment Rate" value={`${Math.round(convR)}%`} sub="fully paid" color={convR>70?C.green:C.amber} icon="💳"/>
        <StatCard label="Orders (30d)" value={last30.length} icon="📅"/>
        <StatCard label="Revenue (30d)" value={fmt(last30.reduce((a,s)=>a+s.total,0))} color={C.green} icon="📈"/>
      </div>
      <div style={{display:"flex",flexDirection:isMobile?"column":"row",gap:11}}>
        <Card style={{flex:1}} title="Orders by Status">
          {bySt.map(({st,count,value})=>(
            <div key={st} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${C.border}`}}>
              <div style={{display:"flex",alignItems:"center",gap:6}}><Badge label={st}/><span style={{fontSize:11,color:C.muted}}>{count}</span></div>
              <span style={{fontFamily:"'DM Mono',monospace",fontSize:12,fontWeight:700}}>{fmt(value)}</span>
            </div>
          ))}
        </Card>
        <Card style={{flex:1}} title="Best Sellers (Units)">
          {bs.map(([name,qty],i)=>(
            <div key={name} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${C.border}`}}>
              <div style={{display:"flex",gap:6,alignItems:"center"}}><span style={{width:16,height:16,borderRadius:"50%",background:C.accentSoft,color:C.accent,display:"flex",alignItems:"center",justifyContent:"center",fontSize:8,fontWeight:800}}>{i+1}</span><span style={{fontSize:12,fontWeight:600}}>{name}</span></div>
              <span style={{fontFamily:"'DM Mono',monospace",fontWeight:700,color:C.accent,fontSize:11}}>{qty} units</span>
            </div>
          ))}
        </Card>
        <Card style={{flex:1}} title="Expense Breakdown">
          {byCat.map(({cat,total})=>{const pct=Math.round((total/expT)*100);return(
            <div key={cat} style={{marginBottom:8}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:3,fontSize:11}}><span style={{fontWeight:600}}>{cat}</span><span style={{color:C.muted,fontSize:10}}>{pct}%</span></div>
              <div style={{height:4,background:C.bg,borderRadius:3}}><div style={{height:"100%",width:`${pct}%`,background:C.red,borderRadius:3}}/></div>
            </div>
          );})}
        </Card>
      </div>
    </div>
  );
}

/* ── Audit Log ── */
function AuditLog({log}){
  const isMobile=useIsMobile();
  const [f,setF]=useState("");
  const fil=log.filter(e=>!f||(e.entity+e.detail+e.user).toLowerCase().includes(f.toLowerCase()));
  const doExport=()=>exportToExcel(fil.map(e=>({Timestamp:e.ts,Action:e.action,Entity:e.entity,Detail:e.detail,User:e.user})),"natrio_audit.xlsx");
  return(
    <div style={{display:"flex",flexDirection:"column",gap:11}}>
      <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
        <Input value={f} onChange={e=>setF(e.target.value)} placeholder="Search audit log…" style={{flex:1,minWidth:160}}/>
        <ImportExportBar exportFn={doExport} label="audit log"/>
        <span style={{fontSize:11,color:C.muted,whiteSpace:"nowrap"}}>{fil.length} entries</span>
      </div>
      {isMobile?(
        <div style={{display:"flex",flexDirection:"column",gap:7}}>
          {fil.length===0&&<div style={{textAlign:"center",color:C.muted,padding:24,fontSize:13}}>No audit entries yet. Actions appear here as you use the app.</div>}
          {fil.slice(0,100).map(e=>(
            <Card key={e.id}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:4}}>
                <div style={{display:"flex",gap:6,alignItems:"center"}}><Badge label={e.action}/><span style={{fontSize:12,fontWeight:700}}>{e.entity}</span></div>
                <span style={{fontSize:10,color:C.accent,fontWeight:600}}>{e.user}</span>
              </div>
              <div style={{fontSize:11,color:C.muted,marginBottom:3}}>{e.detail}</div>
              <div style={{fontSize:10,color:C.subtle,fontFamily:"'DM Mono',monospace"}}>{new Date(e.ts).toLocaleString()}</div>
            </Card>
          ))}
        </div>
      ):(
        <Card><table style={{width:"100%",borderCollapse:"collapse"}}>
          <thead><tr><Th>Time</Th><Th>Action</Th><Th>Entity</Th><Th>Detail</Th><Th>User</Th></tr></thead>
          <tbody>
            {fil.length===0&&<tr><td colSpan={5} style={{padding:"22px",textAlign:"center",color:C.muted,fontSize:12}}>No audit entries yet. Actions appear here as you use the app.</td></tr>}
            {fil.slice(0,300).map(e=>(
              <tr key={e.id} style={{borderBottom:`1px solid ${C.border}`}}>
                <Td style={{color:C.muted,fontSize:10,fontFamily:"'DM Mono',monospace",whiteSpace:"nowrap"}}>{new Date(e.ts).toLocaleString()}</Td>
                <Td><Badge label={e.action}/></Td><Td style={{fontWeight:600,fontSize:12}}>{e.entity}</Td>
                <Td style={{color:C.muted,fontSize:11}}>{e.detail}</Td>
                <Td style={{color:C.accent,fontSize:11,fontWeight:600}}>{e.user}</Td>
              </tr>
            ))}
          </tbody>
        </table></Card>
      )}
    </div>
  );
}

/* ── Settings — FIXED controlled inputs ── */
function Settings({settings,setSettings,inventory,sales,expenses}){
  const [form,setForm]=useState({...settings});
  const [saved,setSaved]=useState(false);
  const [logo,setLogo]=useState(settings.logoDataUrl||"");
  const [tab,setTab]=useState("company");
  /* Sync when settings prop changes (e.g. on re-login) */
  useEffect(()=>{setForm({...settings});setLogo(settings.logoDataUrl||"");},[settings]);
  const handleLogo=e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>{setLogo(ev.target.result);setForm(x=>({...x,logoDataUrl:ev.target.result}));};r.readAsDataURL(f);};
  const save=()=>{setSettings(form);localStorage.setItem("natrio_settings",JSON.stringify(form));setSaved(true);setTimeout(()=>setSaved(false),2500);};
  const exportData=()=>{const d={inventory,sales,expenses,exportedAt:new Date().toISOString()};const b=new Blob([JSON.stringify(d,null,2)],{type:"application/json"});const a=document.createElement("a");a.href=URL.createObjectURL(b);a.download="natrio_backup.json";a.click();};
  /* Each field uses its own updater to avoid stale closure — fixes one-letter bug */
  const field=(key)=>({value:form[key]||"",onChange:e=>setForm(prev=>({...prev,[key]:e.target.value}))});
  const tabs=["company","tax","bank","invoice","backup"];
  return(
    <div style={{display:"flex",flexDirection:"column",gap:12}}>
      <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:4}}>
        {tabs.map(t=><button key={t} onClick={()=>setTab(t)} style={{background:tab===t?C.accent:"transparent",color:tab===t?"#fff":C.muted,border:`1px solid ${tab===t?C.accent:C.border}`,borderRadius:6,padding:"5px 12px",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",textTransform:"capitalize"}}>{t}</button>)}
      </div>
      <div style={{display:"flex",gap:16,alignItems:"flex-start",flexWrap:"wrap"}}>
        <div style={{flex:2,minWidth:280,display:"flex",flexDirection:"column",gap:11}}>
          {tab==="company"&&<Card title="Company Information"><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9}}>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Company Name</div><Input {...field("companyName")}/></div>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Tagline</div><Input {...field("tagline")}/></div>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Address Line 1</div><Input {...field("address1")}/></div>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Address Line 2</div><Input {...field("address2")}/></div>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>City</div><Input {...field("city")}/></div>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Country</div><Input {...field("country")}/></div>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Phone</div><Input {...field("phone")}/></div>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Email</div><Input {...field("email")}/></div>
            <div style={{gridColumn:"1/-1"}}><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Website</div><Input {...field("website")}/></div>
          </div></Card>}
          {tab==="tax"&&<Card title="Tax & Registration"><div style={{display:"grid",gridTemplateColumns:"1fr 2fr 1fr 2fr",gap:9}}>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Label 1</div><Input {...field("taxLabel")} placeholder="NTN"/></div>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Number 1</div><Input {...field("taxNumber")}/></div>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Label 2</div><Input {...field("taxLabel2")} placeholder="STRN"/></div>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Number 2</div><Input {...field("taxNumber2")}/></div>
          </div></Card>}
          {tab==="bank"&&<Card title="Bank / Payment Details"><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:9}}>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Bank Name</div><Input {...field("bankName")}/></div>
            <div><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Account Number</div><Input {...field("bankAccount")}/></div>
            <div style={{gridColumn:"1/-1"}}><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>IBAN</div><Input {...field("bankIBAN")} placeholder="PK00XXXX0000000000000000"/></div>
          </div></Card>}
          {tab==="invoice"&&<Card title="Invoice Footer"><div style={{fontSize:10,color:C.muted,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Footer Message</div><Textarea rows={3} {...field("invoiceNotes")}/></Card>}
          {tab==="backup"&&<Card title="Data Backup"><div style={{display:"flex",flexDirection:"column",gap:10}}>
            <div style={{fontSize:12,color:C.muted,lineHeight:1.7}}>Export all data (inventory, sales, expenses) as JSON for backup or migration.</div>
            <Btn onClick={exportData} variant="teal" style={{alignSelf:"flex-start"}}>⬇ Export All Data (JSON)</Btn>
            <div style={{fontSize:11,color:C.muted,padding:"8px 11px",background:C.bg,borderRadius:7,border:`1px solid ${C.border}`}}>⚠ Data lives in your browser session. Export regularly to avoid data loss.</div>
          </div></Card>}
          <div style={{display:"flex",justifyContent:"flex-end"}}><Btn onClick={save} style={{padding:"9px 22px"}}>{saved?"✓ Saved!":"Save Settings"}</Btn></div>
        </div>
        <div style={{width:210,display:"flex",flexDirection:"column",gap:9}}>
          <Card title="Invoice Logo">
            <div style={{width:"100%",height:100,borderRadius:7,border:`2px dashed ${C.border}`,display:"flex",alignItems:"center",justifyContent:"center",background:C.bg,marginBottom:8,overflow:"hidden"}}>
              {logo?<img src={logo} alt="Logo" style={{maxWidth:"86%",maxHeight:94,objectFit:"contain"}}/>:<div style={{textAlign:"center",color:C.muted}}><div style={{fontSize:20,marginBottom:3}}>🖼</div><div style={{fontSize:10}}>No logo</div></div>}
            </div>
            <label style={{display:"block",background:C.accent,color:"#fff",borderRadius:7,padding:"7px 12px",fontSize:12,fontWeight:600,cursor:"pointer",textAlign:"center"}}>Upload Logo<input type="file" accept="image/*" onChange={handleLogo} style={{display:"none"}}/></label>
            {logo&&<button onClick={()=>{setLogo("");setForm(f=>({...f,logoDataUrl:""}));}} style={{width:"100%",marginTop:4,background:"none",border:`1px solid ${C.border}`,borderRadius:6,padding:"5px",fontSize:10,color:C.muted,cursor:"pointer",fontFamily:"inherit"}}>Remove</button>}
          </Card>
          <Card title="Preview">
            <div style={{background:C.bg,borderRadius:6,padding:9,fontSize:11,lineHeight:1.7}}>
              <div style={{fontWeight:800,color:C.accent,marginBottom:2}}>{form.companyName||"Company Name"}</div>
              {form.address1&&<div style={{color:C.muted,fontSize:10}}>{form.address1}</div>}
              {form.city&&<div style={{color:C.muted,fontSize:10}}>{form.city}{form.country?`, ${form.country}`:""}</div>}
              {form.taxNumber&&<div style={{color:C.muted,fontSize:10}}>{form.taxLabel}: {form.taxNumber}</div>}
            </div>
          </Card>
        </div>
      </div>
    </div>
  );
}

/* ── User Management ── */
function UserManagement({users,setUsers,currentUser,addAudit}){
  const [add,setAdd]=useState(false);const [form,setForm]=useState({});const [eid,setEid]=useState(null);const [ef,setEf]=useState({});
  const create=()=>{if(!form.name||!form.username||!form.password||!form.role)return;if(users.find(u=>u.username===form.username.trim().toLowerCase())){alert("Username exists.");return;}setUsers(u=>[...u,{id:uid(),...form,username:form.username.trim().toLowerCase()}]);addAudit("create","User",form.username,currentUser.name);setAdd(false);setForm({});};
  const saveEdit=id=>{setUsers(u=>u.map(x=>x.id===id?{...x,...ef,username:ef.username.trim().toLowerCase()}:x));addAudit("update","User",ef.username,currentUser.name);setEid(null);};
  const del=id=>{if(id===currentUser.id){alert("Cannot delete yourself.");return;}if(!window.confirm("Delete?"))return;addAudit("delete","User",users.find(u=>u.id===id)?.username||"",currentUser.name);setUsers(u=>u.filter(x=>x.id!==id));};
  return(
    <div style={{display:"flex",flexDirection:"column",gap:11}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:14,fontWeight:800}}>User Accounts</div><div style={{fontSize:11,color:C.muted,marginTop:2}}>Manage access to Natrio</div></div><Btn onClick={()=>{setAdd(true);setForm({role:"viewer"});}}>+ Add User</Btn></div>
      <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{[["Admin","Full access — all modules, settings, user management"],["Sales","POS, Orders, Quotations, Customers, Deliveries, AI"],["Viewer","Dashboard, Orders, Reports, Analytics — read only"]].map(([r,d])=><Card key={r} style={{flex:1,minWidth:180}}><div style={{marginBottom:5}}><Badge label={r}/></div><div style={{fontSize:10,color:C.muted}}>{d}</div></Card>)}</div>
      {add&&<Card style={{border:`1px solid ${C.accent}`}}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:8,marginBottom:8}}><Input placeholder="Full Name" value={form.name||""} onChange={e=>setForm(x=>({...x,name:e.target.value}))}/><Input placeholder="Username" value={form.username||""} onChange={e=>setForm(x=>({...x,username:e.target.value}))}/><Input type="password" placeholder="Password" value={form.password||""} onChange={e=>setForm(x=>({...x,password:e.target.value}))}/><Sel value={form.role||"viewer"} onChange={e=>setForm(x=>({...x,role:e.target.value}))}><option value="viewer">Viewer</option><option value="sales">Sales</option><option value="admin">Admin</option></Sel></div><div style={{display:"flex",gap:7}}><Btn onClick={create} variant="success">Create</Btn><Btn onClick={()=>{setAdd(false);setForm({});}} variant="ghost">Cancel</Btn></div></Card>}
      <Card><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
        <thead><tr><Th>Name</Th><Th>Username</Th><Th>Role</Th><Th>Password</Th><Th>Actions</Th></tr></thead>
        <tbody>{users.map(u=>(
          <tr key={u.id} style={{borderBottom:`1px solid ${C.border}`,background:u.id===currentUser.id?C.accentSoft:"transparent"}}>
            {eid===u.id?(<><td colSpan={4} style={{padding:"7px 9px"}}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:5}}><Input value={ef.name||""} onChange={e=>setEf(x=>({...x,name:e.target.value}))} style={{fontSize:11,padding:"3px 7px"}}/><Input value={ef.username||""} onChange={e=>setEf(x=>({...x,username:e.target.value}))} style={{fontSize:11,padding:"3px 7px"}}/><Input type="password" value={ef.password||""} placeholder="New pwd" onChange={e=>setEf(x=>({...x,password:e.target.value}))} style={{fontSize:11,padding:"3px 7px"}}/><Sel value={ef.role||"viewer"} onChange={e=>setEf(x=>({...x,role:e.target.value}))} style={{fontSize:11,padding:"3px 7px"}}><option value="viewer">Viewer</option><option value="sales">Sales</option><option value="admin">Admin</option></Sel></div></td><td style={{padding:"7px 9px",whiteSpace:"nowrap"}}><Btn onClick={()=>saveEdit(u.id)} variant="success" style={{fontSize:10,padding:"3px 7px",marginRight:3}}>Save</Btn><Btn onClick={()=>setEid(null)} variant="ghost" style={{fontSize:10,padding:"3px 7px"}}>✕</Btn></td></>
            ):(<><Td><div style={{fontWeight:700,fontSize:12}}>{u.name}</div>{u.id===currentUser.id&&<div style={{fontSize:9,color:C.accent}}>● You</div>}</Td><Td style={{fontFamily:"'DM Mono',monospace",color:C.muted,fontSize:11}}>{u.username}</Td><Td><Badge label={u.role.charAt(0).toUpperCase()+u.role.slice(1)}/></Td><Td style={{fontFamily:"'DM Mono',monospace",color:C.subtle}}>{"•".repeat(Math.min(u.password.length,10))}</Td><Td><Btn onClick={()=>{setEid(u.id);setEf({...u});}} variant="ghost" style={{fontSize:10,padding:"3px 6px",marginRight:3}}>Edit</Btn><Btn onClick={()=>del(u.id)} variant="danger" style={{fontSize:10,padding:"3px 6px"}} disabled={u.id===currentUser.id}>Del</Btn></Td></>)}
          </tr>
        ))}</tbody>
      </table></div></Card>
    </div>
  );
}

/* ── Root App ── */
export default function App(){
  const isMobile=useIsMobile();
  const [user,setUser]=useState(null);
  const [tab,setTab]=useState("dashboard");
  const [navOpen,setNavOpen]=useState(false);
  const [inventory,setInventory]=useState([]);
  const [customers,setCustomers]=useState([]);
  const [sales,setSales]=useState([]);
  const [users,setUsers]=useState([{id:1,name:"Admin User",username:"admin",password:"admin123",role:"admin"}]);
  const [suppliers]=useState([]);
  const [pos,setPOs]=useState([]);
  const [expenses,setExpenses]=useState([]);
  const [deliveries,setDeliveries]=useState([]);
  const [auditLog,setAuditLog]=useState([]);
  const [settings,setSettings]=useState(()=>{try{const s=localStorage.getItem("natrio_settings");return s?{...DEFAULT_SETTINGS,...JSON.parse(s)}:DEFAULT_SETTINGS;}catch{return DEFAULT_SETTINGS;}});
  const [showNotif,setShowNotif]=useState(false);

  /* ── Database state ── */
  const [dbLoading,setDbLoading]=useState(isConnected());
  const [syncStatus,setSyncStatus]=useState("idle");
  const [syncError,setSyncError]=useState("");
  const initialized=useRef(false);
  const syncedIds=useRef({inventory:[],customers:[],sales:[],expenses:[],deliveries:[],pos:[],auditLog:[],users:[]});

  /* ── Load all data from Supabase on first mount ── */
  useEffect(()=>{
    if(!isConnected()){
      setAuditLog(loadAudit());
      initialized.current=true;
      setDbLoading(false);
      setSyncStatus("offline");
      return;
    }
    (async()=>{
      setSyncStatus("syncing");
      /* 1. Test the connection first — shows a clear error if misconfigured */
      const test=await import("./supabase.js").then(m=>m.testConnection());
      if(!test.ok){
        setSyncError(test.error);
        setSyncStatus("error");
        setAuditLog(loadAudit());
        initialized.current=true;
        setDbLoading(false);
        return;
      }
      try{
        const [inv,custs,sls,exps,dels,purchOrds,auditEntries,dbUsers,settingsData]=await Promise.all([
          db.load("inventory"),db.load("customers"),db.load("sales"),
          db.load("expenses"),db.load("deliveries"),db.load("purchase_orders"),
          db.load("audit_log"),db.load("users"),db.loadSetting("main_settings"),
        ]);
        /* Load whatever is in the DB — no seed data fallbacks */
        setInventory(inv);      syncedIds.current.inventory=inv.map(i=>String(i.id));
        setCustomers(custs);    syncedIds.current.customers=custs.map(i=>String(i.id));
        setSales(sls);          syncedIds.current.sales=sls.map(i=>String(i.id));
        setExpenses(exps);      syncedIds.current.expenses=exps.map(i=>String(i.id));
        setDeliveries(dels);    syncedIds.current.deliveries=dels.map(i=>String(i.id));
        setPOs(purchOrds);      syncedIds.current.pos=purchOrds.map(i=>String(i.id));
        if(dbUsers.length){setUsers(dbUsers);syncedIds.current.users=dbUsers.map(i=>String(i.id));}
        else{
          /* First run — push default admin user to DB */
          const defaultUsers=[{id:1,name:"Admin User",username:"admin",password:"admin123",role:"admin"}];
          setUsers(defaultUsers);
          await db.upsertMany("users",defaultUsers);
          syncedIds.current.users=defaultUsers.map(i=>String(i.id));
        }
        if(auditEntries.length){setAuditLog(auditEntries);syncedIds.current.auditLog=auditEntries.map(i=>String(i.id));}
        else{setAuditLog(loadAudit());}
        if(settingsData){setSettings(s=>({...DEFAULT_SETTINGS,...s,...settingsData}));}
        setSyncStatus("saved");
        setSyncError("");
      }catch(e){
        console.error("[DB] Load failed:",e);
        setSyncStatus("error");
        setSyncError(e.message||"Unknown error");
        setAuditLog(loadAudit());
      }
      initialized.current=true;
      setDbLoading(false);
    })();
  },[]);

  /* ── Generic sync: upsert all + delete removed ── */
  const syncCollection=useCallback(async(table,items,refKey)=>{
    if(!isConnected()||!initialized.current)return;
    try{
      setSyncStatus("syncing");
      const currentIds=new Set(items.map(i=>String(i.id)));
      const deletedIds=syncedIds.current[refKey].filter(id=>!currentIds.has(id));
      await db.upsertMany(table,items);
      if(deletedIds.length)await db.deleteMany(table,deletedIds);
      syncedIds.current[refKey]=[...currentIds];
      setSyncStatus("saved");
      setSyncError("");
    }catch(e){
      console.error(`[DB] Sync ${table}:`,e);
      setSyncStatus("error");
      setSyncError(e.message||"Sync failed");
    }
  },[]);

  /* Force manual sync of everything */
  const forceSync=useCallback(async()=>{
    if(!isConnected())return;
    setSyncStatus("syncing");
    try{
      await Promise.all([
        db.upsertMany("inventory",inventory),
        db.upsertMany("customers",customers),
        db.upsertMany("sales",sales),
        db.upsertMany("expenses",expenses),
        db.upsertMany("deliveries",deliveries),
        db.upsertMany("purchase_orders",pos),
        db.upsertMany("users",users),
      ]);
      syncedIds.current.inventory=inventory.map(i=>String(i.id));
      syncedIds.current.customers=customers.map(i=>String(i.id));
      syncedIds.current.sales=sales.map(i=>String(i.id));
      syncedIds.current.expenses=expenses.map(i=>String(i.id));
      syncedIds.current.deliveries=deliveries.map(i=>String(i.id));
      syncedIds.current.pos=pos.map(i=>String(i.id));
      syncedIds.current.users=users.map(i=>String(i.id));
      setSyncStatus("saved");setSyncError("");
    }catch(e){setSyncStatus("error");setSyncError(e.message||"Force sync failed");}
  },[inventory,customers,sales,expenses,deliveries,pos]);

  /* ── Debounced sync on state change ── */
  useEffect(()=>{if(!initialized.current)return;const t=setTimeout(()=>syncCollection("inventory",inventory,"inventory"),800);return()=>clearTimeout(t);},[inventory]);
  useEffect(()=>{if(!initialized.current)return;const t=setTimeout(()=>syncCollection("customers",customers,"customers"),800);return()=>clearTimeout(t);},[customers]);
  useEffect(()=>{if(!initialized.current)return;const t=setTimeout(()=>syncCollection("sales",sales,"sales"),800);return()=>clearTimeout(t);},[sales]);
  useEffect(()=>{if(!initialized.current)return;const t=setTimeout(()=>syncCollection("expenses",expenses,"expenses"),800);return()=>clearTimeout(t);},[expenses]);
  useEffect(()=>{if(!initialized.current)return;const t=setTimeout(()=>syncCollection("deliveries",deliveries,"deliveries"),800);return()=>clearTimeout(t);},[deliveries]);
  useEffect(()=>{if(!initialized.current)return;const t=setTimeout(()=>syncCollection("purchase_orders",pos,"pos"),800);return()=>clearTimeout(t);},[pos]);
  useEffect(()=>{if(!initialized.current)return;const t=setTimeout(()=>syncCollection("users",users,"users"),800);return()=>clearTimeout(t);},[users]);
  useEffect(()=>{
    if(!initialized.current)return;
    const t=setTimeout(()=>db.saveSetting("main_settings",settings).catch(e=>console.error("[DB] Settings:",e)),800);
    return()=>clearTimeout(t);
  },[settings]);

  /* ── Audit log — append only ── */
  useEffect(()=>{
    if(!initialized.current||!auditLog.length)return;
    const newest=auditLog[0];
    if(isConnected()){
      db.upsertOne("audit_log",newest).catch(e=>console.error("[DB] Audit:",e));
      if(!syncedIds.current.auditLog.includes(String(newest.id)))
        syncedIds.current.auditLog=[String(newest.id),...syncedIds.current.auditLog];
    } else { saveAudit(auditLog); }
  },[auditLog]);

  /* ── Session restore ── */
  useEffect(()=>{const s=sessionStorage.getItem("natrio_user");if(s)setUser(JSON.parse(s));},[]);

  const addAudit=useCallback((action,entity,detail,who)=>{
    const e={id:uid(),ts:new Date().toISOString(),action,entity,detail,user:who||"system"};
    setAuditLog(a=>[e,...a].slice(0,1000));
  },[]);

  const login=u=>{sessionStorage.setItem("natrio_user",JSON.stringify(u));setUser(u);addAudit("login","Auth",`${u.name} signed in`,u.name);setTab("dashboard");setNavOpen(false);};
  const logout=()=>{sessionStorage.removeItem("natrio_user");setUser(null);setNavOpen(false);};

  /* ── DB loading screen ── */
  if(dbLoading)return(
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",background:C.bg,fontFamily:"'DM Sans','Segoe UI',sans-serif",flexDirection:"column",gap:16,padding:24}}>
      <img src="/logo.jpeg" alt="Natrio" style={{width:88,height:88,objectFit:"contain",borderRadius:12}}/>
      <div style={{fontSize:16,fontWeight:800,color:C.text}}>Loading Natrio…</div>
      <div style={{fontSize:12,color:C.muted}}>Connecting to database, please wait</div>
      <div style={{width:200,height:5,background:C.border,borderRadius:3,overflow:"hidden"}}>
        <div style={{height:"100%",background:C.accent,borderRadius:3,animation:"load 1.4s ease-in-out infinite"}}/>
      </div>
      <style>{`@keyframes load{0%{width:0%;margin-left:0%}50%{width:55%;margin-left:22%}100%{width:0%;margin-left:100%}}`}</style>
    </div>
  );

  if(!user)return <LoginScreen users={users} onLogin={login}/>;

  const allowed=ROLE_TABS[user.role]||["dashboard"];
  const NAV=ALL_NAV.filter(n=>allowed.includes(n.id));
  const canEdit=user.role!=="viewer";
  const active=allowed.includes(tab)?tab:allowed[0];

  const lowN=inventory.filter(i=>i.qty<=i.reorder).length;
  const overN=sales.filter(s=>s.status==="Due").length;
  const notifs=[
    ...inventory.filter(i=>i.qty<=i.reorder).map(i=>({type:"stock",msg:`Low stock: ${i.name} (${i.qty} left)`})),
    ...sales.filter(s=>s.status==="Due").map(s=>({type:"pay",msg:`Overdue: ${s.customer} — ${fmt(s.total)}`})),
  ];
  const roleFg={admin:C.accent,sales:C.green,viewer:C.amber}[user.role];
  const roleBg={admin:C.accentSoft,sales:C.greenSoft,viewer:C.amberSoft}[user.role];
  const cp={currentUser:user,addAudit};

  const switchTab=t=>{setTab(t);setNavOpen(false);};

  const SidebarContent=()=>(
    <>
      <div style={{padding:"12px 12px 9px",borderBottom:`1px solid ${C.border}`}}>
        <img src="/logo.jpeg" alt="Natrio" style={{width:"100%",maxWidth:120,display:"block",margin:"0 auto 4px",objectFit:"contain",borderRadius:6}}/>
        <div style={{fontSize:9,color:C.muted,textAlign:"center"}}>Distribution Suite</div>
      </div>
      <nav style={{flex:1,padding:"4px 5px",overflowY:"auto"}}>
        {NAV.map(n=>(
          <button key={n.id} onClick={()=>switchTab(n.id)} style={{display:"flex",alignItems:"center",gap:7,width:"100%",padding:"8px 8px",borderRadius:6,marginBottom:1,background:active===n.id?C.accentSoft:"transparent",color:active===n.id?C.accent:C.muted,border:active===n.id?`1px solid ${C.accent}22`:"1px solid transparent",fontSize:12,fontWeight:active===n.id?700:500,cursor:"pointer",fontFamily:"inherit",textAlign:"left",transition:"all .1s"}}>
            <span style={{fontSize:13,width:17,textAlign:"center",flexShrink:0}}>{n.icon}</span>
            <span style={{flex:1,fontSize:11}}>{n.label}</span>
            {n.id==="orders"&&overN>0&&<span style={{background:C.red,color:"#fff",borderRadius:"50%",minWidth:15,height:15,fontSize:8,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center",padding:"0 3px"}}>{overN}</span>}
            {n.id==="inventory"&&lowN>0&&<span style={{background:C.amber,color:"#fff",borderRadius:"50%",minWidth:15,height:15,fontSize:8,fontWeight:800,display:"flex",alignItems:"center",justifyContent:"center",padding:"0 3px"}}>{lowN}</span>}
          </button>
        ))}
      </nav>
      <div style={{padding:"9px 9px",borderTop:`1px solid ${C.border}`}}>
        <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6}}>
          <div style={{width:26,height:26,borderRadius:"50%",background:roleBg,border:`1px solid ${roleFg}44`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,fontWeight:800,color:roleFg,flexShrink:0}}>{user.name.charAt(0).toUpperCase()}</div>
          <div style={{overflow:"hidden",flex:1}}><div style={{fontSize:11,fontWeight:700,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{user.name}</div><Badge label={user.role.charAt(0).toUpperCase()+user.role.slice(1)}/></div>
        </div>
        <button onClick={logout} style={{width:"100%",background:C.redSoft,color:C.red,border:`1px solid ${C.red}22`,borderRadius:6,padding:"5px",fontSize:10,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>Sign Out</button>
      </div>
    </>
  );

  return(
    <div style={{display:"flex",height:"100vh",background:C.bg,fontFamily:"'DM Sans','Segoe UI',sans-serif",color:C.text,overflow:"hidden"}}>
      {/* Desktop sidebar */}
      {!isMobile&&(
        <div style={{width:198,background:C.surface,borderRight:`1px solid ${C.border}`,display:"flex",flexDirection:"column",flexShrink:0,boxShadow:"2px 0 6px #0000000a"}}>
          <SidebarContent/>
        </div>
      )}
      {/* Mobile drawer overlay */}
      {isMobile&&navOpen&&(
        <div style={{position:"fixed",inset:0,zIndex:500}} onClick={()=>setNavOpen(false)}>
          <div style={{position:"absolute",inset:0,background:"#00000055"}}/>
          <div style={{position:"absolute",left:0,top:0,bottom:0,width:240,background:C.surface,display:"flex",flexDirection:"column",boxShadow:"4px 0 20px #0000001a"}} onClick={e=>e.stopPropagation()}>
            <SidebarContent/>
          </div>
        </div>
      )}
      {/* Main */}
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <div style={{padding:isMobile?"10px 14px":"11px 20px",borderBottom:`1px solid ${C.border}`,display:"flex",alignItems:"center",justifyContent:"space-between",background:C.surface,boxShadow:"0 1px 3px #0000000a",gap:8}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            {isMobile&&<button onClick={()=>setNavOpen(v=>!v)} style={{background:"none",border:`1px solid ${C.border}`,borderRadius:7,padding:"6px 10px",fontSize:16,cursor:"pointer",lineHeight:1}}>☰</button>}
            <div>
              <div style={{fontSize:isMobile?14:16,fontWeight:800,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:isMobile?160:400}}>{ALL_NAV.find(n=>n.id===active)?.label}</div>
              {!isMobile&&<div style={{fontSize:10,color:C.muted,marginTop:1}}>{new Date().toLocaleDateString("en-US",{weekday:"long",year:"numeric",month:"long",day:"numeric"})}</div>}
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:7}}>
            <div style={{position:"relative"}}>
              <button onClick={()=>setShowNotif(v=>!v)} style={{background:notifs.length?C.amberSoft:"transparent",border:`1px solid ${notifs.length?C.amber:C.border}`,borderRadius:6,padding:"5px 9px",fontSize:11,cursor:"pointer",fontFamily:"inherit",color:notifs.length?C.amber:C.muted,fontWeight:600}}>
                🔔{notifs.length>0&&<span style={{fontWeight:800,marginLeft:3}}>{notifs.length}</span>}
              </button>
              {showNotif&&notifs.length>0&&(
                <div style={{position:"fixed",right:isMobile?8:16,top:55,width:isMobile?"calc(100vw - 16px)":270,maxWidth:340,background:C.surface,border:`1px solid ${C.border}`,borderRadius:9,boxShadow:"0 8px 24px #00000014",zIndex:200,padding:10}}>
                  <div style={{fontSize:9,fontWeight:800,color:C.muted,textTransform:"uppercase",letterSpacing:1,marginBottom:6}}>Notifications</div>
                  {notifs.map((n,i)=><div key={i} style={{padding:"5px 0",borderBottom:`1px solid ${C.border}`,fontSize:11,color:n.type==="pay"?C.red:C.amber}}>{n.msg}</div>)}
                </div>
              )}
            </div>
            {!isMobile&&user.role==="viewer"&&<div style={{background:C.amberSoft,border:`1px solid ${C.amber}44`,borderRadius:6,padding:"4px 9px",fontSize:10,color:C.amber,fontWeight:600}}>👁 View Only</div>}
            {!isMobile&&<div style={{background:C.bg,border:`1px solid ${C.border}`,borderRadius:6,padding:"4px 10px",fontSize:10,color:C.muted}}>{inventory.length} products · {customers.length} customers · {sales.length} orders</div>}
            {/* Sync status */}
            {isConnected()&&(
              <div style={{display:"flex",alignItems:"center",gap:6}}>
                {syncStatus==="error"&&syncError&&(
                  <div title={syncError} style={{fontSize:10,color:C.red,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",cursor:"help"}}>⚠ {syncError.split("\n")[0]}</div>
                )}
                <div style={{fontSize:10,fontWeight:600,padding:"4px 9px",borderRadius:6,border:"1px solid",cursor:syncStatus==="error"?"pointer":"default",
                  background:syncStatus==="saved"?C.greenSoft:syncStatus==="syncing"?C.accentSoft:syncStatus==="error"?C.redSoft:C.bg,
                  color:syncStatus==="saved"?C.green:syncStatus==="syncing"?C.accent:syncStatus==="error"?C.red:C.muted,
                  borderColor:syncStatus==="saved"?C.green+"44":syncStatus==="syncing"?C.accent+"44":syncStatus==="error"?C.red+"44":C.border,
                  whiteSpace:"nowrap"}}
                  onClick={syncStatus==="error"?forceSync:undefined}
                  title={syncStatus==="error"?`Error: ${syncError}\n\nClick to retry`:"Database sync status"}>
                  {syncStatus==="saved"?"✓ Saved":syncStatus==="syncing"?"⟳ Saving…":syncStatus==="error"?"↻ Retry":"●"}
                </div>
              </div>
            )}
            {!isConnected()&&<div style={{fontSize:10,color:C.muted,padding:"4px 9px",border:`1px solid ${C.border}`,borderRadius:6}}>⚡ Offline</div>}
          </div>
        </div>
        <div style={{flex:1,overflow:"auto",padding:active==="ai"?0:isMobile?12:20}} onClick={()=>showNotif&&setShowNotif(false)}>
          {active==="dashboard"  && <Dashboard inventory={inventory} customers={customers} sales={sales} expenses={expenses} deliveries={deliveries}/>}
          {active==="pos"        && canEdit && <POS inventory={inventory} setInventory={setInventory} customers={customers} sales={sales} setSales={setSales} {...cp}/>}
          {active==="orders"     && <Orders sales={sales} setSales={setSales} customers={customers} settings={settings} canEdit={canEdit} {...cp}/>}
          {active==="quotations" && canEdit && <Quotations inventory={inventory} customers={customers} setSales={setSales} setInventory={setInventory} settings={settings} {...cp}/>}
          {active==="inventory"  && canEdit && <Inventory inventory={inventory} setInventory={setInventory} {...cp}/>}
          {active==="purchase"   && user.role==="admin" && <PurchaseOrders pos={pos} setPOs={setPOs} inventory={inventory} setInventory={setInventory} suppliers={suppliers} {...cp}/>}
          {active==="crm"        && canEdit && <CRM customers={customers} setCustomers={setCustomers} sales={sales} {...cp}/>}
          {active==="delivery"   && <DeliveryTracker deliveries={deliveries} setDeliveries={setDeliveries} sales={sales} {...cp}/>}
          {active==="expenses"   && user.role==="admin" && <Expenses expenses={expenses} setExpenses={setExpenses} {...cp}/>}
          {active==="reports"    && <Reports inventory={inventory} customers={customers} sales={sales} expenses={expenses}/>}
          {active==="analytics"  && <Analytics sales={sales} expenses={expenses} inventory={inventory}/>}
          {active==="ai"         && <AIAssistant inventory={inventory} customers={customers} sales={sales} expenses={expenses}/>}
          {active==="users"      && user.role==="admin" && <UserManagement users={users} setUsers={setUsers} {...cp}/>}
          {active==="audit"      && user.role==="admin" && <AuditLog log={auditLog}/>}
          {active==="settings"   && user.role==="admin" && <Settings settings={settings} setSettings={setSettings} inventory={inventory} sales={sales} expenses={expenses}/>}
        </div>
      </div>
    </div>
  );
}
