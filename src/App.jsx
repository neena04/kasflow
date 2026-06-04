import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

// ── Supabase ──────────────────────────────────────────────────────────────────
const SUPABASE_URL = "https://mbsgydpeylhnaiodvdaj.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1ic2d5ZHBleWxobmFpb2R2ZGFqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk0MTk1NjEsImV4cCI6MjA5NDk5NTU2MX0.ocqwqozFFbmp9DAVbSezKVQBmcYjMPPjb0dkaq8BPlY";
const sb = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// ── Password ──────────────────────────────────────────────────────────────────
const APP_PASSWORD = "MabocAlways2026!";

// ── Defaults ──────────────────────────────────────────────────────────────────
const DEFAULT_CATEGORIES = [
  "IPL Collection","Cleaning Supplies","Supplier Payment","E-commerce Purchase",
  "Salary & Payroll","Utility & Overhead","Customer Refund","Software & Subscriptions",
  "Inter-account Transfer","Client Deposit","Client Disbursement","Other Income","Other Expense","Uncategorized",
];
const DEFAULT_SOURCES = [
  {name:"BCA",color:"#1a3a5c",import_type:"file"},
  {name:"BCA Credit Card",color:"#2c5f8a",import_type:"file"},
  {name:"BNC",color:"#7a1c1c",import_type:"file"},
  {name:"Xendit",color:"#2d1a5c",import_type:"file"},
  {name:"GoPay",color:"#1a4a5c",import_type:"screenshot"},
  {name:"Astro",color:"#5c3a1a",import_type:"screenshot"},
  {name:"ShopeePay",color:"#5c1a1a",import_type:"screenshot"},
  {name:"Manual",color:"#3a3a3a",import_type:"file"},
];
const INKCOLS = ["#1a3a5c","#2c5f8a","#7a1c1c","#2d1a5c","#1a4a5c","#5c3a1a","#5c1a1a","#3a3a3a","#1a5c3a","#3a1a5c","#5c4a1a"];
const MONTH_ID = {JAN:1,FEB:2,MAR:3,APR:4,MEI:5,JUN:6,JUL:7,AGS:8,SEP:9,OKT:10,NOV:11,DES:12,MAY:5,AUG:8,OCT:10,DEC:12};

const fmt = (n) => new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",maximumFractionDigits:0}).format(n);
const todayStr = () => new Date().toLocaleDateString("en-US",{weekday:"long",year:"numeric",month:"long",day:"numeric"}).toUpperCase();

// ── Helpers ───────────────────────────────────────────────────────────────────
const parseCSVDate=(str)=>{
  if(!str) return "";
  if(typeof str==="number"){const d=XLSX.SSF.parse_date_code(str);if(d)return `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;}
  const s=String(str).trim();
  const parts=s.split(/[\/\-\.]/);
  if(parts.length===3){
    // YYYY-MM-DD or YYYY/MM/DD
    if(parts[0].length===4)return `${parts[0]}-${parts[1].padStart(2,"0")}-${parts[2].padStart(2,"0")}`;
    // DD/MM/YYYY — Indonesian format, always treat first part as day
    if(parts[2].length===4)return `${parts[2]}-${parts[1].padStart(2,"0")}-${parts[0].padStart(2,"0")}`;
  }
  // Last resort: let JS parse (YYYY-MM-DD strings only)
  const d=new Date(s);if(!isNaN(d)&&s.includes("-"))return d.toISOString().split("T")[0];
  return "";
};
const parseAmount=(str)=>{
  if(str===null||str===undefined||str==="")return 0;
  if(typeof str==="number")return str;
  const s=String(str).replace(/[^0-9,.\-]/g,"");
  // Detect format: if ends with .XX (2 decimal digits), dot is decimal separator
  // e.g. "70,000,000.00" => remove commas => "70000000.00"
  // e.g. "25.000,00" (European) => remove dots, replace comma => "25000.00"
  if(/\.\d{2}$/.test(s)){return parseFloat(s.replace(/,/g,""))||0;}
  if(/,\d{2}$/.test(s)){return parseFloat(s.replace(/\./g,"").replace(",","."))||0;}
  return parseFloat(s.replace(/,/g,"").replace(/\./g,""))||0;
};

// ── PDF.js ────────────────────────────────────────────────────────────────────
const loadPdfJs=()=>new Promise((resolve,reject)=>{
  if(window.pdfjsLib){resolve(window.pdfjsLib);return;}
  const s=document.createElement("script");
  s.src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
  s.onload=()=>{window.pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";resolve(window.pdfjsLib);};
  s.onerror=reject;document.head.appendChild(s);
});
const extractPDFText=async(arrayBuffer)=>{
  const lib=await loadPdfJs();
  const pdf=await lib.getDocument({data:arrayBuffer}).promise;
  let text="";
  for(let p=1;p<=pdf.numPages;p++){
    const page=await pdf.getPage(p);
    const content=await page.getTextContent();
    const items=content.items.slice().sort((a,b)=>{const dy=Math.round(b.transform[5])-Math.round(a.transform[5]);return dy!==0?dy:a.transform[4]-b.transform[4];});
    text+=items.map(i=>i.str).join(" ")+"\n";
  }
  return text;
};
const resolveDate=(ddmmm)=>{
  const parts=ddmmm.split(/[-\/]/);if(parts.length<2)return "";
  const day=parts[0].padStart(2,"0");const mon=MONTH_ID[parts[1].toUpperCase()];if(!mon)return "";
  const now=new Date();const year=mon>now.getMonth()+2?now.getFullYear()-1:now.getFullYear();
  return `${year}-${String(mon).padStart(2,"0")}-${day}`;
};
const parseBCACreditCardPDF=(text,source)=>{
  const rows=[];
  const parseIDRAmount=(s)=>parseFloat(s.replace(/\./g,""))||0;
  const isSkip=(s)=>/subtotal|saldo sebelumnya|tagihan baru|kredit limit|batas tarik|tunggakan|visa sq|suku bunga|informasi/i.test(s);
  const dateRe=/^\d{2}-[A-Z]{3}$/i;
  const amtRe=/^([\d.]+)\s*(CR)?$/i;

  // PDF.js may extract each cell separately, so we work token by token
  const tokens=text.split(/\s+/).filter(Boolean);
  let i=0;
  while(i<tokens.length){
    // Look for a date token DD-MMM
    if(!dateRe.test(tokens[i])){i++;continue;}
    const date=resolveDate(tokens[i]);
    if(!date){i++;continue;}
    i++;
    // Skip optional second date token
    if(i<tokens.length&&dateRe.test(tokens[i]))i++;
    // Collect description tokens until we hit an amount
    const descTokens=[];
    while(i<tokens.length&&!amtRe.test(tokens[i])&&!dateRe.test(tokens[i])){
      descTokens.push(tokens[i]);i++;
    }
    const desc=descTokens.join(" ").trim();
    if(!desc||isSkip(desc))continue;
    // Now expect amount
    if(i>=tokens.length)continue;
    const amtMatch=tokens[i].match(/^([\d.]+)$/);
    if(!amtMatch){continue;}
    const amount=parseIDRAmount(amtMatch[1]);
    i++;
    // Check for CR marker
    const isCredit=i<tokens.length&&tokens[i].toUpperCase()==="CR";
    if(isCredit)i++;
    if(amount===0)continue;
    rows.push({id:crypto.randomUUID(),date,description:desc,amount:isCredit?amount:-amount,type:isCredit?"in":"out",source,category:"Uncategorized"});
  }
  return rows;
};
const parseGenericPDF=(text,source)=>{
  const rows=[];const lineRe=/(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}|\d{2}[-\/][A-Z]{3}[-\/]?\d{0,4})\s+(.{3,60}?)\s+([\d.,]{3,})\s*(CR|DB|K|D)?/i;
  for(const line of text.split("\n")){
    const m=line.trim().match(lineRe);if(!m)continue;
    const date=parseCSVDate(m[1])||"";if(!date)continue;
    const desc=m[2].trim();if(desc.length<2)continue;
    const amount=parseAmount(m[3]);if(amount===0)continue;
    const tag=(m[4]||"").toUpperCase();const isCredit=tag==="CR"||tag==="K";
    rows.push({id:crypto.randomUUID(),date,description:desc,amount:isCredit?amount:-amount,type:isCredit?"in":"out",source,category:"Uncategorized"});
  }
  return rows;
};
const parsePDFBuffer=async(arrayBuffer,source)=>{
  const text=await extractPDFText(arrayBuffer);
  if(/REKENING KARTU KREDIT|KARTU KREDIT BCA/i.test(text))return parseBCACreditCardPDF(text,source);
  return parseGenericPDF(text,source);
};
// BCA Corporate (Mutasi Rekening) - fields wrapped in quotes, separated by ","
const parseBCACorporateCSV=(text,source)=>{
  const rows=[];
  for(const line of text.split("\n")){
    const l=line.trim().replace(/^"|"$/g,"");
    const cols=l.split('","').map(c=>c.replace(/^"|"$/g,"").trim());
    if(cols.length<4)continue;
    const dateStr=parseCSVDate(cols[0]);if(!dateStr)continue;
    const desc=cols[1]||"";
    // Amount col may itself contain commas: "25,000.00 CR"
    const amtRaw=cols.slice(3).join(",");
    const amtMatch=amtRaw.match(/([\d,]+\.?\d*)\s*(CR|DB)/i);
    if(!amtMatch)continue;
    const amount=parseAmount(amtMatch[1]);if(amount===0)continue;
    const isCredit=amtMatch[2].toUpperCase()==="CR";
    rows.push({id:crypto.randomUUID(),date:dateStr,description:desc,amount:isCredit?amount:-amount,type:isCredit?"in":"out",source,category:"Uncategorized"});
  }
  return rows;
};
const parseBCACSV=(text,source)=>text.trim().split("\n").filter(l=>l.trim()).flatMap(line=>{
  const cols=line.split(",").map(c=>c.replace(/^"|"$/g,"").trim());
  if(cols.length<4)return [];
  const dateStr=parseCSVDate(cols[0]);if(!dateStr)return [];
  const desc=cols[1]||cols[2]||"";
  let debit=0,credit=0;
  if(cols.length>=5){debit=parseAmount(cols[2]);credit=parseAmount(cols[3]);}
  else if(cols.length===4){const amt=parseAmount(cols[2]);const t=(cols[3]||"").toLowerCase();if(t.includes("db")||t.includes("debit"))debit=amt;else credit=amt;}
  if(debit===0&&credit===0)return [];
  return [{id:crypto.randomUUID(),date:dateStr,description:desc,amount:credit>0?credit:-debit,type:credit>0?"in":"out",source,category:"Uncategorized"}];
});
const parseGenericCSV=(text,source)=>{
  const lines=text.trim().split("\n").filter(l=>l.trim());if(lines.length<2)return [];
  const headers=lines[0].split(",").map(h=>h.replace(/^"|"$/g,"").trim().toLowerCase());
  return lines.slice(1).flatMap(line=>{
    const cols=line.split(",").map(c=>c.replace(/^"|"$/g,"").trim());
    const obj={};headers.forEach((h,i)=>{obj[h]=cols[i]||"";});
    const dateKey=headers.find(h=>h.includes("date")||h.includes("tanggal")||h.includes("tgl"));
    const descKey=headers.find(h=>h.includes("desc")||h.includes("keterangan")||h.includes("narasi")||h.includes("ket"));
    const amtKey=headers.find(h=>h.includes("amount")||h.includes("jumlah")||h.includes("mutasi")||h.includes("nominal"));
    const debitKey=headers.find(h=>h.includes("debit")||h==="db");
    const creditKey=headers.find(h=>h.includes("credit")||h.includes("kredit")||h==="cr");
    const date=parseCSVDate(obj[dateKey]||"");if(!date)return [];
    let amount=0;
    if(debitKey&&creditKey){const dv=parseAmount(obj[debitKey]);const cv=parseAmount(obj[creditKey]);amount=cv>0?cv:-dv;}
    else if(amtKey){amount=parseAmount(obj[amtKey]);}
    if(amount===0)return [];
    return [{id:crypto.randomUUID(),date,description:obj[descKey]||"",amount,type:amount>=0?"in":"out",source,category:"Uncategorized"}];
  });
};
const parseCSVText=(text,source)=>{
  if((source==="BCA"||source==="BCA Credit Card")&&/Mutasi Rekening|Kode Mata Uang|TRSF E-BANKING|BI-FAST/i.test(text))return parseBCACorporateCSV(text,source);
  if(source==="BCA"||source==="BCA Credit Card")return parseBCACSV(text,source);
  return parseGenericCSV(text,source);
};
const parseXLSXBuffer=(buffer,source)=>{
  const wb=XLSX.read(buffer,{type:"array",cellDates:true});
  const sheet=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(sheet,{header:1,raw:false,dateNF:"YYYY-MM-DD"});
  if(rows.length<2)return [];
  let headerIdx=0;
  for(let i=0;i<Math.min(10,rows.length);i++){if(rows[i].filter(Boolean).length>=3){headerIdx=i;break;}}
  const headers=rows[headerIdx].map(h=>String(h||"").trim().toLowerCase());
  const dateKey=headers.findIndex(h=>h.includes("date")||h.includes("tanggal")||h.includes("tgl"));
  const descKey=headers.findIndex(h=>h.includes("desc")||h.includes("keterangan")||h.includes("narasi")||h.includes("remark"));
  const amtKey=headers.findIndex(h=>h.includes("amount")||h.includes("jumlah")||h.includes("mutasi")||h.includes("nominal"));
  const debitKey=headers.findIndex(h=>h.includes("debit")||h==="db");
  const creditKey=headers.findIndex(h=>h.includes("credit")||h.includes("kredit")||h==="cr");
  return rows.slice(headerIdx+1).flatMap(row=>{
    if(!row||row.filter(Boolean).length<2)return [];
    const date=parseCSVDate(dateKey>=0?row[dateKey]:row[0]);if(!date)return [];
    const desc=String(descKey>=0?(row[descKey]||""):(row[1]||"")).trim();
    let amount=0;
    if(debitKey>=0&&creditKey>=0){const dv=parseAmount(row[debitKey]);const cv=parseAmount(row[creditKey]);amount=cv>0?cv:-dv;}
    else if(amtKey>=0){amount=parseAmount(row[amtKey]);}
    else{for(let i=1;i<row.length;i++){const v=parseAmount(row[i]);if(v!==0){amount=v;break;}}}
    if(amount===0)return [];
    return [{id:crypto.randomUUID(),date,description:desc,amount,type:amount>=0?"in":"out",source,category:"Uncategorized"}];
  });
};
const makeAutoCategory=(categories)=>(desc,type)=>{
  const d=desc.toLowerCase();
  if(d.includes("ipl")||d.includes("pengelolaan"))return categories.includes("IPL Collection")?"IPL Collection":"Uncategorized";
  if(d.includes("cleaning")||d.includes("sabun")||d.includes("deterjen"))return categories.includes("Cleaning Supplies")?"Cleaning Supplies":"Uncategorized";
  if(d.includes("gaji")||d.includes("salary")||d.includes("payroll"))return categories.includes("Salary & Payroll")?"Salary & Payroll":"Uncategorized";
  if(d.includes("listrik")||d.includes("pln")||d.includes("pdam")||d.includes("telkom"))return categories.includes("Utility & Overhead")?"Utility & Overhead":"Uncategorized";
  if(d.includes("shopee")||d.includes("tokopedia")||d.includes("lazada"))return categories.includes("E-commerce Purchase")?"E-commerce Purchase":"Uncategorized";
  if(d.includes("refund")||d.includes("pengembalian")||d.includes("pembayaran"))return type==="in"?(categories.includes("Customer Refund")?"Customer Refund":"Other Income"):"Uncategorized";
  if(d.includes("supplier")||d.includes("vendor"))return categories.includes("Supplier Payment")?"Supplier Payment":"Uncategorized";
  if(d.includes("claude")||d.includes("openai")||d.includes("chatgpt")||d.includes("google one")||d.includes("google workspace")||d.includes("digitalocean")||d.includes("mailgun")||d.includes("managewp")||d.includes("kommo")||d.includes("apple")||d.includes("playstation"))return categories.includes("Software & Subscriptions")?"Software & Subscriptions":"Uncategorized";
  if(d.includes("bunga")||d.includes("biaya")||d.includes("meterai")||d.includes("keterlambatan"))return categories.includes("Utility & Overhead")?"Utility & Overhead":"Uncategorized";
  return type==="in"?(categories.includes("Other Income")?"Other Income":"Uncategorized"):(categories.includes("Other Expense")?"Other Expense":"Uncategorized");
};
const extractViaAI=async(base64,mediaType)=>{
  const response=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},
    body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:2000,messages:[{role:"user",content:[
      {type:"image",source:{type:"base64",media_type:mediaType,data:base64}},
      {type:"text",text:`Extract ALL financial transactions from this e-wallet screenshot. Return ONLY a valid JSON array, no markdown. Each object: {"date":"YYYY-MM-DD","description":"string","amount":number,"type":"in or out"}. Amount always positive. If year missing assume ${new Date().getFullYear()}. If none return [].`}
    ]}]})});
  const data=await response.json();
  return JSON.parse((data.content||[]).map(b=>b.text||"").join("").replace(/```json|```/g,"").trim());
};

// ── Password Gate ─────────────────────────────────────────────────────────────
function PasswordGate({onUnlock}){
  const [input,setInput]=useState("");
  const [error,setError]=useState(false);
  const [shake,setShake]=useState(false);
  const submit=()=>{
    if(input===APP_PASSWORD){onUnlock();}
    else{setError(true);setShake(true);setInput("");setTimeout(()=>setShake(false),600);}
  };
  return(
    <div style={{fontFamily:"'EB Garamond',Georgia,serif",background:"#faf8f4",minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=EB+Garamond:ital,wght@0,400;0,600;0,700;1,400&family=UnifrakturMaguntia&display=swap');*{box-sizing:border-box;margin:0;padding:0;}@keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-8px)}40%,80%{transform:translateX(8px)}}.shake{animation:shake 0.5s ease;}`}</style>
      <div style={{textAlign:"center",marginBottom:40}}>
        <div style={{fontFamily:"'UnifrakturMaguntia',cursive",fontSize:52,lineHeight:1,color:"#1a1a1a",marginBottom:8}}>The Finance Ledger</div>
        <div style={{fontSize:11,letterSpacing:"0.18em",textTransform:"uppercase",color:"#aaa"}}>Restricted Access</div>
      </div>
      <div style={{borderTop:"2px solid #1a1a1a",borderBottom:"1px solid #1a1a1a",padding:"32px 48px",width:340,textAlign:"center"}} className={shake?"shake":""}>
        <div style={{fontSize:11,letterSpacing:"0.14em",textTransform:"uppercase",color:"#777",marginBottom:20}}>Enter Password to Continue</div>
        <input type="password" value={input} onChange={e=>{setInput(e.target.value);setError(false);}} onKeyDown={e=>e.key==="Enter"&&submit()} placeholder="········" autoFocus
          style={{width:"100%",fontFamily:"'EB Garamond',Georgia,serif",background:"#faf8f4",border:"none",borderBottom:`1px solid ${error?"#4a1a1a":"#1a1a1a"}`,padding:"10px 4px",fontSize:18,outline:"none",textAlign:"center",letterSpacing:"0.2em",color:error?"#4a1a1a":"#1a1a1a",marginBottom:8}}/>
        {error&&<div style={{fontSize:12,color:"#4a1a1a",fontStyle:"italic",marginBottom:12}}>Incorrect password. Try again.</div>}
        {!error&&<div style={{marginBottom:12}}/>}
        <button onClick={submit} style={{width:"100%",fontFamily:"'EB Garamond',Georgia,serif",background:"#1a1a1a",color:"#faf8f4",border:"none",padding:"12px",fontSize:13,letterSpacing:"0.1em",textTransform:"uppercase",cursor:"pointer"}}>Enter</button>
      </div>
      <div style={{marginTop:32,fontSize:11,color:"#ccc",letterSpacing:"0.1em",textTransform:"uppercase"}}>The Finance Ledger &nbsp;·&nbsp; Phase I</div>
    </div>
  );
}

// ── Entity Picker ─────────────────────────────────────────────────────────────
function EntityPicker({entities,onSelect}){
  return(
    <div style={{fontFamily:"'EB Garamond',Georgia,serif",background:"#faf8f4",minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=EB+Garamond:ital,wght@0,400;0,600;0,700;1,400&family=UnifrakturMaguntia&display=swap');*{box-sizing:border-box;margin:0;padding:0;}.entity-card{border:1px solid #ccc;padding:28px 40px;cursor:pointer;transition:all 0.2s;text-align:center;background:#faf8f4;min-width:260px;}.entity-card:hover{border-color:#1a1a1a;background:#f0ede6;}`}</style>
      <div style={{textAlign:"center",marginBottom:48}}>
        <div style={{fontFamily:"'UnifrakturMaguntia',cursive",fontSize:52,lineHeight:1,color:"#1a1a1a",marginBottom:8}}>The Finance Ledger</div>
        <div style={{fontSize:11,letterSpacing:"0.18em",textTransform:"uppercase",color:"#aaa"}}>Select Entity to Continue</div>
      </div>
      <div style={{borderTop:"2px solid #1a1a1a",width:"100%",maxWidth:620,marginBottom:32}}/>
      <div style={{display:"flex",gap:24,flexWrap:"wrap",justifyContent:"center"}}>
        {entities.map(e=>(
          <div key={e.name} className="entity-card" onClick={()=>onSelect(e)}>
            <div style={{width:12,height:12,borderRadius:"50%",background:e.color,margin:"0 auto 14px"}}/>
            <div style={{fontSize:11,letterSpacing:"0.12em",textTransform:"uppercase",color:"#aaa",marginBottom:8}}>Entity</div>
            <div style={{fontSize:20,fontWeight:600,lineHeight:1.3}}>{e.name}</div>
            <div style={{fontSize:13,color:"#999",marginTop:6,fontStyle:"italic"}}>{e.short_name}</div>
          </div>
        ))}
      </div>
      <div style={{borderBottom:"1px solid #1a1a1a",width:"100%",maxWidth:620,marginTop:32}}/>
      <div style={{marginTop:24,fontSize:11,color:"#ccc",letterSpacing:"0.1em",textTransform:"uppercase"}}>The Finance Ledger &nbsp;·&nbsp; Phase I</div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
export default function FinanceHub(){
  const [unlocked,setUnlocked]=useState(()=>sessionStorage.getItem("kf_auth")==="1");
  const [entities,setEntities]=useState([]);
  const [currentEntity,setCurrentEntity]=useState(()=>{
    const saved=sessionStorage.getItem("kf_entity");
    return saved?JSON.parse(saved):null;
  });
  const [transactions,setTransactions]=useState([]);
  const [sources,setSources]=useState([]);
  const [categories,setCategories]=useState([]);
  const [loading,setLoading]=useState(true);
  const [activeTab,setActiveTab]=useState("dashboard");
  const [uploadSource,setUploadSource]=useState("BCA Credit Card");
  const [uploading,setUploading]=useState(false);
  const [uploadingPDF,setUploadingPDF]=useState(false);
  const [extracting,setExtractingAI]=useState(false);
  const [filterSource,setFilterSource]=useState("All");
  const [filterType,setFilterType]=useState("All");
  const [filterCategory,setFilterCategory]=useState("All");
  const [searchQuery,setSearchQuery]=useState("");
  const [selected,setSelected]=useState(new Set());
  const [bulkCat,setBulkCat]=useState("");
  const [manualForm,setManualForm]=useState({date:"",description:"",amount:"",type:"out",source:"BCA Credit Card",category:"Uncategorized"});
  const [toast,setToast]=useState(null);
  const [newSourceName,setNewSourceName]=useState("");
  const [newSourceImportType,setNewSourceImportType]=useState("file");
  const [newSourceColor,setNewSourceColor]=useState("#1a3a5c");
  const [newCatName,setNewCatName]=useState("");
  const fileRef=useRef();const imgRef=useRef();const pdfRef=useRef();

  const showToast=(msg,type="success")=>{setToast({msg,type});setTimeout(()=>setToast(null),3800);};
  const autoCategory=makeAutoCategory(categories);

  // Load entities on mount
  useEffect(()=>{
    sb.from("entities").select("*").order("sort_order").then(({data})=>{
      if(data&&data.length)setEntities(data);
    });
  },[]);

  // Load data when entity changes
  useEffect(()=>{
    if(!currentEntity)return;
    const init=async()=>{
      setLoading(true);
      setTransactions([]);
      try{
        let {data:srcData}=await sb.from("sources").select("*").order("sort_order");
        if(!srcData||!srcData.length){
          const {data:inserted}=await sb.from("sources").insert(DEFAULT_SOURCES.map((s,i)=>({...s,sort_order:i}))).select();
          srcData=inserted||DEFAULT_SOURCES;
        }
        setSources(srcData||[]);
        let {data:catData}=await sb.from("categories").select("*").order("sort_order");
        if(!catData||!catData.length){
          const {data:inserted}=await sb.from("categories").insert(DEFAULT_CATEGORIES.map((name,i)=>({name,sort_order:i}))).select();
          catData=inserted||DEFAULT_CATEGORIES.map(name=>({name}));
        }
        setCategories((catData||[]).map(c=>c.name));
        const {data:txData}=await sb.from("transactions").select("*").eq("entity",currentEntity.name).order("date",{ascending:false});
        setTransactions(txData||[]);
      }catch(e){showToast("Could not connect to database.","error");}
      setLoading(false);
    };
    init();
  },[currentEntity]);

  const saveTransactions=async(rows)=>{
    const catList=categories.length>0?categories:DEFAULT_CATEGORIES;
    const ac=makeAutoCategory(catList);
    const toInsert=rows.map(r=>({
      id:r.id||crypto.randomUUID(),date:r.date,description:r.description,
      amount:r.amount,type:r.type,source:r.source,entity:currentEntity.name,
      category:r.category==="Uncategorized"?ac(r.description,r.type):r.category,
    }));
    const {data,error}=await sb.from("transactions").upsert(toInsert).select();
    if(error){showToast("DB error: "+error.message,"error");console.error("Supabase error:",error);return;}
    setTransactions(prev=>{
      const ids=new Set((data||[]).map(t=>t.id));
      return [...prev.filter(t=>!ids.has(t.id)),...(data||[])].sort((a,b)=>b.date.localeCompare(a.date));
    });
    showToast(`Saved ${toInsert.length} transaction${toInsert.length===1?"":"s"}.`);
  };

  const updateCategoryDB=async(id,cat)=>{
    await sb.from("transactions").update({category:cat}).eq("id",id);
    setTransactions(prev=>prev.map(t=>t.id===id?{...t,category:cat}:t));
  };
  const deleteTransactionDB=async(id)=>{
    await sb.from("transactions").delete().eq("id",id);
    setTransactions(prev=>prev.filter(t=>t.id!==id));
  };
  const applyBulkCatDB=async()=>{
    if(!bulkCat){showToast("Select a category first.","error");return;}
    const ids=[...selected];
    await sb.from("transactions").update({category:bulkCat}).in("id",ids);
    setTransactions(prev=>prev.map(t=>selected.has(t.id)?{...t,category:bulkCat}:t));
    showToast(`Updated ${ids.length} entr${ids.length===1?"y":"ies"} to "${bulkCat}".`);
    setSelected(new Set());setBulkCat("");
  };
  const addSourceDB=async()=>{
    const name=newSourceName.trim();
    if(!name){showToast("Enter a source name.","error");return;}
    if(sources.find(s=>s.name===name)){showToast("Source already exists.","error");return;}
    const {data}=await sb.from("sources").insert({name,color:newSourceColor,import_type:newSourceImportType,sort_order:sources.length}).select();
    if(data)setSources(prev=>[...prev,...data]);
    setNewSourceName("");showToast(`Source "${name}" added.`);
  };
  const removeSourceDB=async(name)=>{
    if(["BCA","Manual"].includes(name)){showToast("Cannot remove default sources.","error");return;}
    await sb.from("sources").delete().eq("name",name);
    setSources(prev=>prev.filter(s=>s.name!==name));showToast(`Source "${name}" removed.`);
  };
  const addCategoryDB=async()=>{
    const name=newCatName.trim();
    if(!name){showToast("Enter a category name.","error");return;}
    if(categories.includes(name)){showToast("Category already exists.","error");return;}
    const {data}=await sb.from("categories").insert({name,sort_order:categories.length}).select();
    if(data)setCategories(prev=>[...prev.filter(c=>c!=="Uncategorized"),name,"Uncategorized"]);
    setNewCatName("");showToast(`Category "${name}" added.`);
  };
  const removeCategoryDB=async(name)=>{
    if(name==="Uncategorized"){showToast("Cannot remove Uncategorized.","error");return;}
    await sb.from("categories").delete().eq("name",name);
    await sb.from("transactions").update({category:"Uncategorized"}).eq("category",name).eq("entity",currentEntity.name);
    setCategories(prev=>prev.filter(c=>c!==name));
    setTransactions(prev=>prev.map(t=>t.category===name?{...t,category:"Uncategorized"}:t));
    showToast(`"${name}" removed.`);
  };

  const handleFileUpload=async(e)=>{
    const file=e.target.files[0];if(!file)return;
    const ext=file.name.split(".").pop().toLowerCase();
    setUploading(true);
    try{
      let rows=[];
      if(ext==="csv"||ext==="txt"){rows=parseCSVText(await file.text(),uploadSource);}
      else if(["xls","xlsx","xlsm"].includes(ext)){rows=parseXLSXBuffer(new Uint8Array(await file.arrayBuffer()),uploadSource);}
      else{showToast("Use CSV, XLS, or XLSX.","error");setUploading(false);return;}
      if(!rows.length)showToast("No transactions found in file.","error");
      else await saveTransactions(rows);
    }catch(err){showToast("Parse error: "+err.message,"error");console.error("CSV parse error:",err);}
    setUploading(false);e.target.value="";
  };
  const handlePDFUpload=async(e)=>{
    const file=e.target.files[0];if(!file)return;
    setUploadingPDF(true);
    try{
      const rows=await parsePDFBuffer(await file.arrayBuffer(),uploadSource);
      if(!rows.length)showToast("No transactions found in PDF.","error");
      else await saveTransactions(rows);
    }catch{showToast("PDF parsing failed. Try CSV or XLS instead.","error");}
    setUploadingPDF(false);e.target.value="";
  };
  const handleScreenshotUpload=async(e)=>{
    const file=e.target.files[0];if(!file)return;
    setExtractingAI(true);
    try{
      const base64=await new Promise((res,rej)=>{const r=new FileReader();r.onload=()=>res(r.result.split(",")[1]);r.onerror=rej;r.readAsDataURL(file);});
      const parsed=await extractViaAI(base64,file.type||"image/jpeg");
      const rows=parsed.map(r=>({id:crypto.randomUUID(),date:r.date,description:r.description,amount:r.type==="out"?-Math.abs(r.amount):Math.abs(r.amount),type:r.type,source:uploadSource,category:"Uncategorized"}));
      if(!rows.length)showToast("No transactions found.","error");
      else await saveTransactions(rows);
    }catch{showToast("Extraction failed. Try manual entry.","error");}
    setExtractingAI(false);e.target.value="";
  };
  const handleManualAdd=async()=>{
    if(!manualForm.date||!manualForm.description||!manualForm.amount){showToast("All fields required.","error");return;}
    const amt=parseFloat(manualForm.amount);
    await saveTransactions([{id:crypto.randomUUID(),date:manualForm.date,description:manualForm.description,
      amount:manualForm.type==="out"?-Math.abs(amt):Math.abs(amt),type:manualForm.type,source:manualForm.source,category:manualForm.category}]);
    setManualForm(f=>({...f,date:"",description:"",amount:""}));
  };
  const exportAccurate=()=>{
    if(!transactions.length){showToast("No transactions to export.","error");return;}
    const header="Tanggal,Keterangan,Debit,Kredit,Kategori,Sumber,Entitas";
    const rows=transactions.map(t=>`${t.date},"${t.description}",${t.type==="out"?Math.abs(t.amount):0},${t.type==="in"?Math.abs(t.amount):0},${t.category},${t.source},${t.entity}`);
    const blob=new Blob([[header,...rows].join("\n")],{type:"text/csv"});
    const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`ledger-${currentEntity.short_name}-${new Date().toISOString().split("T")[0]}.csv`;a.click();
    showToast("Ledger exported for Accurate.id.");
  };

  const toggleSelect=(id)=>setSelected(prev=>{const s=new Set(prev);s.has(id)?s.delete(id):s.add(id);return s;});
  const toggleAll=(ids)=>setSelected(prev=>prev.size===ids.length?new Set():new Set(ids));

  const filtered=transactions.filter(t=>{
    if(filterSource!=="All"&&t.source!==filterSource)return false;
    if(filterType!=="All"&&t.type!==filterType)return false;
    if(filterCategory!=="All"&&t.category!==filterCategory)return false;
    if(searchQuery&&!t.description.toLowerCase().includes(searchQuery.toLowerCase()))return false;
    return true;
  });
  const filteredIds=filtered.map(t=>t.id);
  // Revenue = only transactions in categories containing "Revenue"
  const totalIn=transactions.filter(t=>t.category.toLowerCase().includes("revenue")&&!EXCLUDED_CATS.includes(t.category)).reduce((s,t)=>s+Math.abs(t.amount),0);
  // Expenditure excludes inter-account transfers (not real expenses)
  const EXCLUDED_CATS=["Inter-account Transfer","Client Deposit","Client Disbursement"];
  const totalOut=transactions.filter(t=>t.type==="out"&&!EXCLUDED_CATS.includes(t.category)).reduce((s,t)=>s+Math.abs(t.amount),0);
  // Net cash = all money in minus all money out INCLUDING transfers (real cash movement)
  const allIn=transactions.filter(t=>t.type==="in").reduce((s,t)=>s+Math.abs(t.amount),0);
  const allOut=transactions.filter(t=>t.type==="out").reduce((s,t)=>s+Math.abs(t.amount),0);
  const netCash=allIn-allOut;
  const sourceMap=Object.fromEntries(sources.map(s=>[s.name,s]));
  const bySource=sources.filter(s=>s.name!=="Manual").map(s=>({...s,
    in:transactions.filter(t=>t.source===s.name&&t.type==="in").reduce((sum,t)=>sum+Math.abs(t.amount),0),
    out:transactions.filter(t=>t.source===s.name&&t.type==="out").reduce((sum,t)=>sum+Math.abs(t.amount),0),
    count:transactions.filter(t=>t.source===s.name).length
  })).filter(s=>s.count>0);
  const byCategory=categories.map(c=>({category:c,
    total:transactions.filter(t=>t.category===c).reduce((s,t)=>s+Math.abs(t.amount),0),
    count:transactions.filter(t=>t.category===c).length
  })).filter(c=>c.count>0).sort((a,b)=>b.total-a.total);
  const isScreenshotOnly=sources.find(s=>s.name===uploadSource)?.import_type==="screenshot";

  const inputStyle={fontFamily:"'EB Garamond',Georgia,serif",background:"#faf8f4",border:"none",borderBottom:"1px solid #1a1a1a",color:"#1a1a1a",padding:"8px 4px",fontSize:15,outline:"none",width:"100%",borderRadius:0};
  const selectStyle={...inputStyle,cursor:"pointer"};

  // ── Gates ────────────────────────────────────────────────────────────────
  if(!unlocked)return <PasswordGate onUnlock={()=>{sessionStorage.setItem("kf_auth","1");setUnlocked(true);}}/>;
  if(!currentEntity)return <EntityPicker entities={entities.length?entities:[{name:"PT Hidup Lebih Tentram",short_name:"Tentram",color:"#1a3a5c"},{name:"PT Semangat Solusi Digital",short_name:"Solusi Digital",color:"#7a1c1c"}]} onSelect={e=>{sessionStorage.setItem("kf_entity",JSON.stringify(e));setCurrentEntity(e);}}/>;
  if(loading)return(
    <div style={{fontFamily:"'EB Garamond',Georgia,serif",background:"#faf8f4",minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16}}>
      <div style={{fontFamily:"'UnifrakturMaguntia',cursive",fontSize:48,color:"#1a1a1a"}}>The Finance Ledger</div>
      <div style={{fontSize:14,color:"#aaa",fontStyle:"italic",letterSpacing:"0.1em"}}>Loading {currentEntity.name}…</div>
    </div>
  );

  return(
    <div style={{fontFamily:"'EB Garamond',Georgia,serif",background:"#faf8f4",minHeight:"100vh",color:"#1a1a1a"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=EB+Garamond:ital,wght@0,400;0,500;0,600;0,700;0,800;1,400;1,500&family=UnifrakturMaguntia&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        ::-webkit-scrollbar{width:3px;} ::-webkit-scrollbar-thumb{background:#ccc;}
        button{cursor:pointer;font-family:'EB Garamond',Georgia,serif;}
        select,input{font-family:'EB Garamond',Georgia,serif;}
        select option{background:#faf8f4;color:#1a1a1a;}
        .nav-link{background:none;border:none;font-size:13px;letter-spacing:0.06em;text-transform:uppercase;padding:6px 14px;transition:all 0.15s;color:#1a1a1a;}
        .nav-link:hover{background:#1a1a1a;color:#faf8f4;}
        .nav-link.active{background:#1a1a1a;color:#faf8f4;}
        .src-tag{display:inline-block;font-size:10px;letter-spacing:0.08em;padding:2px 7px;border:1px solid currentColor;text-transform:uppercase;}
        .tr-row:hover td{background:#ede9e0;}
        .upload-zone{border:1px solid #aaa;padding:28px 24px;text-align:center;cursor:pointer;transition:all 0.2s;}
        .upload-zone:hover{border-color:#1a1a1a;background:#f0ede6;}
        .upload-zone.busy{opacity:0.6;pointer-events:none;border-color:#1a1a1a;}
        .action-btn{border:1px solid #1a1a1a;background:none;color:#1a1a1a;padding:9px 20px;font-size:13px;letter-spacing:0.06em;text-transform:uppercase;transition:all 0.2s;}
        .action-btn:hover{background:#1a1a1a;color:#faf8f4;}
        .action-btn.primary{background:#1a1a1a;color:#faf8f4;}
        .action-btn.primary:hover{background:#333;}
        .src-btn{border:1px solid #ccc;background:none;color:#666;padding:6px 14px;font-size:12px;letter-spacing:0.06em;text-transform:uppercase;transition:all 0.2s;margin-right:6px;margin-bottom:6px;}
        .src-btn:hover,.src-btn.active{border-color:#1a1a1a;color:#1a1a1a;background:#f0ede6;}
        .field-label{font-size:11px;letter-spacing:0.12em;text-transform:uppercase;color:#666;display:block;margin-bottom:6px;}
        .rm-btn{background:none;border:none;color:#ccc;font-size:16px;padding:0 4px;transition:color 0.15s;line-height:1;}
        .rm-btn:hover{color:#4a1a1a;}
        .tag-row{display:flex;align-items:center;justify-content:space-between;padding:9px 0;border-bottom:1px solid #e8e4dc;}
        .chk{width:15px;height:15px;cursor:pointer;accent-color:#1a1a1a;}
        tr.sel td{background:#f0ede6 !important;}
        .fmt-badge{display:inline-block;font-size:10px;letter-spacing:0.08em;padding:1px 6px;border:1px solid #ccc;text-transform:uppercase;color:#888;}
        .switch-btn{background:none;border:1px solid #ccc;font-size:12px;letter-spacing:0.06em;text-transform:uppercase;padding:4px 12px;color:#888;transition:all 0.15s;}
        .switch-btn:hover{border-color:#1a1a1a;color:#1a1a1a;}
      `}</style>

      {/* Masthead */}
      <div style={{borderBottom:"3px double #1a1a1a",padding:"16px 40px 0"}}>
        <div style={{textAlign:"center",borderBottom:"1px solid #1a1a1a",paddingBottom:12,marginBottom:10}}>
          <div style={{fontSize:11,letterSpacing:"0.2em",textTransform:"uppercase",color:"#555",marginBottom:8}}>{todayStr()}</div>
          <div style={{fontFamily:"'UnifrakturMaguntia',cursive",fontSize:56,lineHeight:1,color:"#1a1a1a"}}>The Finance Ledger</div>
          {/* Entity badge */}
          <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:10,marginTop:8}}>
            <div style={{width:8,height:8,borderRadius:"50%",background:currentEntity.color}}/>
            <div style={{fontSize:12,letterSpacing:"0.12em",textTransform:"uppercase",color:"#555"}}>{currentEntity.name}</div>
            <button className="switch-btn" onClick={()=>{sessionStorage.removeItem("kf_entity");setCurrentEntity(null);setTransactions([]);}}>Switch Entity</button>
          </div>
        </div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingBottom:10}}>
          <div style={{display:"flex",gap:0}}>
            {[["dashboard","Overview"],["ledger","Ledger"],["upload","Import"],["manual","Add Entry"],["settings","Settings"]].map(([id,label])=>(
              <button key={id} className={`nav-link${activeTab===id?" active":""}`} onClick={()=>setActiveTab(id)}>{label}</button>
            ))}
          </div>
          <button className="action-btn" onClick={exportAccurate} style={{fontSize:12}}>↓ Export to Accurate.id</button>
        </div>
      </div>

      <div style={{padding:"28px 40px"}}>

        {/* DASHBOARD */}
        {activeTab==="dashboard"&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1px 1fr 1px 1fr 1px 1fr",borderTop:"2px solid #1a1a1a",borderBottom:"1px solid #1a1a1a",padding:"18px 0",marginBottom:28}}>
              {[
                {label:"Total Revenue",value:fmt(totalIn),sub:`${transactions.filter(t=>t.category.toLowerCase().includes("revenue")).length} entries`,pos:true},
                null,
                {label:"Total Expenditure",value:fmt(totalOut),sub:`${transactions.filter(t=>t.type==="out"&&!EXCLUDED_CATS.includes(t.category)).length} payments`,pos:false},
                null,
                {label:"Net Cash Position",value:fmt(netCash),sub:netCash>=0?"Surplus":"Deficit",pos:netCash>=0},
                null,
                {label:"Entries on Record",value:transactions.length,sub:`across ${bySource.length} source${bySource.length!==1?"s":""}`,pos:null},
              ].map((k,i)=>k===null?<div key={i} style={{borderLeft:"1px solid #ccc"}}/>:(
                <div key={k.label} style={{padding:"0 24px",textAlign:"center"}}>
                  <div style={{fontSize:11,letterSpacing:"0.12em",textTransform:"uppercase",color:"#777",marginBottom:10}}>{k.label}</div>
                  <div style={{fontSize:30,fontWeight:700,fontStyle:"italic",lineHeight:1,color:k.pos===true?"#1a4a1a":k.pos===false?"#4a1a1a":"#1a1a1a"}}>{k.value}</div>
                  <div style={{fontSize:12,color:"#999",marginTop:6}}>{k.sub}</div>
                </div>
              ))}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1px 1fr"}}>
              <div style={{paddingRight:36}}>
                <div style={{borderTop:"2px solid #1a1a1a",borderBottom:"1px solid #1a1a1a",padding:"6px 0",marginBottom:20,textAlign:"center"}}>
                  <span style={{fontSize:11,letterSpacing:"0.14em",textTransform:"uppercase"}}>Cash Flow by Source</span>
                </div>
                {bySource.length===0?(
                  <div style={{textAlign:"center",color:"#bbb",fontSize:15,fontStyle:"italic",padding:"44px 0",lineHeight:1.8}}>No data yet.<br/>Begin by importing a statement.</div>
                ):bySource.map(s=>(
                  <div key={s.name} style={{marginBottom:20}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:6}}>
                      <span style={{fontSize:16,fontWeight:600}}>{s.name}</span>
                      <span style={{fontSize:13,color:"#777",fontStyle:"italic"}}>{s.count} entries</span>
                    </div>
                    <div style={{display:"flex",gap:2,marginBottom:5,height:5,background:"#e8e4dc"}}>
                      <div style={{background:"#1a4a1a",width:`${totalIn+totalOut>0?(s.in/(totalIn+totalOut))*100:0}%`}}/>
                      <div style={{background:"#4a1a1a",width:`${totalIn+totalOut>0?(s.out/(totalIn+totalOut))*100:0}%`}}/>
                    </div>
                    <div style={{display:"flex",gap:20,fontSize:13,color:"#666"}}>
                      <span>In: <em>{fmt(s.in)}</em></span><span>Out: <em>{fmt(s.out)}</em></span>
                    </div>
                    <div style={{borderBottom:"1px solid #e8e4dc",marginTop:16}}/>
                  </div>
                ))}
              </div>
              <div style={{borderLeft:"1px solid #ccc"}}/>
              <div style={{paddingLeft:36}}>
                <div style={{borderTop:"2px solid #1a1a1a",borderBottom:"1px solid #1a1a1a",padding:"6px 0",marginBottom:20,textAlign:"center"}}>
                  <span style={{fontSize:11,letterSpacing:"0.14em",textTransform:"uppercase"}}>Expenditure by Category</span>
                </div>
                {byCategory.length===0?(
                  <div style={{textAlign:"center",color:"#bbb",fontSize:15,fontStyle:"italic",padding:"44px 0",lineHeight:1.8}}>Categories appear once<br/>transactions are recorded.</div>
                ):byCategory.slice(0,8).map((c,i)=>{
                  const max=byCategory[0].total;
                  return(
                    <div key={c.category} style={{marginBottom:16}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:5}}>
                        <div style={{display:"flex",alignItems:"baseline",gap:10}}>
                          <span style={{fontSize:12,color:"#bbb",minWidth:18}}>{i+1}.</span>
                          <span style={{fontSize:16}}>{c.category}</span>
                        </div>
                        <span style={{fontSize:14,fontStyle:"italic"}}>{fmt(c.total)}</span>
                      </div>
                      <div style={{height:3,background:"#e8e4dc"}}>
                        <div style={{height:"100%",width:`${(c.total/max)*100}%`,background:"#1a1a1a"}}/>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}

        {/* LEDGER */}
        {activeTab==="ledger"&&(
          <div>
            <div style={{borderTop:"2px solid #1a1a1a",borderBottom:"1px solid #1a1a1a",padding:"6px 0",marginBottom:20,textAlign:"center"}}>
              <span style={{fontSize:11,letterSpacing:"0.14em",textTransform:"uppercase"}}>General Ledger — {currentEntity.name}</span>
            </div>
            <div style={{display:"flex",gap:12,marginBottom:16,alignItems:"flex-end",flexWrap:"wrap",borderBottom:"1px solid #e0dcd4",paddingBottom:14}}>
              <div style={{flex:1,minWidth:160}}><label className="field-label">Search</label><input placeholder="Search entries…" value={searchQuery} onChange={e=>setSearchQuery(e.target.value)} style={inputStyle}/></div>
              <div><label className="field-label">Source</label>
                <select value={filterSource} onChange={e=>setFilterSource(e.target.value)} style={{...selectStyle,width:150}}>
                  <option value="All">All Sources</option>{sources.map(s=><option key={s.name}>{s.name}</option>)}
                </select>
              </div>
              <div><label className="field-label">Type</label>
                <select value={filterType} onChange={e=>setFilterType(e.target.value)} style={{...selectStyle,width:120}}>
                  <option value="All">All Types</option><option value="in">Receipts</option><option value="out">Payments</option>
                </select>
              </div>
              <div><label className="field-label">Category</label>
                <select value={filterCategory} onChange={e=>setFilterCategory(e.target.value)} style={{...selectStyle,width:200}}>
                  <option value="All">All Categories</option>{categories.map(c=><option key={c}>{c}</option>)}
                </select>
              </div>
              <span style={{fontSize:13,color:"#999",fontStyle:"italic",paddingBottom:8}}>{filtered.length} entries</span>
            </div>
            {selected.size>0&&(
              <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:14,padding:"10px 14px",background:"#f0ede6",border:"1px solid #ccc"}}>
                <span style={{fontSize:14,fontStyle:"italic",color:"#555"}}>{selected.size} entr{selected.size===1?"y":"ies"} selected</span>
                <select value={bulkCat} onChange={e=>setBulkCat(e.target.value)} style={{...selectStyle,width:220,fontSize:13,padding:"4px 4px"}}>
                  <option value="">— Assign category —</option>{categories.map(c=><option key={c}>{c}</option>)}
                </select>
                <button className="action-btn primary" onClick={applyBulkCatDB} style={{padding:"6px 16px",fontSize:12}}>Apply</button>
                <button className="action-btn" onClick={()=>setSelected(new Set())} style={{padding:"6px 14px",fontSize:12}}>Clear</button>
              </div>
            )}
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead>
                <tr style={{borderTop:"1px solid #1a1a1a",borderBottom:"2px solid #1a1a1a"}}>
                  <th style={{padding:"8px 10px",width:32}}><input type="checkbox" className="chk" checked={filteredIds.length>0&&filteredIds.every(id=>selected.has(id))} onChange={()=>toggleAll(filteredIds)}/></th>
                  {["Date","Description","Source","Category","","Amount",""].map((h,i)=>(
                    <th key={i} style={{padding:"8px 10px",textAlign:i===5?"right":"left",fontSize:11,letterSpacing:"0.1em",textTransform:"uppercase",fontWeight:600}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.length===0?(
                  <tr><td colSpan={8} style={{padding:"56px 0",textAlign:"center",color:"#bbb",fontStyle:"italic",fontSize:17,lineHeight:1.8}}>The ledger is empty.<br/>Import statements or add entries manually.</td></tr>
                ):filtered.map((t,idx)=>(
                  <tr key={t.id} className={`tr-row${selected.has(t.id)?" sel":""}`} style={{borderBottom:"1px solid #e8e4dc",background:selected.has(t.id)?"#f0ede6":idx%2===0?"transparent":"#f5f2ed"}}>
                    <td style={{padding:"10px 10px"}}><input type="checkbox" className="chk" checked={selected.has(t.id)} onChange={()=>toggleSelect(t.id)}/></td>
                    <td style={{padding:"10px 10px",fontSize:13,color:"#666",whiteSpace:"nowrap"}}>{t.date}</td>
                    <td style={{padding:"10px 10px",fontSize:15,maxWidth:360,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",cursor:"default"}} title={t.description}>{t.description}</td>
                    <td style={{padding:"10px 10px"}}><span className="src-tag" style={{color:sourceMap[t.source]?.color||"#555",borderColor:sourceMap[t.source]?.color||"#ccc"}}>{t.source}</span></td>
                    <td style={{padding:"10px 10px"}}>
                      <select value={t.category} onChange={e=>updateCategoryDB(t.id,e.target.value)} style={{...selectStyle,fontSize:13,width:"auto",padding:"2px 4px",color:"#555"}}>
                        {categories.map(c=><option key={c}>{c}</option>)}
                      </select>
                    </td>
                    <td style={{padding:"10px 6px"}}><span style={{fontSize:12,letterSpacing:"0.05em",textTransform:"uppercase",color:t.type==="in"?"#1a4a1a":"#4a1a1a",fontStyle:"italic"}}>{t.type==="in"?"CR":"DB"}</span></td>
                    <td style={{padding:"10px 10px",fontSize:15,fontWeight:700,fontStyle:"italic",color:t.type==="in"?"#1a4a1a":"#4a1a1a",textAlign:"right",whiteSpace:"nowrap"}}>
                      {t.type==="out"?"-":"+"}{fmt(Math.abs(t.amount))}
                    </td>
                    <td style={{padding:"10px 6px",textAlign:"center"}}><button className="rm-btn" onClick={()=>deleteTransactionDB(t.id)}>×</button></td>
                  </tr>
                ))}
              </tbody>
              {filtered.length>0&&(
                <tfoot>
                  <tr style={{borderTop:"2px solid #1a1a1a"}}>
                    <td/><td colSpan={5} style={{padding:"10px 10px",fontSize:12,letterSpacing:"0.08em",textTransform:"uppercase",color:"#777",fontStyle:"italic"}}>Net total — {filtered.length} entries</td>
                    <td style={{padding:"10px 10px",textAlign:"right",fontSize:17,fontWeight:700,fontStyle:"italic",color:filtered.reduce((s,t)=>s+(t.type==="in"?Math.abs(t.amount):-Math.abs(t.amount)),0)>=0?"#1a4a1a":"#4a1a1a"}}>
                      {fmt(filtered.reduce((s,t)=>s+(t.type==="in"?Math.abs(t.amount):-Math.abs(t.amount)),0))}
                    </td>
                    <td/>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
        )}

        {/* UPLOAD */}
        {activeTab==="upload"&&(
          <div style={{maxWidth:620}}>
            <div style={{borderTop:"2px solid #1a1a1a",borderBottom:"1px solid #1a1a1a",padding:"6px 0",marginBottom:26,textAlign:"center"}}>
              <span style={{fontSize:11,letterSpacing:"0.14em",textTransform:"uppercase"}}>Import Statements — {currentEntity.short_name}</span>
            </div>
            <div style={{marginBottom:24}}>
              <label className="field-label">Select Payment Source</label>
              <div style={{display:"flex",flexWrap:"wrap",marginTop:4}}>
                {sources.filter(s=>s.name!=="Manual").map(s=>(
                  <button key={s.name} className={`src-btn${uploadSource===s.name?" active":""}`} onClick={()=>setUploadSource(s.name)}>{s.name}</button>
                ))}
              </div>
            </div>
            <div style={{borderTop:"1px solid #1a1a1a",marginBottom:28}}/>
            {!isScreenshotOnly&&(<>
              <div style={{marginBottom:28}}>
                <div style={{display:"flex",alignItems:"baseline",gap:10,marginBottom:4}}>
                  <div style={{fontSize:19,fontWeight:600}}>Upload Statement File</div>
                  <span className="fmt-badge">CSV</span><span className="fmt-badge">XLS</span><span className="fmt-badge">XLSX</span>
                </div>
                <div style={{fontSize:14,color:"#777",marginBottom:16,fontStyle:"italic",lineHeight:1.7}}>Export your {uploadSource} statement. Columns detected automatically.</div>
                <div className={`upload-zone${uploading?" busy":""}`} onClick={()=>fileRef.current?.click()}>
                  <div style={{fontSize:32,marginBottom:8,color:"#ccc"}}>☰</div>
                  <div style={{fontSize:16}}>{uploading?"Parsing file…":"Click to select file"}</div>
                  <div style={{fontSize:13,color:"#aaa",marginTop:4,fontStyle:"italic"}}>CSV, XLS, XLSX accepted</div>
                </div>
                <input ref={fileRef} type="file" accept=".csv,.txt,.xls,.xlsx,.xlsm" style={{display:"none"}} onChange={handleFileUpload}/>
              </div>
              <div style={{marginBottom:28}}>
                <div style={{display:"flex",alignItems:"baseline",gap:10,marginBottom:4}}>
                  <div style={{fontSize:19,fontWeight:600}}>Upload PDF Statement</div>
                  <span className="fmt-badge">PDF</span>
                  <span style={{fontSize:11,letterSpacing:"0.08em",textTransform:"uppercase",border:"1px solid #aaa",padding:"2px 8px",color:"#777"}}>Client-side</span>
                </div>
                <div style={{fontSize:14,color:"#777",marginBottom:16,fontStyle:"italic",lineHeight:1.7}}>Digital PDF bank statements. Text read directly in browser. Optimised for BCA Credit Card.</div>
                <div className={`upload-zone${uploadingPDF?" busy":""}`} onClick={()=>pdfRef.current?.click()}>
                  <div style={{fontSize:32,marginBottom:8,color:"#ccc"}}>⊞</div>
                  <div style={{fontSize:16}}>{uploadingPDF?"Reading PDF…":"Click to select PDF"}</div>
                  <div style={{fontSize:13,color:"#aaa",marginTop:4,fontStyle:"italic"}}>Digital PDFs only</div>
                </div>
                <input ref={pdfRef} type="file" accept=".pdf" style={{display:"none"}} onChange={handlePDFUpload}/>
              </div>
            </>)}
            <div>
              <div style={{display:"flex",alignItems:"baseline",gap:10,marginBottom:4}}>
                <div style={{fontSize:19,fontWeight:600}}>Upload Screenshot</div>
                <span style={{fontSize:11,letterSpacing:"0.08em",textTransform:"uppercase",border:"1px solid #aaa",padding:"2px 8px",color:"#777"}}>AI Powered</span>
              </div>
              <div style={{fontSize:14,color:"#777",marginBottom:16,fontStyle:"italic",lineHeight:1.7}}>
                {isScreenshotOnly?`${uploadSource} has no file export. Screenshot your transaction history and AI will extract each entry.`:"For e-wallets without file export, or scanned PDFs, upload a screenshot."}
              </div>
              <div className={`upload-zone${extracting?" busy":""}`} onClick={()=>imgRef.current?.click()}>
                <div style={{fontSize:32,marginBottom:8,color:"#ccc"}}>◈</div>
                <div style={{fontSize:16}}>{extracting?"AI extracting transactions…":"Click to select screenshot"}</div>
                <div style={{fontSize:13,color:"#aaa",marginTop:4,fontStyle:"italic"}}>JPG, PNG — reads GoPay, ShopeePay, Astro, and more</div>
              </div>
              <input ref={imgRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleScreenshotUpload}/>
            </div>
          </div>
        )}

        {/* MANUAL */}
        {activeTab==="manual"&&(
          <div style={{maxWidth:520}}>
            <div style={{borderTop:"2px solid #1a1a1a",borderBottom:"1px solid #1a1a1a",padding:"6px 0",marginBottom:30,textAlign:"center"}}>
              <span style={{fontSize:11,letterSpacing:"0.14em",textTransform:"uppercase"}}>Record Entry — {currentEntity.short_name}</span>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:28,marginBottom:24}}>
              <div><label className="field-label">Date</label><input type="date" value={manualForm.date} onChange={e=>setManualForm(f=>({...f,date:e.target.value}))} style={inputStyle}/></div>
              <div><label className="field-label">Type</label>
                <select value={manualForm.type} onChange={e=>setManualForm(f=>({...f,type:e.target.value}))} style={selectStyle}>
                  <option value="in">Receipt (Payment In)</option><option value="out">Payment (Payment Out)</option>
                </select>
              </div>
            </div>
            <div style={{marginBottom:24}}><label className="field-label">Description</label><input placeholder="e.g. IPL Unit 12A — May 2026" value={manualForm.description} onChange={e=>setManualForm(f=>({...f,description:e.target.value}))} style={inputStyle}/></div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:28,marginBottom:24}}>
              <div><label className="field-label">Amount (IDR)</label><input type="number" placeholder="2500000" value={manualForm.amount} onChange={e=>setManualForm(f=>({...f,amount:e.target.value}))} style={inputStyle}/></div>
              <div><label className="field-label">Source</label>
                <select value={manualForm.source} onChange={e=>setManualForm(f=>({...f,source:e.target.value}))} style={selectStyle}>
                  {sources.map(s=><option key={s.name}>{s.name}</option>)}
                </select>
              </div>
            </div>
            <div style={{marginBottom:32}}><label className="field-label">Category</label>
              <select value={manualForm.category} onChange={e=>setManualForm(f=>({...f,category:e.target.value}))} style={selectStyle}>
                {categories.map(c=><option key={c}>{c}</option>)}
              </select>
            </div>
            <div style={{borderTop:"2px solid #1a1a1a",marginBottom:24}}/>
            <button className="action-btn primary" onClick={handleManualAdd} style={{width:"100%",padding:"14px",fontSize:14,letterSpacing:"0.1em"}}>Record Entry</button>
          </div>
        )}

        {/* SETTINGS */}
        {activeTab==="settings"&&(
          <div style={{display:"grid",gridTemplateColumns:"1fr 1px 1fr",gap:0}}>
            <div style={{paddingRight:40}}>
              <div style={{borderTop:"2px solid #1a1a1a",borderBottom:"1px solid #1a1a1a",padding:"6px 0",marginBottom:24,textAlign:"center"}}>
                <span style={{fontSize:11,letterSpacing:"0.14em",textTransform:"uppercase"}}>Payment Sources</span>
              </div>
              {sources.map(s=>(
                <div key={s.name} className="tag-row">
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <div style={{width:10,height:10,borderRadius:"50%",background:s.color,flexShrink:0}}/>
                    <span style={{fontSize:16}}>{s.name}</span>
                    <span className="fmt-badge">{s.import_type==="screenshot"?"screenshot":"file"}</span>
                  </div>
                  <button className="rm-btn" onClick={()=>removeSourceDB(s.name)}>×</button>
                </div>
              ))}
              <div style={{marginTop:24,paddingTop:20,borderTop:"1px dashed #ccc"}}>
                <div style={{fontSize:14,fontWeight:600,marginBottom:14}}>Add New Source</div>
                <div style={{marginBottom:12}}><label className="field-label">Source Name</label>
                  <input placeholder="e.g. Dana, OVO, Jenius…" value={newSourceName} onChange={e=>setNewSourceName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addSourceDB()} style={inputStyle}/>
                </div>
                <div style={{display:"flex",gap:20,marginBottom:14,alignItems:"flex-start"}}>
                  <div style={{flex:1}}><label className="field-label">Colour</label>
                    <div style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:4}}>
                      {INKCOLS.map(c=><div key={c} onClick={()=>setNewSourceColor(c)} style={{width:20,height:20,borderRadius:"50%",background:c,cursor:"pointer",outline:newSourceColor===c?"2px solid #1a1a1a":"2px solid transparent",outlineOffset:2}}/>)}
                    </div>
                  </div>
                  <div><label className="field-label">Import via</label>
                    <div style={{display:"flex",flexDirection:"column",gap:6,marginTop:4,fontSize:14}}>
                      {[["file","File (CSV/XLS/PDF)"],["screenshot","Screenshot only"]].map(([v,l])=>(
                        <label key={v} style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer"}}>
                          <input type="radio" checked={newSourceImportType===v} onChange={()=>setNewSourceImportType(v)} style={{accentColor:"#1a1a1a"}}/>{l}
                        </label>
                      ))}
                    </div>
                  </div>
                </div>
                <button className="action-btn primary" onClick={addSourceDB} style={{width:"100%",padding:"10px",fontSize:13,letterSpacing:"0.08em"}}>+ Add Source</button>
              </div>
            </div>
            <div style={{borderLeft:"1px solid #ccc"}}/>
            <div style={{paddingLeft:40}}>
              <div style={{borderTop:"2px solid #1a1a1a",borderBottom:"1px solid #1a1a1a",padding:"6px 0",marginBottom:24,textAlign:"center"}}>
                <span style={{fontSize:11,letterSpacing:"0.14em",textTransform:"uppercase"}}>Transaction Categories</span>
              </div>
              {categories.map(c=>(
                <div key={c} className="tag-row">
                  <span style={{fontSize:16}}>{c}</span>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontSize:12,color:"#aaa",fontStyle:"italic"}}>{transactions.filter(t=>t.category===c).length} entries</span>
                    {c!=="Uncategorized"&&<button className="rm-btn" onClick={()=>removeCategoryDB(c)}>×</button>}
                  </div>
                </div>
              ))}
              <div style={{marginTop:24,paddingTop:20,borderTop:"1px dashed #ccc"}}>
                <div style={{fontSize:14,fontWeight:600,marginBottom:14}}>Add New Category</div>
                <div style={{marginBottom:12}}><label className="field-label">Category Name</label>
                  <input placeholder="e.g. Transport, Marketing…" value={newCatName} onChange={e=>setNewCatName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addCategoryDB()} style={inputStyle}/>
                </div>
                <button className="action-btn primary" onClick={addCategoryDB} style={{width:"100%",padding:"10px",fontSize:13,letterSpacing:"0.08em"}}>+ Add Category</button>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Footer */}
      <div style={{borderTop:"3px double #1a1a1a",margin:"0 40px",padding:"14px 0",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div style={{fontSize:11,color:"#bbb",letterSpacing:"0.1em",textTransform:"uppercase"}}>
          The Finance Ledger &nbsp;·&nbsp; {currentEntity.name} &nbsp;·&nbsp; Phase I
        </div>
        <button className="switch-btn" onClick={()=>{sessionStorage.removeItem("kf_entity");setCurrentEntity(null);setTransactions([]);}} style={{fontSize:11}}>
          Switch Entity
        </button>
      </div>

      {/* Toast */}
      {toast&&(
        <div style={{position:"fixed",bottom:28,right:28,background:"#faf8f4",border:toast.type==="error"?"2px solid #4a1a1a":"2px solid #1a4a1a",
          color:toast.type==="error"?"#4a1a1a":"#1a4a1a",padding:"13px 22px",fontSize:15,
          fontFamily:"'EB Garamond',Georgia,serif",fontStyle:"italic",boxShadow:"4px 4px 0 #1a1a1a",zIndex:999}}>
          {toast.msg}
        </div>
      )}
    </div>
  );
}
