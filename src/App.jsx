import React, { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import {
  getFirestore,
  collection,
  getDocs,
  setDoc,
  deleteDoc,
  doc,
  writeBatch,
  onSnapshot,
} from "firebase/firestore";

// ── Firebase 設定 ──────────────────────────────────────────
const firebaseConfig = {
  apiKey: "AIzaSyDJzrytA_fhcx-LaoFaRO1116RrehvKh38",
  authDomain: "directory-456ac.firebaseapp.com",
  projectId: "directory-456ac",
  storageBucket: "directory-456ac.firebasestorage.app",
  messagingSenderId: "904480821084",
  appId: "1:904480821084:web:69c578cc14d774deaffb99",
  measurementId: "G-NM2N2326BY",
};
const fbApp = initializeApp(firebaseConfig);
const db = getFirestore(fbApp);
const COL = "directory";
const META_DOC = "meta_info";

// ── Google Fonts ───────────────────────────────────────────
if (typeof document !== "undefined" && !document.getElementById("noto-serif-hk-font")) {
  const link = document.createElement("link");
  link.id = "noto-serif-hk-font";
  link.rel = "stylesheet";
  link.href = "https://fonts.googleapis.com/css2?family=Noto+Serif+HK:wght@900&display=swap";
  document.head.appendChild(link);
}

// ── 常數 ──────────────────────────────────────────────────
const FIELDS = [
  { key: "agent",   label: "代號" },
  { key: "vendor",  label: "廠商" },
  { key: "fullname",label: "全名", required: true },
  { key: "taxid",   label: "統一編號" },
  { key: "address", label: "地址" },
];
const CFIELDS = [
  { key: "contact", label: "聯絡人/負責人" },
  { key: "phone",   label: "電話" },
  { key: "fax",     label: "傳真" },
  { key: "mobile",  label: "手機" },
  { key: "email",   label: "信箱" },
];
const KEY_MAP = {
  "代號":"agent","廠商":"vendor","全名":"fullname","統一編號":"taxid","聯絡人/負責人":"contact",
  "聯絡人":"contact","負責人":"contact","電話":"phone","傳真":"fax","手機":"mobile",
  "地址":"address","信箱":"email","電子郵件":"email","公司名稱":"fullname","統編":"taxid",
  "行動電話":"mobile","通訊地址":"address",
};
const emptyBase = () => ({ agent:"", vendor:"", fullname:"", taxid:"", address:"" });
const emptyC    = () => ({ contact:"", phone:"", fax:"", mobile:"", email:"" });
const padAgent  = v => /^\d+$/.test((v||"").trim()) ? (v||"").trim().padStart(4,"0") : (v||"").trim();
const sortKey   = c => { const a=padAgent(c.agent); const n=(c.fullname||c.vendor||"").trim(); return a?a+n:"zzz"+n; };
const cell      = v => v||"—";
const nowStr    = () => new Date().toLocaleString("zh-TW",{year:"numeric",month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit",second:"2-digit"});
const ADMIN_USER = "ch888";
const ADMIN_PASS  = "24980525";

// ── 主元件 ────────────────────────────────────────────────
export default function App() {
  const [customers,    setCustomers]    = useState([]);
  const [lastUpdated,  setLastUpdated]  = useState(null);
  const [loading,      setLoading]      = useState(true);
  const [syncStatus,   setSyncStatus]   = useState(""); // "", "saving", "saved", "error"

  const [showModal,   setShowModal]   = useState(false);
  const [showConfirm, setShowConfirm] = useState(false);
  const [showLogin,   setShowLogin]   = useState(false);
  const [showPDF,     setShowPDF]     = useState(false);
  const [isAdmin,     setIsAdmin]     = useState(false);
  const [loginForm,   setLoginForm]   = useState({ user:"", pass:"" });
  const [loginErr,    setLoginErr]    = useState("");
  const [editGroup,   setEditGroup]   = useState(null);
  const [search,      setSearch]      = useState("");
  const [importMsg,   setImportMsg]   = useState("");
  const fileRef = useRef();

  // ── Firestore 即時監聽 ─────────────────────────────────
  useEffect(() => {
    setLoading(true);
    const unsub = onSnapshot(collection(db, COL), (snapshot) => {
      const docs = [];
      snapshot.forEach((d) => {
        if (d.id !== META_DOC) docs.push({ ...d.data(), _docId: d.id });
      });
      setCustomers(docs);

      // 讀取最後更新時間（存在 __meta__ 文件）
      const metaSnap = snapshot.docs.find(d => d.id === META_DOC);
      if (metaSnap) setLastUpdated(metaSnap.data().lastUpdated || null);

      setLoading(false);
    }, (err) => {
      console.error("Firebase 監聽錯誤:", err);
      setLoading(false);
    });
    return () => unsub();
  }, []);

  // ── 更新資料到 Firestore ───────────────────────────────
  const saveToFirestore = async (nextList) => {
    setSyncStatus("saving");
    try {
      const t = nowStr();
      const batch = writeBatch(db);

      // 取得目前所有 doc ids
      const snapshot = await getDocs(collection(db, COL));
      const existingIds = new Set();
      snapshot.forEach(d => { if (d.id !== META_DOC) existingIds.add(d.id); });

      // 新 list 的 ids
      const newIds = new Set(nextList.map(c => String(c.id)));

      // 刪除已移除的
      existingIds.forEach(id => {
        if (!newIds.has(id)) batch.delete(doc(db, COL, id));
      });

      // 寫入/更新所有記錄
      nextList.forEach(c => {
        const { _docId, ...data } = c;
        batch.set(doc(db, COL, String(c.id)), data);
      });

      // 更新 meta
      batch.set(doc(db, COL, META_DOC), { lastUpdated: t });

      await batch.commit();
      setLastUpdated(t);
      setSyncStatus("saved");
      setTimeout(() => setSyncStatus(""), 2000);
    } catch (err) {
      console.error("儲存失敗:", err);
      setSyncStatus("error");
      setTimeout(() => setSyncStatus(""), 4000);
    }
  };

  const updateCustomers = async (fn) => {
    setCustomers(prev => {
      const next = typeof fn === "function" ? fn(prev) : fn;
      saveToFirestore(next);
      return next;
    });
  };

  // ── 排序 / 分組 ────────────────────────────────────────
  const sorted = useMemo(() =>
    [...customers].sort((a,b) => sortKey(a).localeCompare(sortKey(b),"zh-Hant-TW",{numeric:true})),
  [customers]);

  const grouped = useMemo(() => {
    const q = search.toLowerCase();
    const f = sorted.filter(c =>
      [c.agent,c.vendor,c.fullname,c.taxid,c.contact,c.phone,c.mobile,c.email]
        .some(v => (v||"").toLowerCase().includes(q))
    );
    const map = new Map();
    f.forEach(c => {
      const k = c.fullname||c.vendor||"(未命名)";
      if (!map.has(k)) map.set(k, []);
      map.get(k).push(c);
    });
    return Array.from(map.entries());
  }, [sorted, search]);

  const totalCo = useMemo(() =>
    new Set(customers.map(c => c.fullname||c.vendor)).size,
  [customers]);

  // ── 登入 ───────────────────────────────────────────────
  const handleLogin = () => {
    if (loginForm.user === ADMIN_USER && loginForm.pass === ADMIN_PASS) {
      setIsAdmin(true); setShowLogin(false); setLoginErr(""); setLoginForm({user:"",pass:""});
    } else { setLoginErr("帳號或密碼錯誤！"); }
  };

  // ── 新增 / 編輯 ────────────────────────────────────────
  const openAdd  = () => { setEditGroup({base:emptyBase(),contacts:[emptyC()],isNew:true}); setShowModal(true); };
  const openEdit = rows => {
    const f = rows[0];
    setEditGroup({
      base: {agent:f.agent,vendor:f.vendor,fullname:f.fullname,taxid:f.taxid,address:f.address},
      contacts: rows.map(r => ({_id:r.id, ...emptyC(), ...Object.fromEntries(CFIELDS.map(cf=>[cf.key,r[cf.key]]))})),
      isNew: false,
      origKey: f.fullname||f.vendor,
    });
    setShowModal(true);
  };
  const saveGroup = () => {
    const { base, contacts, isNew, origKey } = editGroup;
    if (!base.fullname.trim() && !base.vendor.trim()) return alert("請填寫全名或廠商");
    const pb = { ...base, agent: padAgent(base.agent) };
    const rows = contacts
      .filter(c => c.contact||c.phone||c.mobile||c.email)
      .map((c,i) => ({ ...pb, ...c, id: c._id||Date.now()+i }));
    if (!rows.length) rows.push({ ...pb, ...emptyC(), id: Date.now() });
    updateCustomers(cs => {
      const kept = isNew ? cs : cs.filter(c => (c.fullname||c.vendor) !== origKey);
      return [...kept, ...rows];
    });
    setShowModal(false);
  };
  const delGroup = k => {
    if (window.confirm(`確定刪除「${k}」？`))
      updateCustomers(cs => cs.filter(c => (c.fullname||c.vendor) !== k));
  };
  const updBase = (k,v) => setEditGroup(g => ({...g, base:{...g.base,[k]:v}}));
  const updC    = (i,k,v) => setEditGroup(g => ({...g, contacts:g.contacts.map((c,ci)=>ci===i?{...c,[k]:v}:c)}));
  const addC    = () => setEditGroup(g => ({...g, contacts:[...g.contacts,emptyC()]}));
  const rmC     = i => setEditGroup(g => ({...g, contacts:g.contacts.filter((_,ci)=>ci!==i)}));

  // ── 匯入 Excel / CSV ───────────────────────────────────
  const handleImport = e => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const wb   = XLSX.read(evt.target.result, {type:"array"});
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:""});
        if (!rows.length) { setImportMsg("❌ 找不到資料"); return; }
        const imported = [];
        rows.forEach((row,i) => {
          const base = {agent:"",vendor:"",fullname:"",taxid:"",address:"",contact:"",phone:"",fax:"",mobile:"",email:""};
          Object.entries(row).forEach(([k,v]) => { const key=KEY_MAP[k.trim()]; if(key) base[key]=String(v??""); });
          const sk   = ["contact","phone","fax","mobile","email"];
          const arrs = sk.map(k => base[k]?base[k].split(/\r?\n/).map(s=>s.trim()).filter(Boolean):[""]);
          const max  = Math.max(...arrs.map(a=>a.length));
          for (let j=0;j<max;j++) {
            const rec = {...base, id:Date.now()+i*100+j};
            sk.forEach((k,ki) => rec[k] = arrs[ki][j]||"");
            rec.agent = padAgent(rec.agent);
            imported.push(rec);
          }
        });
        const valid = imported.filter(r => r.fullname||r.vendor||r.contact||r.phone||r.email);
        if (!valid.length) { setImportMsg("❌ 無法對應欄位，請確認標題列"); return; }
        updateCustomers(cs => [...cs, ...valid]);
        setImportMsg(`✅ 成功匯入 ${valid.length} 筆資料！`);
        setTimeout(() => setImportMsg(""), 4000);
      } catch { setImportMsg("❌ 讀取失敗"); }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  // ── 匯出 ───────────────────────────────────────────────
  const handleExport = type => {
    const map = new Map();
    sorted.forEach(r => {
      const k = r.fullname||r.vendor||"";
      if (!map.has(k)) map.set(k, {...r, _c:[]});
      map.get(k)._c.push(r);
    });
    const data = Array.from(map.values()).map(g => ({
      "代號":g.agent,"廠商":g.vendor,"全名":g.fullname,"統一編號":g.taxid,
      "聯絡人/負責人":g._c.map(c=>c.contact).join("\n"),"電話":g._c.map(c=>c.phone).join("\n"),
      "傳真":g._c.map(c=>c.fax).join("\n"),"手機":g._c.map(c=>c.mobile).join("\n"),
      "地址":g.address,"信箱":g._c.map(c=>c.email).join("\n"),
    }));
    const ws   = XLSX.utils.json_to_sheet(data);
    const wb   = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "廠商資料");
    const fname = `廠商資料_${new Date().toISOString().slice(0,10)}`;
    if (type==="xlsx") XLSX.writeFile(wb, fname+".xlsx");
    else XLSX.writeFile(wb, fname+".csv", {bookType:"csv"});
  };

  const pdfCompanies = () => {
    const map = new Map();
    sorted.forEach(r => {
      const k = r.fullname||r.vendor||"";
      if (!map.has(k)) map.set(k, {...r, contacts:[]});
      map.get(k).contacts.push(r);
    });
    return Array.from(map.values());
  };

  const handleDownloadPDF = () => {
    const cos  = pdfCompanies();
    const date = new Date().toLocaleDateString("zh-TW",{year:"numeric",month:"long",day:"numeric"});
    const cards = cos.map(g=>`<div class="card"><div class="hd"><span class="ag">${g.agent||""}</span>${g.vendor?`<span class="vd">${g.vendor}</span>`:""}<span class="fn">${g.fullname||g.vendor}</span></div><div class="bd">${g.taxid?`<div class="info">🔢 統編：${g.taxid}</div>`:""}${g.address?`<div class="info">📍 ${g.address}</div>`:""}${g.contacts.map((c,ci)=>`<div class="${ci>0?"ct":""}"><div class="row"><span class="lbl">👤 聯絡人</span><span class="val">${c.contact||"—"}</span></div>${c.phone?`<div class="row"><span class="lbl">📞 電話</span><span class="val">${c.phone}</span></div>`:""} ${c.fax?`<div class="row"><span class="lbl">📠 傳真</span><span class="val">${c.fax}</span></div>`:""} ${c.mobile?`<div class="row"><span class="lbl">📱 手機</span><span class="val">${c.mobile}</span></div>`:""}</div>`).join("")}</div></div>`).join("");
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><link href="https://fonts.googleapis.com/css2?family=Noto+Sans+HK:wght@400;700;900&display=swap" rel="stylesheet"><style>
      *{margin:0;padding:0;box-sizing:border-box;}
      body{font-family:'Noto Sans HK','Microsoft JhengHei',sans-serif;padding:6mm 7mm;background:#efefef;}
      h1{text-align:center;font-size:22px;font-weight:900;color:#1a1a1a;letter-spacing:6px;margin-bottom:3px;}
      .sub{text-align:center;font-size:11px;color:#888;margin-bottom:8px;}
      hr{border:none;border-top:3px solid #555;margin-bottom:10px;}
      .grid{display:grid;grid-template-columns:1fr 1fr;gap:7px;}
      .card{border:1.5px solid #d4d4d4;border-radius:7px;overflow:hidden;break-inside:avoid;page-break-inside:avoid;background:#fff;}
      .hd{background:linear-gradient(135deg,#4f46e5,#7c3aed);padding:6px 10px;display:flex;align-items:baseline;gap:6px;flex-wrap:wrap;}
      .ag{font-size:11px;color:#c7d2fe;font-weight:900;}.vd{font-size:11px;color:#c7d2fe;font-weight:700;}.fn{font-size:12.5px;color:#fff;font-weight:900;}
      .bd{padding:7px 10px;background:#fff;}
      .info{font-size:10.5px;color:#555;padding-bottom:4px;border-bottom:1px dashed #ddd;margin-bottom:5px;line-height:1.6;}
      .ct{border-top:1px dashed #ddd;padding-top:4px;margin-top:4px;}
      .row{display:grid;grid-template-columns:54px 1fr;margin-bottom:2px;}
      .lbl{font-size:10px;color:#999;}.val{font-size:11px;font-weight:700;color:#1a1a1a;}
      .ft{text-align:center;font-size:10px;color:#aaa;margin-top:10px;}
      @media print{@page{size:A4 portrait;margin:6mm 7mm;}body{-webkit-print-color-adjust:exact;print-color-adjust:exact;padding:0;background:#efefef;}}
    </style></head><body>
    <h1>廠 商 通 訊 錄</h1>
    <p class="sub">${date} 製　|　最後更新：${lastUpdated||"—"}</p>
    <hr/>
    <div class="grid">${cards}</div>
    <p class="ft">共 ${cos.length} 家廠商・${customers.length} 筆聯絡人</p>
    </body></html>`;
    const blob = new Blob([html], {type:"text/html;charset=utf-8"});
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement("a");
    a.href = url; a.download = `廠商通訊錄_${new Date().toISOString().slice(0,10)}.html`;
    document.body.appendChild(a); a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
  };

  // ── UI 輔助 ────────────────────────────────────────────
  const sticky = (left,bg) => ({position:"sticky",left,background:bg,zIndex:1,boxShadow:"2px 0 4px rgba(0,0,0,0.06)"});
  const DlIcon = () => (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="white" style={{marginRight:6,verticalAlign:"middle"}}>
      <path d="M12 16l-6-6h4V4h4v6h4l-6 6z"/><rect x="4" y="18" width="16" height="2.5" rx="1.2"/>
    </svg>
  );
  const btn = (color, onClick, label, dl=false) => (
    <button onClick={onClick} style={{background:color,color:"#fff",border:"none",borderRadius:8,padding:"10px 16px",cursor:"pointer",fontWeight:700,fontSize:13,display:"inline-flex",alignItems:"center",fontFamily:"'Noto Sans HK',sans-serif"}}>
      {dl && <DlIcon/>}{label}
    </button>
  );

  // 同步狀態標示
  const SyncBadge = () => {
    if (!syncStatus) return null;
    const cfg = {
      saving: {bg:"#fef9c3",color:"#854d0e",text:"⏳ 儲存中..."},
      saved:  {bg:"#d1fae5",color:"#065f46",text:"✅ 已同步到雲端"},
      error:  {bg:"#fee2e2",color:"#991b1b",text:"❌ 同步失敗，請重試"},
    }[syncStatus];
    if (!cfg) return null;
    return (
      <div style={{background:cfg.bg,color:cfg.color,padding:"6px 12px",borderRadius:8,fontSize:12,fontWeight:700,display:"inline-flex",alignItems:"center",gap:4}}>
        {cfg.text}
      </div>
    );
  };

  const cos = pdfCompanies();

  // ── 載入中畫面 ─────────────────────────────────────────
  if (loading) return (
    <div style={{minHeight:"100vh",background:"#e8e8e8",display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16}}>
      <div style={{width:48,height:48,borderRadius:"50%",border:"4px solid #4f46e5",borderTopColor:"transparent",animation:"spin 0.8s linear infinite"}}/>
      <p style={{color:"#4f46e5",fontWeight:900,fontSize:16,fontFamily:"'Noto Serif HK',serif"}}>連線 Firebase 中...</p>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );

  // ── 主畫面 ─────────────────────────────────────────────
  return (
    <>
      <style>{`html,body,#root{width:100%;height:100%;margin:0;padding:0;overflow-x:hidden;}`}</style>
    <div style={{fontFamily:"'Noto Sans HK', sans-serif",minHeight:"100vh",width:"100vw",background:"#eef0f5",display:"flex",flexDirection:"column",padding:0}}>

      {/* Header */}
      <div style={{background:"#eef0f5",padding:"14px 24px",display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10,boxShadow:"0 2px 8px rgba(0,0,0,0.08)",borderBottom:"1px solid #d8dce6"}}>
        <div>
          <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:2}}>
            <div style={{width:44,height:44,background:"#e0e4ef",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",border:"1.5px solid #c8cfe0",flexShrink:0}}>
              <svg width="26" height="26" viewBox="0 0 24 24" fill="none">
                <rect x="3" y="10" width="18" height="11" rx="1" stroke="#4f46e5" strokeWidth="1.8" fill="#e8eaf6"/>
                <path d="M9 21V15h6v6" stroke="#4f46e5" strokeWidth="1.8" strokeLinejoin="round"/>
                <path d="M1 10L12 2l11 8" stroke="#4f46e5" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/>
                <rect x="7" y="12" width="3" height="3" rx="0.5" fill="#4f46e5" opacity="0.7"/>
                <rect x="14" y="12" width="3" height="3" rx="0.5" fill="#4f46e5" opacity="0.7"/>
              </svg>
            </div>
            <div>
              <h1 style={{margin:0,fontSize:22,color:"#1a1a1a",fontWeight:900,fontFamily:"'Noto Serif HK',serif",letterSpacing:3,lineHeight:1.1}}>供應商管理</h1>
            </div>
          </div>
          <p style={{margin:"3px 0 0",fontSize:12,color:"#666"}}>共 {totalCo} 家供應商・{customers.length} 筆聯絡人</p>
          {lastUpdated && <p style={{margin:"2px 0 0",fontSize:11,color:"#94a3b8"}}>🕐 最後更新：{lastUpdated}</p>}
          <div style={{marginTop:4}}><SyncBadge/></div>
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          {btn("#7c3aed", ()=>handleExport("xlsx"), "匯出 Excel", true)}
          {btn("#0891b2", ()=>handleExport("csv"),  "匯出 CSV",   true)}
          {btn("#4f46e5", ()=>setShowPDF(true),     "匯出通訊錄", true)}
          {isAdmin ? (<>
            <div style={{width:1,height:32,background:"#c8cfe0",margin:"0 4px"}}/>
            {btn("#059669", ()=>fileRef.current.click(), "📂 匯入")}
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleImport}/>
            {btn("#2563eb", openAdd, "＋ 新增")}
            {btn("#64748b", ()=>setShowConfirm(true), "🗑️ 清空")}
            <button onClick={()=>setIsAdmin(false)} style={{background:"#e0e4ef",border:"1.5px solid #c8cfe0",borderRadius:8,padding:"9px 14px",cursor:"pointer",fontSize:13,color:"#4f46e5",fontWeight:700,display:"inline-flex",alignItems:"center",gap:6}}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="8" r="4" stroke="#4f46e5" strokeWidth="2"/><path d="M4 20c0-4 3.6-7 8-7s8 3 8 7" stroke="#4f46e5" strokeWidth="2" strokeLinecap="round"/></svg>
              登出
            </button>
          </>) : (<>
            <div style={{width:1,height:32,background:"#c8cfe0",margin:"0 4px"}}/>
            <button onClick={()=>setShowLogin(true)} style={{background:"#e0e4ef",border:"1.5px solid #c8cfe0",borderRadius:8,padding:"9px 14px",cursor:"pointer",fontSize:13,color:"#4f46e5",fontWeight:700,display:"inline-flex",alignItems:"center",gap:6}}>
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none"><circle cx="12" cy="8" r="4" stroke="#4f46e5" strokeWidth="2"/><path d="M4 20c0-4 3.6-7 8-7s8 3 8 7" stroke="#4f46e5" strokeWidth="2" strokeLinecap="round"/></svg>
              登入
            </button>
          </>)}
        </div>
      </div>

      {importMsg && (
        <div style={{margin:"12px 20px 0",background:importMsg.startsWith("✅")?"#d1fae5":"#fee2e2",color:importMsg.startsWith("✅")?"#065f46":"#991b1b",padding:"9px 14px",borderRadius:8,fontSize:13}}>
          {importMsg}
        </div>
      )}

      <div style={{padding:"12px 20px 20px",display:"flex",flexDirection:"column",flex:1,gap:12}}>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 搜尋全名、廠商、聯絡人、電話..."
        style={{width:"100%",padding:"10px 16px",borderRadius:10,border:"1px solid #d1d5db",fontSize:14,boxSizing:"border-box",fontFamily:"'Noto Sans HK',sans-serif",background:"#fff",boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}/>

      {/* Table */}
      <div style={{background:"#fff",borderRadius:12,boxShadow:"0 2px 12px rgba(0,0,0,0.08)",overflow:"auto",maxHeight:"calc(100vh - 220px)"}}>
        {grouped.length===0 ? (
          <div style={{padding:40,textAlign:"center",color:"#999"}}>尚無資料，請匯入或新增</div>
        ) : (
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,minWidth:900}}>
            <thead>
              <tr style={{backgroundImage:"linear-gradient(135deg,#4f46e5,#7c3aed)"}}>
                {["代號","廠商","全名","統一編號","聯絡人/負責人","電話","傳真","手機","地址","信箱",...(isAdmin?["操作"]:[])].map((h,i) => (
                  <th key={h} style={{padding:"11px 12px",textAlign:"left",color:"#fff",fontWeight:700,whiteSpace:"nowrap",fontSize:13,position:"sticky",top:0,zIndex:i<3?4:2,background:"#2d2d2d",...(i===0?{left:0}:i===1?{left:52}:i===2?{left:104}:{})}}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {grouped.map(([gk,rows],gi) => {
                const bg = gi%2===0?"#ffffff":"#f5f5f5";
                return rows.map((c,ri) => (
                  <tr key={c.id} style={{borderBottom:ri===rows.length-1?"2px solid #d4d4d4":"none",background:bg}}>
                    {ri===0 && <>
                      <td rowSpan={rows.length} style={{padding:"10px 12px",color:"#1a1a1a",verticalAlign:"middle",...sticky(0,bg),fontWeight:900}}>{cell(c.agent)}</td>
                      <td rowSpan={rows.length} style={{padding:"10px 12px",color:"#1a1a1a",verticalAlign:"middle",...sticky(52,bg),fontWeight:900}}>{cell(c.vendor)}</td>
                      <td rowSpan={rows.length} style={{padding:"10px 12px",fontWeight:900,color:"#1a1a1a",whiteSpace:"nowrap",verticalAlign:"middle",...sticky(104,bg),boxShadow:"3px 0 8px rgba(0,0,0,0.08)"}}>{cell(c.fullname)}</td>
                      <td rowSpan={rows.length} style={{padding:"10px 12px",color:"#444",verticalAlign:"middle",fontWeight:700}}>{cell(c.taxid)}</td>
                    </>}
                    <td style={{padding:"10px 12px",color:"#333",borderTop:ri>0?"1.5px solid #e0e0e0":"none"}}>{cell(c.contact)}</td>
                    <td style={{padding:"10px 12px",color:"#333",whiteSpace:"nowrap",borderTop:ri>0?"1.5px solid #e0e0e0":"none"}}>{cell(c.phone)}</td>
                    <td style={{padding:"10px 12px",color:"#333",whiteSpace:"nowrap",borderTop:ri>0?"1.5px solid #e0e0e0":"none"}}>{cell(c.fax)}</td>
                    <td style={{padding:"10px 12px",color:"#333",whiteSpace:"nowrap",borderTop:ri>0?"1.5px solid #e0e0e0":"none"}}>{cell(c.mobile)}</td>
                    {ri===0 && <td rowSpan={rows.length} style={{padding:"10px 12px",color:"#444",minWidth:180,verticalAlign:"middle"}}>{cell(c.address)}</td>}
                    <td style={{padding:"10px 12px",color:"#4f46e5",borderTop:ri>0?"1.5px solid #e0e0e0":"none"}}>{cell(c.email)}</td>
                    {isAdmin && ri===0 && (
                      <td rowSpan={rows.length} style={{padding:"10px 12px",whiteSpace:"nowrap",verticalAlign:"middle"}}>
                        <button onClick={()=>openEdit(rows)} style={{background:"#e0f2fe",color:"#0369a1",border:"none",borderRadius:6,padding:"5px 10px",cursor:"pointer",marginRight:6,fontSize:12,fontFamily:"'Noto Sans HK',sans-serif"}}>編輯</button>
                        <button onClick={()=>delGroup(gk)}   style={{background:"#fee2e2",color:"#dc2626",border:"none",borderRadius:6,padding:"5px 10px",cursor:"pointer",fontSize:12,fontFamily:"'Noto Sans HK',sans-serif"}}>刪除</button>
                      </td>
                    )}
                  </tr>
                ));
              })}
            </tbody>
          </table>
        )}
      </div>

      </div>

      {/* 通訊錄預覽 */}
      {showPDF && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:300,padding:16}}>
          <div style={{background:"#efefef",borderRadius:16,width:"92vw",maxWidth:760,maxHeight:"92vh",display:"flex",flexDirection:"column",boxShadow:"0 8px 40px rgba(0,0,0,0.3)"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 18px",borderBottom:"1px solid #ddd",background:"#fff",borderRadius:"16px 16px 0 0"}}>
              <h3 style={{margin:0,color:"#1a1a1a",fontSize:15,fontWeight:900}}>👁 通訊錄預覽</h3>
              <div style={{display:"flex",gap:8}}>
                {btn("#4f46e5", handleDownloadPDF, "匯出通訊錄 PDF", true)}
                <button onClick={()=>setShowPDF(false)} style={{background:"#e5e5e5",color:"#444",border:"none",borderRadius:8,padding:"7px 14px",cursor:"pointer",fontSize:13,fontFamily:"'Noto Sans HK',sans-serif"}}>✕ 關閉</button>
              </div>
            </div>
            <div style={{overflowY:"auto",padding:20,flex:1}}>
              <h2 style={{textAlign:"center",fontSize:22,fontWeight:900,color:"#1a1a1a",letterSpacing:6,marginBottom:4}}>廠 商 通 訊 錄</h2>
              <p style={{textAlign:"center",fontSize:11,color:"#888",marginBottom:10}}>{new Date().toLocaleDateString("zh-TW",{year:"numeric",month:"long",day:"numeric"})} 製　|　最後更新：{lastUpdated||"—"}</p>
              <hr style={{border:"none",borderTop:"3px solid #555",marginBottom:12}}/>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                {cos.map((g,gi) => (
                  <div key={gi} style={{border:"1.5px solid #d4d4d4",borderRadius:8,overflow:"hidden",background:"#fff"}}>
                    <div style={{background:"#2d2d2d",padding:"6px 10px",display:"flex",alignItems:"baseline",gap:6,flexWrap:"wrap"}}>
                      {g.agent  && <span style={{fontSize:11,color:"#c7d2fe",fontWeight:900}}>{g.agent}</span>}
                      {g.vendor && <span style={{fontSize:11,color:"#c7d2fe",fontWeight:700}}>{g.vendor}</span>}
                      <span style={{fontSize:12.5,color:"#fff",fontWeight:900}}>{g.fullname||g.vendor}</span>
                    </div>
                    <div style={{padding:"7px 10px"}}>
                      {g.taxid   && <div style={{fontSize:10.5,color:"#555",paddingBottom:4,borderBottom:"1px dashed #ddd",marginBottom:4}}>🔢 統編：{g.taxid}</div>}
                      {g.address && <div style={{fontSize:10.5,color:"#555",paddingBottom:4,borderBottom:"1px dashed #ddd",marginBottom:5,lineHeight:1.6}}>📍 {g.address}</div>}
                      {g.contacts.map((c,ci) => (
                        <div key={ci} style={{borderTop:ci>0?"1px dashed #ddd":"none",paddingTop:ci>0?4:0,marginTop:ci>0?4:0}}>
                          <div style={{display:"grid",gridTemplateColumns:"54px 1fr",rowGap:2}}>
                            <span style={{fontSize:10,color:"#999"}}>👤 聯絡人</span><span style={{fontSize:11,fontWeight:700,color:"#1a1a1a"}}>{c.contact||"—"}</span>
                            {c.phone  && <><span style={{fontSize:10,color:"#999"}}>📞 電話</span> <span style={{fontSize:11,color:"#1a1a1a"}}>{c.phone}</span></>}
                            {c.fax    && <><span style={{fontSize:10,color:"#999"}}>📠 傳真</span> <span style={{fontSize:11,color:"#1a1a1a"}}>{c.fax}</span></>}
                            {c.mobile && <><span style={{fontSize:10,color:"#999"}}>📱 手機</span> <span style={{fontSize:11,color:"#1a1a1a"}}>{c.mobile}</span></>}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
              <p style={{textAlign:"center",fontSize:10,color:"#aaa",marginTop:12}}>共 {cos.length} 家廠商・{customers.length} 筆聯絡人</p>
            </div>
          </div>
        </div>
      )}

      {/* 登入 */}
      {showLogin && (
        <div style={{position:"fixed",inset:0,background:"#e8e8e8",display:"flex",alignItems:"center",justifyContent:"center",zIndex:300}}>
          <div style={{background:"#e8e8e8",borderRadius:16,padding:32,width:320,boxShadow:"0 8px 32px rgba(0,0,0,0.12)",border:"1px solid #d4d4d4"}}>
            <h2 style={{margin:"0 0 6px",color:"#1a1a1a",fontSize:18,textAlign:"center",fontWeight:900,fontFamily:"'Noto Serif HK',serif"}}>🔒 管理員登入</h2>
            <p style={{margin:"0 0 20px",color:"#94a3b8",fontSize:12,textAlign:"center"}}>請輸入帳號與密碼</p>
            <div style={{marginBottom:12}}>
              <label style={{fontSize:12,fontWeight:700,color:"#555",display:"block",marginBottom:4}}>帳號</label>
              <input value={loginForm.user} onChange={e=>setLoginForm(f=>({...f,user:e.target.value}))} onKeyDown={e=>e.key==="Enter"&&handleLogin()} placeholder="請輸入帳號"
                style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1px solid #ccc",fontSize:14,boxSizing:"border-box",background:"#f5f5f5"}}/>
            </div>
            <div style={{marginBottom:16}}>
              <label style={{fontSize:12,fontWeight:700,color:"#555",display:"block",marginBottom:4}}>密碼</label>
              <input type="password" value={loginForm.pass} onChange={e=>setLoginForm(f=>({...f,pass:e.target.value}))} onKeyDown={e=>e.key==="Enter"&&handleLogin()} placeholder="請輸入密碼"
                style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1px solid #ccc",fontSize:14,boxSizing:"border-box",background:"#f5f5f5"}}/>
            </div>
            {loginErr && <p style={{color:"#dc2626",fontSize:12,margin:"0 0 12px",textAlign:"center"}}>{loginErr}</p>}
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>{setShowLogin(false);setLoginErr("");setLoginForm({user:"",pass:""}); }} style={{flex:1,padding:10,borderRadius:8,border:"none",background:"#e8e8e8",color:"#5b7cbf",cursor:"pointer",fontWeight:900,fontSize:14,boxShadow:"0 2px 6px rgba(0,0,0,0.1)"}}>取消</button>
              <button onClick={handleLogin} style={{flex:2,padding:10,borderRadius:8,border:"none",background:"#e8e8e8",color:"#3d5fa8",cursor:"pointer",fontWeight:900,fontSize:14,boxShadow:"0 2px 6px rgba(0,0,0,0.12)"}}>登入</button>
            </div>
          </div>
        </div>
      )}

      {/* 清空確認 */}
      {showConfirm && (
        <div style={{position:"fixed",inset:0,background:"#e8e8e8",display:"flex",alignItems:"center",justifyContent:"center",zIndex:200}}>
          <div style={{background:"#e8e8e8",borderRadius:16,padding:28,width:320,textAlign:"center",boxShadow:"0 8px 32px rgba(0,0,0,0.12)",border:"1px solid #d4d4d4"}}>
            <p style={{fontSize:36,margin:"0 0 8px"}}>🗑️</p>
            <h3 style={{margin:"0 0 8px",color:"#1a1a1a",fontWeight:900,fontFamily:"'Noto Serif HK',serif"}}>確定清空所有資料？</h3>
            <p style={{color:"#666",fontSize:13,margin:"0 0 18px"}}>此操作無法復原</p>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setShowConfirm(false)} style={{flex:1,padding:10,borderRadius:8,border:"none",background:"#e8e8e8",color:"#3d5fa8",cursor:"pointer",fontWeight:900,fontSize:14,boxShadow:"0 2px 6px rgba(0,0,0,0.1)"}}>取消</button>
              <button onClick={()=>{ updateCustomers([]); setShowConfirm(false); }} style={{flex:1,padding:10,borderRadius:8,border:"none",background:"#e8e8e8",color:"#3d5fa8",cursor:"pointer",fontWeight:900,fontSize:14,boxShadow:"0 2px 6px rgba(0,0,0,0.1)"}}>確定清空</button>
            </div>
          </div>
        </div>
      )}

      {/* 編輯 Modal */}
      {showModal && editGroup && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100,padding:16}}>
          <div style={{background:"#fff",borderRadius:16,padding:26,width:560,maxWidth:"95vw",maxHeight:"90vh",overflowY:"auto",boxShadow:"0 8px 32px rgba(0,0,0,0.15)"}}>
            <h2 style={{margin:"0 0 14px",fontSize:17,color:"#1a1a1a",fontWeight:900}}>{editGroup.isNew?"新增廠商":`編輯：${editGroup.base.fullname||editGroup.base.vendor}`}</h2>
            <div style={{background:"#f5f5f5",borderRadius:10,padding:14,marginBottom:18}}>
              <p style={{margin:"0 0 10px",fontWeight:900,color:"#1a1a1a",fontSize:12}}>📋 廠商資料</p>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px 14px"}}>
                {FIELDS.map(f => (
                  <div key={f.key} style={{gridColumn:f.key==="address"||f.key==="fullname"?"span 2":"span 1"}}>
                    <label style={{display:"block",fontSize:12,fontWeight:700,color:"#555",marginBottom:3}}>{f.label}{f.required&&<span style={{color:"#dc2626"}}> *</span>}</label>
                    <input value={editGroup.base[f.key]} onChange={e=>updBase(f.key,e.target.value)}
                      style={{width:"100%",padding:"7px 10px",borderRadius:7,border:"1px solid #ddd",fontSize:13,boxSizing:"border-box"}}/>
                  </div>
                ))}
              </div>
            </div>
            <p style={{margin:"0 0 8px",fontWeight:900,color:"#1a1a1a",fontSize:12}}>👤 聯絡人（共 {editGroup.contacts.length} 位）</p>
            {editGroup.contacts.map((ct,i) => (
              <div key={i} style={{border:"1px solid #e0e0e0",borderRadius:10,padding:12,marginBottom:10,background:"#fafafa"}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                  <span style={{fontSize:12,fontWeight:900,color:"#666"}}>聯絡人 {i+1}</span>
                  {editGroup.contacts.length>1 && (
                    <button onClick={()=>rmC(i)} style={{background:"#fee2e2",color:"#dc2626",border:"none",borderRadius:6,padding:"3px 10px",cursor:"pointer",fontSize:12}}>移除</button>
                  )}
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"7px 14px"}}>
                  {CFIELDS.map(f => (
                    <div key={f.key} style={{gridColumn:f.key==="email"?"span 2":"span 1"}}>
                      <label style={{display:"block",fontSize:12,fontWeight:700,color:"#555",marginBottom:3}}>{f.label}</label>
                      <input value={ct[f.key]} onChange={e=>updC(i,f.key,e.target.value)}
                        style={{width:"100%",padding:"7px 10px",borderRadius:7,border:"1px solid #ddd",fontSize:13,boxSizing:"border-box"}}/>
                    </div>
                  ))}
                </div>
              </div>
            ))}
            <button onClick={addC} style={{width:"100%",padding:"8px",borderRadius:8,border:"2px dashed #ddd",background:"#fff",color:"#888",cursor:"pointer",fontSize:13,marginBottom:14}}>＋ 新增聯絡人</button>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setShowModal(false)} style={{flex:1,padding:10,borderRadius:8,border:"1px solid #ddd",background:"#fff",cursor:"pointer",fontSize:14}}>取消</button>
              <button onClick={saveGroup} style={{flex:2,padding:10,borderRadius:8,border:"none",background:"#4f46e5",color:"#fff",cursor:"pointer",fontWeight:900,fontSize:14}}>{editGroup.isNew?"新增廠商":"儲存變更"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
    </>
  );
}