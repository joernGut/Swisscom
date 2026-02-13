import { useState, useMemo, useCallback, useEffect } from "react";

const TASK_TYPES = ["MIG","SUP","PRV","SPM","PO","FER","AUS","NPL"];
const TASK_META = {
  MIG: { bg: "#3b82f6", text: "#fff", label: "Migration" },
  SUP: { bg: "#f59e0b", text: "#18181b", label: "Partner-Support" },
  PRV: { bg: "#22c55e", text: "#fff", label: "Provisioning" },
  SPM: { bg: "#a855f7", text: "#fff", label: "SPM" },
  PO:  { bg: "#ec4899", text: "#fff", label: "Product Owner" },
  FER: { bg: "#06b6d4", text: "#fff", label: "Ferien" },
  AUS: { bg: "#52525b", text: "#a1a1aa", label: "Austritt" },
  NPL: { bg: "#27272a", text: "#71717a", label: "Nicht geplant" },
};
const MAX_FER = 10;
const TEAMS = ["Migration","Provisioning","SPM","PO"];
const WD = ["05.01","12.01","19.01","26.01","02.02","09.02","16.02","23.02","02.03","09.03","16.03","23.03","30.03","06.04","13.04","20.04","27.04","04.05","11.05","18.05","25.05","01.06","08.06","15.06","22.06","29.06","06.07","13.07","20.07","27.07","03.08","10.08","17.08","24.08","31.08","07.09","14.09","21.09","28.09","05.10","12.10","19.10","26.10","02.11","09.11","16.11","23.11","30.11","07.12","14.12","21.12","28.12"];
const W52 = Array.from({length:52},(_,i)=>i);

const RAW = `6144166|Brun, Marc|SPM|100|SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM
7012283|Meier, Adrian|Provisioning|100|PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV
7059193|Simon, René|Migration|100|MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS
7064784|Martinello, Roman Pascal|PO|25|PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO
7072184|Blunier, Jérémy|Provisioning|100|PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV
7085107|Michienzi, Stefano|Provisioning|100|PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV
7091933|Stamenkovic, Stefan|PO|50|PO;PO;PO;PO;PO;PO;PO;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS
7098365|Palmiero, Vincenzo|Migration|100|SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG
7099106|Roncoroni, Simone|Migration|25|SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP
7103412|Kreis, Johann|Migration|100|SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG
7105489|Valle, Fabrizio|Migration|100|MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP
7106797|Meli, Pierre|SPM|50|SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM
7108147|Bruni, Thierry|Migration|100|MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP
7110347|Kamber, Michael|SPM|0|NPL;NPL;NPL;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS
7114123|Kunz, Bruno|SPM|100|SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM
7127751|Trotti, Laurent|Provisioning|100|PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV
7130577|Schnarwiler, Philipp|PO|25|PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO;PO
7134441|Tharatori, Ilir|SPM|25|SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM
7141365|Müller, Florian Lars|Migration|40|MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS
7142133|De Nicola, Vincenzo|Migration|100|SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP
7142250|Di Nicola, Alexandro|Provisioning|25|PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV;PRV
7146765|Selmani, Hazbi|Migration|100|SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP
7149134|Stendardo, David|Migration|80|MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP
7149584|Huon, Florian|Migration|100|SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP
7149991|Gutting, Jörn|SPM|100|SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM
7151803|Madonia, Marco|Migration|80|MIG;SUP;SUP;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS
7151804|Donzé, Semjon|SPM|60|SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;SPM;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS
7151806|Thammavongsa, Erich|Migration|80|AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS;AUS
8000418|Teixeira da Silva, Vitson Mário|Migration|100|MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG;MIG;SUP;SUP;MIG`;

function parseData() {
  return RAW.trim().split("\n").map(line => {
    const [pnr,name,team,pct,sh] = line.split("|");
    const shifts = sh.split(";");
    return { id:pnr, pnr:Number(pnr), name, team, pct:Number(pct), shifts:[...shifts], base:[...shifts] };
  });
}
const F="'IBM Plex Sans','Segoe UI',system-ui,sans-serif",M="'IBM Plex Mono','Fira Code',monospace";
const S={
  root:{fontFamily:F,background:"#09090b",color:"#fafafa",minHeight:"100vh",display:"flex",flexDirection:"column"},
  nav:{background:"#18181b",borderBottom:"1px solid #27272a",padding:"0 24px",display:"flex",alignItems:"center",height:52,flexShrink:0},
  nb:a=>({padding:"14px 20px",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:F,color:a?"#fafafa":"#71717a",background:"transparent",border:"none",borderBottom:a?"2px solid #3b82f6":"2px solid transparent"}),
  nt:{fontSize:15,fontWeight:700,color:"#fafafa",marginRight:32,letterSpacing:"-0.03em"},
  card:{background:"#18181b",border:"1px solid #27272a",borderRadius:10,padding:20,marginBottom:16},
  lbl:{fontSize:12,fontWeight:600,color:"#a1a1aa",marginBottom:6,display:"block"},
  inp:{width:"100%",padding:"8px 12px",background:"#09090b",border:"1px solid #3f3f46",borderRadius:6,color:"#fafafa",fontSize:13,fontFamily:F,outline:"none",boxSizing:"border-box"},
  sel:{width:"100%",padding:"8px 12px",background:"#09090b",border:"1px solid #3f3f46",borderRadius:6,color:"#fafafa",fontSize:13,fontFamily:F,outline:"none",boxSizing:"border-box"},
  bp:{padding:"9px 20px",background:"#3b82f6",color:"#fff",border:"none",borderRadius:6,fontSize:13,fontWeight:600,cursor:"pointer",fontFamily:F},
  bg:{padding:"9px 20px",background:"transparent",color:"#a1a1aa",border:"1px solid #3f3f46",borderRadius:6,fontSize:13,fontWeight:600,cursor:"pointer",fontFamily:F},
  tag:(b,c)=>({display:"inline-flex",alignItems:"center",gap:4,padding:"3px 10px",borderRadius:999,fontSize:11,fontWeight:600,background:b,color:c}),
};

function TL({shifts}){return <div style={{display:"flex",gap:1}}>{shifts.map((s,i)=><div key={i} title={`KW${i+1}: ${s}`} style={{flex:1,height:18,borderRadius:2,background:TASK_META[s]?.bg||"#27272a",minWidth:0}}/>)}</div>}
function WP({label,value,onChange}){return <div><label style={S.lbl}>{label}</label><select value={value??""} onChange={e=>onChange(e.target.value===""?null:Number(e.target.value))} style={S.sel}><option value="">— kein Datum —</option>{W52.map(i=><option key={i} value={i}>KW {i+1} ({WD[i]})</option>)}</select></div>}
const teamDef=t=>t==="Migration"?"MIG":t==="Provisioning"?"PRV":t==="SPM"?"SPM":"PO";


function EintrittPage({data,setData}){
  const[sid,setSid]=useState(null),[eE,sEE]=useState(null),[eX,sEX]=useState(null),[q,sQ]=useState("");
  const p=data.find(d=>d.id===sid);
  useEffect(()=>{if(!p)return;const f=p.shifts.findIndex(s=>s!=="AUS"&&s!=="NPL");const r=[...p.shifts].reverse().findIndex(s=>s!=="AUS"&&s!=="NPL");const l=r>=0?51-r:-1;sEE(f>=0?f:null);sEX(l>=0&&l<51&&p.shifts[51]==="AUS"?l:null);},[sid]);
  const fl=data.filter(d=>!q||d.name.toLowerCase().includes(q.toLowerCase()));
  const gs=p=>{if(p.shifts.every(s=>s==="AUS"))return{l:"Ausgetreten",b:"#3f3f46",c:"#a1a1aa"};const f=p.shifts.findIndex(s=>s!=="AUS"&&s!=="NPL");if(f>0)return{l:`Eintritt KW${f+1}`,b:"#164e63",c:"#22d3ee"};const r=[...p.shifts].reverse().findIndex(s=>s!=="AUS"&&s!=="NPL");if(r>=0&&(51-r)<51&&p.shifts[51]==="AUS")return{l:`Austritt KW${52-r}`,b:"#4c1d24",c:"#f87171"};return{l:"Ganzjährig",b:"#14532d",c:"#4ade80"}};
  const apply=()=>{if(!p)return;setData(pr=>pr.map(x=>{if(x.id!==sid)return x;const n=[...x.shifts];const d=teamDef(x.team);if(eE!=null&&eE>0)for(let i=0;i<eE;i++)n[i]="AUS";const s=eE??0,e=eX??51;for(let i=s;i<=e;i++)if(n[i]==="AUS"||n[i]==="NPL")n[i]=d;if(eX!=null&&eX<51)for(let i=eX+1;i<52;i++)n[i]="AUS";return{...x,shifts:n}}))};
  return <div style={{display:"flex",height:"calc(100vh - 52px)"}}>
    <div style={{width:380,borderRight:"1px solid #27272a",display:"flex",flexDirection:"column",flexShrink:0}}>
      <div style={{padding:"16px 16px 12px"}}><input placeholder="Suchen..." value={q} onChange={e=>sQ(e.target.value)} style={{...S.inp,background:"#27272a"}}/></div>
      <div style={{flex:1,overflowY:"auto"}}>{fl.map(x=>{const st=gs(x);return <div key={x.id} onClick={()=>setSid(x.id)} style={{padding:"12px 16px",cursor:"pointer",borderBottom:"1px solid #27272a",background:x.id===sid?"#27272a":"transparent",borderLeft:x.id===sid?"3px solid #3b82f6":"3px solid transparent"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:13,fontWeight:600}}>{x.name}</div><div style={{fontSize:11,color:"#71717a",marginTop:2}}>{x.team} · {x.pct}%</div></div><span style={S.tag(st.b,st.c)}>{st.l}</span></div></div>})}</div>
    </div>
    <div style={{flex:1,overflowY:"auto",padding:28}}>
      {!p?<div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100%",color:"#52525b"}}><div style={{textAlign:"center"}}><div style={{fontSize:40,marginBottom:12}}>&#8592;</div>Mitarbeiter auswählen</div></div>:
      <div style={{maxWidth:720}}>
        <h2 style={{margin:"0 0 8px",fontSize:22,fontWeight:700}}>{p.name}</h2>
        <div style={{display:"flex",gap:10,marginBottom:24,flexWrap:"wrap"}}><span style={S.tag("#1e3a5f","#60a5fa")}>{p.team}</span><span style={S.tag("#27272a","#a1a1aa")}>{p.pct}%</span><span style={S.tag("#164e63","#06b6d4")}>{p.shifts.filter(s=>s==="FER").length}/{MAX_FER} Ferien</span></div>
        <div style={S.card}><div style={{fontSize:12,fontWeight:600,color:"#a1a1aa",marginBottom:6}}>JAHRESÜBERSICHT</div><TL shifts={p.shifts}/></div>
        <div style={S.card}><div style={{fontSize:14,fontWeight:700,marginBottom:16}}>Eintritt / Austritt</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}><WP label="EINTRITT" value={eE} onChange={sEE}/><WP label="AUSTRITT" value={eX} onChange={sEX}/></div>
          <div style={{display:"flex",gap:10}}><button style={S.bp} onClick={apply}>Anwenden</button><button style={S.bg} onClick={()=>{sEE(null);sEX(null)}}>Reset</button></div></div>
        <div style={S.card}><div style={{fontSize:14,fontWeight:700,marginBottom:12}}>Pensum / Team</div>
          <div style={{display:"flex",gap:12}}><div style={{flex:1}}><label style={S.lbl}>PENSUM %</label><input type="number" min={0} max={100} value={p.pct} onChange={e=>{const v=Math.min(100,Math.max(0,Number(e.target.value)));setData(pr=>pr.map(x=>x.id===sid?{...x,pct:v}:x))}} style={{...S.inp,width:120}}/></div>
          <div style={{flex:1}}><label style={S.lbl}>TEAM</label><select value={p.team} onChange={e=>setData(pr=>pr.map(x=>x.id===sid?{...x,team:e.target.value}:x))} style={S.sel}>{TEAMS.map(t=><option key={t}>{t}</option>)}</select></div></div></div>
      </div>}
    </div></div>}


function FerienPage({data,setData}){
  const[sid,setSid]=useState(null),[q,sQ]=useState(""),[ft,sFt]=useState("Alle");
  const p=data.find(d=>d.id===sid);
  const act=useMemo(()=>data.filter(x=>!x.shifts.every(s=>s==="AUS")),[data]);
  const fl=useMemo(()=>act.filter(d=>(ft==="Alle"||d.team===ft)&&(!q||d.name.toLowerCase().includes(q.toLowerCase()))),[act,ft,q]);
  const fc=useCallback(x=>x.shifts.filter(s=>s==="FER").length,[]);
  const tog=wi=>{if(!p)return;const c=p.shifts[wi];if(c==="AUS"||c==="NPL")return;
    setData(pr=>pr.map(x=>{if(x.id!==sid)return x;const n=[...x.shifts];
      if(n[wi]==="FER"){n[wi]=x.base[wi]!=="FER"?x.base[wi]:teamDef(x.team)}
      else{if(fc(x)>=MAX_FER)return x;n[wi]="FER"}return{...x,shifts:n}}))};
  const clr=()=>{if(!p)return;setData(pr=>pr.map(x=>x.id!==sid?x:{...x,shifts:x.shifts.map((s,i)=>s==="FER"?(x.base[i]!=="FER"?x.base[i]:teamDef(x.team)):s)}))};
  const hm=useMemo(()=>W52.map(wi=>{let c=0;act.forEach(x=>{if(x.shifts[wi]==="FER")c++});return c}),[act]);
  const mx=Math.max(...hm,1);
  return <div style={{display:"flex",height:"calc(100vh - 52px)"}}>
    <div style={{width:360,borderRight:"1px solid #27272a",display:"flex",flexDirection:"column",flexShrink:0}}>
      <div style={{padding:"12px 16px",borderBottom:"1px solid #27272a"}}>
        <input placeholder="Suchen..." value={q} onChange={e=>sQ(e.target.value)} style={{...S.inp,background:"#27272a",marginBottom:8}}/>
        <div style={{display:"flex",gap:3}}>{["Alle",...TEAMS].map(t=><button key={t} onClick={()=>sFt(t)} style={{padding:"4px 10px",borderRadius:4,border:"none",cursor:"pointer",fontSize:10,fontWeight:600,fontFamily:F,background:ft===t?"#3b82f6":"#27272a",color:ft===t?"#fff":"#71717a"}}>{t}</button>)}</div></div>
      <div style={{flex:1,overflowY:"auto"}}>{fl.map(x=>{const c=fc(x);return <div key={x.id} onClick={()=>setSid(x.id)} style={{padding:"10px 16px",cursor:"pointer",borderBottom:"1px solid #27272a",background:x.id===sid?"#27272a":"transparent",borderLeft:x.id===sid?"3px solid #06b6d4":"3px solid transparent"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:13,fontWeight:600}}>{x.name}</div><div style={{fontSize:10,color:"#71717a"}}>{x.team}</div></div>
        <div style={{textAlign:"right"}}><div style={{fontSize:18,fontWeight:700,color:c>=MAX_FER?"#f87171":"#06b6d4",fontFamily:M}}>{c}</div><div style={{fontSize:9,color:"#52525b"}}>/ {MAX_FER}</div></div></div>
        <div style={{display:"flex",height:4,gap:1,marginTop:6,borderRadius:2,overflow:"hidden"}}>{Array.from({length:MAX_FER}).map((_,i)=><div key={i} style={{flex:1,background:i<c?"#06b6d4":"#27272a",borderRadius:1}}/>)}</div></div>})}</div></div>
    <div style={{flex:1,overflowY:"auto",padding:28}}>
      <div style={S.card}><div style={{fontSize:14,fontWeight:700,marginBottom:4}}>Team-Ferienkalender</div><div style={{fontSize:11,color:"#52525b",marginBottom:12}}>Personen in Ferien pro KW</div>
        <div style={{display:"flex",gap:1}}>{hm.map((c,i)=><div key={i} title={`KW${i+1}: ${c}`} style={{flex:1,height:28,borderRadius:2,minWidth:0,display:"flex",alignItems:"center",justifyContent:"center",background:c===0?"#18181b":`rgba(6,182,212,${Math.max(0.15,c/mx)})`,fontSize:8,fontWeight:700,fontFamily:M,color:c>0?"#fff":"#3f3f46"}}>{c>0?c:""}</div>)}</div>
        <div style={{display:"flex",justifyContent:"space-between",marginTop:4}}>{[1,13,26,39,52].map(w=><span key={w} style={{fontSize:9,color:"#52525b",fontFamily:M}}>KW{w}</span>)}</div></div>
      {!p?<div style={{textAlign:"center",color:"#52525b",padding:60}}><div style={{fontSize:36,marginBottom:8}}>&#127958;</div><div style={{fontSize:14}}>Mitarbeiter auswählen</div><div style={{fontSize:12,marginTop:4}}>Klick auf Woche zum Setzen/Entfernen (max. {MAX_FER})</div></div>:
      <div style={{maxWidth:800}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20}}>
          <div><h2 style={{margin:"0 0 6px",fontSize:22,fontWeight:700}}>{p.name}</h2><div style={{display:"flex",gap:8}}><span style={S.tag("#1e3a5f","#60a5fa")}>{p.team}</span><span style={S.tag("#27272a","#a1a1aa")}>{p.pct}%</span></div></div>
          <div style={{textAlign:"right"}}><div style={{fontSize:36,fontWeight:800,color:fc(p)>=MAX_FER?"#f87171":"#06b6d4",fontFamily:M,lineHeight:1}}>{fc(p)}</div><div style={{fontSize:11,color:"#71717a"}}>von {MAX_FER}</div></div></div>
        <div style={{background:"#27272a",borderRadius:6,height:8,marginBottom:20,overflow:"hidden"}}><div style={{height:"100%",borderRadius:6,transition:"width 0.3s",width:`${(fc(p)/MAX_FER)*100}%`,background:fc(p)>=MAX_FER?"#dc2626":"#06b6d4"}}/></div>
        <div style={S.card}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><div><div style={{fontSize:14,fontWeight:700}}>Ferienwochen wählen</div><div style={{fontSize:11,color:"#52525b"}}>Max. {MAX_FER} Wochen</div></div><button onClick={clr} style={{...S.bg,padding:"6px 14px",fontSize:11}}>Alle löschen</button></div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(13,1fr)",gap:4}}>{W52.map(i=>{const s=p.shifts[i],iF=s==="FER",isO=s==="AUS"||s==="NPL",ok=!isO&&(iF||fc(p)<MAX_FER);
            return <div key={i} onClick={()=>ok&&tog(i)} style={{padding:"8px 4px",borderRadius:6,textAlign:"center",cursor:ok?"pointer":"default",background:iF?"#06b6d4":isO?"#1a1a1e":"#27272a",border:iF?"2px solid #22d3ee":"2px solid transparent",opacity:isO?0.3:(!iF&&fc(p)>=MAX_FER)?0.5:1,transition:"all 0.15s"}}>
              <div style={{fontSize:10,fontWeight:700,color:iF?"#fff":"#a1a1aa",fontFamily:M}}>KW{i+1}</div>
              <div style={{fontSize:8,color:iF?"#cffafe":"#52525b",marginTop:1}}>{WD[i]}</div>
              {iF&&<div style={{fontSize:11,marginTop:2}}>&#127958;</div>}
              {!iF&&!isO&&<div style={{fontSize:8,marginTop:2,fontWeight:700,color:TASK_META[s]?.bg,fontFamily:M}}>{s}</div>}
            </div>})}</div></div>
        {fc(p)>0&&<div style={S.card}><div style={{fontSize:12,fontWeight:700,color:"#a1a1aa",marginBottom:8}}>GEPLANTE FERIEN</div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{p.shifts.map((s,i)=>s==="FER"?<span key={i} onClick={()=>tog(i)} style={{...S.tag("#164e63","#22d3ee"),cursor:"pointer",fontSize:12,padding:"5px 12px"}}>KW{i+1} ({WD[i]}) ✕</span>:null)}</div></div>}
      </div>}
    </div></div>}


function SchichtPage({data,setData}){
  const[fT,sFT]=useState("Alle"),[fN,sFN]=useState(""),[ec,sEC]=useState(null),[bm,sBM]=useState(false),[bt,sBT]=useState("MIG"),[sel,sSel]=useState(new Set()),[dr,sDr]=useState(false);
  const fl=useMemo(()=>data.filter(r=>(fT==="Alle"||r.team===fT)&&(!fN||r.name.toLowerCase().includes(fN.toLowerCase()))),[data,fT,fN]);
  const cs=useCallback((ri,ci,v)=>{const x=fl[ri];setData(pr=>pr.map(d=>d.id!==x.id?d:{...d,shifts:d.shifts.map((s,j)=>j===ci?v:s)}));},[fl,setData]);
  const ms=useMemo(()=>W52.map(wi=>{let f=0;data.forEach(r=>{if(r.shifts[wi]==="MIG")f+=r.pct/100});return{f:Math.round(f*100)/100,ok:f>=4}}),[data]);
  const ab=()=>{setData(pr=>{const n=pr.map(r=>({...r,shifts:[...r.shifts]}));sel.forEach(k=>{const[ri,ci]=k.split("-").map(Number);const x=fl[ri];const idx=n.findIndex(d=>d.id===x.id);if(idx>=0)n[idx].shifts[ci]=bt});return n});sSel(new Set())};
  useEffect(()=>{const u=()=>sDr(false);window.addEventListener("mouseup",u);return()=>window.removeEventListener("mouseup",u)},[]);
  return <div style={{display:"flex",flexDirection:"column",height:"calc(100vh - 52px)"}}>
    <div style={{padding:"10px 20px",borderBottom:"1px solid #27272a",background:"#18181b",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
      <div style={{display:"flex",gap:3}}>{["Alle",...TEAMS].map(t=><button key={t} onClick={()=>sFT(t)} style={{padding:"5px 12px",borderRadius:5,border:"none",cursor:"pointer",fontSize:11,fontWeight:600,fontFamily:F,background:fT===t?"#3b82f6":"#27272a",color:fT===t?"#fff":"#71717a"}}>{t}</button>)}</div>
      <input placeholder="Name..." value={fN} onChange={e=>sFN(e.target.value)} style={{...S.inp,width:150,background:"#27272a",padding:"5px 10px",fontSize:11}}/>
      <div style={{borderLeft:"1px solid #3f3f46",paddingLeft:8,display:"flex",gap:6,alignItems:"center"}}>
        <button onClick={()=>{sBM(!bm);sSel(new Set())}} style={{...S.bg,padding:"5px 12px",fontSize:11,background:bm?"#7f1d1d":"transparent",color:bm?"#fca5a5":"#71717a",borderColor:bm?"#991b1b":"#3f3f46"}}>{bm?"✕ Bulk AUS":"Bulk-Modus"}</button>
        {bm&&<><select value={bt} onChange={e=>sBT(e.target.value)} style={{...S.sel,width:"auto",padding:"4px 8px",fontSize:11}}>{TASK_TYPES.map(t=><option key={t}>{t}</option>)}</select><button onClick={ab} style={{...S.bp,padding:"5px 14px",fontSize:11,background:"#16a34a"}}>Anwenden ({sel.size})</button></>}
      </div>
      <div style={{marginLeft:"auto",display:"flex",gap:12,fontSize:10}}><span style={{color:"#4ade80"}}>MIG≥4: {ms.filter(s=>s.ok).length}W</span><span style={{color:"#f87171"}}>MIG&lt;4: {ms.filter(s=>!s.ok).length}W</span></div>
    </div>
    <div style={{display:"flex",gap:10,padding:"5px 20px",background:"#0f0f12",borderBottom:"1px solid #27272a",alignItems:"center"}}>
      <span style={{fontSize:9,color:"#52525b",fontWeight:700}}>LEGENDE:</span>
      {TASK_TYPES.map(t=><div key={t} style={{display:"flex",alignItems:"center",gap:3}}><span style={{width:10,height:10,borderRadius:2,background:TASK_META[t].bg}}/><span style={{fontSize:9,color:"#71717a"}}>{t}</span></div>)}
    </div>
    <div style={{flex:1,overflow:"auto"}}>
      <table style={{borderCollapse:"separate",borderSpacing:0,width:"max-content",fontSize:10,fontFamily:M,tableLayout:"fixed"}}>
        <thead>
          <tr style={{position:"sticky",top:0,zIndex:20}}>
            <th style={{position:"sticky",left:0,zIndex:30,background:"#09090b",padding:"3px 6px",borderBottom:"1px solid #27272a",textAlign:"left",color:"#52525b",fontSize:8,width:170,minWidth:170}}>MIG FTE</th>
            <th style={{position:"sticky",left:170,zIndex:30,background:"#09090b",width:30,minWidth:30,borderBottom:"1px solid #27272a"}}/>
            {W52.map(i=><th key={i} style={{background:ms[i].ok?"#052e16":"#2c0b0e",borderBottom:"1px solid #27272a",padding:"3px 0",textAlign:"center",fontSize:9,fontWeight:700,width:34,minWidth:34,color:ms[i].ok?"#4ade80":"#f87171"}}>{ms[i].f}</th>)}
          </tr>
          <tr style={{position:"sticky",top:22,zIndex:20}}>
            <th style={{position:"sticky",left:0,zIndex:30,background:"#18181b",padding:"5px 8px",borderBottom:"2px solid #3b82f6",textAlign:"left",color:"#a1a1aa",fontSize:10,width:170,minWidth:170}}>Name</th>
            <th style={{position:"sticky",left:170,zIndex:30,background:"#18181b",padding:"5px 2px",borderBottom:"2px solid #3b82f6",textAlign:"center",color:"#a1a1aa",fontSize:9,width:30,minWidth:30}}>%</th>
            {W52.map(i=><th key={i} style={{background:"#18181b",padding:"3px 0",borderBottom:"2px solid #3b82f6",textAlign:"center",color:"#52525b",fontSize:8,width:34,minWidth:34,lineHeight:1.2}}><div>W{i+1}</div><div style={{fontSize:7,color:"#3f3f46"}}>{WD[i]}</div></th>)}
          </tr>
        </thead>
        <tbody>{fl.map((row,ri)=><tr key={row.id}>
          <td style={{position:"sticky",left:0,zIndex:10,background:"#09090b",padding:"3px 8px",fontSize:10,fontWeight:600,color:"#e4e4e7",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",borderBottom:"1px solid #1a1a1e",fontFamily:F}} title={row.name}>{row.name}</td>
          <td style={{position:"sticky",left:170,zIndex:10,background:"#09090b",textAlign:"center",fontSize:9,color:"#52525b",borderBottom:"1px solid #1a1a1e",borderRight:"2px solid #27272a"}}>{row.pct}</td>
          {row.shifts.map((s,ci)=>{const m=TASK_META[s]||TASK_META.NPL,isSel=sel.has(`${ri}-${ci}`),isEd=ec?.r===ri&&ec?.c===ci;
            return <td key={ci} onMouseDown={()=>{if(bm){sDr(true);sSel(p=>{const n=new Set(p);n.has(`${ri}-${ci}`)?n.delete(`${ri}-${ci}`):n.add(`${ri}-${ci}`);return n})}}}
              onMouseEnter={()=>{if(dr&&bm)sSel(p=>new Set([...p,`${ri}-${ci}`]))}}
              onClick={()=>{if(!bm)sEC(isEd?null:{r:ri,c:ci})}}
              style={{padding:1,textAlign:"center",cursor:"pointer",borderBottom:"1px solid #18181b11",borderRight:ci%4===3?"1px solid #27272a":"none",position:"relative",userSelect:"none"}}>
              <div style={{background:m.bg,color:m.text,borderRadius:3,padding:"3px 0",fontSize:8,fontWeight:700,outline:isSel?"2px solid #eab308":isEd?"2px solid #3b82f6":"none",outlineOffset:-1,transform:isEd?"scale(1.15)":"none",transition:"transform 0.08s",position:"relative",zIndex:isEd?5:0}}>{s}</div>
              {isEd&&!bm&&<div style={{position:"absolute",top:"100%",left:-20,zIndex:100,background:"#18181b",border:"1px solid #3f3f46",borderRadius:8,boxShadow:"0 12px 40px rgba(0,0,0,0.6)",padding:4,minWidth:120}}>
                {TASK_TYPES.map(t=><button key={t} onClick={e=>{e.stopPropagation();cs(ri,ci,t);sEC(null)}} style={{display:"flex",alignItems:"center",gap:6,padding:"5px 8px",width:"100%",background:s===t?"#27272a":"transparent",border:"none",borderRadius:4,cursor:"pointer",color:"#e4e4e7",fontSize:11,fontFamily:M,textAlign:"left"}}>
                  <span style={{width:12,height:12,borderRadius:3,background:TASK_META[t].bg,flexShrink:0}}/><span style={{fontWeight:700,width:28}}>{t}</span><span style={{color:"#71717a",fontSize:10,fontFamily:F}}>{TASK_META[t].label}</span></button>)}</div>}
            </td>})}
        </tr>)}</tbody>
      </table>
    </div></div>}


function DashPage({data}){
  const ts=useMemo(()=>{const s={};TEAMS.forEach(t=>{s[t]={tot:0,act:0,fte:0,ex:0,fer:0}});data.forEach(p=>{const t=p.team;if(!s[t])return;s[t].tot++;if(p.shifts.every(x=>x==="AUS"))s[t].ex++;else{s[t].act++;s[t].fte+=p.pct/100}s[t].fer+=p.shifts.filter(x=>x==="FER").length});return s},[data]);
  const mw=useMemo(()=>W52.map(wi=>{let f=0;data.forEach(r=>{if(r.shifts[wi]==="MIG")f+=r.pct/100});return Math.round(f*100)/100}),[data]);
  const mx=Math.max(...mw,1);
  const fw=useMemo(()=>W52.map(wi=>{let c=0;data.forEach(r=>{if(r.shifts[wi]==="FER")c++});return c}),[data]);
  const mxf=Math.max(...fw,1);
  const tf=data.reduce((s,p)=>s+p.shifts.filter(x=>x==="FER").length,0);
  const actN=data.filter(p=>!p.shifts.every(s=>s==="AUS")).length;
  return <div style={{padding:28,overflowY:"auto",height:"calc(100vh - 52px)"}}>
    <h2 style={{margin:"0 0 24px",fontSize:20,fontWeight:700}}>Dashboard</h2>
    <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:28}}>
      {TEAMS.map(t=><div key={t} style={S.card}><div style={{fontSize:11,fontWeight:600,color:"#71717a",textTransform:"uppercase",letterSpacing:"0.05em"}}>{t}</div><div style={{fontSize:28,fontWeight:700}}>{ts[t].act}</div><div style={{fontSize:11,color:"#52525b",marginTop:2}}>{ts[t].fte.toFixed(1)} FTE · {ts[t].ex} ausg. · {ts[t].fer}W Ferien</div></div>)}
    </div>
    <div style={{...S.card,background:"linear-gradient(135deg,#164e63 0%,#18181b 100%)",borderColor:"#0e7490",marginBottom:28}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><div><div style={{fontSize:14,fontWeight:700}}>&#127958; Ferien gesamt</div><div style={{fontSize:11,color:"#67e8f9",marginTop:2}}>{tf} Wochen geplant · Ø {(tf/Math.max(1,actN)).toFixed(1)} pro Person</div></div><div style={{fontSize:36,fontWeight:800,color:"#06b6d4",fontFamily:M}}>{tf}</div></div></div>
    <div style={S.card}><div style={{fontSize:14,fontWeight:700,marginBottom:4}}>Ferien pro Woche</div><div style={{fontSize:11,color:"#52525b",marginBottom:16}}>Personen abwesend</div>
      <div style={{position:"relative",height:120}}><div style={{display:"flex",gap:1,height:"100%",alignItems:"flex-end"}}>{fw.map((v,i)=><div key={i} title={`KW${i+1}: ${v}`} style={{flex:1,background:v>0?"#06b6d4":"#27272a",borderRadius:"3px 3px 0 0",height:`${Math.max((v/mxf)*100,v>0?8:2)}%`,minWidth:0}}/>)}</div></div>
      <div style={{display:"flex",justifyContent:"space-between",marginTop:4}}>{[1,13,26,39,52].map(w=><span key={w} style={{fontSize:9,color:"#52525b",fontFamily:M}}>KW{w}</span>)}</div></div>
    <div style={S.card}><div style={{fontSize:14,fontWeight:700,marginBottom:4}}>Migration FTE</div><div style={{fontSize:11,color:"#52525b",marginBottom:16}}>Min. 4.0 FTE</div>
      <div style={{position:"relative",height:160}}><div style={{position:"absolute",left:0,right:0,bottom:`${(4/mx)*100}%`,borderTop:"1.5px dashed #dc2626",zIndex:2}}><span style={{position:"absolute",right:0,top:-14,fontSize:9,color:"#dc2626",fontFamily:M}}>4.0</span></div>
        <div style={{display:"flex",gap:1,height:"100%",alignItems:"flex-end"}}>{mw.map((v,i)=><div key={i} title={`KW${i+1}: ${v}`} style={{flex:1,background:v>=4?"#166534":"#991b1b",borderRadius:"3px 3px 0 0",height:`${Math.max((v/mx)*100,2)}%`,minWidth:0}}/>)}</div></div>
      <div style={{display:"flex",justifyContent:"space-between",marginTop:4}}>{[1,13,26,39,52].map(w=><span key={w} style={{fontSize:9,color:"#52525b",fontFamily:M}}>KW{w}</span>)}</div></div>
  </div>}


export default function App(){
  const[data,setData]=useState(parseData);
  const[page,setPage]=useState("eintritt");
  return <div style={S.root}>
    <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600;700&display=swap" rel="stylesheet"/>
    <nav style={S.nav}>
      <span style={S.nt}>Produktionsplanung 2026</span>
      <button style={S.nb(page==="eintritt")} onClick={()=>setPage("eintritt")}>Eintritt / Austritt</button>
      <button style={S.nb(page==="ferien")} onClick={()=>setPage("ferien")}>Ferien</button>
      <button style={S.nb(page==="schicht")} onClick={()=>setPage("schicht")}>Schichtplan</button>
      <button style={S.nb(page==="dash")} onClick={()=>setPage("dash")}>Dashboard</button>
    </nav>
    <div style={{flex:1,overflow:"hidden"}}>
      {page==="eintritt"&&<EintrittPage data={data} setData={setData}/>}
      {page==="ferien"&&<FerienPage data={data} setData={setData}/>}
      {page==="schicht"&&<SchichtPage data={data} setData={setData}/>}
      {page==="dash"&&<DashPage data={data}/>}
    </div>
  </div>}
