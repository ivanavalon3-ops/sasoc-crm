import { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, LineChart, Line, AreaChart, Area, PieChart, Pie, Cell,
         XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, ReferenceLine } from "recharts";

// ─── STORAGE ──────────────────────────────────────────────────────
const K = {
  ventas:"sasoc:ventas:v1", clientes:"sasoc:clientes:v1",
  proveedores:"sasoc:proveedores:v1", facsP:"sasoc:facturas_prov:v1",
  gastos:"sasoc:gastos:v1", pagos:"sasoc:pagos:v1",
  cuentas:"sasoc:cf:cuentas:v1", movs:"sasoc:cf:movimientos:v1",
  cotizaciones:"sasoc:cotizaciones:v1",
};
const gl = async (k) => { try { const r=await window.storage.get(k,true); return r?JSON.parse(r.value):[]; } catch{return[];} };
const sl = async (k,d) => { try{await window.storage.set(k,JSON.stringify(d),true);}catch(e){console.error(e);} };

// ─── HELPERS ──────────────────────────────────────────────────────
const uid   = () => Date.now().toString(36)+Math.random().toString(36).slice(2,6);
const hoy   = () => new Date().toISOString().slice(0,10);
const $     = (n) => { if(n==null||isNaN(n))return"—"; const a=Math.abs(Number(n)); const s="$"+a.toLocaleString("es-AR",{maximumFractionDigits:0}); return Number(n)<0?"-"+s:s; };
const fecha = (s) => { if(!s)return"—"; const[y,m,d]=s.split("-"); return`${d}/${m}/${y}`; };
const pct   = (n) => isNaN(n)?"—":n.toFixed(1)+"%";
const mesL  = (s) => { if(!s)return""; const[y,m]=s.split("-"); return["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"][parseInt(m)-1]+(" "+y.slice(2)); };
const pN    = (v) => parseFloat(v)||0;
const hoyS  = hoy();

// ─── CONSTANTES ───────────────────────────────────────────────────
const VENDEDORES   = ["Ivan Diaz","Brunella","Otro"];
const ESTADOS_V    = ["Cobrado","Parcial","Pendiente","Ret. IIBB","NC","Anulado"];
const TIPOS_FC     = ["A","B","C"];
const IVA_OPTS     = ["0","2.5","5","10.5","21","27"];
const MEDIOS_C     = ["Banco Macro","Banco Comafi","Transferencia","Cheque","Efectivo"];
const ESTADOS_COT  = ["Seguimiento","Sin respuesta","Cerrado","Ganado","Perdido"];
const CATS_GASTO   = ["Logística / Flete","Impuestos","Sueldos","Alquiler","Marketing","Servicios","Bancarios","Personales","Varios"];
const MEDIOS_G     = ["Transferencia","Tarjeta","Cheque","Efectivo","Débito automático","Otro"];
const T_CHEQUE     = ["A cobrar","A pagar","En custodia","Depositado","Rechazado"];
const BANCOS       = ["Banco Macro","Banco Comafi","Banco Galicia","Banco Nación","BBVA","HSBC","Brubank","Mercado Pago","Otro"];
const MONEDAS      = ["ARS","USD","EUR"];
const T_CUENTA     = ["Cuenta corriente","Caja de ahorro","FCI","Caja efectivo","Cuenta USD","Billetera virtual","Otro"];
const T_MOV        = ["Cobro cliente","Pago proveedor","Gasto operativo","Impuesto","Sueldo","Transferencia interna","Suscripción FCI","Rescate FCI","Otro ingreso","Otro egreso"];
const EGRESOS_MOV  = ["Pago proveedor","Gasto operativo","Impuesto","Sueldo","Suscripción FCI","Otro egreso"];
const CATS_PROV    = ["Productos","Impresión","Logística","Servicios","Insumos","Otro"];
const FORMAS_P     = ["Cheque","Transferencia","Efectivo","Sasoc","Ivan Diaz","Otro"];
const ESTADOS_FC2  = ["Pendiente","Pagado","Vencido","En disputa"];
const ESTADOS_PROV = ["Activo","Inactivo","Suspendido"];
const CONDS_PAGO   = ["Contado","30 días","45 días","60 días","90 días","A convenir"];
const RUBROS       = ["Obra Social","Empresa Privada","PyME","Agro","Retail","Salud","Tecnología","Construcción","Educación","Otro"];
const ORIGENES     = ["Directo","Referido","Web","Red Social","Vendedor","Otro"];
const VERDE="#10b981"; const ROJO="#f87171"; const AZUL="#3b82f6"; const AMBAR="#f59e0b";
const COLORS=["#10b981","#3b82f6","#f59e0b","#8b5cf6","#ef4444","#06b6d4","#84cc16","#f97316"];
const TT={contentStyle:{background:"#1e293b",border:"1px solid rgba(255,255,255,.1)",borderRadius:8,color:"#e2e8f0",fontSize:12},labelStyle:{color:"#94a3b8"}};

// ─── SHARED UI ────────────────────────────────────────────────────
const IS = {background:"rgba(255,255,255,.06)",border:"1px solid rgba(255,255,255,.1)",borderRadius:8,padding:"8px 11px",color:"#e2e8f0",fontSize:13,width:"100%",boxSizing:"border-box",fontFamily:"inherit"};
const I = (p) => <input {...p} style={{...IS,...p.style}}/>;
const S = ({children,...p}) => <select {...p} style={{...IS,...p.style}}>{children}</select>;
const T = (p) => <textarea {...p} rows={p.rows||2} style={{...IS,resize:"vertical",...p.style}}/>;

function Fld({label,children,s2,hint}){
  return (
    <div style={{gridColumn:s2?"span 2":"span 1",display:"flex",flexDirection:"column",gap:4}}>
      <div style={{display:"flex",alignItems:"center",gap:6}}>
        <label style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:".5px"}}>{label}</label>
        {hint&&<span style={{fontSize:10,color:"#1e3a5f",background:"rgba(30,58,95,.3)",padding:"1px 6px",borderRadius:20}}>opcional</span>}
      </div>
      {children}
    </div>
  );
}

function Bdg({t,sm}){
  const m={Activo:["#065f46","#34d399"],Cobrado:["#065f46","#34d399"],Pagado:["#065f46","#34d399"],Depositado:["#065f46","#34d399"],Cerrado:["#065f46","#34d399"],Ganado:["#065f46","#34d399"],"A cobrar":["#065f46","#34d399"],Inactivo:["#78350f","#fcd34d"],Parcial:["#78350f","#fcd34d"],Seguimiento:["#78350f","#fcd34d"],"Sin respuesta":["#78350f","#fcd34d"],"Ret. IIBB":["#78350f","#fcd34d"],"A pagar":["#7f1d1d","#fca5a5"],NC:["#7f1d1d","#fca5a5"],Vencido:["#7f1d1d","#fca5a5"],Rechazado:["#7f1d1d","#fca5a5"],Suspendido:["#7f1d1d","#fca5a5"],Pendiente:["#1e3a5f","#93c5fd"],"En custodia":["#1e3a5f","#93c5fd"],Prospecto:["#1e293b","#94a3b8"],Perdido:["#7f1d1d","#fca5a5"],Anulado:["#1e293b","#64748b"]};
  const[bg,co]=m[t]||["#1e293b","#94a3b8"];
  return <span style={{background:bg,color:co,padding:sm?"1px 7px":"2px 9px",borderRadius:20,fontSize:sm?10:11,fontWeight:600,whiteSpace:"nowrap"}}>{t}</span>;
}

function Av({name,size=38}){
  const i=(name||"?").split(" ").slice(0,2).map(w=>w[0]).join("").toUpperCase();
  const c=["#065f46","#1e3a5f","#78350f","#4c1d95","#164e63"];
  return <div style={{width:size,height:size,borderRadius:"50%",background:c[(name?.charCodeAt(0)||0)%c.length],display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:700,fontSize:size*.35,flexShrink:0}}>{i}</div>;
}

function Mdl({title,onClose,children,wide}){
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.75)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100,padding:20}}>
      <div style={{background:"#131f35",border:"1px solid rgba(255,255,255,.1)",borderRadius:16,width:"100%",maxWidth:wide?860:640,maxHeight:"92vh",overflowY:"auto",boxShadow:"0 24px 64px rgba(0,0,0,.6)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"17px 22px",borderBottom:"1px solid rgba(255,255,255,.08)",position:"sticky",top:0,background:"#131f35",zIndex:10}}>
          <h3 style={{color:"#f1f5f9",fontSize:15,fontWeight:700,margin:0}}>{title}</h3>
          <button onClick={onClose} style={{background:"transparent",border:"none",color:"#64748b",fontSize:20,cursor:"pointer"}}>✕</button>
        </div>
        <div style={{padding:22}}>{children}</div>
      </div>
    </div>
  );
}

function Cfm({msg,onOk,onCancel}){
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.75)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:200}}>
      <div style={{background:"#131f35",border:"1px solid rgba(255,255,255,.1)",borderRadius:14,padding:28,maxWidth:360,width:"90%",textAlign:"center"}}>
        <p style={{color:"#e2e8f0",fontSize:15,marginBottom:22}}>{msg}</p>
        <div style={{display:"flex",gap:10,justifyContent:"center"}}>
          <button onClick={onCancel} style={{padding:"8px 22px",borderRadius:8,border:"1px solid rgba(255,255,255,.12)",background:"transparent",color:"#94a3b8",cursor:"pointer"}}>Cancelar</button>
          <button onClick={onOk} style={{padding:"8px 22px",borderRadius:8,border:"1px solid #ef4444",background:"#7f1d1d",color:"#fca5a5",cursor:"pointer",fontWeight:700}}>Eliminar</button>
        </div>
      </div>
    </div>
  );
}

function Tst({msg,type}){
  const ok=type!=="error";
  return <div style={{position:"fixed",bottom:28,right:28,background:ok?"#065f46":"#7f1d1d",color:ok?"#34d399":"#fca5a5",border:`1px solid ${ok?"#10b981":"#ef4444"}`,borderRadius:10,padding:"11px 18px",fontSize:13,fontWeight:700,zIndex:300,boxShadow:"0 8px 24px rgba(0,0,0,.4)"}}>{ok?"✓ ":"✕ "}{msg}</div>;
}

function Met({label,val,sub,up,dn}){
  return (
    <div style={{background:"#131f35",borderRadius:10,padding:"14px 16px",border:"1px solid rgba(255,255,255,.07)"}}>
      <p style={{fontSize:10,color:"#475569",margin:"0 0 6px",textTransform:"uppercase",letterSpacing:".6px",fontWeight:700}}>{label}</p>
      <p style={{fontSize:20,fontWeight:700,color:"#f1f5f9",margin:0}}>{val}</p>
      {sub&&<p style={{fontSize:11,margin:"4px 0 0",color:up?VERDE:dn?ROJO:"#64748b"}}>{sub}</p>}
    </div>
  );
}

function Crd({children,mb}){
  return <div style={{background:"#131f35",border:"1px solid rgba(255,255,255,.07)",borderRadius:14,padding:"18px 20px",marginBottom:mb??16}}>{children}</div>;
}

function Spin(){
  return (
    <div style={{display:"flex",alignItems:"center",justifyContent:"center",minHeight:300}}>
      <div style={{textAlign:"center"}}>
        <div style={{width:36,height:36,border:"3px solid rgba(16,185,129,.3)",borderTopColor:VERDE,borderRadius:"50%",margin:"0 auto 12px",animation:"spin .8s linear infinite"}}/>
        <p style={{color:"#64748b",fontSize:13}}>Cargando...</p>
      </div>
    </div>
  );
}

function Th({children}){return <th style={{textAlign:"left",padding:"7px 10px",fontSize:10,fontWeight:700,color:"#334155",textTransform:"uppercase",letterSpacing:".5px",borderBottom:"1px solid rgba(255,255,255,.07)",whiteSpace:"nowrap"}}>{children}</th>;}
function Td({children,g,r,mono,fw}){return <td style={{padding:"9px 10px",fontSize:13,color:g?VERDE:r?ROJO:"#cbd5e1",fontWeight:fw?600:400,fontFamily:mono?"monospace":"inherit",borderBottom:"1px solid rgba(255,255,255,.04)",whiteSpace:"nowrap"}}>{children}</td>;}

const btnPri = {padding:"8px 18px",borderRadius:8,border:"1px solid #10b981",background:"#065f46",color:"#6ee7b7",cursor:"pointer",fontSize:13,fontWeight:700,fontFamily:"inherit"};
const btnSec = {padding:"8px 16px",borderRadius:8,border:"1px solid rgba(255,255,255,.12)",background:"transparent",color:"#94a3b8",cursor:"pointer",fontSize:13,fontFamily:"inherit"};


// ─── EXPORTAR A EXCEL ─────────────────────────────────────────────
function exportExcel(data, filename, sheetName = "Datos") {
  if (!data || !data.length) return;
  const ws = XLSX.utils.json_to_sheet(data);
  // Auto column widths
  const cols = Object.keys(data[0]);
  ws['!cols'] = cols.map(k => ({
    wch: Math.max(k.length, ...data.map(r => String(r[k] ?? "").length).slice(0,50)) + 2
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, `${filename}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

function BtnExport({ onClick, label = "Exportar Excel" }) {
  return (
    <button onClick={onClick} style={{padding:"7px 14px",borderRadius:8,border:"1px solid rgba(255,255,255,.12)",background:"rgba(255,255,255,.04)",color:"#94a3b8",cursor:"pointer",fontSize:12,fontWeight:600,fontFamily:"inherit",display:"flex",alignItems:"center",gap:6}}>
      ↓ {label}
    </button>
  );
}

// ─── HOOK TOAST ───────────────────────────────────────────────────
function useToast(){
  const[t,setT]=useState(null);
  const show=(msg,type="ok")=>{setT({msg,type});setTimeout(()=>setT(null),2800);};
  return[t,show];
}

// ══════════════════════════════════════════════════════════════════
// PAGE: DASHBOARD
// ══════════════════════════════════════════════════════════════════
function PageDashboard(){
  const[data,setData]=useState(null);
  useEffect(()=>{
    Promise.all([gl(K.ventas),gl(K.gastos),gl(K.cuentas),gl(K.facsP)])
      .then(([v,g,c,fp])=>setData({v,g,c,fp}));
  },[]);
  if(!data)return(<Spin/>);
  const{v,g,c,fp}=data;
  const tv=v.reduce((a,x)=>a+pN(x.totalVenta),0);
  const tp=v.reduce((a,x)=>a+pN(x.pendiente),0);
  const tg=g.reduce((a,x)=>a+pN(x.total),0);
  const sd=c.filter(x=>x.moneda==="ARS").reduce((a,x)=>a+pN(x.saldo),0);
  const tfp=fp.reduce((a,x)=>a+pN(x.total),0);
  const pagfp=fp.filter(x=>x.estado==="Pagado").reduce((a,x)=>a+pN(x.total),0);
  const vMes={};
  v.forEach(x=>{const m=x.fecha?.slice(0,7);if(!m)return;vMes[m]=(vMes[m]||0)+pN(x.totalVenta);});
  const chart=Object.entries(vMes).sort().slice(-9).map(([m,val])=>({l:mesL(m),v:Math.round(val)}));
  const cli={};
  v.forEach(x=>{if(!x.cliente)return;cli[x.cliente]=(cli[x.cliente]||0)+pN(x.totalVenta);});
  const topC=Object.entries(cli).sort((a,b)=>b[1]-a[1]).slice(0,5);
  const alertas=fp.filter(f=>f.estado==="Pendiente"&&f.fechaVto&&f.fechaVto<hoyS).length;
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
        <Met label="Total vendido"    val={$(tv)}   sub={`${v.length} facturas`}   up/>
        <Met label="Por cobrar"       val={$(tp)}   sub="Pendiente de cobro"       dn={tp>0}/>
        <Met label="Saldo disponible" val={$(sd)}   sub="Cuentas ARS"              up/>
        <Met label="Gastos totales"   val={$(tg)}   sub={`${g.length} registros`}  dn/>
      </div>
      {alertas>0&&(
        <div style={{background:"rgba(248,113,113,.07)",border:"1px solid rgba(248,113,113,.2)",borderRadius:12,padding:"12px 16px",marginBottom:16,display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:18}}>⚠️</span>
          <p style={{color:"#fca5a5",fontSize:13,margin:0,fontWeight:600}}>{alertas} factura{alertas>1?"s":""} de proveedor vencida{alertas>1?"s":""}  — revisá el módulo de Proveedores.</p>
        </div>
      )}
      <div style={{display:"grid",gridTemplateColumns:"3fr 2fr",gap:14,marginBottom:16}}>
        <Crd mb={0}>
          <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Facturación mensual</p>
          {chart.length>0?(
            <ResponsiveContainer width="100%" height={180}>
              <BarChart data={chart} margin={{top:0,right:10,left:0,bottom:0}}>
                <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,.05)"/>
                <XAxis dataKey="l" tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false}/>
                <YAxis tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false} tickFormatter={n=>n>=1e6?"$"+(n/1e6).toFixed(1)+"M":"$"+(n/1e3).toFixed(0)+"K"}/>
                <Tooltip {...TT} formatter={n=>[$(n),"Ventas"]}/>
                <Bar dataKey="v" fill={VERDE} radius={[4,4,0,0]}/>
              </BarChart>
            </ResponsiveContainer>
          ):<p style={{color:"#334155",fontSize:13,textAlign:"center",padding:"40px 0"}}>Cargá ventas para ver el gráfico</p>}
        </Crd>
        <Crd mb={0}>
          <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Top clientes</p>
          {topC.length>0?topC.map(([nom,mon])=>(
            <div key={nom} style={{marginBottom:10}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                <span style={{fontSize:12,color:"#cbd5e1",maxWidth:130,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{nom}</span>
                <span style={{fontSize:12,fontWeight:700,color:VERDE,flexShrink:0}}>{$(mon)}</span>
              </div>
              <div style={{height:4,background:"rgba(255,255,255,.07)",borderRadius:2}}>
                <div style={{height:"100%",width:Math.round(mon/topC[0][1]*100)+"%",background:VERDE,borderRadius:2}}/>
              </div>
            </div>
          )):<p style={{color:"#334155",fontSize:13,textAlign:"center",padding:"20px 0"}}>Sin datos aún</p>}
        </Crd>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
        <Met label="Compras a proveedores" val={$(tfp)} sub={`Pagado: ${$(pagfp)}`}/>
        <Met label="Clientes únicos" val={Object.keys(cli).length} sub="Con ventas registradas"/>
        <Met label="Proveedores registrados" val={0} sub="Ver módulo Proveedores"/>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// PAGE: VENTAS
// ══════════════════════════════════════════════════════════════════
const EV={fecha:hoy(),cliente:"",cuit:"",tipo:"A",nroFactura:"",vendedor:"Ivan Diaz",montoNeto:"",iva:"21",totalVenta:"",cobrado:"",pendiente:"",medioCobro:"Banco Macro",estado:"Pendiente",productos:"",comentario:"",fechaCobro:""};

function calcTotal(f){
  const n=pN(f.montoNeto); const r=pN(f.iva)/100;
  return{totalVenta:(n*(1+r)).toFixed(2),pendiente:((n*(1+r))-pN(f.cobrado)).toFixed(2)};
}

function FormVenta({init,ventas,facsP,provs,onSave,onClose}){
  const[f,setF]=useState(init||EV);
  const[err,setErr]=useState({});
  const set=(k,v)=>setF(p=>{const n={...p,[k]:v};if(k==="montoNeto"||k==="iva"||k==="cobrado"){const t=calcTotal(n);n.totalVenta=t.totalVenta;n.pendiente=t.pendiente;}return n;});
  const sub=()=>{const e={};if(!f.cliente.trim())e.cliente="Requerido";if(!f.montoNeto||isNaN(f.montoNeto))e.montoNeto="Inválido";if(Object.keys(e).length){setErr(e);return;}onSave({...f,id:f.id||uid()});};
  const costos=f.id?facsP.filter(x=>x.ventaId===f.id):[];
  const totalCostos=costos.reduce((a,x)=>a+pN(x.total),0);
  const margen=pN(f.montoNeto)-totalCostos;
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:13}}>
        <Fld label="Cliente *" s2>{err.cliente&&<span style={{color:ROJO,fontSize:11}}>{err.cliente}</span>}<I value={f.cliente} onChange={e=>set("cliente",e.target.value)} style={err.cliente?{borderColor:"#ef4444"}:{}}/></Fld>
        <Fld label="CUIT" hint><I value={f.cuit} onChange={e=>set("cuit",e.target.value)} placeholder="30-00000000-0"/></Fld>
        <Fld label="Fecha"><I type="date" value={f.fecha} onChange={e=>set("fecha",e.target.value)}/></Fld>
        <Fld label="Tipo FC"><S value={f.tipo} onChange={e=>set("tipo",e.target.value)}>{TIPOS_FC.map(x=><option key={x}>{x}</option>)}</S></Fld>
        <Fld label="N° Factura" hint><I value={f.nroFactura} onChange={e=>set("nroFactura",e.target.value)} placeholder="0010-00001234"/></Fld>
        <Fld label="Vendedor"><S value={f.vendedor} onChange={e=>set("vendedor",e.target.value)}>{VENDEDORES.map(x=><option key={x}>{x}</option>)}</S></Fld>
        <Fld label="Monto neto s/IVA *">{err.montoNeto&&<span style={{color:ROJO,fontSize:11}}>{err.montoNeto}</span>}<I type="number" value={f.montoNeto} onChange={e=>set("montoNeto",e.target.value)} style={err.montoNeto?{borderColor:"#ef4444"}:{}}/></Fld>
        <Fld label="IVA %"><S value={f.iva} onChange={e=>set("iva",e.target.value)}>{IVA_OPTS.map(x=><option key={x} value={x}>{x}%</option>)}</S></Fld>
        <Fld label="Total (calculado)"><I value={$(f.totalVenta)} readOnly style={{opacity:.6}}/></Fld>
        <Fld label="Estado"><S value={f.estado} onChange={e=>set("estado",e.target.value)}>{ESTADOS_V.map(x=><option key={x}>{x}</option>)}</S></Fld>
        <Fld label="Cobrado"><I type="number" value={f.cobrado} onChange={e=>set("cobrado",e.target.value)}/></Fld>
        <Fld label="Pendiente (calculado)"><I value={$(f.pendiente)} readOnly style={{opacity:.6}}/></Fld>
        <Fld label="Medio de cobro"><S value={f.medioCobro} onChange={e=>set("medioCobro",e.target.value)}>{MEDIOS_C.map(x=><option key={x}>{x}</option>)}</S></Fld>
        <Fld label="Fecha de cobro" hint><I type="date" value={f.fechaCobro} onChange={e=>set("fechaCobro",e.target.value)}/></Fld>
        <Fld label="Productos / descripción" s2><T value={f.productos} onChange={e=>set("productos",e.target.value)} placeholder="Cuadernos x100, Bolígrafos x50..."/></Fld>
        <Fld label="Comentario" s2 hint><T value={f.comentario} onChange={e=>set("comentario",e.target.value)} placeholder="Retenciones, observaciones..."/></Fld>
      </div>
      {costos.length>0&&(
        <div style={{marginTop:16,padding:"12px 14px",background:"rgba(255,255,255,.03)",borderRadius:10,border:"1px solid rgba(255,255,255,.06)"}}>
          <p style={{fontSize:11,color:VERDE,fontWeight:700,textTransform:"uppercase",margin:"0 0 8px"}}>Costos asociados ({costos.length})</p>
          {costos.map((c,i)=>{const p=provs.find(x=>x.id===c.proveedorId);return<div key={i} style={{display:"flex",justifyContent:"space-between",fontSize:12,color:"#94a3b8",marginBottom:4}}><span>{p?.nombre||"Proveedor"} — {c.productos||"S/D"}</span><span style={{color:ROJO,fontWeight:600}}>{$(c.total)}</span></div>;})}
          <div style={{borderTop:"1px solid rgba(255,255,255,.08)",marginTop:8,paddingTop:8,display:"flex",justifyContent:"space-between"}}>
            <span style={{fontSize:13,color:"#e2e8f0",fontWeight:600}}>Margen estimado</span>
            <span style={{fontSize:13,fontWeight:700,color:margen>=0?VERDE:ROJO}}>{$(margen)} ({pN(f.montoNeto)?pct(margen/pN(f.montoNeto)*100):"—"})</span>
          </div>
        </div>
      )}
      <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:22,borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:18}}>
        <button style={btnSec} onClick={onClose}>Cancelar</button>
        <button style={btnPri} onClick={sub}>Guardar venta</button>
      </div>
    </div>
  );
}

function PageVentas(){
  const[ventas,setVentas]=useState([]);const[facsP,setFacsP]=useState([]);const[provs,setProvs]=useState([]);
  const[loading,setLoading]=useState(true);const[modal,setModal]=useState(null);const[sel,setSel]=useState(null);
  const[cfm,setCfm]=useState(null);const[search,setSearch]=useState("");const[fEst,setFEst]=useState("Todos");
  const[fVnd,setFVnd]=useState("Todos");const[tab,setTab]=useState("lista");const[toast,showToast]=useToast();
  useEffect(()=>{Promise.all([gl(K.ventas),gl(K.facsP),gl(K.proveedores)]).then(([v,f,p])=>{setVentas(v);setFacsP(f);setProvs(p);setLoading(false);});},[]);
  const save=async(v)=>{const n=ventas.some(x=>x.id===v.id)?ventas.map(x=>x.id===v.id?v:x):[v,...ventas];setVentas(n);await sl(K.ventas,n);setModal(null);showToast("Venta guardada");};
  const del=async(id)=>{const n=ventas.filter(x=>x.id!==id);setVentas(n);await sl(K.ventas,n);setModal(null);setCfm(null);showToast("Venta eliminada");};
  const tv=ventas.reduce((a,v)=>a+pN(v.totalVenta),0);
  const tc=ventas.reduce((a,v)=>a+pN(v.cobrado),0);
  const tp=ventas.reduce((a,v)=>a+pN(v.pendiente),0);
  const filt=ventas.filter(v=>{const q=search.toLowerCase();const mQ=!search||v.cliente?.toLowerCase().includes(q)||v.nroFactura?.includes(q)||v.cuit?.includes(q);const mE=fEst==="Todos"||v.estado===fEst;const mV=fVnd==="Todos"||v.vendedor===fVnd;return mQ&&mE&&mV;}).sort((a,b)=>(b.fecha||"").localeCompare(a.fecha||""));
  const vMes={};ventas.forEach(v=>{const m=v.fecha?.slice(0,7);if(!m)return;if(!vMes[m])vMes[m]={mes:m,vendido:0,cobrado:0,cant:0};vMes[m].vendido+=pN(v.totalVenta);vMes[m].cobrado+=pN(v.cobrado);vMes[m].cant++;});
  const histData=(()=>{let a=0;return Object.values(vMes).sort((x,y)=>x.mes.localeCompare(y.mes)).map(d=>{a+=d.vendido;return{...d,l:mesL(d.mes),acum:Math.round(a)};});})();
  const prom=histData.length?Math.round(tv/histData.length):0;
  if(loading)return(<Spin/>);
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
        <Met label="Total vendido" val={$(tv)} sub={`${ventas.length} facturas`} up/>
        <Met label="Total cobrado" val={$(tc)} sub={pct(tv?tc/tv*100:0)+" cobrado"} up/>
        <Met label="Por cobrar" val={$(tp)} sub="Pendiente" dn={tp>0}/>
        <Met label="Ticket promedio" val={$(ventas.length?tv/ventas.length:0)} sub="Por factura"/>
      </div>
      <div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap"}}>
        {["lista","estadisticas"].map(t=><button key={t} onClick={()=>setTab(t)} style={{padding:"7px 16px",borderRadius:8,border:"none",background:tab===t?"rgba(16,185,129,.12)":"rgba(255,255,255,.04)",color:tab===t?VERDE:"#64748b",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>{t==="lista"?"≡ Lista":"∿ Historial"}</button>)}
        <div style={{marginLeft:"auto",display:"flex",gap:8}}>
          <BtnExport onClick={()=>exportExcel(ventas.map(v=>({Fecha:v.fecha,Cliente:v.cliente,CUIT:v.cuit,"N° Factura":v.nroFactura,Tipo:v.tipo,Vendedor:v.vendedor,"Monto Neto":v.montoNeto,IVA:v.iva+"%" ,"Total":v.totalVenta,Cobrado:v.cobrado,Pendiente:v.pendiente,Estado:v.estado,Productos:v.productos})),"ventas_sasoc","Ventas")}/>
          <button style={btnPri} onClick={()=>{setSel(null);setModal("form");}}>+ Nueva venta</button>
        </div>
      </div>
      {tab==="lista"&&(
        <Crd>
          <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap"}}>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar cliente, factura, CUIT..." style={{...IS,flex:1,minWidth:200}}/>
            <select value={fEst} onChange={e=>setFEst(e.target.value)} style={{...IS,width:150}}><option value="Todos">Todos los estados</option>{ESTADOS_V.map(x=><option key={x}>{x}</option>)}</select>
            <select value={fVnd} onChange={e=>setFVnd(e.target.value)} style={{...IS,width:140}}><option value="Todos">Todos</option>{VENDEDORES.map(x=><option key={x}>{x}</option>)}</select>
          </div>
          {ventas.length===0?(<div style={{textAlign:"center",padding:"48px 0"}}><p style={{fontSize:32,marginBottom:10}}>📋</p><p style={{color:"#334155",fontSize:14,fontWeight:600}}>No hay ventas cargadas</p><button style={{...btnPri,marginTop:12}} onClick={()=>setModal("form")}>+ Cargar primera venta</button></div>):(
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr><Th>Fecha</Th><Th>Cliente</Th><Th>N° Factura</Th><Th>Vendedor</Th><Th>Total</Th><Th>Cobrado</Th><Th>Pendiente</Th><Th>Estado</Th><Th/></tr></thead>
                <tbody>{filt.map(v=>(
                  <tr key={v.id} style={{cursor:"pointer"}} onClick={()=>{setSel(v);setModal("det");}}>
                    <Td>{fecha(v.fecha)}</Td><Td fw>{v.cliente}</Td><Td mono>{v.nroFactura||"—"}</Td><Td>{v.vendedor}</Td>
                    <Td fw>{$(v.totalVenta)}</Td><Td g fw>{$(v.cobrado)}</Td><Td r={pN(v.pendiente)>0} fw>{$(v.pendiente)}</Td>
                    <Td><Bdg t={v.estado} sm/></Td>
                    <td style={{padding:"9px 8px",borderBottom:"1px solid rgba(255,255,255,.04)"}} onClick={e=>e.stopPropagation()}>
                      <div className="ra" style={{display:"flex",gap:4}}>
                        <button className="ib" onClick={()=>{setSel(v);setModal("form");}}>✏</button>
                        <button className="ib" style={{color:ROJO}} onClick={()=>setCfm(v.id)}>🗑</button>
                      </div>
                    </td>
                  </tr>
                ))}</tbody>
              </table>
              <p style={{fontSize:11,color:"#334155",margin:"10px 4px 0",textAlign:"right"}}>{filt.length} de {ventas.length} registros</p>
            </div>
          )}
        </Crd>
      )}
      {tab==="estadisticas"&&histData.length>0&&(
        <Crd>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:10}}>
            <div><p style={{fontSize:14,fontWeight:700,color:"#e2e8f0",margin:"0 0 2px"}}>Facturación histórica mensual</p><p style={{fontSize:11,color:"#334155",margin:0}}>Barras = mes · Línea azul = acumulado · 🏆 mejor mes</p></div>
            <div style={{display:"flex",gap:20}}><div style={{textAlign:"right"}}><p style={{fontSize:10,color:"#475569",margin:"0 0 2px",textTransform:"uppercase",fontWeight:700}}>Promedio mensual</p><p style={{fontSize:15,fontWeight:700,color:"#f1f5f9",margin:0}}>{$(prom)}</p></div></div>
          </div>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={histData} margin={{top:10,right:55,left:0,bottom:0}}>
              <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,.05)"/>
              <XAxis dataKey="l" tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false}/>
              <YAxis yAxisId="m" orientation="left" tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false} tickFormatter={n=>n>=1e6?"$"+(n/1e6).toFixed(1)+"M":"$"+(n/1e3).toFixed(0)+"K"}/>
              <YAxis yAxisId="a" orientation="right" tick={{fill:"#334155",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={n=>n>=1e6?"$"+(n/1e6).toFixed(1)+"M":"$"+(n/1e3).toFixed(0)+"K"}/>
              <Tooltip {...TT} formatter={(v,name)=>[$(v),name==="vendido"?"Facturación":"Acumulado"]}/>
              <ReferenceLine yAxisId="m" y={prom} stroke="rgba(245,158,11,.45)" strokeDasharray="6 3"/>
              <Bar yAxisId="m" dataKey="vendido" name="vendido" radius={[5,5,0,0]}>
                {histData.map((d,i)=><Cell key={i} fill={i===histData.length-1?"#34d399":d.vendido===Math.max(...histData.map(x=>x.vendido))?"#fbbf24":VERDE}/>)}
              </Bar>
              <Line yAxisId="a" type="monotone" dataKey="acum" name="acum" stroke={AZUL} strokeWidth={2} dot={false}/>
            </BarChart>
          </ResponsiveContainer>
          <div style={{marginTop:16,overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr><Th>Mes</Th><Th>Facturas</Th><Th>Facturación</Th><Th>Cobrado</Th><Th>% cobrado</Th><Th>vs promedio</Th><Th>Acumulado</Th></tr></thead>
              <tbody>{histData.map((d,i)=>{const vp=prom?((d.vendido-prom)/prom*100):0;const mejor=d.vendido===Math.max(...histData.map(x=>x.vendido));return(
                <tr key={i} style={{background:mejor?"rgba(251,191,36,.04)":"transparent"}}>
                  <Td fw={mejor}>{d.l} {mejor?"🏆":""}</Td><Td>{d.cant}</Td><Td g fw>{$(d.vendido)}</Td><Td>{$(d.cobrado)}</Td>
                  <Td><span style={{background:d.vendido&&d.cobrado/d.vendido>.9?"rgba(16,185,129,.12)":"rgba(245,158,11,.1)",color:d.vendido&&d.cobrado/d.vendido>.9?VERDE:AMBAR,padding:"1px 7px",borderRadius:20,fontSize:11,fontWeight:600}}>{d.vendido?pct(d.cobrado/d.vendido*100):"—"}</span></Td>
                  <Td><span style={{color:vp>=0?VERDE:ROJO,fontWeight:600}}>{vp>=0?"↑":"↓"}{Math.abs(vp).toFixed(1)}%</span></Td>
                  <Td><span style={{color:AZUL,fontWeight:500}}>{$(d.acum)}</span></Td>
                </tr>
              );})}
              </tbody>
              <tfoot><tr style={{background:"rgba(16,185,129,.05)"}}>
                <td colSpan={1} style={{padding:"8px 10px",fontSize:13,color:VERDE,fontWeight:700}}>TOTAL</td>
                <td style={{padding:"8px 10px",fontSize:13,color:VERDE,fontWeight:700}}>{ventas.length}</td>
                <td style={{padding:"8px 10px",fontSize:13,color:VERDE,fontWeight:700}}>{$(tv)}</td>
                <td style={{padding:"8px 10px",fontSize:13,color:VERDE,fontWeight:700}}>{$(tc)}</td>
                <td style={{padding:"8px 10px",fontSize:13,color:VERDE,fontWeight:700}}>{pct(tv?tc/tv*100:0)}</td>
                <td style={{padding:"8px 10px",fontSize:12,color:"#64748b"}}>{$(prom)}/mes</td>
                <td style={{padding:"8px 10px",fontSize:13,color:AZUL,fontWeight:700}}>{$(tv)}</td>
              </tr></tfoot>
            </table>
          </div>
        </Crd>
      )}
      {modal==="form"&&(<Mdl title={sel?"Editar venta":"Nueva venta"} onClose={()=>setModal(null)}><FormVenta init={sel} ventas={ventas} facsP={facsP} provs={provs} onSave={save} onClose={()=>setModal(null)}/></Mdl>)}
      {modal==="det"&&sel&&(<Mdl title={sel.cliente} onClose={()=>setModal(null)}><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"6px 24px"}}>{[["Fecha",fecha(sel.fecha)],["N° Factura",sel.nroFactura||"—"],["Tipo",sel.tipo],["Vendedor",sel.vendedor],["Monto neto",$(sel.montoNeto)],["IVA",sel.iva+"%"],["Total",$(sel.totalVenta)],["Cobrado",$(sel.cobrado)],["Pendiente",$(sel.pendiente)],["Medio cobro",sel.medioCobro]].map(([l,v])=><div key={l} style={{padding:"7px 0",borderBottom:"1px solid rgba(255,255,255,.05)"}}><p style={{fontSize:11,color:"#475569",margin:0}}>{l}</p><p style={{fontSize:13,color:"#e2e8f0",margin:0,fontWeight:500}}>{v}</p></div>)}</div>{sel.productos&&<div style={{marginTop:10,padding:"10px",background:"rgba(255,255,255,.03)",borderRadius:8}}><p style={{fontSize:11,color:"#475569",margin:"0 0 4px"}}>PRODUCTOS</p><p style={{fontSize:13,color:"#cbd5e1",margin:0}}>{sel.productos}</p></div>}<div style={{display:"flex",justifyContent:"space-between",marginTop:16}}><button style={{...btnSec,color:ROJO,borderColor:"rgba(239,68,68,.3)"}} onClick={()=>setCfm(sel.id)}>🗑 Eliminar</button><button style={btnPri} onClick={()=>setModal("form")}>✏ Editar</button></div></Mdl>)}
      {cfm&&(<Cfm msg="¿Eliminar esta venta?" onOk={()=>del(cfm)} onCancel={()=>setCfm(null)}/>)}
      {toast&&<Tst msg={toast.msg} type={toast.type}/>}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// PAGE: CLIENTES
// ══════════════════════════════════════════════════════════════════
function PageClientes(){
  const[ventas,setVentas]=useState([]);const[extras,setExtras]=useState({});const[loading,setLoading]=useState(true);
  const[modal,setModal]=useState(null);const[sel,setSel]=useState(null);const[search,setSearch]=useState("");
  const[cards,setCards]=useState(false);const[toast,showToast]=useToast();
  useEffect(()=>{Promise.all([gl(K.ventas),gl(K.clientes)]).then(([v,c])=>{setVentas(v);const ex=Array.isArray(c)?{}:c;setExtras(ex||{});setLoading(false);});},[]);
  const saveExtra=async(empresa,data)=>{const n={...extras,[empresa]:data};setExtras(n);await sl(K.clientes,n);setModal(null);showToast("Datos guardados");};
  const empresas=[...new Set(ventas.map(v=>v.cliente?.trim()).filter(Boolean))];
  const clientes=empresas.map(emp=>{const cvs=ventas.filter(v=>v.cliente?.trim()===emp);const tv=cvs.reduce((a,v)=>a+pN(v.totalVenta),0);const tp=cvs.reduce((a,v)=>a+pN(v.pendiente),0);return{empresa:emp,extra:extras[emp]||null,cvs:cvs.length,tv,tp};}).sort((a,b)=>b.tv-a.tv);
  const filt=clientes.filter(c=>{const q=search.toLowerCase();return!search||c.empresa.toLowerCase().includes(q)||c.extra?.contacto?.toLowerCase().includes(q)||c.extra?.tel?.includes(q);});
  const totalV=clientes.reduce((a,c)=>a+c.tv,0);const totalP=clientes.reduce((a,c)=>a+c.tp,0);
  if(loading)return(<Spin/>);
  const EF={contacto:"",cargo:"",tel:"",email:"",ciudad:"",rubro:"Otro",origen:"Directo",estado:"Activo",notas:""};
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
        <Met label="Clientes detectados" val={clientes.length} sub="Desde ventas" up/>
        <Met label="Total facturado" val={$(totalV)} sub="A todos los clientes" up/>
        <Met label="Pendiente de cobro" val={$(totalP)} sub="Por recuperar" dn={totalP>0}/>
        <Met label="Con datos de contacto" val={clientes.filter(c=>c.extra?.contacto||c.extra?.tel).length} sub={`${clientes.filter(c=>!c.extra?.contacto&&!c.extra?.tel).length} sin completar`}/>
      </div>
      <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar empresa, contacto o teléfono..." style={{...IS,flex:1,minWidth:200}}/>
        <BtnExport onClick={()=>exportExcel(clientes.map(c=>({Empresa:c.empresa,Contacto:c.extra?.contacto||"",Teléfono:c.extra?.tel||"",Email:c.extra?.email||"",Ciudad:c.extra?.ciudad||"",Rubro:c.extra?.rubro||"","Total Facturado":c.tv,"Pendiente":c.tp,Facturas:c.cvs,Estado:c.extra?.estado||"Activo"})),"clientes_sasoc","Clientes")}/>
        <div style={{display:"flex",gap:6}}>
          {["lista","cards"].map(t=><button key={t} onClick={()=>setCards(t==="cards")} style={{padding:"7px 14px",borderRadius:8,border:"none",background:(t==="cards")===cards?"rgba(16,185,129,.12)":"rgba(255,255,255,.04)",color:(t==="cards")===cards?VERDE:"#64748b",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>{t==="lista"?"≡":"⊞"}</button>)}
        </div>
      </div>
      {clientes.length===0?(<Crd><div style={{textAlign:"center",padding:"48px 0"}}><p style={{fontSize:32,marginBottom:10}}>👥</p><p style={{color:"#334155",fontSize:14}}>Los clientes aparecen automáticamente al cargar ventas.</p></div></Crd>):!cards?(
        <Crd>
          <div style={{overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr><Th>Empresa</Th><Th>Contacto</Th><Th>Teléfono</Th><Th>Ciudad</Th><Th>Facturas</Th><Th>Total facturado</Th><Th>Pendiente</Th><Th>Estado</Th><Th/></tr></thead>
              <tbody>{filt.map(c=>(
                <tr key={c.empresa} style={{cursor:"pointer"}} onClick={()=>{setSel(c);setModal("det");}}>
                  <Td><div style={{display:"flex",alignItems:"center",gap:8}}><Av name={c.empresa} size={28}/><span style={{fontWeight:600,color:"#f1f5f9"}}>{c.empresa}</span></div></Td>
                  <Td>{c.extra?.contacto||<span style={{color:"#334155",fontSize:11}}>sin datos</span>}</Td>
                  <Td mono>{c.extra?.tel||"—"}</Td><Td>{c.extra?.ciudad||"—"}</Td>
                  <Td>{c.cvs}</Td><Td g fw>{$(c.tv)}</Td><Td r={c.tp>0} fw>{$(c.tp)}</Td>
                  <Td><Bdg t={c.extra?.estado||"Activo"} sm/></Td>
                  <td style={{padding:"9px 8px",borderBottom:"1px solid rgba(255,255,255,.04)"}} onClick={e=>e.stopPropagation()}>
                    <div className="ra"><button className="ib" onClick={()=>{setSel(c);setModal("edit");}}>✏</button></div>
                  </td>
                </tr>
              ))}</tbody>
            </table>
          </div>
        </Crd>
      ):(
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(250px,1fr))",gap:14}}>
          {filt.map(c=>(
            <div key={c.empresa} className="pcard" onClick={()=>{setSel(c);setModal("det");}}>
              <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:10}}>
                <div style={{display:"flex",alignItems:"center",gap:10}}><Av name={c.empresa} size={38}/><div><p style={{fontSize:13,fontWeight:700,color:"#f1f5f9",margin:0}}>{c.empresa}</p><p style={{fontSize:11,color:"#475569",margin:"2px 0 0"}}>{c.extra?.rubro||"Sin rubro"}</p></div></div>
                <Bdg t={c.extra?.estado||"Activo"} sm/>
              </div>
              {c.extra?.contacto&&<p style={{fontSize:12,color:"#64748b",marginBottom:3}}>👤 {c.extra.contacto}</p>}
              {c.extra?.tel&&<p style={{fontSize:12,color:"#64748b",marginBottom:3}}>📞 {c.extra.tel}</p>}
              <div style={{borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:10,marginTop:8,display:"flex",justifyContent:"space-between"}}>
                <div><p style={{fontSize:10,color:"#475569",margin:0,textTransform:"uppercase",fontWeight:700}}>Facturado</p><p style={{fontSize:14,fontWeight:700,color:VERDE,margin:0}}>{$(c.tv)}</p></div>
                <div style={{textAlign:"right"}}><p style={{fontSize:10,color:"#475569",margin:0,textTransform:"uppercase",fontWeight:700}}>Facturas</p><p style={{fontSize:14,fontWeight:700,color:"#94a3b8",margin:0}}>{c.cvs}</p></div>
              </div>
              {c.tp>0&&<div style={{marginTop:8,padding:"5px 10px",background:"rgba(248,113,113,.08)",borderRadius:7,display:"flex",justifyContent:"space-between"}}><span style={{fontSize:11,color:ROJO}}>Pendiente</span><span style={{fontSize:12,fontWeight:700,color:ROJO}}>{$(c.tp)}</span></div>}
            </div>
          ))}
        </div>
      )}
      {modal==="det"&&sel&&(
        <Mdl title="Ficha de cliente" onClose={()=>setModal(null)} wide>
          <div style={{display:"flex",alignItems:"center",gap:16,marginBottom:20,padding:"16px",background:"rgba(255,255,255,.03)",borderRadius:12,border:"1px solid rgba(255,255,255,.07)"}}>
            <Av name={sel.empresa} size={52}/><div style={{flex:1}}><h2 style={{color:"#f1f5f9",fontSize:17,fontWeight:700,margin:"0 0 4px"}}>{sel.empresa}</h2>{sel.extra?.rubro&&<p style={{color:"#64748b",fontSize:12,margin:0}}>{sel.extra.rubro}{sel.extra?.ciudad?` · ${sel.extra.ciudad}`:""}</p>}</div>
            <button style={btnPri} onClick={()=>setModal("edit")}>✏ {sel.extra?"Editar":"Agregar datos"}</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:16}}>
            <Met label="Total facturado" val={$(sel.tv)} up/><Met label="Pendiente" val={$(sel.tp)} dn={sel.tp>0}/><Met label="Facturas" val={sel.cvs}/>
          </div>
          {sel.extra&&<div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
            {[["👤 Contacto",sel.extra.contacto],["💼 Cargo",sel.extra.cargo],["📞 Teléfono",sel.extra.tel],["✉️ Email",sel.extra.email],["📍 Ciudad",sel.extra.ciudad],["🎯 Origen",sel.extra.origen]].filter(([,v])=>v).map(([l,v])=><div key={l} style={{padding:"8px 0",borderBottom:"1px solid rgba(255,255,255,.05)"}}><p style={{fontSize:11,color:"#475569",margin:0}}>{l}</p><p style={{fontSize:13,color:"#e2e8f0",margin:0,fontWeight:500}}>{v}</p></div>)}
          </div>}
          {sel.extra?.notas&&<div style={{marginBottom:16,padding:"10px",background:"rgba(245,158,11,.06)",border:"1px solid rgba(245,158,11,.15)",borderRadius:10}}><p style={{fontSize:11,color:"#78350f",margin:"0 0 4px"}}>NOTAS</p><p style={{fontSize:13,color:"#fcd34d",margin:0}}>{sel.extra.notas}</p></div>}
          <p style={{fontSize:11,color:VERDE,fontWeight:700,textTransform:"uppercase",marginBottom:10}}>Historial de ventas ({ventas.filter(v=>v.cliente?.trim()===sel.empresa).length})</p>
          <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr><Th>Fecha</Th><Th>N° Factura</Th><Th>Total</Th><Th>Cobrado</Th><Th>Estado</Th></tr></thead><tbody>{ventas.filter(v=>v.cliente?.trim()===sel.empresa).sort((a,b)=>(b.fecha||"").localeCompare(a.fecha||"")).map((v,i)=><tr key={i}><Td>{fecha(v.fecha)}</Td><Td mono>{v.nroFactura||"—"}</Td><Td fw>{$(v.totalVenta)}</Td><Td g>{$(v.cobrado)}</Td><Td><Bdg t={v.estado} sm/></Td></tr>)}</tbody></table></div>
        </Mdl>
      )}
      {modal==="edit"&&sel&&(
        <Mdl title="Datos de contacto" onClose={()=>setModal(null)}>
          <div style={{padding:"12px 16px",background:"rgba(16,185,129,.06)",border:"1px solid rgba(16,185,129,.15)",borderRadius:10,marginBottom:18,display:"flex",alignItems:"center",gap:12}}>
            <Av name={sel.empresa} size={40}/><div><p style={{color:"#f1f5f9",fontWeight:700,fontSize:14,margin:0}}>{sel.empresa}</p><p style={{color:"#64748b",fontSize:12,margin:"3px 0 0"}}>Todos los campos son opcionales</p></div>
          </div>
          {(()=>{const[f,setF]=useState(sel.extra||EF);const set=(k,v)=>setF(p=>({...p,[k]:v}));return(
            <div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:13}}>
                <Fld label="Contacto" hint><I value={f.contacto} onChange={e=>set("contacto",e.target.value)} placeholder="Nombre apellido"/></Fld>
                <Fld label="Cargo" hint><I value={f.cargo} onChange={e=>set("cargo",e.target.value)} placeholder="Gerente, Compras..."/></Fld>
                <Fld label="Teléfono" hint><I value={f.tel} onChange={e=>set("tel",e.target.value)} placeholder="11 1234-5678"/></Fld>
                <Fld label="Email" hint><I type="email" value={f.email} onChange={e=>set("email",e.target.value)} placeholder="mail@empresa.com"/></Fld>
                <Fld label="Ciudad" hint><I value={f.ciudad} onChange={e=>set("ciudad",e.target.value)} placeholder="Buenos Aires..."/></Fld>
                <Fld label="Rubro" hint><S value={f.rubro} onChange={e=>set("rubro",e.target.value)}>{RUBROS.map(x=><option key={x}>{x}</option>)}</S></Fld>
                <Fld label="Origen" hint><S value={f.origen} onChange={e=>set("origen",e.target.value)}>{ORIGENES.map(x=><option key={x}>{x}</option>)}</S></Fld>
                <Fld label="Estado" hint><S value={f.estado} onChange={e=>set("estado",e.target.value)}>{"Activo,Inactivo,Prospecto,Suspendido".split(",").map(x=><option key={x}>{x}</option>)}</S></Fld>
                <Fld label="Notas" hint s2><T value={f.notas} onChange={e=>set("notas",e.target.value)} placeholder="Observaciones..."/></Fld>
              </div>
              <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:22,borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:18}}>
                <button style={btnSec} onClick={()=>setModal("det")}>Cancelar</button>
                <button style={btnPri} onClick={()=>saveExtra(sel.empresa,f)}>Guardar</button>
              </div>
            </div>
          );})()}
        </Mdl>
      )}
      {toast&&<Tst msg={toast.msg} type={toast.type}/>}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// PAGE: PROVEEDORES
// ══════════════════════════════════════════════════════════════════
const EP={nombre:"",cuit:"",contacto:"",tel:"",email:"",categoria:"Productos",condPago:"60 días",estado:"Activo",notas:""};
const EFP={proveedorId:"",nroFactura:"",fecha:hoy(),fechaVto:"",montoNeto:"",iva:"21",total:"",formaPago:"Cheque",nroCheque:"",estado:"Pendiente",productos:"",notas:"",ventaId:""};

function PageProveedores(){
  const[provs,setProvs]=useState([]);const[facs,setFacs]=useState([]);const[ventas,setVentas]=useState([]);
  const[loading,setLoading]=useState(true);const[modal,setModal]=useState(null);const[sel,setSel]=useState(null);
  const[selFc,setSelFc]=useState(null);const[cfm,setCfm]=useState(null);const[search,setSearch]=useState("");
  const[tab,setTab]=useState("provs");const[toast,showToast]=useToast();
  useEffect(()=>{Promise.all([gl(K.proveedores),gl(K.facsP),gl(K.ventas)]).then(([p,f,v])=>{setProvs(p);setFacs(f);setVentas(v);setLoading(false);});},[]);
  const saveP=async(p)=>{const n=provs.some(x=>x.id===p.id)?provs.map(x=>x.id===p.id?p:x):[p,...provs];setProvs(n);await sl(K.proveedores,n);setModal(null);showToast("Proveedor guardado");};
  const delP=async(id)=>{const np=provs.filter(x=>x.id!==id);const nf=facs.filter(x=>x.proveedorId!==id);setProvs(np);setFacs(nf);await Promise.all([sl(K.proveedores,np),sl(K.facsP,nf)]);setModal(null);setCfm(null);showToast("Proveedor eliminado");};
  const saveF=async(f)=>{const n=facs.some(x=>x.id===f.id)?facs.map(x=>x.id===f.id?f:x):[f,...facs];setFacs(n);await sl(K.facsP,n);setModal(null);showToast("Factura guardada");};
  const delF=async(id)=>{const n=facs.filter(x=>x.id!==id);setFacs(n);await sl(K.facsP,n);setCfm(null);showToast("Factura eliminada");};
  const pStats=provs.map(p=>{const pf=facs.filter(f=>f.proveedorId===p.id);return{...p,_fcs:pf.length,_tot:pf.reduce((a,f)=>a+pN(f.total),0),_pend:pf.filter(f=>f.estado!=="Pagado").reduce((a,f)=>a+pN(f.total),0)};}).sort((a,b)=>b._tot-a._tot);
  const totComp=facs.reduce((a,f)=>a+pN(f.total),0);const totPag=facs.filter(f=>f.estado==="Pagado").reduce((a,f)=>a+pN(f.total),0);
  const venc=facs.filter(f=>f.estado==="Pendiente"&&f.fechaVto&&f.fechaVto<hoyS).length;
  const filtP=pStats.filter(p=>{const q=search.toLowerCase();return!search||p.nombre?.toLowerCase().includes(q)||p.cuit?.includes(q);});
  const filtF=facs.map(f=>({...f,_p:provs.find(p=>p.id===f.proveedorId)})).filter(f=>!search||f._p?.nombre?.toLowerCase().includes(search.toLowerCase())||f.nroFactura?.includes(search)).sort((a,b)=>(b.fecha||"").localeCompare(a.fecha||""));
  if(loading)return(<Spin/>);
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
        <Met label="Proveedores activos" val={provs.filter(p=>p.estado==="Activo").length} sub={`${provs.length} en total`}/>
        <Met label="Total comprado" val={$(totComp)} up/><Met label="Total pendiente" val={$(totComp-totPag)} dn={totComp>totPag}/>
        <Met label="Facturas vencidas" val={venc} sub="Requieren atención" dn={venc>0}/>
      </div>
      {venc>0&&<div style={{background:"rgba(248,113,113,.07)",border:"1px solid rgba(248,113,113,.2)",borderRadius:12,padding:"12px 16px",marginBottom:16,display:"flex",alignItems:"center",gap:10}}><span style={{fontSize:18}}>⚠️</span><p style={{color:"#fca5a5",fontSize:13,margin:0,fontWeight:600}}>{venc} factura{venc>1?"s":""} vencida{venc>1?"s":""}</p></div>}
      <div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap"}}>
        {["provs","facs"].map(t=><button key={t} onClick={()=>setTab(t)} style={{padding:"7px 16px",borderRadius:8,border:"none",background:tab===t?"rgba(16,185,129,.12)":"rgba(255,255,255,.04)",color:tab===t?VERDE:"#64748b",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>{t==="provs"?`🏭 Proveedores (${provs.length})`:`📄 Facturas (${facs.length})`}</button>)}
        <div style={{marginLeft:"auto",display:"flex",gap:8}}>
          <BtnExport onClick={()=>{if(tab==="provs")exportExcel(provs.map(p=>({Proveedor:p.nombre,CUIT:p.cuit,Categoría:p.categoria,"Cond. Pago":p.condPago,Contacto:p.contacto,Teléfono:p.tel,Estado:p.estado})),"proveedores_sasoc","Proveedores");else exportExcel(facs.map(f=>({Fecha:f.fecha,Proveedor:provs.find(p=>p.id===f.proveedorId)?.nombre||"","N° Factura":f.nroFactura,"Monto Neto":f.montoNeto,IVA:f.iva+"%",Total:f.total,"Forma Pago":f.formaPago,Estado:f.estado,"Fecha Vto":f.fechaVto})),"facturas_prov_sasoc","Facturas Prov");}}/>
          <button style={{...btnPri,background:"rgba(16,185,129,.08)",borderColor:"rgba(16,185,129,.4)",color:"#34d399"}} onClick={()=>{setSelFc(null);setModal("fc");}}>+ Factura</button>
          <button style={btnPri} onClick={()=>{setSel(null);setModal("prov");}}>+ Proveedor</button>
        </div>
      </div>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar..." style={{...IS,marginBottom:14}}/>
      {tab==="provs"&&(
        <Crd>
          {provs.length===0?(<div style={{textAlign:"center",padding:"48px 0"}}><p style={{fontSize:32,marginBottom:10}}>🏭</p><p style={{color:"#334155",fontSize:14,fontWeight:600}}>No hay proveedores cargados</p><button style={{...btnPri,marginTop:12}} onClick={()=>setModal("prov")}>+ Agregar proveedor</button></div>):(
            <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr><Th>Proveedor</Th><Th>CUIT</Th><Th>Categoría</Th><Th>Cond. pago</Th><Th>Facturas</Th><Th>Total comprado</Th><Th>Pendiente</Th><Th>Estado</Th><Th/></tr></thead>
              <tbody>{filtP.map(p=>(
                <tr key={p.id} style={{cursor:"pointer"}} onClick={()=>{setSel(p);setModal("detP");}}>
                  <Td><div style={{display:"flex",alignItems:"center",gap:8}}><Av name={p.nombre} size={28}/><span style={{fontWeight:600,color:"#f1f5f9"}}>{p.nombre}</span></div></Td>
                  <Td mono>{p.cuit||"—"}</Td><Td>{p.categoria}</Td><Td>{p.condPago}</Td><Td>{p._fcs}</Td>
                  <Td r fw>{$(p._tot)}</Td><Td r={p._pend>0} fw>{$(p._pend)}</Td><Td><Bdg t={p.estado} sm/></Td>
                  <td style={{padding:"9px 8px",borderBottom:"1px solid rgba(255,255,255,.04)"}} onClick={e=>e.stopPropagation()}>
                    <div className="ra" style={{display:"flex",gap:4}}>
                      <button className="ib" onClick={()=>{setSel(p);setModal("prov");}}>✏</button>
                      <button className="ib" style={{color:ROJO}} onClick={()=>setCfm({t:"prov",id:p.id})}>🗑</button>
                    </div>
                  </td>
                </tr>
              ))}</tbody>
            </table></div>
          )}
        </Crd>
      )}
      {tab==="facs"&&(
        <Crd>
          {facs.length===0?(<div style={{textAlign:"center",padding:"48px 0"}}><p style={{fontSize:32,marginBottom:10}}>📄</p><p style={{color:"#334155",fontSize:14,fontWeight:600}}>No hay facturas de proveedor</p><button style={{...btnPri,marginTop:12}} onClick={()=>setModal("fc")}>+ Cargar factura</button></div>):(
            <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr><Th>Fecha</Th><Th>Proveedor</Th><Th>N° Factura</Th><Th>Total</Th><Th>Forma pago</Th><Th>Venta asociada</Th><Th>Vence</Th><Th>Estado</Th><Th/></tr></thead>
              <tbody>{filtF.map(f=>{const v=f.ventaId?ventas.find(x=>x.id===f.ventaId):null;const venc2=f.estado==="Pendiente"&&f.fechaVto&&f.fechaVto<hoyS;return(
                <tr key={f.id}>
                  <Td>{fecha(f.fecha)}</Td><Td fw>{f._p?.nombre||"—"}</Td><Td mono>{f.nroFactura||"—"}</Td>
                  <Td r fw>{$(f.total)}</Td><Td>{f.formaPago}{f.nroCheque?` #${f.nroCheque}`:""}</Td>
                  <Td>{v?<span style={{background:"rgba(16,185,129,.1)",color:VERDE,padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:600}}>{v.cliente}</span>:<span style={{color:"#334155",fontSize:11}}>—</span>}</Td>
                  <Td><span style={{color:venc2?ROJO:"#64748b"}}>{fecha(f.fechaVto)}{venc2?" ⚠":""}</span></Td>
                  <Td><Bdg t={venc2?"Vencido":f.estado} sm/></Td>
                  <td style={{padding:"9px 8px",borderBottom:"1px solid rgba(255,255,255,.04)"}}>
                    <div className="ra" style={{display:"flex",gap:4}}>
                      <button className="ib" onClick={()=>{setSelFc(f);setModal("fc");}}>✏</button>
                      <button className="ib" style={{color:ROJO}} onClick={()=>setCfm({t:"fc",id:f.id})}>🗑</button>
                    </div>
                  </td>
                </tr>
              );})}
              </tbody>
            </table></div>
          )}
        </Crd>
      )}
      {modal==="prov"&&(<Mdl title={sel?.nombre?"Editar proveedor":"Nuevo proveedor"} onClose={()=>setModal(null)}>{(()=>{const[f,setF]=useState(sel||EP);const set=(k,v)=>setF(p=>({...p,[k]:v}));return(<div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:13}}><Fld label="Nombre *" s2><I value={f.nombre} onChange={e=>set("nombre",e.target.value)}/></Fld><Fld label="CUIT" hint><I value={f.cuit} onChange={e=>set("cuit",e.target.value)} placeholder="30-00000000-0"/></Fld><Fld label="Categoría"><S value={f.categoria} onChange={e=>set("categoria",e.target.value)}>{CATS_PROV.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Cond. de pago"><S value={f.condPago} onChange={e=>set("condPago",e.target.value)}>{CONDS_PAGO.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Estado"><S value={f.estado} onChange={e=>set("estado",e.target.value)}>{ESTADOS_PROV.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Contacto" hint><I value={f.contacto} onChange={e=>set("contacto",e.target.value)}/></Fld><Fld label="Teléfono" hint><I value={f.tel} onChange={e=>set("tel",e.target.value)}/></Fld><Fld label="Email" hint><I value={f.email} onChange={e=>set("email",e.target.value)}/></Fld><Fld label="Notas" hint s2><T value={f.notas} onChange={e=>set("notas",e.target.value)}/></Fld></div><div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:22,borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:18}}><button style={btnSec} onClick={()=>setModal(null)}>Cancelar</button><button style={btnPri} onClick={()=>{if(!f.nombre.trim())return;saveP({...f,id:f.id||uid()});}}>Guardar</button></div></div>);})()}</Mdl>)}
      {modal==="fc"&&(<Mdl title={selFc?"Editar factura":"Nueva factura de proveedor"} onClose={()=>setModal(null)}>{(()=>{const[f,setF]=useState(selFc||EFP);const set=(k,v)=>setF(p=>{const n={...p,[k]:v};if(k==="montoNeto"||k==="iva"){const nt=pN(n.montoNeto)*(1+pN(n.iva)/100);n.total=nt.toFixed(2);}return n;});return(<div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:13}}><Fld label="Proveedor *" s2><S value={f.proveedorId} onChange={e=>set("proveedorId",e.target.value)}><option value="">— Seleccioná —</option>{provs.map(p=><option key={p.id} value={p.id}>{p.nombre}</option>)}</S></Fld><Fld label="N° Factura" hint><I value={f.nroFactura} onChange={e=>set("nroFactura",e.target.value)} placeholder="00005-00012345"/></Fld><Fld label="Fecha"><I type="date" value={f.fecha} onChange={e=>set("fecha",e.target.value)}/></Fld><Fld label="Monto neto *"><I type="number" value={f.montoNeto} onChange={e=>set("montoNeto",e.target.value)}/></Fld><Fld label="IVA %"><S value={f.iva} onChange={e=>set("iva",e.target.value)}>{IVA_OPTS.map(x=><option key={x} value={x}>{x}%</option>)}</S></Fld><Fld label="Total (calculado)"><I value={$(f.total)} readOnly style={{opacity:.6}}/></Fld><Fld label="Estado"><S value={f.estado} onChange={e=>set("estado",e.target.value)}>{ESTADOS_FC2.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Forma de pago"><S value={f.formaPago} onChange={e=>set("formaPago",e.target.value)}>{FORMAS_P.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="N° cheque / ref." hint><I value={f.nroCheque} onChange={e=>set("nroCheque",e.target.value)}/></Fld><Fld label="Fecha vencimiento" hint><I type="date" value={f.fechaVto} onChange={e=>set("fechaVto",e.target.value)}/></Fld><Fld label="Asociar a venta" hint s2><S value={f.ventaId} onChange={e=>set("ventaId",e.target.value)}><option value="">— Sin asociar —</option>{ventas.map(v=><option key={v.id} value={v.id}>{v.fecha} · {v.cliente}{v.nroFactura?" · "+v.nroFactura:""}</option>)}</S></Fld><Fld label="Productos / detalle" hint s2><T value={f.productos} onChange={e=>set("productos",e.target.value)}/></Fld></div><div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:22,borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:18}}><button style={btnSec} onClick={()=>setModal(null)}>Cancelar</button><button style={btnPri} onClick={()=>{if(!f.proveedorId)return;saveF({...f,id:f.id||uid()});}}>Guardar factura</button></div></div>);})()}</Mdl>)}
      {cfm&&(<Cfm msg={cfm.t==="prov"?"¿Eliminar proveedor y todas sus facturas?":"¿Eliminar esta factura?"} onOk={()=>cfm.t==="prov"?delP(cfm.id):delF(cfm.id)} onCancel={()=>setCfm(null)}/>)}
      {toast&&<Tst msg={toast.msg} type={toast.type}/>}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// PAGE: GASTOS + PAGOS
// ══════════════════════════════════════════════════════════════════
const EG={fecha:hoy(),categoria:"Logística / Flete",tipoComp:"FC A",emisor:"",montoNeto:"",iva:"21",total:"",medioPago:"Transferencia",responsable:"Ivan Diaz",ventaId:"",descripcion:"",notas:""};
const ECH={tipo:"A cobrar",nro:"",banco:"Banco Macro",emisor:"",beneficiario:"",importe:"",moneda:"ARS",fechaEmision:hoy(),fechaVto:"",endosado:"no",endosadoA:"",ventaId:"",notas:""};
const ETR={subtipo:"cobro",fecha:hoy(),banco:"Banco Macro",moneda:"ARS",importe:"",concepto:"",contraparte:"",ventaId:"",notas:""};

function PageGastos(){
  const[gastos,setGastos]=useState([]);const[pagos,setPagos]=useState([]);const[ventas,setVentas]=useState([]);
  const[loading,setLoading]=useState(true);const[tab,setTab]=useState("gastos");const[modal,setModal]=useState(null);
  const[sel,setSel]=useState(null);const[cfm,setCfm]=useState(null);const[search,setSearch]=useState("");
  const[fCat,setFCat]=useState("Todos");const[subTab,setSubTab]=useState("cheques");const[toast,showToast]=useToast();
  useEffect(()=>{Promise.all([gl(K.gastos),gl(K.pagos),gl(K.ventas)]).then(([g,p,v])=>{setGastos(g);setPagos(p);setVentas(v);setLoading(false);});},[]);
  const saveG=async(g)=>{const n=gastos.some(x=>x.id===g.id)?gastos.map(x=>x.id===g.id?g:x):[g,...gastos];setGastos(n);await sl(K.gastos,n);setModal(null);showToast("Gasto guardado");};
  const delG=async(id)=>{const n=gastos.filter(x=>x.id!==id);setGastos(n);await sl(K.gastos,n);setCfm(null);showToast("Eliminado");};
  const saveP=async(p)=>{const n=pagos.some(x=>x.id===p.id)?pagos.map(x=>x.id===p.id?p:x):[p,...pagos];setPagos(n);await sl(K.pagos,n);setModal(null);showToast("Guardado");};
  const delP=async(id)=>{const n=pagos.filter(x=>x.id!==id);setPagos(n);await sl(K.pagos,n);setCfm(null);showToast("Eliminado");};
  const tg=gastos.reduce((a,g)=>a+pN(g.total),0);const tgMes=gastos.filter(g=>g.fecha?.slice(0,7)===hoyS.slice(0,7)).reduce((a,g)=>a+pN(g.total),0);
  const cheques=pagos.filter(p=>p.instrumento==="cheque");const transfs=pagos.filter(p=>p.instrumento==="transferencia");
  const chAC=cheques.filter(c=>c.tipo==="A cobrar").reduce((a,c)=>a+pN(c.importe),0);const chAP=cheques.filter(c=>c.tipo==="A pagar").reduce((a,c)=>a+pN(c.importe),0);
  const tEnt=transfs.filter(t=>t.subtipo==="cobro").reduce((a,t)=>a+pN(t.importe),0);const tSal=transfs.filter(t=>t.subtipo==="pago").reduce((a,t)=>a+pN(t.importe),0);
  const filtG=gastos.filter(g=>{const q=search.toLowerCase();return(!search||g.emisor?.toLowerCase().includes(q)||g.descripcion?.toLowerCase().includes(q))&&(fCat==="Todos"||g.categoria===fCat);}).sort((a,b)=>(b.fecha||"").localeCompare(a.fecha||""));
  const filtCh=cheques.filter(c=>{const q=search.toLowerCase();return!search||c.nro?.includes(q)||c.emisor?.toLowerCase().includes(q)||c.beneficiario?.toLowerCase().includes(q);});
  const filtTr=transfs.filter(t=>{const q=search.toLowerCase();return!search||t.contraparte?.toLowerCase().includes(q)||t.concepto?.toLowerCase().includes(q);});
  if(loading)return(<Spin/>);
  return (
    <div>
      <div style={{display:"flex",gap:6,marginBottom:16}}>
        {["gastos","pagos"].map(t=><button key={t} onClick={()=>setTab(t)} style={{padding:"9px 22px",borderRadius:9,border:tab===t?"1px solid rgba(16,185,129,.25)":"1px solid transparent",background:tab===t?"rgba(16,185,129,.15)":"transparent",color:tab===t?VERDE:"#475569",cursor:"pointer",fontSize:14,fontWeight:700,fontFamily:"inherit"}}>{t==="gastos"?"💸 Gastos":"💳 Pagos / Cobros"}</button>)}
      </div>
      {tab==="gastos"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
            <Met label="Total gastos" val={$(tg)} sub={`${gastos.length} registros`} dn/><Met label="Este mes" val={$(tgMes)} dn/>
            {(()=>{const cat={};gastos.forEach(g=>{cat[g.categoria]=(cat[g.categoria]||0)+pN(g.total);});const top=Object.entries(cat).sort((a,b)=>b[1]-a[1])[0];return top?<Met label="Mayor categoría" val={top[0]} sub={$(top[1])}/>:<Met label="Categorías" val="—"/>;})()}
            <Met label="Registros" val={gastos.length}/>
          </div>
          <Crd>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
              <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,margin:0}}>Registro de gastos</p>
              <button style={btnPri} onClick={()=>{setSel(null);setModal("gasto");}}>+ Cargar gasto</button>
            </div>
            <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap"}}>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar emisor o descripción..." style={{...IS,flex:1,minWidth:200}}/>
              <select value={fCat} onChange={e=>setFCat(e.target.value)} style={{...IS,width:180}}><option value="Todos">Todas las categorías</option>{CATS_GASTO.map(x=><option key={x}>{x}</option>)}</select>
            </div>
            {gastos.length===0?(<div style={{textAlign:"center",padding:"48px 0"}}><p style={{fontSize:32,marginBottom:10}}>💸</p><p style={{color:"#334155",fontSize:14,fontWeight:600}}>No hay gastos registrados</p><button style={{...btnPri,marginTop:12}} onClick={()=>setModal("gasto")}>+ Cargar primer gasto</button></div>):(
              <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr><Th>Fecha</Th><Th>Emisor</Th><Th>Categoría</Th><Th>Monto neto</Th><Th>Total</Th><Th>Medio</Th><Th>Resp.</Th><Th>Venta</Th><Th/></tr></thead>
                <tbody>{filtG.map(g=>{const v=g.ventaId?ventas.find(x=>x.id===g.ventaId):null;return(
                  <tr key={g.id}>
                    <Td>{fecha(g.fecha)}</Td><Td fw>{g.emisor}</Td><Td>{g.categoria}</Td><Td>{$(g.montoNeto)}</Td>
                    <Td r fw>{$(g.total)}</Td><Td>{g.medioPago}</Td><Td>{g.responsable||"—"}</Td>
                    <Td>{v?<span style={{background:"rgba(16,185,129,.1)",color:VERDE,padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>{v.cliente}</span>:<span style={{color:"#334155",fontSize:11}}>—</span>}</Td>
                    <td style={{padding:"9px 8px",borderBottom:"1px solid rgba(255,255,255,.04)"}}>
                      <div className="ra" style={{display:"flex",gap:4}}>
                        <button className="ib" onClick={()=>{setSel(g);setModal("gasto");}}>✏</button>
                        <button className="ib" style={{color:ROJO}} onClick={()=>setCfm({t:"g",id:g.id})}>🗑</button>
                      </div>
                    </td>
                  </tr>
                );})}
                </tbody>
              </table><p style={{fontSize:11,color:"#334155",margin:"10px 4px 0",textAlign:"right"}}>{filtG.length} registros · Total: <strong style={{color:ROJO}}>{$(filtG.reduce((a,g)=>a+pN(g.total),0))}</strong></p></div>
            )}
          </Crd>
        </div>
      )}
      {tab==="pagos"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
            <Met label="Cheques a cobrar" val={$(chAC)} up/><Met label="Cheques a pagar" val={$(chAP)} dn/>
            <Met label="Transf. cobradas" val={$(tEnt)} up/><Met label="Transf. pagadas" val={$(tSal)} dn/>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,flexWrap:"wrap",gap:10}}>
            <div style={{display:"flex",gap:4}}>
              {["cheques","transfs"].map(t=><button key={t} onClick={()=>setSubTab(t)} style={{padding:"7px 14px",borderRadius:8,border:"none",background:subTab===t?"rgba(16,185,129,.12)":"rgba(255,255,255,.04)",color:subTab===t?VERDE:"#64748b",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>{t==="cheques"?`🧾 Cheques (${cheques.length})`:`↔ Transferencias (${transfs.length})`}</button>)}
            </div>
            <div style={{display:"flex",gap:8}}>
              <button style={{...btnPri,background:"rgba(16,185,129,.08)",borderColor:"rgba(16,185,129,.4)",color:"#34d399"}} onClick={()=>{setSel(null);setModal("transf");}}>+ Transferencia</button>
              <button style={btnPri} onClick={()=>{setSel(null);setModal("cheque");}}>+ Cheque</button>
            </div>
          </div>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar..." style={{...IS,marginBottom:14}}/>
          {subTab==="cheques"&&(
            <Crd>
              {cheques.length===0?(<div style={{textAlign:"center",padding:"48px 0"}}><p style={{fontSize:32,marginBottom:10}}>🧾</p><p style={{color:"#334155",fontSize:14}}>No hay cheques registrados</p><button style={{...btnPri,marginTop:12}} onClick={()=>setModal("cheque")}>+ Registrar cheque</button></div>):(
                <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead><tr><Th>Tipo</Th><Th>N° Cheque</Th><Th>Banco</Th><Th>Emisor</Th><Th>Beneficiario</Th><Th>Importe</Th><Th>Emisión</Th><Th>Vence</Th><Th>Venta</Th><Th/></tr></thead>
                  <tbody>{filtCh.map(c=>{const venc2=["A cobrar","A pagar"].includes(c.tipo)&&c.fechaVto&&c.fechaVto<=hoyS;const v=c.ventaId?ventas.find(x=>x.id===c.ventaId):null;return(
                    <tr key={c.id}>
                      <Td><Bdg t={c.tipo} sm/></Td><Td mono fw>{c.nro}</Td><Td>{c.banco}</Td><Td>{c.emisor||"—"}</Td><Td>{c.beneficiario||"—"}</Td>
                      <Td fw><span style={{color:c.tipo==="A cobrar"?VERDE:ROJO}}>{$(c.importe)}</span></Td>
                      <Td>{fecha(c.fechaEmision)}</Td><Td><span style={{color:venc2?ROJO:"#64748b"}}>{fecha(c.fechaVto)}{venc2?" ⚠":""}</span></Td>
                      <Td>{v?<span style={{background:"rgba(16,185,129,.1)",color:VERDE,padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>{v.cliente}</span>:<span style={{color:"#334155",fontSize:11}}>—</span>}</Td>
                      <td style={{padding:"9px 8px",borderBottom:"1px solid rgba(255,255,255,.04)"}}>
                        <div className="ra" style={{display:"flex",gap:4}}>
                          <button className="ib" onClick={()=>{setSel(c);setModal("cheque");}}>✏</button>
                          <button className="ib" style={{color:ROJO}} onClick={()=>setCfm({t:"p",id:c.id})}>🗑</button>
                        </div>
                      </td>
                    </tr>
                  );})}
                  </tbody>
                </table></div>
              )}
            </Crd>
          )}
          {subTab==="transfs"&&(
            <Crd>
              {transfs.length===0?(<div style={{textAlign:"center",padding:"48px 0"}}><p style={{fontSize:32,marginBottom:10}}>↔</p><p style={{color:"#334155",fontSize:14}}>No hay transferencias registradas</p><button style={{...btnPri,marginTop:12}} onClick={()=>setModal("transf")}>+ Registrar transferencia</button></div>):(
                <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead><tr><Th>Fecha</Th><Th>Tipo</Th><Th>Banco</Th><Th>Contraparte</Th><Th>Concepto</Th><Th>Importe</Th><Th>Venta</Th><Th/></tr></thead>
                  <tbody>{filtTr.map(t=>{const v=t.ventaId?ventas.find(x=>x.id===t.ventaId):null;return(
                    <tr key={t.id}>
                      <Td>{fecha(t.fecha)}</Td>
                      <Td><span style={{background:t.subtipo==="cobro"?"rgba(16,185,129,.15)":"rgba(248,113,113,.15)",color:t.subtipo==="cobro"?VERDE:ROJO,padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:600}}>{t.subtipo==="cobro"?"↓ Cobro":"↑ Pago"}</span></Td>
                      <Td>{t.banco}</Td><Td fw>{t.contraparte}</Td><Td>{t.concepto||"—"}</Td>
                      <Td fw><span style={{color:t.subtipo==="cobro"?VERDE:ROJO}}>{t.subtipo==="cobro"?"+ ":"– "}{$(t.importe)}</span></Td>
                      <Td>{v?<span style={{background:"rgba(16,185,129,.1)",color:VERDE,padding:"2px 8px",borderRadius:20,fontSize:10,fontWeight:600}}>{v.cliente}</span>:<span style={{color:"#334155",fontSize:11}}>—</span>}</Td>
                      <td style={{padding:"9px 8px",borderBottom:"1px solid rgba(255,255,255,.04)"}}>
                        <div className="ra" style={{display:"flex",gap:4}}>
                          <button className="ib" onClick={()=>{setSel(t);setModal("transf");}}>✏</button>
                          <button className="ib" style={{color:ROJO}} onClick={()=>setCfm({t:"p",id:t.id})}>🗑</button>
                        </div>
                      </td>
                    </tr>
                  );})}
                  </tbody>
                </table><p style={{fontSize:11,color:"#334155",margin:"10px 4px 0",textAlign:"right"}}>Entradas: <strong style={{color:VERDE}}>{$(tEnt)}</strong> · Salidas: <strong style={{color:ROJO}}>{$(tSal)}</strong> · Neto: <strong style={{color:tEnt-tSal>=0?VERDE:ROJO}}>{$(tEnt-tSal)}</strong></p></div>
              )}
            </Crd>
          )}
        </div>
      )}
      {modal==="gasto"&&(<Mdl title={sel?"Editar gasto":"Nuevo gasto"} onClose={()=>setModal(null)}>{(()=>{const[f,setF]=useState(sel||EG);const set=(k,v)=>setF(p=>{const n={...p,[k]:v};if(k==="montoNeto"||k==="iva"){n.total=(pN(n.montoNeto)*(1+pN(n.iva)/100)).toFixed(2);}return n;});return(<div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:13}}><Fld label="Fecha"><I type="date" value={f.fecha} onChange={e=>set("fecha",e.target.value)}/></Fld><Fld label="Categoría"><S value={f.categoria} onChange={e=>set("categoria",e.target.value)}>{CATS_GASTO.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Emisor *" s2><I value={f.emisor} onChange={e=>set("emisor",e.target.value)} placeholder="Nombre del emisor"/></Fld><Fld label="Monto neto *"><I type="number" value={f.montoNeto} onChange={e=>set("montoNeto",e.target.value)}/></Fld><Fld label="IVA %"><S value={f.iva} onChange={e=>set("iva",e.target.value)}>{IVA_OPTS.map(x=><option key={x} value={x}>{x}%</option>)}</S></Fld><Fld label="Total (calculado)"><I value={$(f.total)} readOnly style={{opacity:.6}}/></Fld><Fld label="Medio de pago"><S value={f.medioPago} onChange={e=>set("medioPago",e.target.value)}>{MEDIOS_G.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Responsable"><I value={f.responsable} onChange={e=>set("responsable",e.target.value)}/></Fld><Fld label="Descripción" s2><I value={f.descripcion} onChange={e=>set("descripcion",e.target.value)} placeholder="Detalle del gasto..."/></Fld><Fld label="Asociar a venta" hint s2><S value={f.ventaId} onChange={e=>set("ventaId",e.target.value)}><option value="">— Sin asociar —</option>{ventas.map(v=><option key={v.id} value={v.id}>{v.fecha} · {v.cliente}</option>)}</S></Fld></div><div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:22,borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:18}}><button style={btnSec} onClick={()=>setModal(null)}>Cancelar</button><button style={btnPri} onClick={()=>{if(!f.emisor.trim()||!f.montoNeto)return;saveG({...f,id:f.id||uid()});}}>Guardar gasto</button></div></div>);})()}</Mdl>)}
      {modal==="cheque"&&(<Mdl title={sel?.nro?"Editar cheque":"Nuevo cheque"} onClose={()=>setModal(null)}>{(()=>{const[f,setF]=useState(sel||ECH);const set=(k,v)=>setF(p=>({...p,[k]:v}));return(<div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:13}}><Fld label="Tipo"><S value={f.tipo} onChange={e=>set("tipo",e.target.value)}>{T_CHEQUE.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="N° cheque *"><I value={f.nro} onChange={e=>set("nro",e.target.value)}/></Fld><Fld label="Banco"><S value={f.banco} onChange={e=>set("banco",e.target.value)}>{BANCOS.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Moneda"><S value={f.moneda} onChange={e=>set("moneda",e.target.value)}>{MONEDAS.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Importe *"><I type="number" value={f.importe} onChange={e=>set("importe",e.target.value)}/></Fld><Fld label="Emisor"><I value={f.emisor} onChange={e=>set("emisor",e.target.value)}/></Fld><Fld label="Beneficiario"><I value={f.beneficiario} onChange={e=>set("beneficiario",e.target.value)}/></Fld><Fld label="Fecha emisión"><I type="date" value={f.fechaEmision} onChange={e=>set("fechaEmision",e.target.value)}/></Fld><Fld label="Fecha vencimiento"><I type="date" value={f.fechaVto} onChange={e=>set("fechaVto",e.target.value)}/></Fld><Fld label="¿Endosado?"><S value={f.endosado} onChange={e=>set("endosado",e.target.value)}><option value="no">No</option><option value="si">Sí</option></S></Fld>{f.endosado==="si"&&<Fld label="Endosado a"><I value={f.endosadoA} onChange={e=>set("endosadoA",e.target.value)}/></Fld>}<Fld label="Asociar a venta" hint s2><S value={f.ventaId} onChange={e=>set("ventaId",e.target.value)}><option value="">— Sin asociar —</option>{ventas.map(v=><option key={v.id} value={v.id}>{v.fecha} · {v.cliente}</option>)}</S></Fld></div><div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:22,borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:18}}><button style={btnSec} onClick={()=>setModal(null)}>Cancelar</button><button style={btnPri} onClick={()=>{if(!f.nro.trim()||!f.importe)return;saveP({...f,id:f.id||uid(),instrumento:"cheque"});}}>Guardar cheque</button></div></div>);})()}</Mdl>)}
      {modal==="transf"&&(<Mdl title={sel?.contraparte?"Editar transferencia":"Nueva transferencia"} onClose={()=>setModal(null)}>{(()=>{const[f,setF]=useState(sel||ETR);const set=(k,v)=>setF(p=>({...p,[k]:v}));return(<div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:13}}><Fld label="Tipo"><S value={f.subtipo} onChange={e=>set("subtipo",e.target.value)}><option value="cobro">↓ Cobro (entrada)</option><option value="pago">↑ Pago (salida)</option></S></Fld><Fld label="Fecha"><I type="date" value={f.fecha} onChange={e=>set("fecha",e.target.value)}/></Fld><Fld label="Banco"><S value={f.banco} onChange={e=>set("banco",e.target.value)}>{BANCOS.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Moneda"><S value={f.moneda} onChange={e=>set("moneda",e.target.value)}>{MONEDAS.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Importe *"><I type="number" value={f.importe} onChange={e=>set("importe",e.target.value)}/></Fld><Fld label={f.subtipo==="cobro"?"Pagado por *":"Pagado a *"}><I value={f.contraparte} onChange={e=>set("contraparte",e.target.value)}/></Fld><Fld label="Concepto" s2><I value={f.concepto} onChange={e=>set("concepto",e.target.value)} placeholder="Cobro FC..., Pago proveedor..."/></Fld><Fld label="Asociar a venta" hint s2><S value={f.ventaId} onChange={e=>set("ventaId",e.target.value)}><option value="">— Sin asociar —</option>{ventas.map(v=><option key={v.id} value={v.id}>{v.fecha} · {v.cliente}</option>)}</S></Fld></div><div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:22,borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:18}}><button style={btnSec} onClick={()=>setModal(null)}>Cancelar</button><button style={btnPri} onClick={()=>{if(!f.contraparte.trim()||!f.importe)return;saveP({...f,id:f.id||uid(),instrumento:"transferencia"});}}>Guardar</button></div></div>);})()}</Mdl>)}
      {cfm&&(<Cfm msg="¿Eliminar este registro?" onOk={()=>cfm.t==="g"?delG(cfm.id):delP(cfm.id)} onCancel={()=>setCfm(null)}/>)}
      {toast&&<Tst msg={toast.msg} type={toast.type}/>}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// PAGE: CASH FLOW
// ══════════════════════════════════════════════════════════════════
const EC={nombre:"",tipo:"Cuenta corriente",banco:"Banco Macro",moneda:"ARS",saldo:"",notas:"",fechaActualizacion:hoy()};
const EM={fecha:hoy(),tipo:"Cobro cliente",descripcion:"",importe:"",esEgreso:false,cuentaId:"",ventaId:"",notas:""};

function PageCashFlow(){
  const[cuentas,setCuentas]=useState([]);const[movs,setMovs]=useState([]);const[ventas,setVentas]=useState([]);
  const[loading,setLoading]=useState(true);const[tab,setTab]=useState("res");const[modal,setModal]=useState(null);
  const[sel,setSel]=useState(null);const[cfm,setCfm]=useState(null);const[search,setSearch]=useState("");
  const[fTipo,setFTipo]=useState("Todos");const[toast,showToast]=useToast();
  useEffect(()=>{Promise.all([gl(K.cuentas),gl(K.movs),gl(K.ventas)]).then(([c,m,v])=>{setCuentas(c);setMovs(m);setVentas(v);setLoading(false);});},[]);
  const saveC=async(c)=>{const n=cuentas.some(x=>x.id===c.id)?cuentas.map(x=>x.id===c.id?c:x):[c,...cuentas];setCuentas(n);await sl(K.cuentas,n);setModal(null);showToast("Cuenta guardada");};
  const delC=async(id)=>{const n=cuentas.filter(x=>x.id!==id);setCuentas(n);await sl(K.cuentas,n);setCfm(null);showToast("Cuenta eliminada");};
  const saveM=async(m)=>{const n=movs.some(x=>x.id===m.id)?movs.map(x=>x.id===m.id?m:x):[m,...movs];setMovs(n);await sl(K.movs,n);setModal(null);showToast("Movimiento guardado");};
  const delM=async(id)=>{const n=movs.filter(x=>x.id!==id);setMovs(n);await sl(K.movs,n);setCfm(null);showToast("Eliminado");};
  const reales=movs.filter(m=>!m.proyectado);const proyect=movs.filter(m=>m.proyectado).sort((a,b)=>(a.fecha||"").localeCompare(b.fecha||""));
  const saldoARS=cuentas.filter(c=>c.moneda==="ARS").reduce((a,c)=>a+pN(c.saldo),0);const saldoFCI=cuentas.filter(c=>c.tipo==="FCI"&&c.moneda==="ARS").reduce((a,c)=>a+pN(c.saldo),0);
  const saldoUSD=cuentas.filter(c=>c.moneda==="USD").reduce((a,c)=>a+pN(c.saldo),0);
  const tIng=reales.filter(m=>!m.esEgreso).reduce((a,m)=>a+pN(m.importe),0);const tEgr=reales.filter(m=>m.esEgreso).reduce((a,m)=>a+pN(m.importe),0);
  const byMes={};reales.forEach(m=>{const mes=m.fecha?.slice(0,7);if(!mes)return;if(!byMes[mes])byMes[mes]={mes,ingresos:0,egresos:0};if(m.esEgreso)byMes[mes].egresos+=pN(m.importe);else byMes[mes].ingresos+=pN(m.importe);});
  const barData=Object.values(byMes).sort((a,b)=>a.mes.localeCompare(b.mes)).slice(-6).map(d=>({...d,l:mesL(d.mes)}));
  let spAcum=saldoARS;const proyData=[{l:"Hoy",s:saldoARS}];proyect.forEach(m=>{spAcum+=m.esEgreso?-pN(m.importe):pN(m.importe);proyData.push({l:fecha(m.fecha),s:spAcum,desc:m.descripcion});});
  const filtM=reales.filter(m=>{const q=search.toLowerCase();const mQ=!search||m.descripcion?.toLowerCase().includes(q)||m.tipo?.toLowerCase().includes(q);const mT=fTipo==="Todos"||(fTipo==="Ingresos"&&!m.esEgreso)||(fTipo==="Egresos"&&m.esEgreso)||(m.tipo===fTipo);return mQ&&mT;}).sort((a,b)=>(b.fecha||"").localeCompare(a.fecha||""));
  if(loading)return(<Spin/>);
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
        <Met label="Saldo total ARS" val={$(saldoARS)} up/><Met label="En FCI" val={$(saldoFCI)} up/>
        <Met label="Saldo USD" val={`U$D ${$(saldoUSD)}`}/><Met label="Neto movimientos" val={$(tIng-tEgr)} up={tIng>=tEgr} dn={tIng<tEgr}/>
      </div>
      <div style={{display:"flex",gap:6,marginBottom:14}}>
        {["res","movs","proy"].map(t=><button key={t} onClick={()=>setTab(t)} style={{padding:"7px 16px",borderRadius:8,border:"none",background:tab===t?"rgba(16,185,129,.12)":"rgba(255,255,255,.04)",color:tab===t?VERDE:"#64748b",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>{t==="res"?"📊 Resumen":t==="movs"?"↔ Movimientos":"🔮 Proyección"}</button>)}
        <div style={{marginLeft:"auto",display:"flex",gap:8}}>
          <button style={{...btnPri,background:"rgba(16,185,129,.08)",borderColor:"rgba(16,185,129,.4)",color:"#34d399"}} onClick={()=>{setSel(null);setModal("mov");}}>+ Movimiento</button>
          <button style={btnPri} onClick={()=>{setSel(null);setModal("cuenta");}}>+ Cuenta</button>
        </div>
      </div>
      {tab==="res"&&(
        <div>
          {cuentas.length===0?(<Crd><div style={{textAlign:"center",padding:"40px 0"}}><p style={{fontSize:32,marginBottom:10}}>🏦</p><p style={{color:"#334155",fontSize:14,fontWeight:600}}>No hay cuentas registradas</p><button style={{...btnPri,marginTop:12}} onClick={()=>setModal("cuenta")}>+ Agregar primera cuenta</button></div></Crd>):(
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:12,marginBottom:16}}>
              {cuentas.map(c=>(
                <div key={c.id} style={{background:"#131f35",borderRadius:14,padding:"18px 20px",border:`1px solid ${c.tipo==="FCI"?"rgba(16,185,129,.25)":c.moneda!=="ARS"?"rgba(245,158,11,.2)":"rgba(255,255,255,.07)"}`,position:"relative"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
                    <div><p style={{fontSize:10,color:"#475569",margin:"0 0 2px",textTransform:"uppercase",fontWeight:700}}>{c.tipo}</p><p style={{fontSize:14,fontWeight:700,color:"#f1f5f9",margin:0}}>{c.nombre}</p>{c.banco&&<p style={{fontSize:11,color:"#475569",margin:"2px 0 0"}}>{c.banco}</p>}</div>
                    <div style={{display:"flex",gap:4}}>
                      <button className="ib" onClick={()=>{setSel(c);setModal("cuenta");}}>✏</button>
                      <button className="ib" style={{color:ROJO}} onClick={()=>setCfm({t:"c",id:c.id})}>🗑</button>
                    </div>
                  </div>
                  <p style={{fontSize:22,fontWeight:700,color:pN(c.saldo)<0?ROJO:c.tipo==="FCI"?VERDE:c.moneda!=="ARS"?AMBAR:"#f1f5f9",margin:"0 0 4px"}}>{c.moneda!=="ARS"?c.moneda+" ":""}{$(c.saldo)}</p>
                  <p style={{fontSize:10,color:"#334155",margin:0}}>Act.: {fecha(c.fechaActualizacion||hoyS)}</p>
                </div>
              ))}
            </div>
          )}
          {barData.length>0&&(
            <Crd>
              <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Ingresos vs egresos por mes</p>
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={barData} margin={{top:0,right:10,left:0,bottom:0}}>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,.05)"/>
                  <XAxis dataKey="l" tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false} tickFormatter={n=>n>=1e6?"$"+(n/1e6).toFixed(1)+"M":n>=1e3?"$"+(n/1e3).toFixed(0)+"K":"$"+n}/>
                  <Tooltip {...TT} formatter={(v,name)=>[$(v),name==="ingresos"?"Ingresos":"Egresos"]}/>
                  <Bar dataKey="ingresos" fill={VERDE} radius={[4,4,0,0]} name="ingresos"/>
                  <Bar dataKey="egresos" fill={ROJO} radius={[4,4,0,0]} name="egresos"/>
                </BarChart>
              </ResponsiveContainer>
              <div style={{display:"flex",gap:16,marginTop:10,justifyContent:"flex-end"}}>
                <span style={{fontSize:12,color:VERDE}}>● Ingresos: {$(tIng)}</span>
                <span style={{fontSize:12,color:ROJO}}>● Egresos: {$(tEgr)}</span>
                <span style={{fontSize:12,fontWeight:700,color:tIng>=tEgr?VERDE:ROJO}}>Neto: {$(tIng-tEgr)}</span>
              </div>
            </Crd>
          )}
          {reales.length>0&&(
            <Crd mb={0}>
              <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Últimos movimientos</p>
              {[...reales].sort((a,b)=>(b.fecha||"").localeCompare(a.fecha||"")).slice(0,6).map((m,i)=>{const c=cuentas.find(x=>x.id===m.cuentaId);return(
                <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 0",borderBottom:i<5?"1px solid rgba(255,255,255,.05)":"none"}}>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <div style={{width:30,height:30,borderRadius:8,background:m.esEgreso?"rgba(248,113,113,.1)":"rgba(16,185,129,.1)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,flexShrink:0}}>{m.esEgreso?"↑":"↓"}</div>
                    <div><p style={{fontSize:13,color:"#e2e8f0",margin:0,fontWeight:500}}>{m.descripcion}</p><p style={{fontSize:11,color:"#475569",margin:0}}>{fecha(m.fecha)}{c?" · "+c.nombre:""}</p></div>
                  </div>
                  <p style={{fontSize:14,fontWeight:700,margin:0,color:m.esEgreso?ROJO:VERDE}}>{m.esEgreso?"– ":"+ "}{$(m.importe)}</p>
                </div>
              );})}
            </Crd>
          )}
        </div>
      )}
      {tab==="movs"&&(
        <Crd>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <p style={{margin:0,fontSize:13,color:"#64748b"}}>{reales.length} movimientos</p>
            <BtnExport onClick={()=>exportExcel(reales.map(m=>({Fecha:m.fecha,Descripción:m.descripcion,Tipo:m.tipo,Dirección:m.esEgreso?"Egreso":"Ingreso",Importe:m.importe,Cuenta:cuentas.find(c=>c.id===m.cuentaId)?.nombre||""})),"cashflow_sasoc","Movimientos")}/>
          </div>
          <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap"}}>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar descripción o tipo..." style={{...IS,flex:1,minWidth:200}}/>
            <select value={fTipo} onChange={e=>setFTipo(e.target.value)} style={{...IS,width:160}}><option value="Todos">Todos</option><option value="Ingresos">Solo ingresos</option><option value="Egresos">Solo egresos</option>{T_MOV.map(x=><option key={x}>{x}</option>)}</select>
          </div>
          {reales.length===0?(<div style={{textAlign:"center",padding:"48px 0"}}><p style={{fontSize:32,marginBottom:10}}>↔</p><p style={{color:"#334155",fontSize:14}}>No hay movimientos registrados</p><button style={{...btnPri,marginTop:12}} onClick={()=>setModal("mov")}>+ Registrar primer movimiento</button></div>):(
            <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr><Th/><Th>Fecha</Th><Th>Descripción</Th><Th>Tipo</Th><Th>Cuenta</Th><Th>Importe</Th><Th/></tr></thead>
              <tbody>{filtM.map(m=>{const c=cuentas.find(x=>x.id===m.cuentaId);return(
                <tr key={m.id}>
                  <td style={{padding:"8px 10px",borderBottom:"1px solid rgba(255,255,255,.04)",width:32}}><div style={{width:28,height:28,borderRadius:7,background:m.esEgreso?"rgba(248,113,113,.1)":"rgba(16,185,129,.1)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13}}>{m.esEgreso?"↑":"↓"}</div></td>
                  <Td>{fecha(m.fecha)}</Td><Td fw>{m.descripcion}</Td><Td>{m.tipo}</Td><Td>{c?.nombre||"—"}</Td>
                  <Td fw><span style={{color:m.esEgreso?ROJO:VERDE}}>{m.esEgreso?"– ":"+ "}{$(m.importe)}</span></Td>
                  <td style={{padding:"9px 8px",borderBottom:"1px solid rgba(255,255,255,.04)"}}>
                    <div className="ra" style={{display:"flex",gap:4}}>
                      <button className="ib" onClick={()=>{setSel(m);setModal("mov");}}>✏</button>
                      <button className="ib" style={{color:ROJO}} onClick={()=>setCfm({t:"m",id:m.id})}>🗑</button>
                    </div>
                  </td>
                </tr>
              );})}
              </tbody>
            </table></div>
          )}
        </Crd>
      )}
      {tab==="proy"&&(
        <div>
          {proyData.length>1&&(
            <Crd>
              <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:16}}>
                <Met label="Saldo hoy" val={$(saldoARS)}/><Met label="Ingresos esperados" val={$(proyect.filter(m=>!m.esEgreso).reduce((a,m)=>a+pN(m.importe),0))} up/>
                <Met label="Egresos esperados" val={$(proyect.filter(m=>m.esEgreso).reduce((a,m)=>a+pN(m.importe),0))} dn/>
                <Met label="Saldo proyectado" val={$(proyData[proyData.length-1]?.s||saldoARS)} up={proyData[proyData.length-1]?.s>=saldoARS}/>
              </div>
              <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Evolución proyectada</p>
              <ResponsiveContainer width="100%" height={200}>
                <LineChart data={proyData} margin={{top:5,right:10,left:0,bottom:0}}>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,.05)"/>
                  <XAxis dataKey="l" tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false} tickFormatter={n=>n>=1e6?"$"+(n/1e6).toFixed(1)+"M":"$"+(n/1e3).toFixed(0)+"K"}/>
                  <Tooltip {...TT} formatter={v=>[$(v),"Saldo proyectado"]}/>
                  <ReferenceLine y={0} stroke="rgba(248,113,113,.3)" strokeDasharray="4 4"/>
                  <Line type="monotone" dataKey="s" stroke={VERDE} strokeWidth={2} dot={{fill:VERDE,r:4}}/>
                </LineChart>
              </ResponsiveContainer>
            </Crd>
          )}
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,margin:0}}>Ítems proyectados ({proyect.length})</p>
            <div style={{display:"flex",gap:8}}>
              {proyect.length>0&&<button onClick={async()=>{const n=movs.filter(m=>!m.proyectado);setMovs(n);await sl(K.movs,n);showToast("Proyección limpiada");}} style={{...btnSec,color:ROJO,borderColor:"rgba(239,68,68,.3)"}}>🗑 Limpiar</button>}
              <button style={btnPri} onClick={()=>setModal("proy")}>+ Agregar ítems</button>
            </div>
          </div>
          <Crd mb={0}>
            {proyect.length===0?(<div style={{textAlign:"center",padding:"36px 0"}}><p style={{fontSize:28,marginBottom:8}}>🔮</p><p style={{color:"#334155",fontSize:13}}>Agregá ingresos y egresos futuros para ver la proyección</p></div>):(
              proyect.map(m=>(
                <div key={m.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 0",borderBottom:"1px solid rgba(255,255,255,.05)"}}>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <div style={{width:28,height:28,borderRadius:7,background:m.esEgreso?"rgba(248,113,113,.1)":"rgba(16,185,129,.1)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,flexShrink:0}}>{m.esEgreso?"↑":"↓"}</div>
                    <div><p style={{fontSize:13,color:"#e2e8f0",margin:0}}>{m.descripcion}</p><p style={{fontSize:11,color:"#475569",margin:0}}>{fecha(m.fecha)}</p></div>
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:12}}>
                    <p style={{fontSize:14,fontWeight:700,margin:0,color:m.esEgreso?ROJO:VERDE}}>{m.esEgreso?"– ":"+ "}{$(m.importe)}</p>
                    <button className="ib" style={{color:ROJO}} onClick={()=>setCfm({t:"m",id:m.id})}>🗑</button>
                  </div>
                </div>
              ))
            )}
          </Crd>
        </div>
      )}
      {modal==="cuenta"&&(<Mdl title={sel?.nombre?"Editar cuenta":"Nueva cuenta"} onClose={()=>setModal(null)}>{(()=>{const[f,setF]=useState(sel||EC);const set=(k,v)=>setF(p=>({...p,[k]:v}));return(<div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:13}}><Fld label="Nombre *" s2><I value={f.nombre} onChange={e=>set("nombre",e.target.value)} placeholder="Ej: Cuenta Macro, FCI Balanz..."/></Fld><Fld label="Tipo"><S value={f.tipo} onChange={e=>set("tipo",e.target.value)}>{T_CUENTA.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Banco"><S value={f.banco} onChange={e=>set("banco",e.target.value)}>{BANCOS.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Moneda"><S value={f.moneda} onChange={e=>set("moneda",e.target.value)}>{MONEDAS.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Saldo actual *"><I type="number" value={f.saldo} onChange={e=>set("saldo",e.target.value)}/></Fld><Fld label="Fecha actualización"><I type="date" value={f.fechaActualizacion} onChange={e=>set("fechaActualizacion",e.target.value)}/></Fld><Fld label="Notas" hint s2><T value={f.notas} onChange={e=>set("notas",e.target.value)} placeholder="CBU, alias..."/></Fld></div><div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:22,borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:18}}><button style={btnSec} onClick={()=>setModal(null)}>Cancelar</button><button style={btnPri} onClick={()=>{if(!f.nombre.trim())return;saveC({...f,id:f.id||uid()});}}>Guardar cuenta</button></div></div>);})()}</Mdl>)}
      {modal==="mov"&&(<Mdl title={sel?.descripcion?"Editar movimiento":"Nuevo movimiento"} onClose={()=>setModal(null)}>{(()=>{const[f,setF]=useState(sel||EM);const set=(k,v)=>setF(p=>({...p,[k]:v}));return(<div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:13}}><Fld label="Fecha"><I type="date" value={f.fecha} onChange={e=>set("fecha",e.target.value)}/></Fld><Fld label="Tipo"><S value={f.tipo} onChange={e=>{const v=e.target.value;setF(p=>({...p,tipo:v,esEgreso:EGRESOS_MOV.includes(v)}))}}>{ T_MOV.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Descripción *" s2><I value={f.descripcion} onChange={e=>set("descripcion",e.target.value)} placeholder="Cobro OSPESE FC 1151..."/></Fld><Fld label="Importe *"><I type="number" value={f.importe} onChange={e=>set("importe",e.target.value)}/></Fld><Fld label="Dirección"><div style={{display:"flex",gap:8}}><button type="button" onClick={()=>set("esEgreso",false)} style={{flex:1,padding:"8px",borderRadius:8,border:`1px solid ${!f.esEgreso?VERDE:"rgba(255,255,255,.1)"}`,background:!f.esEgreso?"rgba(16,185,129,.12)":"transparent",color:!f.esEgreso?VERDE:"#64748b",cursor:"pointer",fontSize:13,fontWeight:600}}>↓ Ingreso</button><button type="button" onClick={()=>set("esEgreso",true)} style={{flex:1,padding:"8px",borderRadius:8,border:`1px solid ${f.esEgreso?"#ef4444":"rgba(255,255,255,.1)"}`,background:f.esEgreso?"rgba(239,68,68,.1)":"transparent",color:f.esEgreso?ROJO:"#64748b",cursor:"pointer",fontSize:13,fontWeight:600}}>↑ Egreso</button></div></Fld><Fld label="Cuenta" hint><S value={f.cuentaId} onChange={e=>set("cuentaId",e.target.value)}><option value="">— Sin especificar —</option>{cuentas.map(c=><option key={c.id} value={c.id}>{c.nombre}</option>)}</S></Fld><Fld label="Asociar a venta" hint s2><S value={f.ventaId} onChange={e=>set("ventaId",e.target.value)}><option value="">— Sin asociar —</option>{ventas.map(v=><option key={v.id} value={v.id}>{v.fecha} · {v.cliente}</option>)}</S></Fld></div><div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:22,borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:18}}><button style={btnSec} onClick={()=>setModal(null)}>Cancelar</button><button style={btnPri} onClick={()=>{if(!f.descripcion.trim()||!f.importe)return;saveM({...f,id:f.id||uid()});}}>Guardar</button></div></div>);})()}</Mdl>)}
      {modal==="proy"&&(<Mdl title="Agregar ítems a proyección" onClose={()=>setModal(null)}>{(()=>{const[items,setItems]=useState([{id:uid(),fecha:hoy(),descripcion:"",importe:"",esEgreso:false}]);const addI=()=>setItems(p=>[...p,{id:uid(),fecha:hoy(),descripcion:"",importe:"",esEgreso:false}]);const setI=(id,k,v)=>setItems(p=>p.map(x=>x.id===id?{...x,[k]:v}:x));const remI=(id)=>setItems(p=>p.filter(x=>x.id!==id));return(<div><p style={{fontSize:13,color:"#64748b",marginBottom:16}}>Agregá ingresos y egresos futuros esperados.</p>{items.map((it,idx)=><div key={it.id} style={{background:"rgba(255,255,255,.03)",borderRadius:10,padding:14,marginBottom:10,border:"1px solid rgba(255,255,255,.06)"}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}><span style={{fontSize:12,color:"#64748b",fontWeight:700}}>Ítem {idx+1}</span>{items.length>1&&<button onClick={()=>remI(it.id)} style={{background:"transparent",border:"none",color:ROJO,cursor:"pointer",fontSize:12}}>✕ Quitar</button>}</div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr auto",gap:10,alignItems:"end"}}><div style={{display:"flex",flexDirection:"column",gap:4}}><label style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>Fecha</label><I type="date" value={it.fecha} onChange={e=>setI(it.id,"fecha",e.target.value)}/></div><div style={{display:"flex",flexDirection:"column",gap:4}}><label style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>Descripción</label><I value={it.descripcion} onChange={e=>setI(it.id,"descripcion",e.target.value)} placeholder="Ej: Cobro THAR SA"/></div><div style={{display:"flex",flexDirection:"column",gap:4}}><label style={{fontSize:11,color:"#64748b",fontWeight:700,textTransform:"uppercase"}}>Importe</label><I type="number" value={it.importe} onChange={e=>setI(it.id,"importe",e.target.value)}/></div><div style={{display:"flex",gap:6,paddingBottom:1}}><button type="button" onClick={()=>setI(it.id,"esEgreso",false)} style={{padding:"8px 10px",borderRadius:7,border:`1px solid ${!it.esEgreso?VERDE:"rgba(255,255,255,.1)"}`,background:!it.esEgreso?"rgba(16,185,129,.12)":"transparent",color:!it.esEgreso?VERDE:"#64748b",cursor:"pointer",fontSize:12,fontWeight:600}}>↓</button><button type="button" onClick={()=>setI(it.id,"esEgreso",true)} style={{padding:"8px 10px",borderRadius:7,border:`1px solid ${it.esEgreso?"#ef4444":"rgba(255,255,255,.1)"}`,background:it.esEgreso?"rgba(239,68,68,.1)":"transparent",color:it.esEgreso?ROJO:"#64748b",cursor:"pointer",fontSize:12,fontWeight:600}}>↑</button></div></div></div>)}<button onClick={addI} style={{width:"100%",padding:9,borderRadius:9,border:"1px dashed rgba(255,255,255,.12)",background:"transparent",color:"#64748b",cursor:"pointer",fontSize:13,marginBottom:16}}>+ Agregar ítem</button><div style={{display:"flex",gap:10,justifyContent:"flex-end",borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:18}}><button style={btnSec} onClick={()=>setModal(null)}>Cancelar</button><button style={btnPri} onClick={async()=>{const v=items.filter(x=>x.descripcion.trim()&&x.importe&&!isNaN(x.importe)).map(x=>({...x,id:uid(),tipo:"Proyectado",proyectado:true}));if(!v.length)return;const n=[...movs,...v];setMovs(n);await sl(K.movs,n);setModal(null);showToast(`${v.length} ítem${v.length>1?"s":""} agregados`);}}>Agregar a proyección</button></div></div>);})()}</Mdl>)}
      {cfm&&(<Cfm msg="¿Eliminar?" onOk={()=>{if(cfm.t==="c")delC(cfm.id);else delM(cfm.id);}} onCancel={()=>setCfm(null)}/>)}
      {toast&&<Tst msg={toast.msg} type={toast.type}/>}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// PAGE: ESTADÍSTICAS
// ══════════════════════════════════════════════════════════════════
function PageEstadisticas(){
  const[data,setData]=useState(null);const[sec,setSec]=useState("gral");
  useEffect(()=>{Promise.all([gl(K.ventas),gl(K.gastos),gl(K.cuentas),gl(K.facsP)]).then(([v,g,c,fp])=>setData({v,g,c,fp}));},[]);
  if(!data)return(<Spin/>);
  const{v,g,c,fp}=data;
  const tv=v.reduce((a,x)=>a+pN(x.totalVenta),0);const tc=v.reduce((a,x)=>a+pN(x.cobrado),0);const tg=g.reduce((a,x)=>a+pN(x.total),0);
  const tfp=fp.reduce((a,x)=>a+pN(x.total),0);const pagfp=fp.filter(x=>x.estado==="Pagado").reduce((a,x)=>a+pN(x.total),0);
  const sd=c.filter(x=>x.moneda==="ARS").reduce((a,x)=>a+pN(x.saldo),0);
  const costoAsoc=fp.filter(f=>f.ventaId).reduce((a,f)=>a+pN(f.total),0);const gastoAsoc=g.filter(x=>x.ventaId).reduce((a,x)=>a+pN(x.total),0);
  const margenBruto=tv-costoAsoc-gastoAsoc;
  const vMes={};v.forEach(x=>{const m=x.fecha?.slice(0,7);if(!m)return;if(!vMes[m])vMes[m]={mes:m,vendido:0,cobrado:0,cant:0};vMes[m].vendido+=pN(x.totalVenta);vMes[m].cobrado+=pN(x.cobrado);vMes[m].cant++;});
  const vMesArr=Object.values(vMes).sort((a,b)=>a.mes.localeCompare(b.mes)).slice(-9);
  const gMes={};g.forEach(x=>{const m=x.fecha?.slice(0,7);if(!m)return;gMes[m]=(gMes[m]||0)+pN(x.total);});
  const meses=[...new Set([...Object.keys(vMes),...Object.keys(gMes)])].sort().slice(-9);
  const vsData=meses.map(m=>({l:mesL(m),ventas:Math.round(vMes[m]?.vendido||0),gastos:Math.round(gMes[m]||0)}));
  const cliMap={};v.forEach(x=>{if(!x.cliente)return;if(!cliMap[x.cliente])cliMap[x.cliente]={n:x.cliente,tv:0,tc:0,fc:0};cliMap[x.cliente].tv+=pN(x.totalVenta);cliMap[x.cliente].tc+=pN(x.cobrado);cliMap[x.cliente].fc++;});
  const topCli=Object.values(cliMap).sort((a,b)=>b.tv-a.tv).slice(0,8);
  const vndMap={};v.forEach(x=>{const vnd=x.vendedor||"Sin asignar";if(!vndMap[vnd])vndMap[vnd]={n:vnd,tv:0,fc:0};vndMap[vnd].tv+=pN(x.totalVenta);vndMap[vnd].fc++;});
  const porVnd=Object.values(vndMap).sort((a,b)=>b.tv-a.tv);
  const estMap={};v.forEach(x=>{estMap[x.estado]=(estMap[x.estado]||0)+1;});
  const estData=Object.entries(estMap).map(([name,value])=>({name,value}));
  const catMap={};g.forEach(x=>{catMap[x.categoria]=(catMap[x.categoria]||0)+pN(x.total);});
  const catData=Object.entries(catMap).sort((a,b)=>b[1]-a[1]).map(([name,value])=>({name,value}));
  const rent=v.map(x=>{const co=fp.filter(f=>f.ventaId===x.id).reduce((a,f)=>a+pN(f.total),0);const ga=g.filter(f=>f.ventaId===x.id).reduce((a,f)=>a+pN(f.total),0);const n=pN(x.montoNeto);const m=n-co-ga;return{c:x.cliente,f:x.nroFactura,d:x.fecha,v:n,co:co+ga,m,p:n?m/n*100:0};}).filter(r=>r.co>0).sort((a,b)=>b.m-a.m);
  const SECS=["gral","ventas","finanzas","clientes","rentabilidad"];
  const SLBL=["🏠 General","📈 Ventas","💰 Finanzas","👥 Clientes","📊 Rentabilidad"];
  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:10}}>
        <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
          {SECS.map((s,i)=><button key={s} onClick={()=>setSec(s)} style={{padding:"8px 16px",borderRadius:9,border:sec===s?"1px solid rgba(16,185,129,.25)":"1px solid transparent",background:sec===s?"rgba(16,185,129,.15)":"transparent",color:sec===s?VERDE:"#475569",cursor:"pointer",fontSize:13,fontWeight:600,fontFamily:"inherit"}}>{SLBL[i]}</button>)}
        </div>
        <BtnExport label="Exportar todo" onClick={()=>{const wb=XLSX.utils.book_new();[["Ventas",v.map(x=>({Fecha:x.fecha,Cliente:x.cliente,Total:x.totalVenta,Cobrado:x.cobrado,Estado:x.estado}))],["Gastos",g.map(x=>({Fecha:x.fecha,Emisor:x.emisor,Categoría:x.categoria,Total:x.total}))],["Clientes",Object.entries(Object.fromEntries(v.map(x=>[x.cliente,x]))).map(([n])=>({Cliente:n}))]].forEach(([name,data])=>{if(data.length)XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(data),name);});XLSX.writeFile(wb,`sasoc_reporte_${new Date().toISOString().slice(0,10)}.xlsx`);}}/>
      </div>
      {sec==="gral"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
            <Met label="Total vendido" val={$(tv)} sub={`${v.length} facturas`} up/><Met label="Total gastos" val={$(tg)} sub={`${g.length} registros`} dn/>
            <Met label="Saldo disponible" val={$(sd)} up/><Met label="Margen bruto est." val={pct(tv?margenBruto/tv*100:0)} sub={$(margenBruto)} up={margenBruto>0} dn={margenBruto<=0}/>
          </div>
          {vsData.length>0&&(
            <Crd>
              <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Ventas vs Gastos por mes</p>
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={vsData} margin={{top:0,right:10,left:0,bottom:0}}>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,.05)"/>
                  <XAxis dataKey="l" tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false}/>
                  <YAxis tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false} tickFormatter={n=>n>=1e6?"$"+(n/1e6).toFixed(1)+"M":n>=1e3?"$"+(n/1e3).toFixed(0)+"K":"$"+n}/>
                  <Tooltip {...TT} formatter={(v,name)=>[$(v),name==="ventas"?"Ventas":"Gastos"]}/>
                  <Bar dataKey="ventas" fill={VERDE} radius={[4,4,0,0]} name="ventas"/>
                  <Bar dataKey="gastos" fill={ROJO} radius={[4,4,0,0]} name="gastos"/>
                </BarChart>
              </ResponsiveContainer>
            </Crd>
          )}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <Crd mb={0}>
              <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Estados de ventas</p>
              {estData.length>0?(
                <div style={{display:"flex",alignItems:"center",gap:16}}>
                  <ResponsiveContainer width={120} height={120}><PieChart><Pie data={estData} cx="50%" cy="50%" innerRadius={32} outerRadius={54} dataKey="value" paddingAngle={3}>{estData.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Pie><Tooltip {...TT}/></PieChart></ResponsiveContainer>
                  <div style={{flex:1}}>{estData.map((d,i)=><div key={d.name} style={{display:"flex",justifyContent:"space-between",marginBottom:6}}><div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:9,height:9,borderRadius:2,background:COLORS[i%COLORS.length],flexShrink:0}}/><span style={{fontSize:12,color:"#94a3b8"}}>{d.name}</span></div><span style={{fontSize:12,fontWeight:700,color:"#e2e8f0"}}>{d.value}</span></div>)}</div>
                </div>
              ):<p style={{color:"#334155",fontSize:13,textAlign:"center",padding:"20px 0"}}>Sin datos</p>}
            </Crd>
            <Crd mb={0}>
              <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Gastos por categoría</p>
              {catData.length>0?(
                <div style={{display:"flex",alignItems:"center",gap:16}}>
                  <ResponsiveContainer width={120} height={120}><PieChart><Pie data={catData} cx="50%" cy="50%" innerRadius={32} outerRadius={54} dataKey="value" paddingAngle={3}>{catData.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Pie><Tooltip {...TT} formatter={v=>[$(v),"Total"]}/></PieChart></ResponsiveContainer>
                  <div style={{flex:1}}>{catData.slice(0,6).map((d,i)=><div key={d.name} style={{display:"flex",justifyContent:"space-between",marginBottom:6}}><div style={{display:"flex",alignItems:"center",gap:6}}><div style={{width:9,height:9,borderRadius:2,background:COLORS[i%COLORS.length],flexShrink:0}}/><span style={{fontSize:11,color:"#94a3b8",maxWidth:100,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{d.name}</span></div><span style={{fontSize:11,fontWeight:700,color:ROJO}}>{$(d.value)}</span></div>)}</div>
                </div>
              ):<p style={{color:"#334155",fontSize:13,textAlign:"center",padding:"20px 0"}}>Sin datos</p>}
            </Crd>
          </div>
        </div>
      )}
      {sec==="ventas"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
            <Met label="Total vendido" val={$(tv)} up/><Met label="Total cobrado" val={$(tc)} up/>
            <Met label="Por cobrar" val={$(tv-tc)} dn={tv>tc}/><Met label="Ticket promedio" val={$(v.length?tv/v.length:0)}/>
          </div>
          {vMesArr.length>0&&(()=>{let a=0;const hd=vMesArr.map(d=>{a+=d.vendido;return{...d,l:mesL(d.mes),acum:Math.round(a)};});const pm=Math.round(tv/vMesArr.length);const mejor=Math.max(...hd.map(d=>d.vendido));return(
            <Crd>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:10}}>
                <div><p style={{fontSize:14,fontWeight:700,color:"#e2e8f0",margin:"0 0 2px"}}>Facturación histórica mensual</p><p style={{fontSize:11,color:"#334155",margin:0}}>Barras = mes · Línea = acumulado · 🏆 mejor</p></div>
                <div style={{textAlign:"right"}}><p style={{fontSize:10,color:"#475569",margin:"0 0 2px",textTransform:"uppercase",fontWeight:700}}>Promedio mensual</p><p style={{fontSize:15,fontWeight:700,color:"#f1f5f9",margin:0}}>{$(pm)}</p></div>
              </div>
              <ResponsiveContainer width="100%" height={240}>
                <BarChart data={hd} margin={{top:10,right:55,left:0,bottom:0}}>
                  <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,.05)"/>
                  <XAxis dataKey="l" tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false}/>
                  <YAxis yAxisId="m" orientation="left" tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false} tickFormatter={n=>n>=1e6?"$"+(n/1e6).toFixed(1)+"M":n>=1e3?"$"+(n/1e3).toFixed(0)+"K":"$"+n}/>
                  <YAxis yAxisId="a" orientation="right" tick={{fill:"#334155",fontSize:10}} axisLine={false} tickLine={false} tickFormatter={n=>n>=1e6?"$"+(n/1e6).toFixed(1)+"M":n>=1e3?"$"+(n/1e3).toFixed(0)+"K":"$"+n}/>
                  <Tooltip {...TT} formatter={(v,name)=>[$(v),name==="vendido"?"Facturación":"Acumulado"]}/>
                  <ReferenceLine yAxisId="m" y={pm} stroke="rgba(245,158,11,.45)" strokeDasharray="6 3"/>
                  <Bar yAxisId="m" dataKey="vendido" name="vendido" radius={[5,5,0,0]}>{hd.map((d,i)=><Cell key={i} fill={d.vendido===mejor?"#fbbf24":i===hd.length-1?"#34d399":VERDE}/>)}</Bar>
                  <Line yAxisId="a" type="monotone" dataKey="acum" name="acum" stroke={AZUL} strokeWidth={2} dot={false}/>
                </BarChart>
              </ResponsiveContainer>
              <div style={{marginTop:16,overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr><Th>Mes</Th><Th>Facturas</Th><Th>Facturación</Th><Th>Cobrado</Th><Th>% cobrado</Th><Th>vs prom.</Th><Th>Acumulado</Th></tr></thead><tbody>{hd.map((d,i)=>{const vp=pm?((d.vendido-pm)/pm*100):0;const esM=d.vendido===mejor;return(<tr key={i} style={{background:esM?"rgba(251,191,36,.04)":"transparent"}}><Td fw={esM}>{d.l} {esM?"🏆":""}</Td><Td>{d.cant}</Td><Td g fw>{$(d.vendido)}</Td><Td>{$(d.cobrado)}</Td><Td><span style={{background:d.vendido&&d.cobrado/d.vendido>.9?"rgba(16,185,129,.12)":"rgba(245,158,11,.1)",color:d.vendido&&d.cobrado/d.vendido>.9?VERDE:AMBAR,padding:"1px 7px",borderRadius:20,fontSize:11,fontWeight:600}}>{d.vendido?pct(d.cobrado/d.vendido*100):"—"}</span></Td><Td><span style={{color:vp>=0?VERDE:ROJO,fontWeight:600}}>{vp>=0?"↑":"↓"}{Math.abs(vp).toFixed(1)}%</span></Td><Td><span style={{color:AZUL,fontWeight:500}}>{$(d.acum)}</span></Td></tr>);})}</tbody><tfoot><tr style={{background:"rgba(16,185,129,.05)"}}><td colSpan={1} style={{padding:"8px 10px",fontSize:13,color:VERDE,fontWeight:700}}>TOTAL</td><td style={{padding:"8px 10px",fontSize:13,color:VERDE,fontWeight:700}}>{v.length}</td><td style={{padding:"8px 10px",fontSize:13,color:VERDE,fontWeight:700}}>{$(tv)}</td><td style={{padding:"8px 10px",fontSize:13,color:VERDE,fontWeight:700}}>{$(tc)}</td><td style={{padding:"8px 10px",fontSize:13,color:VERDE,fontWeight:700}}>{pct(tv?tc/tv*100:0)}</td><td style={{padding:"8px 10px",fontSize:12,color:"#64748b"}}>{$(pm)}/mes</td><td style={{padding:"8px 10px",fontSize:13,color:AZUL,fontWeight:700}}>{$(tv)}</td></tr></tfoot></table></div>
            </Crd>
          );})()}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <Crd mb={0}>
              <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Por vendedor</p>
              {porVnd.map(x=>(<div key={x.n} style={{marginBottom:12}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}><span style={{fontSize:13,color:"#cbd5e1"}}>{x.n} <span style={{color:"#475569",fontSize:11}}>({x.fc} fc)</span></span><span style={{fontSize:13,fontWeight:700,color:VERDE}}>{$(x.tv)}</span></div><div style={{height:5,background:"rgba(255,255,255,.07)",borderRadius:3}}><div style={{height:"100%",width:Math.round(x.tv/porVnd[0].tv*100)+"%",background:VERDE,borderRadius:3}}/></div></div>))}
            </Crd>
            <Crd mb={0}>
              <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Facturas por mes</p>
              {vMesArr.length>0&&(<ResponsiveContainer width="100%" height={180}><BarChart data={vMesArr.map(d=>({l:mesL(d.mes),v:d.cant}))} margin={{top:0,right:10,left:-20,bottom:0}}><CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,.05)"/><XAxis dataKey="l" tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false}/><YAxis tick={{fill:"#475569",fontSize:11}} axisLine={false} tickLine={false}/><Tooltip {...TT} formatter={v=>[v,"Facturas"]}/><Bar dataKey="v" fill={AZUL} radius={[4,4,0,0]}/></BarChart></ResponsiveContainer>)}
            </Crd>
          </div>
        </div>
      )}
      {sec==="finanzas"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
            <Met label="Saldo total ARS" val={$(sd)} up/><Met label="Compras proveedores" val={$(tfp)} dn/>
            <Met label="Total pagado prov." val={$(pagfp)} up/><Met label="Pendiente prov." val={$(tfp-pagfp)} dn={tfp>pagfp}/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
            <Crd mb={0}>
              <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Gastos por categoría</p>
              {catData.map((d,i)=>(<div key={d.name} style={{marginBottom:10}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}><span style={{fontSize:13,color:"#cbd5e1"}}>{d.name}</span><span style={{fontSize:13,fontWeight:700,color:COLORS[i%COLORS.length]}}>{$(d.value)}</span></div><div style={{height:5,background:"rgba(255,255,255,.07)",borderRadius:3}}><div style={{height:"100%",width:Math.round(d.value/catData[0]?.value*100)+"%",background:COLORS[i%COLORS.length],borderRadius:3}}/></div></div>))}
            </Crd>
            <Crd mb={0}>
              <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Estado facturas proveedores</p>
              <div style={{display:"flex",flexDirection:"column",gap:10}}>
                {[{l:"Total facturado",v:$(tfp),c:ROJO},{l:"Total pagado",v:$(pagfp),c:VERDE},{l:"Pendiente pago",v:$(tfp-pagfp),c:AMBAR}].map(x=>(<div key={x.l} style={{padding:"12px 14px",background:"rgba(255,255,255,.03)",borderRadius:10,border:"1px solid rgba(255,255,255,.06)"}}><p style={{fontSize:11,color:"#475569",margin:"0 0 4px",textTransform:"uppercase",fontWeight:700}}>{x.l}</p><p style={{fontSize:20,fontWeight:700,color:x.c,margin:0}}>{x.v}</p></div>))}
              </div>
            </Crd>
          </div>
        </div>
      )}
      {sec==="clientes"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
            <Met label="Clientes únicos" val={Object.keys(cliMap).length} up/><Met label="Total vendido" val={$(tv)} up/>
            <Met label="Ticket promedio" val={$(v.length?tv/v.length:0)}/><Met label="Por cobrar total" val={$(tv-tc)} dn={tv>tc}/>
          </div>
          <Crd mb={16}>
            <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Top clientes por facturación</p>
            {topCli.map(x=>(<div key={x.n} style={{marginBottom:12}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}><div><span style={{fontSize:13,color:"#e2e8f0",fontWeight:600}}>{x.n}</span><span style={{fontSize:11,color:"#475569",marginLeft:8}}>{x.fc} factura{x.fc>1?"s":""}</span></div><div style={{textAlign:"right"}}><span style={{fontSize:13,fontWeight:700,color:VERDE}}>{$(x.tv)}</span>{x.tv>x.tc&&<span style={{fontSize:11,color:ROJO,marginLeft:8}}>({$(x.tv-x.tc)} pend.)</span>}</div></div><div style={{height:5,background:"rgba(255,255,255,.07)",borderRadius:3}}><div style={{height:"100%",width:Math.round(x.tv/topCli[0].tv*100)+"%",background:VERDE,borderRadius:3}}/></div></div>))}
          </Crd>
          <Crd mb={0}>
            <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Detalle por cliente</p>
            <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr><Th>Cliente</Th><Th>Facturas</Th><Th>Total vendido</Th><Th>Total cobrado</Th><Th>Pendiente</Th><Th>% cobrado</Th></tr></thead><tbody>{topCli.map(x=>(<tr key={x.n}><Td fw>{x.n}</Td><Td>{x.fc}</Td><Td g fw>{$(x.tv)}</Td><Td>{$(x.tc)}</Td><Td r={x.tv>x.tc} fw>{$(x.tv-x.tc)}</Td><Td><span style={{background:x.tv&&x.tc/x.tv>.9?"rgba(16,185,129,.12)":"rgba(245,158,11,.1)",color:x.tv&&x.tc/x.tv>.9?VERDE:AMBAR,padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:600}}>{x.tv?pct(x.tc/x.tv*100):"—"}</span></Td></tr>))}</tbody></table></div>
          </Crd>
        </div>
      )}
      {sec==="rentabilidad"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
            <Met label="Margen bruto est." val={pct(tv?margenBruto/tv*100:0)} sub={$(margenBruto)} up={margenBruto>0} dn={margenBruto<=0}/>
            <Met label="Costo de ventas" val={$(costoAsoc)} sub="FC proveedor vinculadas" dn/>
            <Met label="Gastos vinculados" val={$(gastoAsoc)} sub="Gastos a ventas" dn/>
            <Met label="Ops. analizadas" val={rent.length} sub="Con costos asociados"/>
          </div>
          {rent.length>0?(
            <Crd>
              <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:14}}>Rentabilidad por operación</p>
              <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}><thead><tr><Th>Cliente</Th><Th>Fecha</Th><Th>Venta s/IVA</Th><Th>Costos</Th><Th>Margen $</Th><Th>Margen %</Th></tr></thead><tbody>{rent.map((r,i)=>(<tr key={i}><Td fw>{r.c}</Td><Td>{fecha(r.d)}</Td><Td>{$(r.v)}</Td><Td r>{$(r.co)}</Td><Td fw><span style={{color:r.m>=0?VERDE:ROJO}}>{$(r.m)}</span></Td><Td><span style={{background:r.p>25?"rgba(16,185,129,.12)":r.p>10?"rgba(245,158,11,.1)":"rgba(248,113,113,.1)",color:r.p>25?VERDE:r.p>10?AMBAR:ROJO,padding:"2px 8px",borderRadius:20,fontSize:11,fontWeight:700}}>{pct(r.p)}</span></Td></tr>))}</tbody></table></div>
            </Crd>
          ):<Crd><div style={{textAlign:"center",padding:"36px 0"}}><p style={{fontSize:28,marginBottom:8}}>📊</p><p style={{color:"#334155",fontSize:13}}>Asociá facturas de proveedor y gastos a tus ventas para ver la rentabilidad por operación.</p></div></Crd>}
        </div>
      )}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// PAGE: COTIZACIONES
// ══════════════════════════════════════════════════════════════════
const ECOT={empresa:"",contacto:"",tel:"",email:"",producto:"",vendedor:"Ivan Diaz",fecha:hoy(),estado:"Seguimiento",notas:""};

function PageCotizaciones(){
  const[cots,setCots]=useState([]);const[loading,setLoading]=useState(true);const[modal,setModal]=useState(null);
  const[sel,setSel]=useState(null);const[cfm,setCfm]=useState(null);const[search,setSearch]=useState("");
  const[fEst,setFEst]=useState("Todos");const[toast,showToast]=useToast();
  useEffect(()=>{gl(K.cotizaciones).then(d=>{setCots(d);setLoading(false);});},[]);
  const save=async(c)=>{const n=cots.some(x=>x.id===c.id)?cots.map(x=>x.id===c.id?c:x):[c,...cots];setCots(n);await sl(K.cotizaciones,n);setModal(null);showToast("Cotización guardada");};
  const del=async(id)=>{const n=cots.filter(x=>x.id!==id);setCots(n);await sl(K.cotizaciones,n);setCfm(null);showToast("Eliminada");};
  const filt=cots.filter(c=>{const q=search.toLowerCase();return(!search||c.empresa?.toLowerCase().includes(q)||c.contacto?.toLowerCase().includes(q)||c.producto?.toLowerCase().includes(q))&&(fEst==="Todos"||c.estado===fEst);}).sort((a,b)=>(b.fecha||"").localeCompare(a.fecha||""));
  if(loading)return(<Spin/>);
  return (
    <div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
        <Met label="Total cotizaciones" val={cots.length}/><Met label="En seguimiento" val={cots.filter(c=>c.estado==="Seguimiento").length} up/>
        <Met label="Ganadas" val={cots.filter(c=>c.estado==="Ganado").length} up/><Met label="Tasa de cierre" val={cots.length?Math.round(cots.filter(c=>c.estado==="Ganado").length/cots.length*100)+"%":"—"} up/>
      </div>
      <Crd>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,margin:0}}>Pipeline de cotizaciones</p>
          <button style={btnPri} onClick={()=>{setSel(null);setModal("form");}}>+ Nueva cotización</button>
        </div>
        <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap"}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 Buscar empresa, contacto o producto..." style={{...IS,flex:1,minWidth:200}}/>
          <select value={fEst} onChange={e=>setFEst(e.target.value)} style={{...IS,width:160}}><option value="Todos">Todos los estados</option>{ESTADOS_COT.map(x=><option key={x}>{x}</option>)}</select>
        </div>
        {cots.length===0?(<div style={{textAlign:"center",padding:"48px 0"}}><p style={{fontSize:32,marginBottom:10}}>📋</p><p style={{color:"#334155",fontSize:14,fontWeight:600}}>No hay cotizaciones cargadas</p><button style={{...btnPri,marginTop:12}} onClick={()=>setModal("form")}>+ Registrar primer oportunidad</button></div>):(
          <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr><Th>Empresa</Th><Th>Contacto</Th><Th>Teléfono</Th><Th>Producto</Th><Th>Vendedor</Th><Th>Fecha</Th><Th>Estado</Th><Th/></tr></thead>
            <tbody>{filt.map(c=>(
              <tr key={c.id} style={{cursor:"pointer"}} onClick={()=>{setSel(c);setModal("form");}}>
                <Td fw>{c.empresa}</Td><Td>{c.contacto||"—"}</Td><Td mono>{c.tel||"—"}</Td>
                <Td>{c.producto||"—"}</Td><Td>{c.vendedor}</Td><Td>{fecha(c.fecha)}</Td>
                <Td><Bdg t={c.estado} sm/></Td>
                <td style={{padding:"9px 8px",borderBottom:"1px solid rgba(255,255,255,.04)"}} onClick={e=>e.stopPropagation()}>
                  <div className="ra" style={{display:"flex",gap:4}}>
                    <button className="ib" style={{color:ROJO}} onClick={()=>setCfm(c.id)}>🗑</button>
                  </div>
                </td>
              </tr>
            ))}</tbody>
          </table><p style={{fontSize:11,color:"#334155",margin:"10px 4px 0",textAlign:"right"}}>{filt.length} de {cots.length}</p></div>
        )}
      </Crd>
      {modal==="form"&&(<Mdl title={sel?"Editar cotización":"Nueva cotización"} onClose={()=>setModal(null)}>{(()=>{const[f,setF]=useState(sel||ECOT);const set=(k,v)=>setF(p=>({...p,[k]:v}));return(<div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:13}}><Fld label="Empresa *" s2><I value={f.empresa} onChange={e=>set("empresa",e.target.value)} placeholder="Nombre de la empresa"/></Fld><Fld label="Contacto" hint><I value={f.contacto} onChange={e=>set("contacto",e.target.value)}/></Fld><Fld label="Teléfono" hint><I value={f.tel} onChange={e=>set("tel",e.target.value)} placeholder="11 1234-5678"/></Fld><Fld label="Email" hint><I type="email" value={f.email} onChange={e=>set("email",e.target.value)}/></Fld><Fld label="Producto / Servicio"><I value={f.producto} onChange={e=>set("producto",e.target.value)} placeholder="Qué se cotizó"/></Fld><Fld label="Vendedor"><I value={f.vendedor} onChange={e=>set("vendedor",e.target.value)}/></Fld><Fld label="Fecha"><I type="date" value={f.fecha} onChange={e=>set("fecha",e.target.value)}/></Fld><Fld label="Estado"><S value={f.estado} onChange={e=>set("estado",e.target.value)}>{ESTADOS_COT.map(x=><option key={x}>{x}</option>)}</S></Fld><Fld label="Notas" hint s2><T value={f.notas} onChange={e=>set("notas",e.target.value)} placeholder="Detalle, presupuesto, seguimiento..."/></Fld></div><div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:22,borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:18}}><button style={btnSec} onClick={()=>setModal(null)}>Cancelar</button><button style={btnPri} onClick={()=>{if(!f.empresa.trim())return;save({...f,id:f.id||uid()});}}>Guardar</button></div></div>);})()}</Mdl>)}
      {cfm&&(<Cfm msg="¿Eliminar esta cotización?" onOk={()=>del(cfm)} onCancel={()=>setCfm(null)}/>)}
      {toast&&<Tst msg={toast.msg} type={toast.type}/>}
    </div>
  );
}


// ══════════════════════════════════════════════════════════════════
// PAGE: IMPORTAR DATOS
// ══════════════════════════════════════════════════════════════════
const IMP_MODULES = {
  ventas: {
    label:"Ventas / Facturas", icon:"🧾", key:K.ventas, color:VERDE,
    fields:[
      {id:"fecha",label:"Fecha",req:true,hint:"DD/MM/YYYY o YYYY-MM-DD"},
      {id:"cliente",label:"Cliente",req:true},{id:"cuit",label:"CUIT",req:false},
      {id:"nroFactura",label:"N° Factura",req:false},{id:"tipo",label:"Tipo FC",req:false,hint:"A, B o C"},
      {id:"vendedor",label:"Vendedor",req:false},{id:"montoNeto",label:"Monto neto s/IVA",req:true,hint:"Número sin $"},
      {id:"iva",label:"IVA %",req:false,hint:"21"},{id:"totalVenta",label:"Total",req:false,hint:"Se calcula si no está"},
      {id:"cobrado",label:"Cobrado",req:false},{id:"estado",label:"Estado",req:false,hint:"Cobrado, Pendiente..."},
      {id:"productos",label:"Productos",req:false},
    ],
    transform:(row)=>{
      const neto=parseFloat(row.montoNeto)||0;
      const ivaR=(parseFloat(row.iva||"21"))/100;
      const total=row.totalVenta?parseFloat(row.totalVenta):neto*(1+ivaR);
      const cobrado=parseFloat(row.cobrado)||0;
      return{id:uid(),fecha:impNormDate(row.fecha)||hoy(),cliente:row.cliente?.trim()||"",cuit:row.cuit?.trim()||"",nroFactura:row.nroFactura?.trim()||"",tipo:row.tipo?.trim()||"A",vendedor:row.vendedor?.trim()||"Ivan Diaz",montoNeto:neto.toFixed(2),iva:row.iva?.toString()||"21",totalVenta:total.toFixed(2),cobrado:cobrado.toFixed(2),pendiente:(total-cobrado).toFixed(2),estado:row.estado?.trim()||(cobrado>=total?"Cobrado":"Pendiente"),productos:row.productos?.trim()||"",comentario:"",fechaCobro:"",medioCobro:"Banco Macro"};
    }
  },
  proveedores:{
    label:"Proveedores",icon:"🏭",key:K.proveedores,color:AZUL,
    fields:[
      {id:"nombre",label:"Nombre",req:true},{id:"cuit",label:"CUIT",req:false},
      {id:"categoria",label:"Categoría",req:false},{id:"contacto",label:"Contacto",req:false},
      {id:"tel",label:"Teléfono",req:false},{id:"condPago",label:"Cond. pago",req:false,hint:"60 días"},
    ],
    transform:(row)=>({id:uid(),nombre:row.nombre?.trim()||"",cuit:row.cuit?.trim()||"",categoria:row.categoria?.trim()||"Productos",contacto:row.contacto?.trim()||"",tel:row.tel?.trim()||"",email:"",condPago:row.condPago?.trim()||"60 días",estado:"Activo",notas:""})
  },
  gastos:{
    label:"Gastos",icon:"💸",key:K.gastos,color:ROJO,
    fields:[
      {id:"fecha",label:"Fecha",req:true},{id:"emisor",label:"Emisor",req:true},
      {id:"categoria",label:"Categoría",req:false},{id:"montoNeto",label:"Monto neto",req:true},
      {id:"iva",label:"IVA %",req:false,hint:"21"},{id:"total",label:"Total",req:false},
      {id:"medioPago",label:"Medio pago",req:false},{id:"descripcion",label:"Descripción",req:false},
    ],
    transform:(row)=>{const neto=parseFloat(row.montoNeto)||0;const ivaR=(parseFloat(row.iva||"21"))/100;const total=row.total?parseFloat(row.total):neto*(1+ivaR);return{id:uid(),fecha:impNormDate(row.fecha)||hoy(),emisor:row.emisor?.trim()||"",categoria:row.categoria?.trim()||"Varios",tipoComp:"FC A",montoNeto:neto.toFixed(2),iva:row.iva?.toString()||"21",total:total.toFixed(2),medioPago:row.medioPago?.trim()||"Transferencia",responsable:"Ivan Diaz",descripcion:row.descripcion?.trim()||"",notas:"",ventaId:""};}
  },
  cuentas:{
    label:"Cuentas bancarias",icon:"🏦",key:K.cuentas,color:AMBAR,
    fields:[
      {id:"nombre",label:"Nombre cuenta",req:true,hint:"Ej: Cuenta Macro"},
      {id:"tipo",label:"Tipo",req:false,hint:"Cuenta corriente, FCI..."},
      {id:"banco",label:"Banco",req:false},{id:"moneda",label:"Moneda",req:false,hint:"ARS, USD"},
      {id:"saldo",label:"Saldo actual",req:true,hint:"Número sin $"},
    ],
    transform:(row)=>({id:uid(),nombre:row.nombre?.trim()||"",tipo:row.tipo?.trim()||"Cuenta corriente",banco:row.banco?.trim()||"Banco Macro",moneda:row.moneda?.trim()||"ARS",saldo:parseFloat(row.saldo)||0,notas:"",fechaActualizacion:hoy()})
  },
};

function impNormDate(val){
  if(!val)return null;
  const s=val.toString().trim();
  const m1=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if(m1)return`${m1[3]}-${m1[2].padStart(2,"0")}-${m1[1].padStart(2,"0")}`;
  if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;
  if(!isNaN(val)){const d=new Date(Math.round((Number(val)-25569)*86400*1000));return d.toISOString().slice(0,10);}
  return null;
}

function autoMap(cols,fields){
  const syns={fecha:["fecha","date","fecha factura","fecha venta"],cliente:["cliente","empresa","client","company"],cuit:["cuit","cuil"],nroFactura:["numero factura","nro factura","n° factura","factura","invoice"],tipo:["tipo","tipo fc","tipo factura"],vendedor:["vendedor","seller"],montoNeto:["neto","monto neto","neto gravado","base imponible","precio neto"],iva:["iva","iva %","tasa iva"],totalVenta:["total","total venta","importe total","total con iva"],cobrado:["cobrado","pagado","monto cobrado"],estado:["estado","status","estado pago"],productos:["producto","productos","descripcion","detalle","items"],emisor:["emisor","proveedor","supplier"],categoria:["categoria","rubro","category"],total:["total","importe total"],medioPago:["medio pago","forma pago","payment"],descripcion:["descripcion","detalle","description","concepto"],nombre:["nombre","name","razon social"],banco:["banco","bank"],moneda:["moneda","currency"],saldo:["saldo","balance","importe"],contacto:["contacto","contact"],tel:["telefono","tel","phone"],condPago:["condicion pago","cond pago","plazo"]};
  const map={};
  cols.forEach(col=>{
    const cn=col.toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g,"");
    for(const[fid,ss]of Object.entries(syns)){
      if(fields.some(f=>f.id===fid)&&!map[fid]){
        if(ss.some(s=>cn.includes(s)||s.includes(cn)))map[fid]=col;
      }
    }
  });
  return map;
}

function PageImportar(){
  const[step,setStep]=useState("inicio");
  const[modKey,setModKey]=useState("ventas");
  const[sheetData,setSheetData]=useState(null);
  const[mapping,setMapping]=useState({});
  const[mergeMode,setMergeMode]=useState("append");
  const[result,setResult]=useState(null);
  const[toast,showToast]=useToast();
  const fileRef=useState(null);

  const mod=IMP_MODULES[modKey];

  const parseFile=(file)=>{
    const reader=new FileReader();
    reader.onload=(e)=>{
      try{
        const wb=XLSX.read(e.target.result,{type:"array",cellDates:false});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        if(raw.length<2){showToast("El archivo está vacío","error");return;}
        const headers=raw[0].map(h=>h?.toString().trim()).filter(Boolean);
        const rows=raw.slice(1).filter(r=>r.some(c=>c!=="")).map(r=>{const o={};headers.forEach((h,i)=>{o[h]=r[i]??"";});return o;});
        setSheetData({fileName:file.name,rows,cols:headers});
        setMapping(autoMap(headers,mod.fields));
        setStep("mapear");
        showToast(`${rows.length} filas detectadas en "${file.name}"`);
      }catch(err){showToast("Error al leer: "+err.message,"error");}
    };
    reader.readAsArrayBuffer(file);
  };

  const doImport=async()=>{
    setStep("cargando");
    const rows=sheetData.rows;
    let ok=0;const errors=[];
    const remapped=rows.map(row=>{const m={};Object.entries(mapping).forEach(([fid,col])=>{if(col)m[fid]=row[col];});return m;});
    const req=mod.fields.filter(f=>f.req).map(f=>f.id);
    const valid=remapped.filter((r,i)=>{const miss=req.filter(f=>!r[f]||r[f].toString().trim()==="");if(miss.length){errors.push(`Fila ${i+2}: falta ${miss.join(", ")}`);return false;}return true;});
    const transformed=valid.map(r=>{try{return mod.transform(r);}catch(e){return null;}}).filter(Boolean);
    try{
      const existing=await gl(mod.key);
      const newData=mergeMode==="replace"?transformed:[...existing,...transformed];
      await sl(mod.key,newData);
      ok=transformed.length;
    }catch(err){errors.push("Error al guardar: "+err.message);}
    setResult({ok,errors:errors.slice(0,15),total:rows.length});
    setStep("done");
  };

  const reset=()=>{setStep("inicio");setSheetData(null);setMapping({});setResult(null);};

  return(
    <div>
      {/* Header steps */}
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:20}}>
        {["inicio","mapear","done"].map((s,i)=>(
          <div key={s} style={{display:"flex",alignItems:"center",gap:8}}>
            <div style={{width:28,height:28,borderRadius:"50%",background:step===s||( s==="done"&&step==="cargando")?"#065f46":"rgba(255,255,255,.06)",border:`1px solid ${step===s?"#10b981":"rgba(255,255,255,.1)"}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,color:step===s?"#34d399":"#475569"}}>{i+1}</div>
            <span style={{fontSize:12,color:step===s?"#34d399":"#475569",fontWeight:step===s?600:400}}>{s==="inicio"?"Seleccionar":s==="mapear"?"Mapear columnas":"Resultado"}</span>
            {i<2&&<div style={{width:20,height:1,background:"rgba(255,255,255,.08)"}}/>}
          </div>
        ))}
      </div>

      {step==="inicio"&&(
        <div>
          <Crd>
            <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:16}}>1. Seleccioná el módulo de destino</p>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:24}}>
              {Object.entries(IMP_MODULES).map(([key,m])=>(
                <button key={key} onClick={()=>setModKey(key)} style={{padding:"12px 16px",borderRadius:12,border:`1px solid ${modKey===key?m.color:"rgba(255,255,255,.08)"}`,background:modKey===key?`rgba(${modKey===key&&m.color==="#10b981"?"16,185,129":m.color==="#3b82f6"?"59,130,246":m.color==="#f87171"?"248,113,113":"245,158,11"},.08)`:"rgba(255,255,255,.03)",cursor:"pointer",textAlign:"left",fontFamily:"inherit",transition:"all .15s"}}>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <span style={{fontSize:22}}>{m.icon}</span>
                    <div>
                      <p style={{color:modKey===key?m.color:"#e2e8f0",fontWeight:600,fontSize:13,margin:0}}>{m.label}</p>
                      <p style={{color:"#475569",fontSize:11,margin:"2px 0 0"}}>{m.fields.length} campos</p>
                    </div>
                  </div>
                </button>
              ))}
            </div>
            <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,marginBottom:12}}>2. Subí tu archivo Excel o CSV</p>
            <div
              onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f)parseFile(f);}}
              onDragOver={e=>e.preventDefault()}
              onClick={()=>{const inp=document.createElement("input");inp.type="file";inp.accept=".xlsx,.xls,.csv";inp.onchange=e=>{if(e.target.files[0])parseFile(e.target.files[0]);};inp.click();}}
              style={{border:"2px dashed rgba(255,255,255,.12)",borderRadius:14,padding:"40px",textAlign:"center",cursor:"pointer",transition:"border-color .15s"}}
              onMouseEnter={e=>e.currentTarget.style.borderColor="#10b981"}
              onMouseLeave={e=>e.currentTarget.style.borderColor="rgba(255,255,255,.12)"}>
              <p style={{fontSize:32,marginBottom:12}}>📤</p>
              <p style={{color:"#e2e8f0",fontWeight:600,fontSize:14,margin:"0 0 6px"}}>Arrastrá tu archivo acá</p>
              <p style={{color:"#475569",fontSize:13,margin:"0 0 14px"}}>o hacé clic para seleccionar</p>
              <div style={{display:"flex",gap:8,justifyContent:"center"}}>
                {[".xlsx",".xls",".csv"].map(ext=><span key={ext} style={{background:"rgba(255,255,255,.06)",padding:"3px 10px",borderRadius:20,fontSize:12,color:"#64748b"}}>{ext}</span>)}
              </div>
            </div>
          </Crd>
          {/* Formato esperado */}
          <Crd mb={0}>
            <p style={{color:"#e2e8f0",fontWeight:700,fontSize:13,marginBottom:12}}>📋 Formato esperado para <span style={{color:mod.color}}>{mod.label}</span></p>
            <div style={{overflowX:"auto"}}>
              <table style={{borderCollapse:"collapse",fontSize:12}}>
                <thead><tr>{mod.fields.map(f=><th key={f.id} style={{padding:"6px 12px",borderBottom:"1px solid rgba(255,255,255,.08)",color:f.req?VERDE:"#64748b",textAlign:"left",whiteSpace:"nowrap",fontWeight:700}}>{f.label}{f.req?" ★":""}</th>)}</tr></thead>
                <tbody><tr>{mod.fields.map(f=><td key={f.id} style={{padding:"6px 12px",color:"#334155",fontSize:11,whiteSpace:"nowrap"}}>{f.hint||"—"}</td>)}</tr></tbody>
              </table>
            </div>
            <p style={{fontSize:11,color:"#334155",marginTop:8}}>★ Campos obligatorios · La primera fila debe ser el encabezado · El mapeo de columnas es automático</p>
          </Crd>
        </div>
      )}

      {step==="mapear"&&sheetData&&(
        <div>
          <Crd>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:18,flexWrap:"wrap",gap:10}}>
              <div>
                <p style={{color:"#e2e8f0",fontWeight:700,fontSize:14,margin:"0 0 4px"}}>Mapeo de columnas — {mod.label}</p>
                <p style={{fontSize:12,color:"#64748b",margin:0}}>📄 {sheetData.fileName} · {sheetData.rows.length} filas · <span style={{color:VERDE}}>✓ Mapeo automático aplicado</span></p>
              </div>
              <span style={{fontSize:11,color:"#64748b"}}>★ = obligatorio</span>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:18}}>
              {mod.fields.map(f=>(
                <div key={f.id} style={{display:"flex",flexDirection:"column",gap:4}}>
                  <label style={{fontSize:11,color:f.req?VERDE:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:".4px"}}>{f.label}{f.req?" ★":""}</label>
                  {f.hint&&<span style={{fontSize:10,color:"#334155"}}>{f.hint}</span>}
                  <select value={mapping[f.id]||""} onChange={e=>setMapping(p=>({...p,[f.id]:e.target.value}))} style={{...IS,border:`1px solid ${mapping[f.id]?VERDE:"rgba(255,255,255,.1)"}`}}>
                    <option value="">— No mapear —</option>
                    {sheetData.cols.map(c=><option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              ))}
            </div>
            {/* Merge mode */}
            <div style={{padding:"12px 14px",background:"rgba(255,255,255,.03)",borderRadius:10,border:"1px solid rgba(255,255,255,.07)",marginBottom:18}}>
              <p style={{fontSize:11,fontWeight:700,color:"#94a3b8",marginBottom:8,textTransform:"uppercase"}}>Modo de importación</p>
              <div style={{display:"flex",gap:10}}>
                {[{v:"append",l:"➕ Agregar",d:"Añade sin borrar lo existente"},{v:"replace",l:"🔄 Reemplazar",d:"Borra todo y reemplaza"}].map(opt=>(
                  <button key={opt.v} onClick={()=>setMergeMode(opt.v)} style={{flex:1,padding:"10px 12px",borderRadius:9,border:`1px solid ${mergeMode===opt.v?(opt.v==="replace"?"#ef4444":"#10b981"):"rgba(255,255,255,.09)"}`,background:mergeMode===opt.v?(opt.v==="replace"?"rgba(239,68,68,.08)":"rgba(16,185,129,.08)"):"transparent",color:mergeMode===opt.v?(opt.v==="replace"?ROJO:VERDE):"#64748b",cursor:"pointer",textAlign:"left",fontFamily:"inherit"}}>
                    <p style={{fontWeight:700,fontSize:13,margin:"0 0 2px"}}>{opt.l}</p>
                    <p style={{fontSize:11,margin:0,opacity:.7}}>{opt.d}</p>
                  </button>
                ))}
              </div>
            </div>
            {/* Preview */}
            <p style={{fontSize:12,color:"#64748b",fontWeight:700,textTransform:"uppercase",marginBottom:8}}>Vista previa (primeras 5 filas)</p>
            <div style={{overflowX:"auto",maxHeight:200,overflowY:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead style={{position:"sticky",top:0,background:"#0d1929"}}>
                  <tr>{mod.fields.filter(f=>mapping[f.id]).map(f=><th key={f.id} style={{textAlign:"left",padding:"6px 10px",fontSize:10,fontWeight:700,color:"#334155",textTransform:"uppercase",borderBottom:"1px solid rgba(255,255,255,.07)",whiteSpace:"nowrap"}}>{f.label}</th>)}</tr>
                </thead>
                <tbody>
                  {sheetData.rows.slice(0,5).map((row,i)=>(
                    <tr key={i}>{mod.fields.filter(f=>mapping[f.id]).map(f=><td key={f.id} style={{padding:"6px 10px",borderBottom:"1px solid rgba(255,255,255,.04)",color:"#cbd5e1",whiteSpace:"nowrap",maxWidth:150,overflow:"hidden",textOverflow:"ellipsis"}}>{row[mapping[f.id]]??""}</td>)}</tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Crd>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
            <button style={btnSec} onClick={reset}>Cancelar</button>
            <button style={btnPri} onClick={doImport}>Importar {sheetData.rows.length} registros →</button>
          </div>
        </div>
      )}

      {step==="cargando"&&(
        <div style={{textAlign:"center",padding:"80px 0"}}>
          <div style={{width:52,height:52,border:"4px solid rgba(16,185,129,.2)",borderTopColor:VERDE,borderRadius:"50%",margin:"0 auto 20px",animation:"spin .8s linear infinite"}}/>
          <p style={{fontSize:17,fontWeight:700,color:"#e2e8f0",margin:"0 0 8px"}}>Importando datos...</p>
          <p style={{color:"#64748b",fontSize:13}}>Procesando {sheetData?.rows.length} registros en {mod.label}</p>
        </div>
      )}

      {step==="done"&&result&&(
        <div>
          <div style={{background:result.ok>0?"rgba(16,185,129,.06)":"rgba(239,68,68,.06)",borderRadius:16,padding:"32px",border:`1px solid ${result.ok>0?"rgba(16,185,129,.2)":"rgba(239,68,68,.2)"}`,textAlign:"center",marginBottom:20}}>
            <p style={{fontSize:48,marginBottom:12}}>{result.ok>0?"✅":"❌"}</p>
            <p style={{fontSize:20,fontWeight:700,color:"#f1f5f9",margin:"0 0 8px"}}>{result.ok} de {result.total} registros importados correctamente</p>
            <p style={{color:"#64748b",fontSize:13,margin:0}}>Módulo: {mod.label} · Modo: {mergeMode==="replace"?"Reemplazar":"Agregar"}</p>
            {result.errors.length>0&&(
              <div style={{marginTop:16,textAlign:"left",background:"rgba(239,68,68,.06)",borderRadius:10,padding:"12px 14px"}}>
                <p style={{fontSize:12,color:ROJO,fontWeight:700,marginBottom:6}}>⚠ {result.errors.length} filas con errores:</p>
                {result.errors.map((e,i)=><p key={i} style={{fontSize:12,color:"#fca5a5",margin:"2px 0"}}>• {e}</p>)}
              </div>
            )}
          </div>
          <div style={{display:"flex",gap:10,justifyContent:"center"}}>
            <button style={btnPri} onClick={reset}>+ Importar otro módulo</button>
          </div>
        </div>
      )}
      {toast&&<Tst msg={toast.msg} type={toast.type}/>}
    </div>
  );
}


// ══════════════════════════════════════════════════════════════════
// PAGE: COMISIONES
// ══════════════════════════════════════════════════════════════════
const K_COM = "sasoc:comisiones_config:v1";

function PageComisiones(){
  const[ventas,setVentas]=useState([]);
  const[config,setConfig]=useState({}); // {vendedor: {pct, nombre}}
  const[loading,setLoading]=useState(true);
  const[mes,setMes]=useState(hoy().slice(0,7));
  const[modalConf,setModalConf]=useState(false);
  const[liquidadas,setLiquidadas]=useState({}); // {ventaId: true}
  const[toast,showToast]=useToast();

  useEffect(()=>{
    Promise.all([gl(K.ventas),gl(K_COM)]).then(([v,c])=>{
      setVentas(v);
      // config es objeto {vendedor: {pct, liquidadas: {ventaId:true}}}
      const cfg=Array.isArray(c)?{}:(c||{});
      setConfig(cfg);
      // Merge all liquidadas
      const liq={};
      Object.values(cfg).forEach(vc=>{ Object.assign(liq, vc.liquidadas||{}); });
      setLiquidadas(liq);
      setLoading(false);
    });
  },[]);

  const saveConfig=async(cfg)=>{
    setConfig(cfg);
    await sl(K_COM, cfg);
  };

  // Vendedores únicos
  const vendedores=[...new Set(ventas.map(v=>v.vendedor).filter(Boolean))];

  // Ventas cobradas del mes seleccionado (estado Cobrado o cobrado > 0)
  const ventasMes=ventas.filter(v=>{
    const enMes=v.fecha?.slice(0,7)===mes;
    const cobrada=(v.estado==="Cobrado"||pN(v.cobrado)>0);
    return enMes&&cobrada;
  });

  // Calcular comisiones por vendedor
  const comisiones=vendedores.map(vnd=>{
    const pct=pN(config[vnd]?.pct||0);
    const vVnd=ventasMes.filter(v=>v.vendedor===vnd);
    const totalCobrado=vVnd.reduce((a,v)=>a+pN(v.cobrado),0);
    const pendLiquidar=vVnd.filter(v=>!liquidadas[v.id]);
    const yaLiquidadas=vVnd.filter(v=>liquidadas[v.id]);
    const montoPend=pendLiquidar.reduce((a,v)=>a+pN(v.cobrado),0);
    const montoLiq=yaLiquidadas.reduce((a,v)=>a+pN(v.cobrado),0);
    const comPend=montoPend*pct/100;
    const comLiq=montoLiq*pct/100;
    return{vnd,pct,vVnd,totalCobrado,pendLiquidar,yaLiquidadas,montoPend,montoLiq,comPend,comLiq,total:totalCobrado*pct/100};
  }).filter(c=>c.vVnd.length>0||config[c.vnd]);

  const totalComPend=comisiones.reduce((a,c)=>a+c.comPend,0);
  const totalComLiq=comisiones.reduce((a,c)=>a+c.comLiq,0);

  const liquidarVendedor=async(vnd)=>{
    const c=comisiones.find(x=>x.vnd===vnd);
    if(!c||!c.pendLiquidar.length)return;
    const nuevasLiq={...liquidadas};
    c.pendLiquidar.forEach(v=>{nuevasLiq[v.id]=true;});
    setLiquidadas(nuevasLiq);
    // Guardar en config del vendedor
    const newCfg={...config,[vnd]:{...config[vnd],liquidadas:{...(config[vnd]?.liquidadas||{}),...Object.fromEntries(c.pendLiquidar.map(v=>[v.id,true]))}}};
    await saveConfig(newCfg);
    showToast(`Comisión de ${vnd} marcada como liquidada`);
  };

  const exportComisiones=()=>{
    const rows=[];
    comisiones.forEach(c=>{
      c.vVnd.forEach(v=>{
        rows.push({Vendedor:c.vnd,Mes:mes,Fecha:v.fecha,Cliente:v.cliente,"N° Factura":v.nroFactura||"—","Monto Cobrado":pN(v.cobrado),"% Comisión":c.pct+"%","Comisión":pN(v.cobrado)*c.pct/100,Estado:liquidadas[v.id]?"Liquidada":"Pendiente"});
      });
    });
    exportExcel(rows,"comisiones_sasoc","Comisiones");
  };

  if(loading)return(<Spin/>);

  return(
    <div>
      {/* Header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20,flexWrap:"wrap",gap:12}}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <input type="month" value={mes} onChange={e=>setMes(e.target.value)} style={{...IS,width:"auto",fontSize:14,fontWeight:600}}/>
          <span style={{fontSize:13,color:"#64748b"}}>{ventasMes.length} ventas cobradas en el período</span>
        </div>
        <div style={{display:"flex",gap:8}}>
          <BtnExport onClick={exportComisiones} label="Exportar liquidación"/>
          <button style={btnPri} onClick={()=>setModalConf(true)}>⚙ Configurar %</button>
        </div>
      </div>

      {/* Métricas */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:20}}>
        <Met label="Total cobrado (mes)" val={$(ventasMes.reduce((a,v)=>a+pN(v.cobrado),0))} up/>
        <Met label="Comisiones a pagar" val={$(totalComPend)} sub="Pendientes de liquidar" dn={totalComPend>0}/>
        <Met label="Ya liquidado" val={$(totalComLiq)} sub="Este mes" up/>
        <Met label="Vendedores activos" val={comisiones.length} sub="Con ventas en el período"/>
      </div>

      {vendedores.length===0?(
        <Crd><div style={{textAlign:"center",padding:"48px 0"}}>
          <p style={{fontSize:32,marginBottom:10}}>💰</p>
          <p style={{color:"#334155",fontSize:14,fontWeight:600}}>No hay ventas registradas aún</p>
          <p style={{color:"#1e3a5f",fontSize:13}}>Cargá ventas con el campo "Vendedor" completo para ver las comisiones.</p>
        </div></Crd>
      ):(
        comisiones.map(c=>(
          <Crd key={c.vnd}>
            {/* Cabecera vendedor */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16,flexWrap:"wrap",gap:10}}>
              <div style={{display:"flex",alignItems:"center",gap:12}}>
                <Av name={c.vnd} size={44}/>
                <div>
                  <h3 style={{color:"#f1f5f9",fontSize:16,fontWeight:700,margin:"0 0 4px"}}>{c.vnd}</h3>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{background:"rgba(16,185,129,.12)",color:VERDE,padding:"2px 10px",borderRadius:20,fontSize:12,fontWeight:600}}>{c.pct}% de comisión</span>
                    <span style={{fontSize:12,color:"#64748b"}}>{c.vVnd.length} ventas cobradas</span>
                  </div>
                </div>
              </div>
              <div style={{display:"flex",gap:10,alignItems:"center"}}>
                <div style={{textAlign:"right"}}>
                  <p style={{fontSize:11,color:"#475569",margin:"0 0 2px",textTransform:"uppercase",fontWeight:700}}>A liquidar</p>
                  <p style={{fontSize:20,fontWeight:700,color:c.comPend>0?AMBAR:"#64748b",margin:0}}>{$(c.comPend)}</p>
                </div>
                {c.pendLiquidar.length>0&&(
                  <button onClick={()=>liquidarVendedor(c.vnd)} style={{...btnPri,background:"rgba(245,158,11,.15)",borderColor:AMBAR,color:AMBAR}}>
                    ✓ Marcar liquidado
                  </button>
                )}
              </div>
            </div>

            {/* Resumen */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:16}}>
              {[
                {l:"Total cobrado",v:$(c.totalCobrado),c:"#e2e8f0"},
                {l:"Comisión total",v:$(c.total),c:VERDE},
                {l:"Ya liquidado",v:$(c.comLiq),c:"#64748b"},
              ].map((m,i)=>(
                <div key={i} style={{background:"rgba(255,255,255,.03)",borderRadius:10,padding:"10px 14px",border:"1px solid rgba(255,255,255,.06)"}}>
                  <p style={{fontSize:10,color:"#475569",margin:"0 0 4px",textTransform:"uppercase",fontWeight:700}}>{m.l}</p>
                  <p style={{fontSize:16,fontWeight:700,color:m.c,margin:0}}>{m.v}</p>
                </div>
              ))}
            </div>

            {/* Detalle ventas */}
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr>
                  <Th>Fecha</Th><Th>Cliente</Th><Th>N° Factura</Th>
                  <Th>Monto cobrado</Th><Th>Comisión ({c.pct}%)</Th><Th>Estado</Th>
                </tr></thead>
                <tbody>
                  {c.vVnd.map(v=>{
                    const liq=liquidadas[v.id];
                    const com=pN(v.cobrado)*c.pct/100;
                    return(
                      <tr key={v.id}>
                        <Td>{fecha(v.fecha)}</Td>
                        <Td fw>{v.cliente}</Td>
                        <Td mono>{v.nroFactura||"—"}</Td>
                        <Td g fw>{$(v.cobrado)}</Td>
                        <Td fw><span style={{color:liq?"#64748b":AMBAR,fontWeight:700}}>{$(com)}</span></Td>
                        <Td>
                          <span style={{background:liq?"rgba(16,185,129,.1)":"rgba(245,158,11,.1)",color:liq?VERDE:AMBAR,padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:600}}>
                            {liq?"✓ Liquidada":"Pendiente"}
                          </span>
                        </Td>
                      </tr>
                    );
                  })}
                </tbody>
                <tfoot>
                  <tr style={{background:"rgba(255,255,255,.03)"}}>
                    <td colSpan={3} style={{padding:"8px 10px",fontSize:12,color:"#64748b",fontWeight:700}}>TOTAL {c.vnd.toUpperCase()}</td>
                    <td style={{padding:"8px 10px",fontSize:13,color:VERDE,fontWeight:700}}>{$(c.totalCobrado)}</td>
                    <td style={{padding:"8px 10px",fontSize:13,color:AMBAR,fontWeight:700}}>{$(c.total)}</td>
                    <td style={{padding:"8px 10px"}}/>
                  </tr>
                </tfoot>
              </table>
            </div>
          </Crd>
        ))
      )}

      {/* Modal configuración % */}
      {modalConf&&(
        <Mdl title="Configurar comisiones por vendedor" onClose={()=>setModalConf(false)}>
          {(()=>{
            const[tmpCfg,setTmpCfg]=useState({...config});
            const allVnd=[...new Set([...vendedores,...Object.keys(config)])];
            return(
              <div>
                <p style={{fontSize:13,color:"#64748b",marginBottom:16}}>Definí el porcentaje de comisión para cada vendedor. Se aplica sobre el monto cobrado de cada venta.</p>
                {allVnd.length===0&&<p style={{color:"#334155",fontSize:13,textAlign:"center",padding:"20px 0"}}>No hay vendedores registrados aún. Cargá ventas con el campo Vendedor completo.</p>}
                {allVnd.map(vnd=>(
                  <div key={vnd} style={{display:"flex",alignItems:"center",gap:12,padding:"12px 0",borderBottom:"1px solid rgba(255,255,255,.06)"}}>
                    <Av name={vnd} size={36}/>
                    <div style={{flex:1}}>
                      <p style={{fontSize:13,fontWeight:600,color:"#e2e8f0",margin:"0 0 2px"}}>{vnd}</p>
                      <p style={{fontSize:11,color:"#475569",margin:0}}>Vendedor</p>
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <I type="number" min="0" max="100" step="0.5"
                        value={tmpCfg[vnd]?.pct||""}
                        onChange={e=>setTmpCfg(p=>({...p,[vnd]:{...(p[vnd]||{}),pct:parseFloat(e.target.value)||0}}))}
                        placeholder="0"
                        style={{width:80,textAlign:"center"}}
                      />
                      <span style={{fontSize:14,color:"#64748b",fontWeight:600}}>%</span>
                    </div>
                  </div>
                ))}
                {/* Agregar vendedor manual */}
                <div style={{marginTop:16,padding:"12px 14px",background:"rgba(255,255,255,.03)",borderRadius:10,border:"1px solid rgba(255,255,255,.06)"}}>
                  <p style={{fontSize:12,color:"#475569",margin:"0 0 8px",fontWeight:700,textTransform:"uppercase"}}>Agregar vendedor manualmente</p>
                  {(()=>{
                    const[nuevoVnd,setNuevoVnd]=useState("");
                    const[nuevoPct,setNuevoPct]=useState("");
                    return(
                      <div style={{display:"flex",gap:8,alignItems:"center"}}>
                        <I value={nuevoVnd} onChange={e=>setNuevoVnd(e.target.value)} placeholder="Nombre del vendedor" style={{flex:1}}/>
                        <I type="number" value={nuevoPct} onChange={e=>setNuevoPct(e.target.value)} placeholder="%" style={{width:70,textAlign:"center"}}/>
                        <button style={btnPri} onClick={()=>{if(!nuevoVnd.trim())return;setTmpCfg(p=>({...p,[nuevoVnd.trim()]:{...(p[nuevoVnd.trim()]||{}),pct:parseFloat(nuevoPct)||0}}));setNuevoVnd("");setNuevoPct("");}}>+</button>
                      </div>
                    );
                  })()}
                </div>
                <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:22,borderTop:"1px solid rgba(255,255,255,.07)",paddingTop:18}}>
                  <button style={btnSec} onClick={()=>setModalConf(false)}>Cancelar</button>
                  <button style={btnPri} onClick={async()=>{await saveConfig(tmpCfg);setModalConf(false);showToast("Configuración guardada");}}>Guardar configuración</button>
                </div>
              </div>
            );
          })()}
        </Mdl>
      )}
      {toast&&<Tst msg={toast.msg} type={toast.type}/>}
    </div>
  );
}

// ══════════════════════════════════════════════════════════════════
// SHELL PRINCIPAL
// ══════════════════════════════════════════════════════════════════
const NAV=[
  {id:"dashboard",label:"Dashboard",icon:"▦"},
  {id:"ventas",label:"Ventas",icon:"🧾",sec:"COMERCIAL"},
  {id:"clientes",label:"Clientes",icon:"👥"},
  {id:"cotizaciones",label:"Cotizaciones",icon:"📋"},
  {id:"proveedores",label:"Proveedores",icon:"🚚",sec:"OPERACIONES"},
  {id:"gastos",label:"Gastos / Pagos",icon:"💳"},
  {id:"cashflow",label:"Cash Flow",icon:"💵",sec:"FINANZAS"},
  {id:"estadisticas",label:"Estadísticas",icon:"📈",sec:"ANÁLISIS"},
  {id:"importar",label:"Importar datos",icon:"📤",sec:"HERRAMIENTAS"},
  {id:"comisiones",label:"Comisiones",icon:"💰",sec:null},
];

const PAGES={
  dashboard:<PageDashboard/>,ventas:<PageVentas/>,clientes:<PageClientes/>,
  cotizaciones:<PageCotizaciones/>,proveedores:<PageProveedores/>,
  gastos:<PageGastos/>,cashflow:<PageCashFlow/>,estadisticas:<PageEstadisticas/>,
  importar:<PageImportar/>,
  comisiones:<PageComisiones/>,
};

const LABELS={dashboard:"Dashboard",ventas:"Ventas",clientes:"Clientes",cotizaciones:"Cotizaciones",proveedores:"Proveedores",gastos:"Gastos / Pagos",cashflow:"Cash Flow",estadisticas:"Estadísticas",importar:"Importar datos",comisiones:"Comisiones"};

export default function App({user=null, onLogout=null}){
  const[page,setPage]=useState("dashboard");
  return (
    <div style={{display:"flex",minHeight:"100vh",background:"#0b1120",fontFamily:"'DM Sans',system-ui,sans-serif",color:"#e2e8f0"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
        @keyframes spin{to{transform:rotate(360deg)}}
        @keyframes fadein{from{opacity:0;transform:translateY(5px)}to{opacity:1;transform:none}}
        ::-webkit-scrollbar{width:4px;height:4px}::-webkit-scrollbar-thumb{background:#1e3a5f;border-radius:2px}
        input,select,textarea{font-family:inherit}
        input:focus,select:focus,textarea:focus{border-color:#10b981!important;outline:none}
        tr:hover td{background:rgba(255,255,255,.025)!important}
        .ra{opacity:0;transition:opacity .12s}.ib{padding:4px 8px;border-radius:6px;border:1px solid rgba(255,255,255,.09);background:rgba(255,255,255,.04);color:#94a3b8;cursor:pointer;font-size:11px}
        tr:hover .ra{opacity:1}
        .pcard,.ccard{background:#131f35;border:1px solid rgba(255,255,255,.07);border-radius:14px;padding:16px;cursor:pointer;transition:border-color .15s,transform .15s}
        .pcard:hover,.ccard:hover{border-color:rgba(16,185,129,.35);transform:translateY(-1px)}
      `}</style>
      {/* Sidebar */}
      <div style={{width:215,background:"#0d1929",borderRight:"1px solid rgba(255,255,255,.07)",display:"flex",flexDirection:"column",flexShrink:0,minHeight:"100vh"}}>
        <div style={{padding:"18px 14px 14px",borderBottom:"1px solid rgba(255,255,255,.07)"}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:34,height:34,borderRadius:9,background:"linear-gradient(135deg,#065f46,#10b981)",display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:800,fontSize:16}}>S</div>
            <div><p style={{color:"#f1f5f9",fontSize:13,fontWeight:700,margin:0}}>SASOC Merch</p><p style={{color:"#334155",fontSize:10,margin:0}}>CRM Empresarial</p></div>
          </div>
        </div>
        <nav style={{flex:1,padding:"10px 8px",overflowY:"auto"}}>
          {NAV.map(item=>(
            <div key={item.id}>
              {item.sec&&<p style={{fontSize:9,color:"#1e3a5f",textTransform:"uppercase",letterSpacing:"0.9px",padding:"10px 10px 3px",fontWeight:700,margin:0}}>{item.sec}</p>}
              <button onClick={()=>setPage(item.id)} style={{display:"flex",alignItems:"center",gap:9,width:"100%",padding:"8px 10px",borderRadius:8,border:"none",background:page===item.id?"rgba(16,185,129,.12)":"transparent",color:page===item.id?"#34d399":"#64748b",fontSize:13,cursor:"pointer",textAlign:"left",fontWeight:page===item.id?700:400,fontFamily:"inherit",transition:"all .12s"}}>
                <span style={{fontSize:14}}>{item.icon}</span>{item.label}
              </button>
            </div>
          ))}
        </nav>
        <div style={{padding:10,borderTop:"1px solid rgba(255,255,255,.07)"}}>
          <div style={{display:"flex",alignItems:"center",gap:8,padding:"8px 10px",background:"rgba(255,255,255,.04)",borderRadius:8}}>
            <div style={{width:28,height:28,borderRadius:"50%",background:"#065f46",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"#6ee7b7",fontWeight:700,flexShrink:0}}>
              {user?.email?.[0]?.toUpperCase()||"U"}
            </div>
            <div style={{flex:1,minWidth:0}}>
              <p style={{color:"#e2e8f0",fontSize:11,fontWeight:700,margin:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{user?.email||"Usuario"}</p>
              <p style={{color:"#334155",fontSize:10,margin:0}}>Administrador</p>
            </div>
            {onLogout&&<button onClick={onLogout} title="Cerrar sesión" style={{background:"transparent",border:"none",color:"#475569",cursor:"pointer",fontSize:14,padding:"2px 4px",flexShrink:0}} >⎋</button>}
          </div>
        </div>
      </div>
      {/* Main */}
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <div style={{background:"#0d1929",borderBottom:"1px solid rgba(255,255,255,.07)",padding:"13px 24px",display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
          <h2 style={{color:"#f1f5f9",fontSize:17,fontWeight:700,margin:0}}>{LABELS[page]}</h2>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <div style={{width:8,height:8,borderRadius:"50%",background:"#10b981"}}/>
            <span style={{fontSize:12,color:"#475569"}}>Datos en tiempo real</span>
          </div>
        </div>
        <div key={page} style={{flex:1,overflowY:"auto",padding:"22px 24px",animation:"fadein .18s ease"}}>
          {PAGES[page]}
        </div>
      </div>
    </div>
  );
}
