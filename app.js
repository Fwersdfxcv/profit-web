/* 利润分析管理软件 · 纯前端版（无后端，数据不出本机）
 * 复用 profit-manager 的 clean_report / dashboard 算法，端口径 1:1 一致。
 * 读取：SheetJS(XLSX)  写出含 live 公式的 xlsx：ExcelJS  看板：ECharts(dashboard.html)
 */
'use strict';

/* ---------------- 常量（与 clean_report.STD_COLS 对应） ---------------- */
const STD = [
  "企业名称","企业编码","姓名","手机号","所属部门","isbn","商品ID","商品名称","兑换时间","兑换单编号",
  "兑换单类型","订单来源","子订单号","CP名/品牌名称","类型","出版社","出版日期","作者","关联标签","一级分类",
  "二级分类","新一级分类","新二级分类","书城分类","购买数量","码洋价","成本价","销售价","实付单价","实付金额",
  "订单金额","税费","退款数量","退款积分小计","消费积分合计","第三方订单编号"
];
const TYPE_ORDER = ["权益","纸书","文创"];
const MONEY_IDX = new Set([25,26,27,28,29,30,31,33,34]); // 金额列（分），写盘时 ÷100
const MONEY_HEADERS = new Set(["码洋价","成本价","销售价","实付单价","实付金额","订单金额","税费",
  "退款积分小计","消费积分合计","码洋价小计","成本小计","利润"]);
const CENTS = 100.0;
const FLAG_PAPER_ZERO_TAX = "全员阅读平台";
const BLACKLIST_SET = new Set((window.BUILTIN_BLACKLIST_IDS || []));
const REASON_HIGH_COST_DISCOUNT = 0.65;
const REASON_LOW_REDEEM_RATE = 0.60;

/* 清洗表表头（39 列，含插入的小计/利润列） */
const CLEAN_HEADER = STD.slice(0,26)
  .concat(["码洋价小计"], STD.slice(26,27), ["成本小计"], STD.slice(27,35), ["利润"], STD.slice(35,36));

/* ---------------- 工具函数 ---------------- */
function toNum(v){
  if(v===null||v===undefined) return 0.0;
  if(typeof v==='boolean') return 0.0;
  if(typeof v==='number'){ if(isNaN(v)) return 0.0; return v; }
  let s=String(v).trim().replace(/,/g,'').replace(/'/g,'').replace(/¥/g,'');
  if(!s||['-','--','None','nan'].includes(s)) return 0.0;
  const f=parseFloat(s); return isNaN(f)?0.0:f;
}
function toInt(v){ return Math.round(toNum(v)); }
function norm(s){ return s===null||s===undefined ? "" : String(s).trim().replace('（','(').replace('）',')'); }
function normId(v){
  let s=norm(v).trim();
  for(const q of ["'",'"',"'",'"',"'",'"']) s=s.replace(q,'');
  s=s.replace(/\s+/g,'');
  if(!s) return "";
  if(/^\d+$/.test(s)) return s;
  const f=parseFloat(s);
  if(!isNaN(f) && f===Math.floor(f)) return String(Math.floor(f));
  return s;
}
function round2(v){ return Math.round(v*100)/100; }

/* ---------------- 地理位置识别（移植自 store.py detect_geo） ----------------
 * 源数据无省/市/支行列，地理位置只能从【文件名】识别：
 *   - GEO_HINTS：城市/支行 关键词词表（命中即归属，省/市优先，支行其次）
 *   - CHINA_GEO：全量省市表（兜底识别城市，使任意城市文件名都能识别）
 *   - 支行兜底：取文件名首个机构/支行 token（如「苏州中行」）
 */
const GEO_HINTS = {
  // 省/市
  "苏州":{"province":"江苏省","city":"苏州市","town":""},
  "南京":{"province":"江苏省","city":"南京市","town":""},
  "无锡":{"province":"江苏省","city":"无锡市","town":""},
  "常州":{"province":"江苏省","city":"常州市","town":""},
  "徐州":{"province":"江苏省","city":"徐州市","town":""},
  "南通":{"province":"江苏省","city":"南通市","town":""},
  "连云港":{"province":"江苏省","city":"连云港市","town":""},
  "淮安":{"province":"江苏省","city":"淮安市","town":""},
  "盐城":{"province":"江苏省","city":"盐城市","town":""},
  "扬州":{"province":"江苏省","city":"扬州市","town":""},
  "镇江":{"province":"江苏省","city":"镇江市","town":""},
  "泰州":{"province":"江苏省","city":"泰州市","town":""},
  "宿迁":{"province":"江苏省","city":"宿迁市","town":""},
  "北京":{"province":"北京市","city":"北京市","town":""},
  "上海":{"province":"上海市","city":"上海市","town":""},
  "广州":{"province":"广东省","city":"广州市","town":""},
  "深圳":{"province":"广东省","city":"深圳市","town":""},
  "杭州":{"province":"浙江省","city":"杭州市","town":""},
  "成都":{"province":"四川省","city":"成都市","town":""},
  "武汉":{"province":"湖北省","city":"武汉市","town":""},
  "西安":{"province":"陕西省","city":"西安市","town":""},
  "重庆":{"province":"重庆市","city":"重庆市","town":""},
  "天津":{"province":"天津市","city":"天津市","town":""},
  // 支行/镇（仅填 town；省/市留空，交由用户手选或上游城市词表补全）
  "姑苏":{"province":"","city":"","town":"姑苏支行"},
  "吴中":{"province":"","city":"","town":"吴中支行"},
  "相城":{"province":"","city":"","town":"相城支行"},
  "园区":{"province":"","city":"","town":"园区支行"},
  "新区":{"province":"","city":"","town":"新区支行"},
  "高新":{"province":"","city":"","town":"新区支行"},
  "张家港":{"province":"","city":"","town":"张家港分行"},
  "常熟":{"province":"","city":"","town":"常熟支行"},
  "昆山":{"province":"","city":"","town":"昆山支行"},
  "太仓":{"province":"","city":"","town":"太仓支行"},
  "吴江":{"province":"","city":"","town":"吴江支行"}
};
function geoShort(full){ return String(full==null?'':full).replace(/(省|市|自治区|自治州|地区|盟|县|区|特别行政区)$/,''); }
function detectGeo(filename){
  const name=String(filename||'').replace(/\.xlsx?$/i,'');
  let province='', city='', town='';
  for(const kw in GEO_HINTS){
    if(name.indexOf(kw)>=0){ const g=GEO_HINTS[kw];
      if(g.province && !province) province=g.province;
      if(g.city && !city) city=g.city;
      if(g.town && !town) town=g.town;
    }
  }
  // 全量省市表兜底识别城市（避免只写城市名而不在词表内时漏识别）
  if(!city && window.CHINA_GEO){
    outer: for(const p in window.CHINA_GEO){
      for(const c of window.CHINA_GEO[p]){ const s=geoShort(c);
        if(s && s.length>=2 && name.indexOf(s)>=0){ city=c; province=p; break outer; } }
    }
  }
  // 支行兜底：文件名首个机构/支行 token；但若与 省/市 同名则视为无支行（避免把城市本身当支行）
  if(!town){
    const head=name.split(/[_\-—\s/]+/)[0].replace(/^\d+/,'');
    const tok=head.split(/(明细|兑换|报表|数据|统计|表|名单|清单|附件|汇总|对应|参考|\d)/)[0].trim();
    if(tok && geoShort(tok)!==geoShort(city) && geoShort(tok)!==geoShort(province)) town=tok;
  }
  // 期号：优先匹配 4 位年（YYYY.M），再退回 2 位（M.D），避免 4 位年份被截断成 26.5
  const m=name.match(/(\d{4}[.\-_]\d{1,2})|(\d{1,2}[.\-_]\d{1,2})/);
  const period=m? m[0].replace(/[-_]/g,'.'):'';
  return {province,city,town,period};
}

/* ---------------- base64 内嵌资源解码（模板/看板/ECharts 已打包进 JS，无需 fetch） ---------------- */
function b64ToU8(b64){ const bin=atob(b64); const u8=new Uint8Array(bin.length);
  for(let i=0;i<bin.length;i++) u8[i]=bin.charCodeAt(i); return u8; }
function b64ToText(b64){ return new TextDecoder('utf-8').decode(b64ToU8(b64)); }

/* 从清洗表数据中推断项目总名称：取「企业名称」出现频次最高的值 */
function inferProjectName(records, fileName){
  if(!records || !records.length) return String(fileName||'未命名').replace(/\.xlsx?$/i,'');
  const map={}; let first='';
  for(const r of records){
    const name=norm(r["企业名称"]);
    if(!name) continue;
    if(!first) first=name;
    map[name]=(map[name]||0)+1;
  }
  let best=first; let bestCnt=0;
  for(const k in map){ if(map[k]>bestCnt){ bestCnt=map[k]; best=k; } }
  return best || String(fileName||'未命名').replace(/\.xlsx?$/i,'');
}

/* ---------------- 读取源 xlsx（SheetJS） ---------------- */
function loadRecords(buf){
  const wb = XLSX.read(new Uint8Array(buf), {type:'array'});
  const records=[]; let rawRows=0; const skipped=[];
  for(const name of wb.SheetNames){
    const ws = wb.Sheets[name];
    if(!ws) continue;
    const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null, raw:false});
    if(!rows.length) continue;
    const header = rows[0].map(h=>norm(h));
    const idx={};
    header.forEach((h,i)=>{ if(h && !(h in idx)) idx[h]=i; });
    const missing = ["类型","消费积分合计","购买数量","码洋价","成本价"].filter(c=>!(norm(c) in idx));
    if(missing.length){ skipped.push(name+'（缺列:'+missing.join('/')+'）'); continue; }
    for(let j=1;j<rows.length;j++){
      const row=rows[j];
      if(row==null || row.every(v=>v===null||v==='')) continue;
      rawRows++;
      const rec={};
      for(const c of STD){ const i=idx[norm(c)]; rec[c]=(i!==undefined && i<row.length)? row[i] : ""; }
      records.push(rec);
    }
  }
  return {records, rawRows, skipped};
}

/* ---------------- 清洗 / 分组 ---------------- */
function cleanRecords(records){
  const kept=[], droppedZero={n:0}, droppedNoType={n:0};
  for(const r of records){
    if(!norm(r["类型"])){ droppedNoType.n++; continue; }
    if(toNum(r["消费积分合计"])===0){ droppedZero.n++; continue; }
    kept.push(r);
  }
  return {kept, droppedZero:droppedZero.n, droppedNoType:droppedNoType.n};
}
function groupByType(records){
  const g={};
  for(const r of records){ const t=norm(r["类型"]); (g[t]=g[t]||[]).push(r); }
  const ordered = TYPE_ORDER.filter(t=>t in g).concat(Object.keys(g).filter(t=>!TYPE_ORDER.includes(t)));
  const out={}; ordered.forEach(t=>out[t]=g[t]); return out;
}
function summarize(rows){
  const orders=rows.length;
  const points=rows.reduce((a,r)=>a+toNum(r["消费积分合计"]),0);
  const cost=rows.reduce((a,r)=>a+toNum(r["成本价"])*toInt(r["购买数量"]),0);
  const mark=rows.reduce((a,r)=>a+toNum(r["码洋价"])*toInt(r["购买数量"]),0);
  return {
    "兑换单数": orders,
    "兑换金额含税": round2(points/CENTS),
    "成本含税": round2(cost/CENTS),
    "码洋合计": round2(mark/CENTS),
    "平均动销成本折扣": mark? round2(cost/mark) : 0.0
  };
}
function detectPaperZero(records){
  const flag=FLAG_PAPER_ZERO_TAX;
  for(const r of records){ for(const k in r){ if(r[k]!==null && r[k]!==undefined && String(r[k]).indexOf(flag)>=0) return true; } }
  return false;
}

/* ---------------- 利润表：税率映射（运行时从模板读取，避免硬编码漂移） ---------------- */
let TAX_BASE = {revRateTotal:0.06, byType:{}};
function parseTaxBase(buf){
  const wb=XLSX.read(new Uint8Array(buf),{type:'array'});
  const ws=wb.Sheets['项目数据分析（总）'];
  if(!ws) return;
  // raw:true —— 税率单元格是百分比格式（0.06 显示为 6%），必须取原始数值而非格式化字符串，
  // 否则 toNum("6%")=6，会把 (1+0.06) 误算成 (1+6)，导致不含税口径差 100 倍。
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
  const byType={}; let revRateTotal=0.06;
  rows.forEach(row=>{
    const label=norm(row[0]); if(!label) return;
    if(label==='总'){ revRateTotal=toNum(row[2]); return; }
    const short=label.replace('订单','');
    if(['权益','纸书','文创'].includes(short)) byType[short]={cost:toNum(row[1]), rev:toNum(row[2])};
  });
  TAX_BASE={revRateTotal, byType};
}
function buildTaxInfo(paperZero){
  const cost={}, rev={};
  for(const t of ['权益','纸书','文创']){ if(TAX_BASE.byType[t]){ cost[t]=TAX_BASE.byType[t].cost; rev[t]=TAX_BASE.byType[t].rev; } }
  if(paperZero && rev['纸书']!==undefined) rev['纸书']=0;
  return {rev_rate:TAX_BASE.revRateTotal, cost_rate_by_type:cost, rev_rate_by_type:rev};
}
function buildCaliberRows(stats, taxInfo){
  const caliber={};
  for(const t in stats){
    const rr=taxInfo.rev_rate_by_type[t]!==undefined?taxInfo.rev_rate_by_type[t]:taxInfo.rev_rate;
    const cr=taxInfo.cost_rate_by_type[t]||0;
    const revInc=stats[t]["兑换金额含税"], costInc=stats[t]["成本含税"];
    const revExcl=rr? revInc/(1+rr):revInc;
    const costExcl=cr? costInc/(1+cr):costInc;
    const profitExcl=revExcl-costExcl;
    caliber[t]={
      "兑换单数":stats[t]["兑换单数"],
      "兑换金额含税":revInc, "成本含税":costInc,
      "利润含税":round2(revInc-costInc), "利润率含税":revInc?round2((revInc-costInc)/revInc):0,
      "兑换金额不含税":round2(revExcl), "成本不含税":round2(costExcl),
      "利润不含税":round2(profitExcl), "利润率不含税":revExcl?round2(profitExcl/revExcl):0,
      "动销折扣":stats[t]["平均动销成本折扣"], "成本税率":cr, "收入税率":rr
    };
  }
  return caliber;
}

/* ---------------- 看板聚合（端口 compute） ---------------- */
function reasonOf(cost_f, mark_f, rev_f, avg_qty){
  const disc = mark_f? cost_f/mark_f : 0;
  const redeem = mark_f? rev_f/mark_f : 0;
  let base;
  if(disc>=REASON_HIGH_COST_DISCOUNT) base='成本高（动销折扣偏高）';
  else if(redeem<REASON_LOW_REDEEM_RATE) base='积分兑换价偏低（兑换折扣过小）';
  else base='成本与定价双高（综合）';
  if(avg_qty>=3) base+=' / 量大';
  return base;
}
function compute(kept, taxInfo){
  const rev_rate = Math.max(0, taxInfo.rev_rate||0);
  const crByType = taxInfo.cost_rate_by_type||{};
  const rrByType = taxInfo.rev_rate_by_type||{};
  const prod={}, daily={}, typeTot={};
  const total={orders:0,rev:0,cost:0,profit:0,qty:0,mark_f:0};
  const totalExcl={rev:0,cost:0,profit:0};
  let negRows=0;

  for(const r of kept){
    const t=norm(r["类型"]);
    if(!t || t==='合计' || t==='动销成本折扣') continue;
    const qty=toInt(r["购买数量"]);
    const costFen=toNum(r["成本价"]), markFen=toNum(r["码洋价"]), pointsFen=toNum(r["消费积分合计"]);
    if(pointsFen===0) continue;
    const rev=pointsFen/CENTS, cst=costFen*qty/CENTS, pf=(pointsFen-costFen*qty)/CENTS;
    const pid=norm(r["商品ID"]), pname=norm(r["商品名称"]);
    const phone=norm(r["手机号"]).replace(/'/g,''), name=norm(r["姓名"]);
    const user=phone||name;
    const dstr=norm(r["兑换时间"]).slice(0,10);
    const costRate=Math.max(0, crByType[t]||0);
    const rRate=Math.max(0, (rrByType[t]!==undefined?rrByType[t]:rev_rate));
    const revExcl=rRate? rev/(1+rRate):rev;
    const cstExcl=costRate? cst/(1+costRate):cst;
    const pfExcl=revExcl-cstExcl;
    const key=t+'|'+pid+'|'+pname;
    let p=prod[key];
    if(!p){ p={type:t,pid,pname,orders:0,qty:0,rev_f:0,cost_f:0,mark_f:0,rev_excl_f:0,cost_excl_f:0,users:{}}; prod[key]=p; }
    p.orders++; p.qty+=qty; p.rev_f+=rev; p.cost_f+=cst; p.mark_f+=markFen*qty/CENTS;
    p.users[user]=(p.users[user]||0)+1; p.rev_excl_f+=revExcl; p.cost_excl_f+=cstExcl;
    let dd=daily[dstr]; if(!dd){ dd={profit:0,revenue:0,orders:0,neg_orders:0,neg_amount:0}; daily[dstr]=dd; }
    dd.profit+=pf; dd.revenue+=rev; dd.orders++; if(pf<0){ dd.neg_orders++; dd.neg_amount+=pf; }
    let tt=typeTot[t]; if(!tt){ tt={orders:0,rev:0,cost:0,profit:0,mark_f:0,rev_excl:0,cost_excl:0,profit_excl:0}; typeTot[t]=tt; }
    tt.orders++; tt.rev+=rev; tt.cost+=cst; tt.profit+=pf; tt.mark_f+=markFen*qty/CENTS;
    tt.rev_excl+=revExcl; tt.cost_excl+=cstExcl; tt.profit_excl+=pfExcl;
    total.orders++; total.rev+=rev; total.cost+=cst; total.profit+=pf; total.qty+=qty; total.mark_f+=markFen*qty;
    totalExcl.rev+=revExcl; totalExcl.cost+=cstExcl; totalExcl.profit+=pfExcl;
  }

  // 负利润 / 刷单 / 全量
  const neg_products=[], brush_products=[], reason_dist={};
  let brush_qty_total=0, brush_orders_total=0, bl_hit_total=0;
  for(const key in prod){
    const p=prod[key];
    const parts=key.split('|'); const t=parts[0], pid=parts[1], pname=parts[2];
    const nid=normId(pid);
    const m_id = BLACKLIST_SET.has(nid)? nid : null;
    if(m_id){ bl_hit_total++;
      if(p.qty>=2){
        const profit=p.rev_f-p.cost_f;
        brush_products.push({type:t,pid,pname,orders:p.orders,qty:p.qty,
          revenue:round2(p.rev_f),cost:round2(p.cost_f),profit:round2(profit),
          users:Object.keys(p.users).length,avg_qty:round2(p.qty/p.orders),neg:profit<0,
          bl_match:m_id,bl_match_type:"商品id",bl_qty:null});
        brush_qty_total+=p.qty; brush_orders_total+=p.orders;
      }
    }
    if(p.rev_f===0) continue;
    const profit=p.rev_f-p.cost_f; if(profit>=0) continue;
    const cost_rate=p.rev_f? p.cost_f/p.rev_f : 0;
    const avg_qty=p.orders? p.qty/p.orders : 0;
    const reason=reasonOf(p.cost_f,p.mark_f,p.rev_f,avg_qty);
    const base_reason=reason.split(' /')[0];
    reason_dist[base_reason]=(reason_dist[base_reason]||0)+1;
    const repeat_users=Object.values(p.users).filter(c=>c>1).length;
    neg_products.push({type:t,pid,pname,orders:p.orders,qty:p.qty,
      revenue:round2(p.rev_f),cost:round2(p.cost_f),profit:round2(profit),
      cost_rate:round2(cost_rate),discount:round2(p.cost_f/p.mark_f||0),
      redeem_rate:round2(p.rev_f/p.mark_f||0),avg_qty:round2(avg_qty),reason,
      users:Object.keys(p.users).length,
      repeat_rate:round2(p.orders? Object.values(p.users).reduce((a,c)=>a+(c>1?c-1:0),0)/p.orders:0),
      max_user_orders:Math.max(...Object.values(p.users)),
      repeat_users,
      top1_share:round2(p.orders&&Object.keys(p.users).length? Math.max(...Object.values(p.users))/p.orders:0)});
  }
  neg_products.sort((a,b)=>a.profit-b.profit);
  brush_products.sort((a,b)=>(-a.qty) || a.pname.localeCompare(b.pname,'zh'));

  const all_products=[];
  for(const key in prod){
    const p=prod[key]; if(p.orders<=0) continue;
    const parts=key.split('|');
    const profit_f=p.rev_f-p.cost_f;
    const rev_excl=p.rev_excl_f, cost_excl=p.cost_excl_f, profit_excl=rev_excl-cost_excl;
    all_products.push({type:parts[0],pid:parts[1],pname:parts[2],orders:p.orders,qty:p.qty,
      revenue:round2(p.rev_f),cost:round2(p.cost_f),profit:round2(profit_f),
      revenue_excl:round2(rev_excl),cost_excl:round2(cost_excl),profit_excl:round2(profit_excl),
      margin_excl:rev_excl?round2(profit_excl/rev_excl):0,
      users:Object.keys(p.users).length,avg_qty:round2(p.qty/p.orders)});
  }
  all_products.sort((a,b)=>b.profit_excl-a.profit_excl);

  const type_summary=[];
  for(const t in typeTot){
    const tt=typeTot[t];
    const neg_in_type=neg_products.filter(x=>x.type===t).length;
    const neg_amt=neg_products.filter(x=>x.type===t).reduce((a,x)=>a+x.profit,0);
    const avg_disc=tt.mark_f? round2(tt.cost/tt.mark_f):0;
    const rev_excl_t=tt.rev_excl, cost_excl_t=tt.cost_excl, profit_excl_t=tt.profit_excl;
    const rev_incl_t=tt.rev, cost_incl_t=tt.cost, profit_incl_t=tt.profit;
    type_summary.push({type:t,orders:tt.orders,
      revenue:round2(rev_excl_t),cost:round2(cost_excl_t),profit:round2(profit_excl_t),
      margin:rev_excl_t?round2(profit_excl_t/rev_excl_t):0,
      revenue_incl:round2(rev_incl_t),cost_incl:round2(cost_incl_t),profit_incl:round2(profit_incl_t),
      margin_incl:rev_incl_t?round2(profit_incl_t/rev_incl_t):0,
      avg_discount:avg_disc,neg_products:neg_in_type,neg_amount:round2(neg_amt)});
  }
  type_summary.sort((a,b)=>b.profit-a.profit);

  const daily_list=[]; let cum=0;
  Object.keys(daily).sort().forEach(d=>{ const dd=daily[d]; cum+=dd.profit;
    daily_list.push({date:d,profit:round2(dd.profit),revenue:round2(dd.revenue),orders:dd.orders,
      neg_orders:dd.neg_orders,neg_amount:round2(dd.neg_amount),cum_profit:round2(cum)}); });

  // 负利润订单行数
  for(const r of kept){
    const t=norm(r["类型"]); if(!t||t==='合计'||t==='动销成本折扣') continue;
    if(toNum(r["消费积分合计"])===0) continue;
    if(toNum(r["消费积分合计"])-toNum(r["成本价"])*toInt(r["购买数量"]) < 0) negRows++;
  }

  const n_days=daily_list.length;
  const rev_excl=totalExcl.rev, cost_excl=totalExcl.cost, profit_excl=totalExcl.profit;
  const margin_excl=rev_excl? profit_excl/rev_excl : 0;
  const kpi={
    "总兑换单数":total.orders,"总购买量":total.qty,
    "总兑换金额(含税)":round2(total.rev),"总兑换金额(不含税)":round2(rev_excl),
    "总成本(含税)":round2(total.cost),"总成本(不含税)":round2(cost_excl),
    "总利润":round2(total.profit),"总利润(不含税)":round2(profit_excl),
    "整体利润率":round2(margin_excl),"收入税率":round2(rev_rate),
    "负利润商品数":neg_products.length,"负利润订单行数":negRows,
    "负利润金额":round2(neg_products.reduce((a,x)=>a+x.profit,0)),
    "刷单对象商品数":brush_products.length,"刷单涉及购买量":brush_qty_total,
    "刷单涉及订单数":brush_orders_total,"黑名单命中商品数":bl_hit_total,
    "数据起期":daily_list.length?daily_list[0].date:'',"数据止期":daily_list.length?daily_list[daily_list.length-1].date:'',
    "天数":n_days,
    "日均利润":n_days?round2(total.profit/n_days):0,
    "日均利润(不含税)":n_days?round2(profit_excl/n_days):0
  };
  return {kpi,type_summary,neg_products,brush_products,products:all_products,reason_dist,
    reason_meta:{unit:"商品数（按「类型 + 商品ID + 商品名称」去重后的负利润 SKU 数，非订单行数）",
      discount_threshold:REASON_HIGH_COST_DISCOUNT,redeem_threshold:REASON_LOW_REDEEM_RATE,qty_threshold:3,
      defs:{"成本高（动销折扣偏高）":"动销折扣 = 成本 ÷ 码洋 ≥ "+REASON_HIGH_COST_DISCOUNT,
        "积分兑换价偏低（兑换折扣过小）":"兑换率 = 收入(积分) ÷ 码洋 < "+REASON_LOW_REDEEM_RATE,
        "成本与定价双高（综合）":"既不满足「成本高」也不满足「兑换价偏低」，由成本与定价共同导致"},
      extra:"若平均购买量 ≥ 3 件商品，会在原因后追加「/ 量大」，表示单笔亏损被放大。"},
    daily:daily_list};
}

/* ---------------- 写出清洗表（ExcelJS，含 live 公式） ---------------- */
async function writeCleanXlsx(groups){
  const wb=new ExcelJS.Workbook();
  for(const tname in groups){
    const rows=groups[tname];
    const ws=wb.addWorksheet(String(tname).slice(0,31));
    ws.addRow(CLEAN_HEADER);
    // 按利润升序排，便于红字集中在底部
    const sorted=rows.slice().sort((a,b)=>(toNum(a["消费积分合计"])-toNum(a["成本价"])*toInt(a["购买数量"]))
      - (toNum(b["消费积分合计"])-toNum(b["成本价"])*toInt(b["购买数量"])));
    sorted.forEach(rec=>{
      const vals=STD.map(c=>rec[c]!==undefined?rec[c]:"");
      vals[24]=toInt(vals[24]);
      for(let i=25;i<=34;i++){ let v=toNum(vals[i]); if(MONEY_IDX.has(i)) v=round2(v/CENTS); vals[i]=v; }
      const out=vals.slice(0,26).concat([null],[vals[26]],[null],vals.slice(27,35),[null],[vals[35]]);
      ws.addRow(out);
    });
    const N=sorted.length, last=N+1, tr=N+2, dr=N+3;
    for(let i=0;i<N;i++){
      const r=i+2;
      ws.getCell(r,27).value={formula:`Z${r}*Y${r}`};
      ws.getCell(r,29).value={formula:`AB${r}*Y${r}`};
      ws.getCell(r,38).value={formula:`AK${r}-AB${r}*Y${r}`};
      const profitFen=toNum(sorted[i]["消费积分合计"])-toNum(sorted[i]["成本价"])*toInt(sorted[i]["购买数量"]);
      if(profitFen<0){
        ws.getCell(r,38).font={color:{argb:'FF9C0006'},bold:true};
        for(let c=1;c<=39;c++) ws.getCell(r,c).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFC7CE'}};
      }
    }
    // 合计 + 动销成本折扣
    ws.getCell(tr,25).value="合计"; ws.getCell(tr,25).font={bold:true};
    [27,29,37,38].forEach(col=>{ const L=colToLetter(col);
      ws.getCell(tr,col).value={formula:`SUM(${L}2:${L}${last})`};
      ws.getCell(tr,col).font={bold:true}; ws.getCell(tr,col).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF2CC'}}; });
    ws.getCell(dr,25).value="动销成本折扣"; ws.getCell(dr,25).font={bold:true};
    ws.getCell(dr,27).value={formula:`AC${tr}/AA${tr}`};
    ws.getCell(dr,27).font={bold:true}; ws.getCell(dr,27).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF2CC'}};
    ws.getCell(dr,27).numFmt='0.0000';
    // 表头样式 + 金额列格式
    for(let c=1;c<=39;c++){
      const cell=ws.getCell(1,c); cell.font={bold:true};
      cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFDDEBF7'}};
      cell.alignment={horizontal:'center',vertical:'center'};
      if(MONEY_HEADERS.has(CLEAN_HEADER[c-1])) ws.getColumn(c).numFmt='0.00';
    }
    ws.views=[{state:'frozen',ySplit:1}];
    for(let c=1;c<=39;c++) ws.getColumn(c).width=14;
    ws.getColumn(8).width=40;
  }
  const buf=await wb.xlsx.writeBuffer();
  return new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function colToLetter(col){ let s=''; col--; while(col>=0){ s=String.fromCharCode(65+(col%26))+s; col=Math.floor(col/26)-1; } return s; }

/* ---------------- 写出利润表（基于模板，含 live 公式） ---------------- */
async function writeProfitXlsx(stats, paperZero, tplBuf){
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.load(tplBuf);
  const ws=wb.getWorksheet('项目数据分析（总）');
  if(!ws) throw new Error('模板缺少 sheet: 项目数据分析（总）');
  const rowOf={权益:4,纸书:5,文创:6};
  for(const t in stats){
    const row=rowOf[t]; if(!row) continue;
    ws.getCell(row,4).value=stats[t]["兑换单数"];
    ws.getCell(row,6).value=stats[t]["兑换金额含税"]; ws.getCell(row,6).numFmt='0.00';
    ws.getCell(row,9).value=stats[t]["成本含税"];     ws.getCell(row,9).numFmt='0.00';
    ws.getCell(row,12).value=stats[t]["平均动销成本折扣"]; ws.getCell(row,12).numFmt='0.0000';
  }
  if(paperZero && rowOf['纸书']){ ws.getCell(rowOf['纸书'],3).value=0; ws.getCell(rowOf['纸书'],3).numFmt='0.00'; }
  const buf=await wb.xlsx.writeBuffer();
  return new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}

/* ---------------- 看板渲染（注入 dashboard.html） ----------------
 * mode='mini'：用于主页 iframe 内嵌，无 header/nav，更紧凑
 * mode='full'：用于下载，独立打开仍是完整网页
 */
async function buildDashboardHtml(data, mode='mini'){
  const key = mode==='full' ? 'EMBED_DASHBOARD_B64' : 'EMBED_DASHBOARD_MINI_B64';
  if(!window[key]) throw new Error('看板模板未加载（embed_dashboard.js）');
  let html=b64ToText(window[key]);
  const json=JSON.stringify(data).replace(/<\/script>/g,'<\\/script>');
  html=html.replace('/*__DATA__*/', json);
  // 将看板内的 echarts 外链替换为内联脚本，避免 file:// 或静态托管下相对路径加载失败
  if(window.EMBED_ECHARTS_B64){
    const echartsText=b64ToText(window.EMBED_ECHARTS_B64);
    html=html.replace('<script src="./echarts.min.js"></script>', '<script>'+echartsText+'</script>');
  }
  return html;
}

/* ---------------- IndexedDB 本地持久化 ---------------- */
const DB_NAME='profit-web', STORE='datasets', CFG='config';
function openDB(){ return new Promise((res,rej)=>{ const r=indexedDB.open(DB_NAME,2);
  r.onupgradeneeded=()=>{ const db=r.result;
    if(!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE,{keyPath:'id'});
    if(!db.objectStoreNames.contains(CFG)) db.createObjectStore(CFG,{keyPath:'key'}); };
  r.onsuccess=()=>res(r.result); r.onerror=()=>rej(r.error); }); }
async function saveDS(ds){ const db=await openDB(); return new Promise((res,rej)=>{ const tx=db.transaction(STORE,'readwrite');
  tx.objectStore(STORE).put(ds); tx.oncomplete=()=>res(); tx.onerror=()=>rej(tx.error); }); }
async function listDS(){ const db=await openDB(); return new Promise((res,rej)=>{ const tx=db.transaction(STORE,'readonly');
  const req=tx.objectStore(STORE).getAll(); req.onsuccess=()=>res(req.result||[]); req.onerror=()=>rej(tx.error); }); }
async function getDS(id){ const db=await openDB(); return new Promise((res,rej)=>{ const tx=db.transaction(STORE,'readonly');
  const req=tx.objectStore(STORE).get(id); req.onsuccess=()=>res(req.result); req.onerror=()=>rej(req.error); }); }
async function delDS(id){ const db=await openDB(); return new Promise((res,rej)=>{ const tx=db.transaction(STORE,'readwrite');
  tx.objectStore(STORE).delete(id); tx.oncomplete=()=>res(); tx.onerror=()=>rej(tx.error); }); }
/* 清空整个数据集仓库（兜底清理：不依赖单条 id，解决异常记录删不掉的问题） */
async function clearAllDS(){
  if(!confirm('确认清空浏览器内全部已归档数据集？此操作不可恢复（不影响你设置的工作目录里的文件）。')) return;
  try{
    const db=await openDB();
    await new Promise((res,rej)=>{ const tx=db.transaction(STORE,'readwrite');
      tx.objectStore(STORE).clear(); tx.oncomplete=()=>res(); tx.onerror=()=>rej(tx.error); });
    await refreshList();
    setStatus('已清空本地全部已归档数据集。');
  }catch(e){ console.error(e); alert('清空失败：'+e.message); }
}
async function saveConfig(key,val){ const db=await openDB(); return new Promise((res,rej)=>{ const tx=db.transaction(CFG,'readwrite');
  tx.objectStore(CFG).put({key,val}); tx.oncomplete=()=>res(); tx.onerror=()=>rej(tx.error); }); }
async function getConfig(key){ const db=await openDB(); return new Promise((res,rej)=>{ const tx=db.transaction(CFG,'readonly');
  const req=tx.objectStore(CFG).get(key); req.onsuccess=()=>res(req.result?req.result.val:null); req.onerror=()=>rej(tx.error); }); }

/* ---------------- 保存目录（File System Access API） ----------------
 * 点「设置保存目录」选一个本地文件夹，之后三个下载直接写进去；
 * 不支持的浏览器（Firefox/Safari 或非 https/localhost）自动退回默认下载。
 */
let SAVE_DIR_HANDLE=null, SAVE_DIR_NAME='';
let WORK_DIR_HANDLE=null, WORK_DIR_NAME='', WORK_DIR_ITEMS=[];
function hasFSAccess(){ return typeof window.showDirectoryPicker==='function'; }
async function setSaveDir(){
  if(!hasFSAccess()){ alert('当前浏览器不支持「选择文件夹」（需 Chrome / Edge 且通过 https 或 localhost 访问）。\n已继续使用浏览器默认下载目录。'); return; }
  try{
    const handle=await window.showDirectoryPicker({mode:'readwrite'});
    SAVE_DIR_HANDLE=handle; SAVE_DIR_NAME=handle.name;
    try{ await saveConfig('saveDir', handle); }catch(e){ console.warn('保存目录句柄持久化失败（不影响本次使用）', e); }
    if($('dir-label')) $('dir-label').textContent='📂 '+handle.name;
    if($('clear-dir')) $('clear-dir').style.display='';
    setStatus('已设置保存目录：'+handle.name+'（下载将直接写入此文件夹）');
  }catch(e){ if(e.name!=='AbortError'){ console.error(e); alert('选择目录失败：'+e.message); } }
}
async function clearSaveDir(){ SAVE_DIR_HANDLE=null; SAVE_DIR_NAME='';
  try{ await saveConfig('saveDir', null); }catch(e){}
  if($('dir-label')) $('dir-label').textContent=''; if($('clear-dir')) $('clear-dir').style.display='none';
  setStatus('已取消自定义保存目录，下载将使用浏览器默认目录。'); }
async function saveBlob(blob, filename){
  if(SAVE_DIR_HANDLE){
    try{ await writeIntoDir(SAVE_DIR_HANDLE, blob, filename); setStatus('已保存到目录【'+SAVE_DIR_NAME+'】→ '+filename); return; }
    catch(e){
      if((e.name==='NotAllowedError'||e.name==='SecurityError')){
        try{ if(await SAVE_DIR_HANDLE.requestPermission({mode:'readwrite'})==='granted'){ await writeIntoDir(SAVE_DIR_HANDLE, blob, filename); setStatus('已保存到目录【'+SAVE_DIR_NAME+'】→ '+filename); return; } }catch(_){}
      }
      console.warn('写入目录失败，退回默认下载', e);
    }
  }
  downloadBlob(blob, filename);
}
async function writeIntoDir(handle, blob, filename){
  const fh=await handle.getFileHandle(filename,{create:true});
  const w=await fh.createWritable(); await w.write(blob); await w.close();
}

/* 取得「当前数据集」在工作目录下的子文件夹句柄（自动建/复用），
   让「手动下载」与「自动归档」落在同一个 {编号}__{名称} 文件夹里 */
async function currentWorkDirHandle(){
  if(!WORK_DIR_HANDLE) return null;
  let folderName = (CURRENT && CURRENT.folderName) || null;
  if(!folderName){
    const shortId = ((CURRENT && CURRENT.savedId) ? CURRENT.savedId : guid()).slice(0,12);
    folderName = shortId + '__' + safeName(CURRENT ? CURRENT.projectName : '未命名');
    if(CURRENT) CURRENT.folderName = folderName;
  }
  return await WORK_DIR_HANDLE.getDirectoryHandle(folderName, {create:true});
}
/* 下载按钮专用：写入工作目录子文件夹；无工作目录时提示先设置 */
async function saveBlobToWorkDir(blob, filename){
  if(!WORK_DIR_HANDLE){
    setStatus('未设置「工作目录」，下载将使用浏览器默认目录。');
    alert('请先点击左侧「设置工作目录」选择一个本地文件夹，之后下载会自动归类到该文件夹下的 {编号}__{名称} 子文件夹中。');
    downloadBlob(blob, filename);
    return;
  }
  const dirH = await currentWorkDirHandle();
  if(!dirH){ downloadBlob(blob, filename); return; }
  try{
    await writeIntoDir(dirH, blob, filename);
    setStatus('已保存到工作目录【'+WORK_DIR_NAME+' / '+(CURRENT?CURRENT.folderName:'')+'】→ '+filename);
    // 若本次是首次通过下载按钮落到本地，回写 IndexedDB 让左侧显示「本地+浏览器」
    if(CURRENT && CURRENT.savedId && CURRENT.folderName){
      try{ await markLocalFolder(CURRENT.savedId, CURRENT.folderName); }catch(_){}
    }
    // 刷新左侧列表，让用户立刻看到本地+浏览器的合并状态
    try{ await refreshWorkDir(); }catch(_){}
    return;
  }catch(e){
    if(e.name==='NotAllowedError'||e.name==='SecurityError'){
      try{
        if(await WORK_DIR_HANDLE.requestPermission({mode:'readwrite'})==='granted'){
          const d=await currentWorkDirHandle();
          await writeIntoDir(d,blob,filename);
          setStatus('已保存（重新授权后）：'+filename);
          if(CURRENT && CURRENT.savedId && CURRENT.folderName){
            try{ await markLocalFolder(CURRENT.savedId, CURRENT.folderName); }catch(_){}
          }
          try{ await refreshWorkDir(); }catch(_){}
          return;
        }
      }catch(_){}
    }
    console.warn('写入子文件夹失败，退回默认下载', e);
    alert('写入工作目录失败（'+e.message+'），已退回浏览器默认下载。');
    downloadBlob(blob, filename);
  }
}

/* ---------------- 工作目录（扫描本地文件夹作为数据源） ---------------- */
async function setWorkDir(){
  if(!hasFSAccess()){ alert('当前浏览器不支持「选择文件夹」（需 Chrome / Edge 且通过 https 或 localhost 访问）。'); return; }
  try{
    const handle=await window.showDirectoryPicker({mode:'readwrite'});
    WORK_DIR_HANDLE=handle; WORK_DIR_NAME=handle.name;
    SAVE_DIR_HANDLE=handle; SAVE_DIR_NAME=handle.name; // 工作目录同时作为下载目录
    try{ await saveConfig('workDir', handle); }catch(e){ console.warn('工作目录句柄持久化失败', e); }
    renderWorkDir();
    setStatus('已设置工作目录：'+handle.name+'，正在扫描…');
    await refreshWorkDir();
    setStatus('已设置工作目录：'+handle.name+'，找到 '+WORK_DIR_ITEMS.filter(i=>i.source==='file').length+' 个 xlsx 文件。');
  }catch(e){ if(e.name!=='AbortError'){ console.error(e); alert('选择目录失败：'+e.message); } }
}
async function clearWorkDir(){ WORK_DIR_HANDLE=null; WORK_DIR_NAME=''; WORK_DIR_ITEMS=[]; SAVE_DIR_HANDLE=null; SAVE_DIR_NAME='';
  try{ await saveConfig('workDir', null); }catch(e){}
  renderWorkDir(); await refreshList(); setStatus('已取消工作目录。'); }
function renderWorkDir(){
  const box=$('dir-box');
  if(!WORK_DIR_HANDLE){ box.innerHTML='<div class="empty">未设置工作目录<br>分析结果将只保存在浏览器内</div><button id="set-workdir" class="btn small ghost" style="width:100%;margin-top:8px;">设置工作目录</button>'; }
  else{ box.innerHTML='<div class="dir-info">📂 '+esc(WORK_DIR_NAME)+'</div>'+
      '<div class="dir-actions"><button id="refresh-workdir" class="btn small ghost">刷新</button><button id="clear-workdir" class="btn small danger">取消</button></div>'; }
  if($('set-workdir')) $('set-workdir').onclick=setWorkDir;
  if($('refresh-workdir')) $('refresh-workdir').onclick=refreshWorkDir;
  if($('clear-workdir')) $('clear-workdir').onclick=clearWorkDir;
}

/* 确保目录句柄拥有写入权限（页面刷新后从 IndexedDB 恢复的句柄会丢失写权限，需重新申请）。
   返回 true 表示已具备写权限；false 表示用户拒绝。 */
async function ensureDirWrite(dirHandle){
  if(!dirHandle || !dirHandle.queryPermission) return true;
  try{
    let st=await dirHandle.queryPermission({mode:'readwrite'});
    if(st!=='granted'){ st=await dirHandle.requestPermission({mode:'readwrite'}); }
    return st==='granted';
  }catch(e){ console.warn('申请目录写权限失败', e); return false; }
}

/* 判断一个 xlsx buffer 是否为「商品兑换明细」源数据 */
async function isSourceFile(buf){
  try{
    const wb=XLSX.read(new Uint8Array(buf),{type:'array'});
    for(const name of wb.SheetNames){
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:null,raw:false});
      if(!rows.length) continue;
      const header=rows[0].map(h=>norm(h));
      const need=["类型","消费积分合计","购买数量","码洋价","成本价"];
      if(need.every(c=>header.includes(c))) return true;
    }
  }catch(e){ console.warn('isSourceFile err', e); }
  return false;
}

/* 读取分析子文件夹内的 meta 信息 */
async function readFolderMeta(dirHandle){
  try{
    const fh=await dirHandle.getFileHandle('data.json');
    const file=await fh.getFile();
    return JSON.parse(await file.text());
  }catch(e){ return null; }
}

/* 扫描工作目录：子文件夹视为历史分析集；根目录 xlsx 视为待分析源文件 */
async function refreshWorkDir(){
  if(!WORK_DIR_HANDLE){ WORK_DIR_ITEMS=[]; await refreshList(); return; }
  const items=[];
  try{
    for await (const [entryName, handle] of WORK_DIR_HANDLE.entries()){
      if(handle.kind==='directory'){
        const meta=await readFolderMeta(handle);
        let displayName=entryName;
        if(entryName.indexOf('__')>=0) displayName=entryName.split('__').slice(1).join('__');
        const kpiText = meta && meta.dashboardData && meta.dashboardData.kpi
          ? `利润率 ${ (meta.dashboardData.kpi['整体利润率']*100).toFixed(1) }%` : '';
        items.push({
          source:'folder', id:'folder-'+entryName, displayName:displayName,
          folderName:entryName, handle, parentHandle:WORK_DIR_HANDLE, meta,
          time: meta ? (meta.uploadedAt||'') : '',
          kpiText,
          sub: (meta && meta.projectName ? meta.projectName : displayName) +
               ' ｜ ' + (meta && meta.uploadedAt ? meta.uploadedAt : '未分析') +
               (kpiText?' ｜ '+kpiText:'') + ' ｜ 本地文件夹'
        });
      }else if(handle.kind==='file' && /\.xlsx?$/i.test(entryName)){
        try{
          const file=await handle.getFile();
          const buf=await file.arrayBuffer();
          const source=await isSourceFile(buf);
          let entName='';
          if(source){
            try{
              const wb=XLSX.read(new Uint8Array(buf),{type:'array'});
              for(const sname of wb.SheetNames){
                const rows=XLSX.utils.sheet_to_json(wb.Sheets[sname],{header:1,defval:null,raw:false});
                if(!rows.length) continue;
                const header=rows[0].map(h=>norm(h));
                const idx={}; header.forEach((h,i)=>{ if(h && !(h in idx)) idx[h]=i; });
                const entIdx=idx['企业名称']; if(entIdx===undefined) continue;
                const counts={};
                for(let r=1;r<rows.length;r++){ const v=norm(rows[r][entIdx]); if(v){ counts[v]=(counts[v]||0)+1; if(!entName) entName=v; } }
                let best='', bestCnt=0; for(const k in counts){ if(counts[k]>bestCnt){ bestCnt=counts[k]; best=k; } }
                if(best){ entName=best; break; }
              }
            }catch(e){ console.warn('read entName', e); }
          }
          items.push({
            source:'file', id:'file-'+entryName, displayName:entryName.replace(/\.xlsx?$/i,''),
            fileName:entryName, handle, parentHandle:WORK_DIR_HANDLE, isSource:source, entName,
            time:new Date(file.lastModified).toLocaleString('zh-CN'),
            kpiText:'',
            sub: (entName? esc(entName)+' ｜ ':'') + new Date(file.lastModified).toLocaleString('zh-CN') + ' ｜ 本地文件'
          });
        }catch(e){ console.warn('扫描文件失败', entryName, e); }
      }
    }
  }catch(e){ console.error('扫描目录失败', e); setStatus('扫描目录失败：'+e.message); }
  WORK_DIR_ITEMS=items.sort((a,b)=>(b.time||'').localeCompare(a.time||''));
  await refreshList();
}

/* 分析工作目录中的某个文件 */
async function analyzeFileHandle(fileHandle){
  let file;
  try{ file=await fileHandle.getFile(); }catch(e){ alert('读取文件失败：'+e.message); return; }
  await runAnalysis(file, {fromWorkDir:true, fileHandle});
}

/* 打开工作目录中的历史分析子文件夹 */
async function openFolderItem(dirHandle, folderName){
  setStatus('正在打开本地分析文件夹…');
  try{
    const meta=await readFolderMeta(dirHandle);
    if(meta && meta.dashboardData){
      let cleanBlob=null, profitBlob=null, dashHtml=null;
      try{
        for await (const [name, h] of dirHandle.entries()){
          if(h.kind!=='file') continue;
          if(/_已清洗\.xlsx$/i.test(name)){ const f=await h.getFile(); cleanBlob=new Blob([await f.arrayBuffer()],{type:f.type}); }
          else if(/^项目数据统计口径参考_.*\.xlsx$/i.test(name)){ const f=await h.getFile(); profitBlob=new Blob([await f.arrayBuffer()],{type:f.type}); }
          else if(/_看板\.html$/i.test(name)){ const f=await h.getFile(); dashHtml=await f.text(); }
        }
      }catch(e){ console.warn('读取输出文件失败', e); }
      CURRENT={ projectName:meta.projectName||meta.name, tag:meta.tag, raw:null, kept:null, stats:meta.stats||{},
        paperZero:meta.paperZero, dashboardHtml:dashHtml, data:meta.dashboardData, uploadedAt:meta.uploadedAt||'' };
      CURRENT.cleanBlob=cleanBlob; CURRENT.profitBlob=profitBlob;
      if(folderName) CURRENT.folderName=folderName; // 让后续「下载」复用同一本地子文件夹
      $('dash-frame').srcdoc=await buildDashboardHtml(meta.dashboardData,'mini');
      $('dl-clean').disabled=!cleanBlob; $('dl-profit').disabled=!profitBlob; $('dl-dash').disabled=!dashHtml;
      $('result-meta').innerHTML=`本地文件夹：<b>${esc(meta.projectName||meta.name)}</b> ｜ ${meta.uploadedAt||''}`+(meta.summary?` ｜ 有效 ${meta.summary.kept} 行`:'');
      setStatus('已打开本地分析文件夹。');
      return;
    }
    // 无 meta：尝试找源文件重新分析
    for await (const [name, h] of dirHandle.entries()){
      if(h.kind==='file' && /\.xlsx?$/i.test(name)){
        const file=await h.getFile();
        const buf=await file.arrayBuffer();
        if(await isSourceFile(buf)){ await runAnalysis(file, {fromWorkDir:true, fileHandle:h}); return; }
      }
    }
    alert('该文件夹内未找到可识别的分析数据或源文件。');
  }catch(e){ console.error(e); alert('打开文件夹失败：'+e.message); }
}

/* 在工作目录下为每次分析创建独立子文件夹，并写入源文件、清洗表、利润表、看板、data.json */
async function saveResultsToWorkDir(projectName, tag, file, rawBuf, cleanBlob, profitBlob, dashHtml, data, stats, paperZero, summary){
  if(!WORK_DIR_HANDLE) return {ok:false, error:'未设置工作目录'};
  try{
    const shortId=guid().slice(0,12);
    const base=safeName(projectName);
    const folderName=shortId+'__'+base;
    const dirHandle=await WORK_DIR_HANDLE.getDirectoryHandle(folderName,{create:true});
    const fileStem=String(file.name||'').replace(/\.xlsx?$/i,'');

    // 1) data.json（看板数据 + 摘要，用于左侧列表与重新打开）
    const meta={
      name:projectName, projectName, tag, uploadedAt:new Date().toLocaleString('zh-CN'),
      fileName:file.name, dashboardData:data, stats, paperZero, summary
    };
    const metaBlob=new Blob([JSON.stringify(meta,null,2)],{type:'application/json'});
    await writeIntoDir(dirHandle, metaBlob, 'data.json');

    // 2) 原始文件副本
    if(rawBuf){
      await writeIntoDir(dirHandle, new Blob([rawBuf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}), safeName(fileStem)+'_源数据.xlsx');
    }

    // 3) 清洗表、利润表、看板网页
    await writeIntoDir(dirHandle, cleanBlob, safeName(fileStem)+'_已清洗.xlsx');
    await writeIntoDir(dirHandle, profitBlob, '项目数据统计口径参考_'+safeName(tag)+'.xlsx');
    await writeIntoDir(dirHandle, new Blob([dashHtml||''],{type:'text/html'}), safeName(fileStem)+'_看板.html');

    return {ok:true, folderName};
  }catch(e){ console.warn('写入工作目录失败', e); return {ok:false, error:e.message}; }
}

/* ---------------- UI ---------------- */
let ECHARTS_TEXT=null, TPL_BUF=null, CURRENT=null;
const $ = id=>document.getElementById(id);

function setStatus(s){ $('status').textContent=s||''; }
function downloadBlob(blob, filename){ const url=URL.createObjectURL(blob);
  const a=document.createElement('a'); a.href=url; a.download=filename; document.body.appendChild(a); a.click();
  setTimeout(()=>{URL.revokeObjectURL(url); a.remove();},1000); }
function guid(){ return Date.now().toString(36)+Math.random().toString(36).slice(2,8); }
function safeName(name){ return String(name||'file').replace(/[\\/:*?"<>|]/g,'_').replace(/\s+/g,' ').trim().slice(0,120) || 'file'; }
const nextFrame=()=>new Promise(r=>setTimeout(r,40));
function setProgress(pct, text){ const box=$('progress'); if(!box) return; box.style.display='';
  $('progress-fill').style.width=Math.max(0,Math.min(100,pct))+'%';
  $('progress-text').textContent=(text||'')+' '+Math.round(pct)+'%'; }
function hideProgress(){ const box=$('progress'); if(box) box.style.display='none'; }
function toggleFullscreen(){
  const el=$('dash-frame');
  const fsEl=document.fullscreenElement||document.webkitFullscreenElement;
  if(fsEl){ const exit=document.exitFullscreen||document.webkitExitFullscreen; if(exit) exit.call(document); }
  else{ const req=el.requestFullscreen||el.webkitRequestFullscreen; if(req) req.call(el); else alert('当前浏览器不支持网页内全屏，请改用浏览器自带全屏（按 F11）。'); }
}

/* ---------------- 归属（省/市/支行）编辑与筛选 ---------------- */
function allCities(){ const a=[]; for(const p in (window.CHINA_GEO||{})) (window.CHINA_GEO[p]||[]).forEach(c=>a.push(c)); return a; }
function fillGeoSelects(){
  const provList=$('list-prov'), cityList=$('list-city');
  if(provList && provList.children.length===0){
    for(const p in (window.CHINA_GEO||{})){ const o=document.createElement('option'); o.value=p; provList.appendChild(o); }
  }
  if(cityList && cityList.children.length===0){ allCities().forEach(c=>{ const o=document.createElement('option'); o.value=c; cityList.appendChild(o); }); }
}
function onGProvChange(){
  const p=$('g-prov').value; const cityList=$('list-city');
  if(cityList){ cityList.innerHTML=''; const cities=p? (window.CHINA_GEO[p]||[]):allCities();
    cities.forEach(c=>{ const o=document.createElement('option'); o.value=c; cityList.appendChild(o); }); }
}
function onFileSelected(file){
  $('run-btn').disabled=!file;
  if(!file){ $('sel-name').textContent=''; return; }
  $('sel-name').textContent=file.name;
  fillGeoSelects();
  const g=detectGeo(file.name);
  $('g-prov').value = (window.CHINA_GEO && g.province && (g.province in window.CHINA_GEO)) ? g.province : (g.province||'');
  onGProvChange();
  if(g.city){ const opts=[...$('g-city').options].map(o=>o.value); $('g-city').value = opts.includes(g.city)? g.city : ''; }
  $('g-town').value=g.town||'';
  $('g-period').value=g.period||'';
  $('geo-edit').style.display='';
}

/* 历史数据集：以「企业名称」为总名称，全部列在左侧 */
let ALL_DS=[];

/* 从 ds 记录推断企业名称（修复老数据用文件名命名的问题） */
function dsDisplayName(ds){
  const rawName=ds.name||'';
  // 如果名字已经明显是企业名称（不含常见文件名片段），直接返回
  if(rawName && !/商品兑换明细|\.xlsx|明细|兑换|报表|统计|数据/.test(rawName)) return rawName;
  // 否则尝试从 kept 记录推断
  if(ds.kept && ds.kept.length){
    const name=inferProjectName(ds.kept, rawName);
    if(name && name!==rawName) return name;
  }
  // 尝试从 raw buffer 推断（较耗，仅兜底）
  if(ds.raw && !ds.kept){
    try{
      const wb=XLSX.read(new Uint8Array(ds.raw),{type:'array'});
      for(const sname of wb.SheetNames){
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[sname],{header:1,defval:null,raw:false});
        if(!rows.length) continue; const header=rows[0].map(h=>norm(h));
        const idx={}; header.forEach((h,i)=>{ if(h && !(h in idx)) idx[h]=i; });
        const entIdx=idx['企业名称']; if(entIdx===undefined) continue;
        const counts={}; for(let r=1;r<rows.length;r++){ const v=norm(rows[r][entIdx]); if(v) counts[v]=(counts[v]||0)+1; }
        let best='', bestCnt=0; for(const k in counts){ if(counts[k]>bestCnt){ bestCnt=counts[k]; best=k; } }
        if(best) return best;
      }
    }catch(e){}
  }
  return rawName || '未命名';
}

async function runAnalysis(file, opts={}){
  opts=opts||{};
  $('run-btn').disabled=true;
  try{
    setProgress(5,'读取文件');
    const buf=await file.arrayBuffer();
    setProgress(20,'解析数据'); await nextFrame();
    const {records, rawRows, skipped}=loadRecords(buf);
    if(!records.length) throw new Error('未解析到有效数据行（请确认列名：类型/消费积分合计/购买数量/码洋价/成本价）。'+(skipped.length?'已跳过: '+skipped.join('; '):''));
    setProgress(40,'清洗数据（'+rawRows+' 行）'); await nextFrame();
    const {kept, droppedZero, droppedNoType}=cleanRecords(records);
    if(!kept.length) throw new Error('清洗后无有效数据（消费积分合计=0 的行已被剔除）。');
    const groups=groupByType(kept);
    const paperZero=detectPaperZero(records);
    const stats={}; for(const t in groups) stats[t]=summarize(groups[t]);
    const taxInfo=buildTaxInfo(paperZero);
    const caliber=buildCaliberRows(stats, taxInfo);
    setProgress(60,'计算看板指标'); await nextFrame();
    const data=compute(kept, taxInfo);
    const fileStem=String(file.name||'未命名').replace(/\.xlsx?$/i,'');
    const projectName=fileStem;          // 初提取：以文件名作为数据集名称
    data.project_name=projectName;
    data.generated_at=new Date().toLocaleString('zh-CN');
    data.detail_file=file.name;
    data.caliber_file='';
    data.blacklist_file='内置黑名单（'+BLACKLIST_SET.size+' 个商品id）';
    data.caliber=caliber;
    data.tax_info=taxInfo;

    const tag=deriveTag(fileStem);
    setProgress(80,'生成清洗表 / 利润表'); await nextFrame();
    const cleanBlob=await writeCleanXlsx(groups);
    const profitBlob=await writeProfitXlsx(stats, paperZero, TPL_BUF);

    // 渲染看板（iframe 内用精简版，无 header/nav；下载用完整版）
    setProgress(93,'渲染看板'); await nextFrame();
    const htmlMini=await buildDashboardHtml(data,'mini');
    const htmlFull=await buildDashboardHtml(data,'full');
    $('dash-frame').srcdoc=htmlMini;

    // 保存当前结果，供下载 / 持久化
    CURRENT={ projectName, tag, file, raw:buf, kept, stats, paperZero,
      cleanBlob, profitBlob, dashboardHtml:htmlFull, data,
      droppedZero, droppedNoType, rawRows,
      uploadedAt:new Date().toLocaleString('zh-CN') };

    // 下载按钮可用
    $('dl-clean').disabled=false; $('dl-profit').disabled=false; $('dl-dash').disabled=false;
    $('result-meta').innerHTML=
      `已处理 <b>${rawRows}</b> 行 ｜ 有效 <b>${kept.length}</b> 行（剔除积分=0: ${droppedZero}，无类型: ${droppedNoType}）`+
      (paperZero?' ｜ <b style="color:#e0922c">检测到【全员阅读平台】，纸书项目税率归 0</b>':'')+
      ` ｜ 整体利润率(不含税) <b>${ (data.kpi['整体利润率']*100).toFixed(1) }%</b>`;
    setProgress(100,'完成');
    setStatus('完成。可下载清洗表 / 利润表 / 看板，或点「保存到本机」留以后查看。');
    // 自动归档到 IndexedDB，刷新网页后仍可在左侧「历史数据集」找回
    try{ await persistCurrent(); }catch(e){ console.warn('自动保存失败', e); }
    // 若设置了工作目录，在目录下新建子文件夹归档
    if(WORK_DIR_HANDLE){
      try{
        const res=await saveResultsToWorkDir(
          CURRENT.projectName, CURRENT.tag, CURRENT.file, CURRENT.raw,
          CURRENT.cleanBlob, CURRENT.profitBlob, CURRENT.dashboardHtml,
          CURRENT.data, CURRENT.stats, CURRENT.paperZero,
          {rawRows:CURRENT.rawRows, kept:CURRENT.kept.length, droppedZero:CURRENT.droppedZero, droppedNoType:CURRENT.droppedNoType}
        );
        if(res.ok){
          CURRENT.folderName = res.folderName;
          try{ await markLocalFolder(CURRENT.savedId, res.folderName); }catch(_){}
          setStatus('已保存到工作目录【'+WORK_DIR_NAME+' / '+res.folderName+'】并归档到浏览器。');
          await refreshWorkDir();
        }
      }catch(e){ console.warn('写入工作目录失败', e); }
    }
  }catch(e){
    console.error(e); setStatus('出错：'+e.message); alert('处理失败：'+e.message);
  }finally{
    $('run-btn').disabled=false; setTimeout(hideProgress, 700);
  }
}
function deriveTag(stem){ const m=stem.match(/(\d{1,2}[.\-_]\d{1,2})/); return m? m[1].replace(/[-_]/g,'.') : new Date().toISOString().slice(0,10); }

async function persistCurrent(){
  if(!CURRENT) return;
  // 归属（省/市/支行/期号）：优先取用户编辑框，缺省回退到文件名自动识别
  const geo={
    province:$('g-prov')?($('g-prov').value||''):'',
    city:$('g-city')?($('g-city').value||''):'',
    town:$('g-town')?($('g-town').value.trim()||''):'',
    period:$('g-period')?($('g-period').value.trim()||''):''
  };
  if(!geo.province && !geo.city && !geo.town && !geo.period){ const g=detectGeo(CURRENT.projectName); Object.assign(geo,g); }
  // 若已自动/手动保存过，则更新同一条记录，避免重复归档
  const id = CURRENT.savedId || guid();
  const ds={ id, name:CURRENT.projectName, uploadedAt:CURRENT.uploadedAt, geo,
    raw:CURRENT.raw, kept:CURRENT.kept, stats:CURRENT.stats, paperZero:CURRENT.paperZero,
    tag:CURRENT.tag, dashboardData:CURRENT.data,
    summary:{rawRows:CURRENT.rawRows, kept:CURRENT.kept.length, droppedZero:CURRENT.droppedZero, droppedNoType:CURRENT.droppedNoType} };
  await saveDS(ds);
  CURRENT.savedId = id;
  await refreshList();
  setStatus('已保存到本机浏览器（IndexedDB）'+(geo.province||geo.city||geo.town?(' ｜ 归属：'+(geo.province||'-')+'/'+(geo.city||'-')+'/'+(geo.town||'-')):'')+'，可在左侧「历史数据集」再次打开。');
}

/* 确保数据集有 geo；老的/空的会自动用文件名识别并写回 IndexedDB */
function ensureGeo(ds){
  const g=ds.geo||{};
  const hasAny=g.province||g.city||g.town||g.period;
  if(!hasAny){ const d=detectGeo(ds.name); ds.geo=d; return true; }
  return false;
}
async function updateDSGeo(ds){ try{ await saveDS(ds); }catch(e){ console.warn('写回 geo 失败', e); } }
/* 自动归档到工作目录后，回写 IndexedDB 记录标记其本地子文件夹名（用于「本地+浏览器」双份识别） */
async function markLocalFolder(id, folderName){
  try{ const ds=await getDS(id); if(ds){ ds.localFolder=folderName; await saveDS(ds); } }catch(e){ console.warn('标记本地文件夹失败', e); }
}

async function refreshList(){
  const all=await listDS();
  ALL_DS=all;

  // 合并 IndexedDB 记录 + 工作目录条目；同一份数据若「本地子文件夹」与「IndexedDB」都有则合并为一条
  const byFolder={};
  for(const ds of all){ if(ds.localFolder) byFolder[ds.localFolder]=ds; }
  const items=[];
  const consumedIdx=new Set();
  for(const it of WORK_DIR_ITEMS){
    if(it.source==='folder'){
      const matched=byFolder[it.folderName];
      const dual=!!matched;
      if(matched) consumedIdx.add(matched.id);
      items.push({
        source:'folder', id:it.id,
        displayName: dual ? dsDisplayName(matched) : it.displayName,
        folderName:it.folderName, time:it.time, handle:it.handle,
        parentHandle:it.parentHandle, meta:it.meta, kpiText:it.kpiText,
        sub: dual ? ((matched.uploadedAt||it.time||'')+' ｜ '+(it.kpiText||'')+' ｜ 本地+浏览器') : it.sub,
        dual, ds:matched
      });
    }else if(it.source==='file'){
      items.push({
        source:'file', id:it.id, displayName:it.displayName, fileName:it.fileName,
        time:it.time, handle:it.handle, parentHandle:it.parentHandle, isSource:it.isSource, entName:it.entName,
        kpiText:'', sub:it.sub
      });
    }
  }
  // IndexedDB 记录（source='indexeddb'）；已合并进本地文件夹的跳过
  for(const ds of all){
    if(consumedIdx.has(ds.id)) continue;
    const display=dsDisplayName(ds);
    const k=ds.dashboardData? ds.dashboardData.kpi['整体利润率']:0;
    const kpiText=`利润率 ${ (k*100).toFixed(1) }%`;
    const dual=!!ds.localFolder;
    items.push({
      source:'indexeddb', id:ds.id, displayName:display, ds:ds,
      time:ds.uploadedAt||'',
      kpiText:kpiText,
      sub: (ds.uploadedAt||'') + ' ｜ ' + kpiText + (dual?' ｜ 本地+浏览器':' ｜ 仅浏览器')
    });
  }
  // 按时间倒序
  items.sort((a,b)=> (b.time||'').localeCompare(a.time||''));

  const box=$('ds-list'); box.innerHTML='';
  // 来源图例：让用户理解「仅浏览器 / 本地文件夹 / 本地+浏览器」
  const legend=document.createElement('div');
  legend.style.cssText='font-size:11.5px;color:var(--sub);background:#f8fafc;border:1px solid var(--line);border-radius:8px;padding:8px 10px;margin-bottom:8px;line-height:1.6;';
  legend.innerHTML='💡 <b>数据来源</b>：💾仅浏览器=只存浏览器本地库(IndexedDB)；📁本地文件夹=只存工作目录；🔗本地+浏览器=两处都有（删任一处不影响另一处）。未设工作目录时全部为「仅浏览器」。';
  box.append(legend);

  if(!items.length){ const e=document.createElement('div'); e.className='empty'; e.textContent='暂无数据集'; box.append(e); return; }
  items.forEach(item=>{
    let el;
    const edId = item.ds ? item.ds.id : item.id;
    const badge = item.source==='file' ? '📄 本地文件·待分析'
                : item.dual ? '🔗 本地+浏览器'
                : item.source==='folder' ? '📁 本地文件夹'
                : '💾 仅浏览器';
    try{
      el=document.createElement('div'); el.className='ds-item';
      el.innerHTML=`<div class="ds-name">${esc(item.displayName||'未命名')}</div>
        <div class="ds-sub">${esc(item.sub||'')}</div>
        <div style="font-size:11.5px;margin-top:4px;"><span style="display:inline-block;padding:1px 8px;border-radius:10px;background:#eef2f8;color:#41506a;">${badge}</span></div>
        <div id="editor-${esc(edId)}" style="display:none;width:100%;"></div>`;
      const acts=document.createElement('div'); acts.className='ds-acts';
      if(item.source==='folder'){
        const openBtn=document.createElement('button'); openBtn.className='btn small'; openBtn.textContent='打开';
        openBtn.onclick=()=>openFolderItem(item.handle, item.folderName);
        acts.append(openBtn);
        if(item.dual && item.ds){
          const rename=document.createElement('button'); rename.className='btn small ghost'; rename.textContent='改名';
          rename.onclick=()=>toggleNameEditor(item.ds);
          acts.append(rename);
        }
        const delBtn=document.createElement('button'); delBtn.className='btn small danger'; delBtn.textContent=item.dual?'删除(两处)':'删除';
        delBtn.onclick=async()=>{
          const msg = item.dual ? '确认同时删除本地文件夹与该数据集（浏览器）？此操作不可恢复。' : '确认删除该分析文件夹及其内所有文件？此操作不可恢复。';
          if(confirm(msg)){
            try{
              const dh=item.parentHandle;
              if(!dh) throw new Error('未获取到工作目录句柄');
              if(!await ensureDirWrite(dh)) throw new Error('未授予写入权限，无法删除（可在弹窗中点击「允许」后重试）');
              if(item.folderName){ await dh.removeEntry(item.folderName,{recursive:true}); }
              if(item.dual && item.ds){ await delDS(item.ds.id); }
              await refreshWorkDir(); setStatus('已删除：'+item.folderName);
            }catch(e){ console.error(e); alert('删除失败：'+e.message); }
          }
        };
        acts.append(delBtn);
      }else if(item.source==='file'){
        const analyze=document.createElement('button'); analyze.className='btn small primary'; analyze.textContent='分析';
        analyze.onclick=()=>analyzeFileHandle(item.handle);
        acts.append(analyze);
        const delBtn=document.createElement('button'); delBtn.className='btn small danger'; delBtn.textContent='删除';
        delBtn.onclick=async()=>{
          if(!confirm('确认从工作目录删除文件「'+item.fileName+'」？此操作不可恢复。')) return;
          try{
            const dh=item.parentHandle;
            if(!dh) throw new Error('未获取到工作目录句柄');
            if(!await ensureDirWrite(dh)) throw new Error('未授予写入权限，无法删除（可在弹窗中点击「允许」后重试）');
            await dh.removeEntry(item.fileName);
            await refreshWorkDir(); setStatus('已删除文件：'+item.fileName);
          }catch(e){ console.error(e); alert('删除文件失败：'+e.message); }
        };
        acts.append(delBtn);
        if(!item.isSource){
          const hint=document.createElement('span'); hint.style.fontSize='11.5px'; hint.style.color='var(--sub)'; hint.textContent='（非源数据）';
          acts.append(hint);
        }
      }else{
        const open=document.createElement('button'); open.className='btn small'; open.textContent='打开';
        open.onclick=()=>openDS(item.id);
        const rename=document.createElement('button'); rename.className='btn small ghost'; rename.textContent='改名';
        rename.onclick=()=>toggleNameEditor(item.ds);
        const del=document.createElement('button'); del.className='btn small danger'; del.textContent='删除';
        del.onclick=async()=>{ if(confirm('确认删除该数据集？')){ try{ await delDS(item.id); await refreshList(); }catch(e){ console.error(e); alert('单条删除失败：'+e.message+'（可改用左上「清空全部」）。'); } } };
        acts.append(open,rename,del);
      }
      el.append(acts);
    }catch(e){ console.warn('渲染单项失败', e, item); el=document.createElement('div'); el.className='ds-item';
      el.innerHTML=`<div class="ds-name">${esc(item.displayName||'未命名')}</div><div class="ds-sub">记录异常，建议「清空全部」</div>`; }
    box.append(el);
  });
}

/* 历史数据集单项名称 inline 编辑 */
function toggleNameEditor(ds){
  const box=$(`editor-${ds.id}`); if(!box) return;
  if(box.style.display==='none'){
    box.innerHTML=`<div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin-top:6px;">
      <input id="en-${ds.id}" value="${esc(ds.name||'')}" placeholder="数据集名称" style="flex:1;min-width:120px;padding:5px 7px;border:1px solid var(--line);border-radius:6px;font-size:12px;">
      <button class="btn small primary" id="esn-${ds.id}">保存</button>
      <button class="btn small ghost" id="encancel-${ds.id}">取消</button>
    </div>`;
    box.style.display='';
    $(`esn-${ds.id}`).onclick=async()=>{
      const newName=$('en-'+ds.id).value.trim();
      if(!newName){ alert('名称不能为空'); return; }
      ds.name=newName;
      await saveDS(ds);
      if(CURRENT && CURRENT.savedId===ds.id) CURRENT.projectName=newName;
      await refreshList();
    };
    $(`encancel-${ds.id}`).onclick=()=>{ box.style.display='none'; box.innerHTML=''; };
  }else{
    box.style.display='none'; box.innerHTML='';
  }
}

/* 历史数据集单项归属 inline 编辑 */
function toggleGeoEditor(ds){
  const box=$(`editor-${ds.id}`); if(!box) return;
  if(box.style.display==='none'){
    const g=ds.geo||{};
    box.innerHTML=`<div style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;">
      <input id="ep-${ds.id}" list="list-prov" value="${esc(g.province||'')}" placeholder="省份" style="width:90px;padding:5px 7px;border:1px solid var(--line);border-radius:6px;font-size:12px;">
      <input id="ec-${ds.id}" list="list-city" value="${esc(g.city||'')}" placeholder="城市" style="width:90px;padding:5px 7px;border:1px solid var(--line);border-radius:6px;font-size:12px;">
      <input id="et-${ds.id}" value="${esc(g.town||'')}" placeholder="支行/县" style="width:110px;padding:5px 7px;border:1px solid var(--line);border-radius:6px;font-size:12px;">
      <input id="epr-${ds.id}" value="${esc(g.period||'')}" placeholder="期号" style="width:70px;padding:5px 7px;border:1px solid var(--line);border-radius:6px;font-size:12px;">
      <button class="btn small primary" id="es-${ds.id}">保存</button>
      <button class="btn small ghost" id="ecancel-${ds.id}">取消</button>
    </div>`;
    box.style.display='';
    $(`es-${ds.id}`).onclick=async()=>{
      ds.geo={
        province:$(`ep-${ds.id}`).value.trim(),
        city:$(`ec-${ds.id}`).value.trim(),
        town:$(`et-${ds.id}`).value.trim(),
        period:$(`epr-${ds.id}`).value.trim()
      };
      await saveDS(ds);
      await refreshList();
    };
    $(`ecancel-${ds.id}`).onclick=()=>{ box.style.display='none'; box.innerHTML=''; };
  }else{
    box.style.display='none'; box.innerHTML='';
  }
}
function esc(s){ return String(s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

async function openDS(id){
  let ds;
  try{ ds=await getDS(id); }catch(e){ alert('读取数据集失败：'+e.message); return; }
  if(!ds){ alert('未找到该数据集（可能已被删除）。'); try{ await refreshList(); }catch(_){} return; }
  if(!ds.kept || !ds.kept.length){ alert('该数据集缺少明细数据，无法打开。\n建议点「删除」移除后重新上传分析；若删除也无效，用左上「清空全部」。'); return; }
  setStatus('从本机加载历史数据集…');
  CURRENT={ projectName:ds.name, tag:ds.tag, raw:ds.raw, kept:ds.kept, stats:ds.stats,
    paperZero:ds.paperZero, dashboardHtml:null, data:ds.dashboardData, uploadedAt:ds.uploadedAt,
    savedId:ds.id };
  if(ds.localFolder) CURRENT.folderName=ds.localFolder; // 若本地已有子文件夹，下载时复用同一处
  // 重新生成 xlsx 供下载
  const groups=groupByType(ds.kept);
  CURRENT.cleanBlob=await writeCleanXlsx(groups);
  CURRENT.profitBlob=await writeProfitXlsx(ds.stats, ds.paperZero, TPL_BUF);
  CURRENT.dashboardHtml=await buildDashboardHtml(ds.dashboardData,'full');
  $('dash-frame').srcdoc=await buildDashboardHtml(ds.dashboardData,'mini');
  $('dl-clean').disabled=false; $('dl-profit').disabled=false; $('dl-dash').disabled=false;
  $('result-meta').innerHTML=`历史数据集：<b>${esc(ds.name)}</b> ｜ ${ds.summary?('有效 '+ds.summary.kept+' 行'):''}`;
  setStatus('已打开历史数据集。');
}

/* 下载按钮 */
function bindDownloads(){
  // 三个下载：写入工作目录下的子文件夹（与自动归档同一处）；未设工作目录则提示并退回默认下载
  const dl = async (btn, blob, filename)=>{
    if(!CURRENT) return;
    btn.disabled=true;
    try{ await saveBlobToWorkDir(blob, filename); }
    catch(e){ console.error(e); alert('保存失败：'+e.message); }
    finally{ btn.disabled=false; }
  };
  $('dl-clean').onclick=()=> dl($('dl-clean'), CURRENT.cleanBlob, safeName((CURRENT.projectName||'清洗')+'_已清洗.xlsx'));
  $('dl-profit').onclick=()=> dl($('dl-profit'), CURRENT.profitBlob, safeName('项目数据统计口径参考_'+(CURRENT.tag||'')+'.xlsx'));
  $('dl-dash').onclick=()=> dl($('dl-dash'), new Blob([CURRENT.dashboardHtml||''],{type:'text/html'}), safeName((CURRENT.projectName||'看板')+'_看板.html'));
  $('save-btn').onclick=persistCurrent;
  if($('set-dir')) $('set-dir').onclick=setSaveDir;
  if($('clear-dir')) $('clear-dir').onclick=clearSaveDir;
  if($('clear-all')) $('clear-all').onclick=clearAllDS;
  if($('dash-full')) $('dash-full').onclick=toggleFullscreen;
  // 文件选择 / 拖拽：统一走 onFileSelected（自动识别归属）
  $('run-btn').onclick=()=>{ const f=$('file-input').files[0]; if(f) runAnalysis(f); };
  $('file-input').addEventListener('change',e=>{ onFileSelected(e.target.files[0]); });
  $('g-prov').addEventListener('change', onGProvChange);
  // 拖拽
  const dz=$('drop');
  ['dragover','dragenter'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('over');}));
  ['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('over');}));
  dz.addEventListener('drop',e=>{ const f=e.dataTransfer.files[0]; if(f){ $('file-input').files=normalizeFileList(f); onFileSelected(f); } });
}
function normalizeFileList(f){ const dt=new DataTransfer(); dt.items.add(f); return dt.files; }

/* 初始化 */
async function init(){
  try{
    if(!window.EMBED_TEMPLATE_B64) throw new Error('模板未加载（embed_template.js）');
    TPL_BUF=b64ToU8(window.EMBED_TEMPLATE_B64);
    parseTaxBase(TPL_BUF);
    ECHARTS_TEXT = window.EMBED_ECHARTS_B64 ? b64ToText(window.EMBED_ECHARTS_B64) : null;
  }catch(e){ console.error('预加载失败', e); alert('初始化失败：'+e.message); setStatus('出错：'+e.message); return; }
  bindDownloads();
  // 预填归属编辑框的 datalist（全量省/市，供识别时自由增改）
  fillGeoSelects();
  // 恢复工作目录（File System Access 句柄可存 IndexedDB）
  renderWorkDir();
  try{
    const h=await getConfig('workDir');
    if(h && h.name){
      WORK_DIR_HANDLE=h; WORK_DIR_NAME=h.name;
      SAVE_DIR_HANDLE=h; SAVE_DIR_NAME=h.name;
      renderWorkDir();
      // 尝试请求权限并扫描
      try{
        const perm=await h.requestPermission({mode:'readwrite'});
        if(perm==='granted'){ await refreshWorkDir(); }
        else{ setStatus('工作目录权限未授予，请点击「刷新」重新授权。'); }
      }catch(e){ console.warn('请求工作目录权限失败', e); setStatus('工作目录权限未授予，请点击「刷新」重新授权。'); }
    }
  }catch(e){ console.warn('恢复工作目录失败', e); }
  await refreshList();
  setStatus('就绪：选择或拖入「商品兑换明细」xlsx 开始。数据仅在你的浏览器内处理，不会上传。');
}
window.addEventListener('DOMContentLoaded', init);
