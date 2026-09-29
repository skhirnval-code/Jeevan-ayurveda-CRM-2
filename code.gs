/*******************************************************
 * HERBS CRM - Google Apps Script Backend
 * -----------------------------------------------------
 * Google Sheets based CRM converted from the supplied
 * CRM feature specification.
 *
 * SETUP:
 * 1. Create a Google Sheet.
 * 2. Extensions > Apps Script.
 * 3. Paste this file as Code.gs.
 * 4. (index.html GitHub Pages par host hai - yahan paste karna zaroori nahi)
 * 5. Run setupCRM() once and authorize.
 * 6. Deploy > New deployment > Web app.
 *    Execute as: Me
 *    Who has access: Anyone with the link (or your org)
 *
 * NOTE:
 * - This is a Sheets-backed CRM implementation.
 * - Shiprocket / India Post / Airtel IQ are prepared as
 *   integration points; add credentials in Settings.
 *******************************************************/

const CRM = {
  SHEETS: {
    USERS: 'Users',
    ORDERS: 'Orders',
    DEALERS: 'Dealers',
    LEDGER: 'Ledger',
    INVOICES: 'Invoices',
    NOTES: 'Notes',
    HISTORY: 'OrderHistory',
    AUDIT: 'AuditLogs',
    SOURCES: 'Sources',
    STATUSES: 'Statuses',
    PRODUCTS: 'Products',
    SETTINGS: 'Settings',
    FOLLOWUPS: 'FollowupRules',
    CALLS: 'Calls'
  },
  ROLES: ['SUPER_ADMIN','MANAGER','ZM','AGENT','VIEWER','DEALER'],
  STATUSES: [
    'New','Confirm Pending','Confirmed','In Transit','Delivered',
    'Callback','Pending','GPO','GPO Pending','GPO Portal','GPO Done',
    'GPO Delivered','Confirm cancel','Cancel pending','Final cancel',
    'Cancelled','Dealer Cancel','Future Delivery','UNA','NDR','Lost',
    'RTO','Double Cancel'
  ],
  SOURCES: [
    'Abandoned Cart','Calling','Discount Lead','IND','IND MANDEEP','Nasha',
    'Nasha Abandoned Cart','Nasha WhatsApp','Orders','Pincode','WhatsApp'
  ]
};

/* ========================= HTTP ENTRY (GitHub Pages frontend) =========================
 * Frontend (GitHub Pages) is file ko fetch() se POST karta hai:
 *   body = JSON {action, payload}
 * Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone
 */
function doGet(e) {
  return jsonOut_({ok:true, message:'HERBS CRM API is running', time:new Date()});
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return jsonOut_({ok:false, error:'Invalid JSON body'}); }
  return jsonOut_(api(body.action, body.payload || {}));
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ========================= SETUP ========================= */

function setupCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the Google Sheet and run setupCRM() from Apps Script.');

  const schemas = {
    Users: ['id','name','username','email','phone','password','role','active','workMode','permissions','createdAt','lastLoginAt'],
    Orders: ['id','orderNo','orderDate','customerName','phone','altPhone','email','product','extra','qty','unitPrice','total','online','balance','paymentMode','paymentStatus','address','pincode','city','state','district','source','status','remark','followupAt','leadOwner','leadOwnerId','agentAssignedAt','dealer','dealerCode','dealerId','dealerAssignedAt','dealerMargin','courier','carrier','awb','shipStatus','shipOrderId','shipmentId','labelUrl','labelPrinted','manifestDone','bookedAt','deliveredAt','lastShipEventAt','statusChangedAt','statusChangedBy','deletedAt','createdAt','updatedAt'],
    Dealers: ['id','code','name','username','firmName','contactPerson','mobile','altMobile','email','gst','pan','address','city','pincode','territory','state','district','creditLimit','openingStock','defaultMargin','minStock','notes','active','zm','zmId','createdAt'],
    Ledger: ['id','dealerCode','date','particular','debit','credit','net','running','reference'],
    Invoices: ['id','invoiceNo','date','dealerCode','dealerName','paymentMode','paidAmount','itemsJson','discount','notes','grandTotal','balance','status','createdAt'],
    Notes: ['id','orderId','phone','text','user','createdAt'],
    OrderHistory: ['id','orderId','field','oldValue','newValue','byUser','system','createdAt'],
    AuditLogs: ['id','time','user','module','action','target','details'],
    Sources: ['id','name','active','sortOrder','createdAt'],
    Statuses: ['id','name','color','terminal','revenue','sortOrder','active'],
    Products: ['id','name','hsn','price','active'],
    Settings: ['key','value'],
    FollowupRules: ['id','status','afterDays'],
    Calls: ['id','time','agent','customer','phone','orderId','direction','outcome','duration','recording','afterStatus','callId','createdAt']
  };

  Object.keys(schemas).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,schemas[name].length).setValues([schemas[name]]);
      sh.setFrozenRows(1);
    } else {
      const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),schemas[name].length)).getValues()[0];
      schemas[name].forEach((h,i) => {
        if (current[i] !== h) sh.getRange(1,i+1).setValue(h);
      });
    }
  });

  seedIfEmpty_('Users', ['id','name','username','email','phone','password','role','active','workMode','permissions','createdAt','lastLoginAt'],
    [[1,'Sandeep Kumar','admin','admin@example.com','','admin123','SUPER_ADMIN','TRUE','OFFICE','{}',new Date(),'']]);

  if (sheetData_('Sources').length === 0) {
    CRM.SOURCES.forEach((x,i)=>appendRowObject_('Sources',{id:uid_('SRC'),name:x,active:true,sortOrder:i,createdAt:new Date()}));
  }
  if (sheetData_('Statuses').length === 0) {
    CRM.STATUSES.forEach((x,i)=>appendRowObject_('Statuses',{id:uid_('STS'),name:x,color:'slate',terminal:['Delivered','GPO Delivered','Cancelled','Final cancel','Lost','RTO'].includes(x),revenue:['Delivered','GPO Delivered'].includes(x),sortOrder:i,active:true}));
  }
  if (sheetData_('Products').length === 0) {
    ['Anti Addiction','Sutra Gold+','Other'].forEach((x,i)=>appendRowObject_('Products',{id:uid_('PRD'),name:x,hsn:'',price:0,active:true}));
  }
  audit_('SYSTEM','setup','CRM','CRM initialized');
  return {ok:true,message:'CRM sheets created/updated successfully.'};
}

/* ========================= WEB API ========================= */

function api(action, payload) {
  payload = payload || {};
  try {
    switch(action) {
      case 'login': return login_(payload);
      case 'bootstrap': return bootstrap_(payload);
      case 'dashboard': return dashboard_(payload);
      case 'orders': return orders_(payload);
      case 'getOrder': return getOrder_(payload);
      case 'saveOrder': return saveOrder_(payload);
      case 'deleteOrder': return deleteOrder_(payload);
      case 'bulkStatus': return bulkStatus_(payload);
      case 'bulkAssign': return bulkAssign_(payload);
      case 'importOrders': return importOrders_(payload);
      case 'dealers': return dealers_(payload);
      case 'saveDealer': return saveDealer_(payload);
      case 'deleteDealer': return deleteDealer_(payload);
      case 'ledger': return ledger_(payload);
      case 'addPayment': return addPayment_(payload);
      case 'invoices': return invoices_(payload);
      case 'saveInvoice': return saveInvoice_(payload);
      case 'reports': return reports_(payload);
      case 'incentive': return incentive_(payload);
      case 'users': return users_(payload);
      case 'saveUser': return saveUser_(payload);
      case 'settings': return settings_(payload);
      case 'saveSettings': return saveSettings_(payload);
      case 'audit': return auditList_(payload);
      case 'calls': return calls_(payload);
      case 'saveNote': return saveNote_(payload);
      case 'health': return health_(payload);
      default: throw new Error('Unknown API action: '+action);
    }
  } catch(e) {
    return {ok:false,error:e.message || String(e)};
  }
}

/* ========================= AUTH ========================= */

function login_(p) {
  const username = String(p.username||'').trim().toLowerCase();
  const password = String(p.password||'');
  const rows = sheetData_('Users');
  const u = rows.find(x => String(x.username).toLowerCase() === username && String(x.password) === password && truthy_(x.active));
  if (!u) {
    audit_('AUTH','login_failed',username,'Invalid credentials');
    throw new Error('Username or password is incorrect.');
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('TOKEN_'+token, JSON.stringify({
    id:u.id,name:u.name,username:u.username,role:u.role,email:u.email||'',workMode:u.workMode||'OFFICE',
    permissions:parseJson_(u.permissions,{})
  }), 21600);
  updateById_('Users',u.id,{lastLoginAt:new Date()});
  audit_(u.username,'login','AUTH',u.username);
  return {ok:true,token,user:safeUser_(u)};
}

function auth_(token) {
  if (!token) throw new Error('Please login.');
  const s = CacheService.getScriptCache().get('TOKEN_'+token);
  if (!s) throw new Error('Session expired. Please login again.');
  return JSON.parse(s);
}

function safeUser_(u) {
  return {id:u.id,name:u.name,username:u.username,email:u.email,phone:u.phone,role:u.role,active:truthy_(u.active),workMode:u.workMode||'OFFICE',permissions:parseJson_(u.permissions,{})};
}

/* ========================= BOOTSTRAP ========================= */

function bootstrap_(p) {
  const user = auth_(p.token);
  return {
    ok:true,user:user,
    statuses:sheetData_('Statuses').filter(x=>truthy_(x.active)),
    sources:sheetData_('Sources').filter(x=>truthy_(x.active)),
    products:sheetData_('Products').filter(x=>truthy_(x.active)),
    dealers:sheetData_('Dealers').filter(x=>truthy_(x.active)).map(safeDealer_),
    users: user.role==='SUPER_ADMIN'||user.role==='MANAGER' ? sheetData_('Users').map(safeUser_) : [],
    settings:settingsObject_()
  };
}

/* ========================= DASHBOARD ========================= */

function dashboard_(p) {
  const user = auth_(p.token);
  const range = dateRange_(p.range,p.from,p.to);
  let orders = sheetData_('Orders').filter(x=>!x.deletedAt);
  orders = orders.filter(x=>inRange_(x.orderDate||x.createdAt,range.from,range.to));
  if (user.role==='AGENT') orders = orders.filter(x=>String(x.leadOwnerId||'')===String(user.id));
  const total = orders.length;
  const revenue = sum_(orders,'total');
  const online = sum_(orders,'online');
  const cod = sum_(orders,'balance');
  const status = groupCount_(orders,'status');
  const source = topGroup_(orders,'source');
  const product = topGroup_(orders,'product');
  const state = topGroup_(orders,'state');
  const today = startOfDay_(new Date());
  const followToday = orders.filter(x=>x.followupAt && sameDay_(new Date(x.followupAt),today)).length;
  const overdue = orders.filter(x=>x.followupAt && new Date(x.followupAt)<today && !isTerminal_(x.status)).length;
  const delivered = orders.filter(x=>['Delivered','GPO Delivered'].includes(String(x.status))).length;
  const cancelled = orders.filter(x=>['Cancelled','Final cancel','Double Cancel','Dealer Cancel'].includes(String(x.status))).length;
  return {ok:true,range,total,revenue,online,cod,confirmed:status['Confirmed']||0,pending:status['Pending']||0,delivered,cancelled,followToday,overdue,status,source,product,state,recent:orders.slice(-10).reverse(),followups:orders.filter(x=>x.followupAt && sameDay_(new Date(x.followupAt),today)).slice(0,20)};
}

/* ========================= ORDERS ========================= */

function orders_(p) {
  const user = auth_(p.token);
  let rows = sheetData_('Orders').filter(x=>!x.deletedAt);
  const f = p.filters||{};
  if (f.status) rows = rows.filter(x=>String(x.status)===String(f.status));
  if (f.source) rows = rows.filter(x=>String(x.source)===String(f.source));
  if (f.paymentStatus) rows = rows.filter(x=>String(x.paymentStatus)===String(f.paymentStatus));
  if (f.state) rows = rows.filter(x=>String(x.state)===String(f.state));
  if (f.city) rows = rows.filter(x=>String(x.city).toLowerCase().includes(String(f.city).toLowerCase()));
  if (f.phone) rows = rows.filter(x=>String(x.phone).includes(String(f.phone)));
  if (f.orderNo) rows = rows.filter(x=>String(x.orderNo).toLowerCase().includes(String(f.orderNo).toLowerCase()));
  if (f.awb) rows = rows.filter(x=>String(x.awb).toLowerCase().includes(String(f.awb).toLowerCase()));
  if (f.customer) rows = rows.filter(x=>String(x.customerName).toLowerCase().includes(String(f.customer).toLowerCase()));
  if (f.product) rows = rows.filter(x=>String(x.product).toLowerCase().includes(String(f.product).toLowerCase()));
  if (f.leadOwner) rows = rows.filter(x=>String(x.leadOwnerId||'')===String(f.leadOwner));
  if (f.dealerCode) rows = rows.filter(x=>String(x.dealerCode||'')===String(f.dealerCode));
  if (f.from || f.to) {
    const r = dateRange_('custom',f.from,f.to);
    rows = rows.filter(x=>inRange_(x.orderDate||x.createdAt,r.from,r.to));
  }
  if (user.role==='AGENT') rows = rows.filter(x=>String(x.leadOwnerId||'')===String(user.id));
  rows.sort((a,b)=>new Date(b.orderDate||b.createdAt)-new Date(a.orderDate||a.createdAt));
  const pageSize = Math.min(Number(p.pageSize||20),500);
  const page = Math.max(Number(p.page||1),1);
  const total = rows.length;
  const start=(page-1)*pageSize;
  return {ok:true,rows:rows.slice(start,start+pageSize),total,page,pageSize,pages:Math.max(Math.ceil(total/pageSize),1)};
}

function getOrder_(p) {
  auth_(p.token);
  const row = sheetData_('Orders').find(x=>String(x.id)===String(p.id)||String(x.orderNo)===String(p.id));
  if (!row) throw new Error('Order not found.');
  return {ok:true,order:row,history:sheetData_('OrderHistory').filter(x=>String(x.orderId)===String(row.id)),notes:sheetData_('Notes').filter(x=>String(x.orderId)===String(row.id)||String(x.phone)===String(row.phone))};
}

function saveOrder_(p) {
  const user = auth_(p.token);
  const o = Object.assign({},p.order||{});
  if (!o.customerName || !o.phone || !o.product) throw new Error('Customer name, phone and product are required.');
  o.qty = Number(o.qty||1);
  o.unitPrice = Number(o.unitPrice||0);
  o.total = round_(o.qty*o.unitPrice);
  o.online = Number(o.online||0);
  o.balance = Math.max(o.total-o.online,0);
  o.updatedAt = new Date();
  if (!o.id) {
    o.id=uid_('ORD');
    o.orderNo=o.orderNo||('ORD-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd')+'-'+String(Date.now()).slice(-6));
    o.createdAt=new Date();
    o.status=o.status||'New';
    o.paymentMode=o.paymentMode||'COD';
    o.paymentStatus=o.paymentStatus||'Pending';
    o.statusChangedAt=new Date();
    o.statusChangedBy=user.username;
    if (!o.leadOwnerId && user.role==='AGENT') {o.leadOwnerId=user.id;o.leadOwner=user.name;o.agentAssignedAt=new Date();}
    appendRowObject_('Orders',o);
    history_(o.id,'create','',o.status,user.username,false);
    audit_(user.username,'order.create',o.orderNo,'Created order');
  } else {
    const old = sheetData_('Orders').find(x=>String(x.id)===String(o.id));
    if (!old) throw new Error('Order not found.');
    if (String(old.status)!==String(o.status)) {
      o.statusChangedAt=new Date();
      o.statusChangedBy=user.username;
      history_(o.id,'status',old.status,o.status,user.username,false);
    }
    updateById_('Orders',o.id,o);
    audit_(user.username,'order.update',o.orderNo,'Updated order');
  }
  return {ok:true,order:o};
}

function deleteOrder_(p) {
  const user=auth_(p.token);
  const id=p.id;
  updateById_('Orders',id,{deletedAt:new Date(),deletedReason:p.reason||'Deleted'});
  audit_(user.username,'order.delete',id,'Soft deleted');
  return {ok:true};
}

function bulkStatus_(p) {
  const user=auth_(p.token);
  (p.ids||[]).forEach(id=>{
    const o=sheetData_('Orders').find(x=>String(x.id)===String(id));
    if (!o) return;
    updateById_('Orders',id,{status:p.status,remark:p.remark||o.remark,statusChangedAt:new Date(),statusChangedBy:user.username});
    history_(id,'status',o.status,p.status,user.username,false);
  });
  audit_(user.username,'order.bulkStatus',(p.ids||[]).join(','),p.status);
  return {ok:true};
}

function bulkAssign_(p) {
  const user=auth_(p.token);
  (p.ids||[]).forEach(id=>updateById_('Orders',id,{leadOwner:p.leadOwner||'',leadOwnerId:p.leadOwnerId||'',agentAssignedAt:new Date()}));
  audit_(user.username,'order.bulkAssign',(p.ids||[]).join(','),p.leadOwner||'Unassigned');
  return {ok:true};
}

function importOrders_(p) {
  const user=auth_(p.token);
  const rows=Array.isArray(p.rows)?p.rows:[];
  let count=0;
  rows.forEach(raw=>{
    const o=Object.assign({},raw);
    o.id=uid_('ORD');
    o.orderNo=o.orderNo||('ORD-'+Date.now()+'-'+count);
    o.orderDate=o.orderDate||new Date();
    o.qty=Number(o.qty||1); o.unitPrice=Number(o.unitPrice||0);
    o.total=Number(o.total||o.qty*o.unitPrice); o.online=Number(o.online||0);
    o.balance=Math.max(o.total-o.online,0);
    o.status=o.status||'New'; o.paymentMode=o.paymentMode||'COD'; o.paymentStatus=o.paymentStatus||'Pending';
    o.createdAt=new Date(); o.updatedAt=new Date(); o.statusChangedAt=new Date();
    appendRowObject_('Orders',o); count++;
  });
  audit_(user.username,'order.import','Orders','Imported '+count);
  return {ok:true,count};
}

/* ========================= DEALERS ========================= */

function dealers_(p) {
  auth_(p.token);
  let rows=sheetData_('Dealers').filter(x=>truthy_(x.active));
  const f=p.filters||{};
  if (f.search) {
    const q=String(f.search).toLowerCase();
    rows=rows.filter(x=>[x.name,x.code,x.username,x.mobile,x.city,x.state].some(v=>String(v||'').toLowerCase().includes(q)));
  }
  if (f.state) rows=rows.filter(x=>x.state===f.state);
  if (f.district) rows=rows.filter(x=>x.district===f.district);
  if (f.zm) rows=rows.filter(x=>x.zm===f.zm);
  return {ok:true,rows:rows.map(safeDealer_)};
}

function saveDealer_(p) {
  const user=auth_(p.token);
  const d=Object.assign({},p.dealer||{});
  if (!d.name) throw new Error('Dealer name is required.');
  if (!d.id) {
    d.id=uid_('DLR'); d.code=d.code||nextDealerCode_(); d.active=true; d.createdAt=new Date();
    appendRowObject_('Dealers',d);
    audit_(user.username,'dealer.create',d.code,'Created dealer');
  } else {
    updateById_('Dealers',d.id,d);
    audit_(user.username,'dealer.update',d.code||d.id,'Updated dealer');
  }
  return {ok:true,dealer:safeDealer_(d)};
}
function deleteDealer_(p) {
  const user=auth_(p.token);
  updateById_('Dealers',p.id,{active:false});
  audit_(user.username,'dealer.disable',p.id,'Dealer disabled');
  return {ok:true};
}
function safeDealer_(d) {
  return Object.assign({},d,{active:truthy_(d.active)});
}
function nextDealerCode_() {
  const rows=sheetData_('Dealers');
  let max=0;
  rows.forEach(d=>{const n=parseInt(String(d.code||'').replace(/\D/g,''),10);if(n>max)max=n;});
  return 'PHD'+String(max+1).padStart(4,'0');
}

/* ========================= LEDGER ========================= */

function ledger_(p) {
  auth_(p.token);
  let rows=sheetData_('Ledger');
  if (p.dealerCode) rows=rows.filter(x=>x.dealerCode===p.dealerCode);
  if (p.from||p.to) {const r=dateRange_('custom',p.from,p.to);rows=rows.filter(x=>inRange_(x.date,r.from,r.to));}
  rows.sort((a,b)=>new Date(a.date)-new Date(b.date));
  let running=0;
  rows=rows.map(x=>{running+=Number(x.credit||0)-Number(x.debit||0);x.running=running;return x;});
  const sale=sum_(sheetData_('Orders').filter(x=>x.dealerCode===p.dealerCode&&!x.deletedAt),'total');
  const paid=sum_(rows,'credit');
  return {ok:true,rows:rows.reverse(),summary:{sale,paid,balance:sale-paid}};
}
function addPayment_(p) {
  const user=auth_(p.token);
  const amount=Number(p.amount||0); if(amount<=0)throw new Error('Payment amount required.');
  appendRowObject_('Ledger',{id:uid_('TXN'),dealerCode:p.dealerCode,date:p.date||new Date(),particular:p.particular||'Payment',debit:0,credit:amount,net:amount,running:0,reference:p.reference||''});
  audit_(user.username,'dealer.payment',p.dealerCode,'Added payment '+amount);
  return {ok:true};
}

/* ========================= INVOICES ========================= */

function invoices_(p) {
  auth_(p.token);
  let rows=sheetData_('Invoices');
  if(p.dealerCode)rows=rows.filter(x=>x.dealerCode===p.dealerCode);
  return {ok:true,rows:rows.sort((a,b)=>new Date(b.date)-new Date(a.date))};
}
function saveInvoice_(p) {
  const user=auth_(p.token);
  const inv=Object.assign({},p.invoice||{});
  if(!inv.dealerCode)throw new Error('Dealer is required.');
  inv.itemsJson=JSON.stringify(inv.items||[]);
  inv.grandTotal=Number(inv.grandTotal||0);inv.paidAmount=Number(inv.paidAmount||0);
  inv.balance=Math.max(inv.grandTotal-inv.paidAmount,0);
  inv.status=inv.balance>0?'Pending':'Paid';
  if(!inv.id){inv.id=uid_('INV');inv.invoiceNo=inv.invoiceNo||('PH-DLR-'+String(Date.now()).slice(-6));inv.date=inv.date||new Date();inv.createdAt=new Date();appendRowObject_('Invoices',inv);}
  else updateById_('Invoices',inv.id,inv);
  audit_(user.username,'dealer.invoice',inv.invoiceNo,'Invoice saved');
  return {ok:true,invoice:inv};
}

/* ========================= REPORTS ========================= */

function reports_(p) {
  auth_(p.token);
  let orders=sheetData_('Orders').filter(x=>!x.deletedAt);
  const r=dateRange_(p.range,p.from,p.to);
  orders=orders.filter(x=>inRange_(x.orderDate||x.createdAt,r.from,r.to));
  if(p.source)orders=orders.filter(x=>x.source===p.source);
  const delivered=orders.filter(x=>['Delivered','GPO Delivered'].includes(x.status));
  const status=groupCount_(orders,'status');
  const sources={};
  orders.forEach(o=>{
    const s=o.source||'Unknown';
    sources[s]=sources[s]||{source:s,total:0,confirmed:0,delivered:0,revenue:0};
    sources[s].total++;
    if(o.status==='Confirmed')sources[s].confirmed++;
    if(['Delivered','GPO Delivered'].includes(o.status)){sources[s].delivered++;sources[s].revenue+=Number(o.total||0);}
  });
  return {ok:true,range:r,total:orders.length,revenue:sum_(orders,'total'),delivered:delivered.length,aov:delivered.length?round_(sum_(delivered,'total')/delivered.length):0,status,sources:Object.values(sources),states:topGroup_(orders,'state'),dealer:dealerReport_(orders)};
}
function incentive_(p) {
  const user=auth_(p.token);
  let orders=sheetData_('Orders').filter(x=>!x.deletedAt&&['Delivered','GPO Delivered'].includes(x.status));
  const r=dateRange_(p.range,p.from,p.to);
  orders=orders.filter(x=>inRange_(x.deliveredAt||x.orderDate||x.createdAt,r.from,r.to));
  if(p.agentId)orders=orders.filter(x=>String(x.leadOwnerId)===String(p.agentId));
  const users=sheetData_('Users');
  const map={};
  orders.forEach(o=>{
    const key=String(o.leadOwnerId||'unassigned');
    const u=users.find(x=>String(x.id)===key);
    const mode=(u&&u.workMode)||'OFFICE';
    const online=Number(o.online||0), cod=Number(o.balance||0);
    let inc=0;
    if(mode==='WFH') inc=online*.15+cod*.10;
    else if(Number(o.total||0)>1000) inc=online*.15+Math.max(cod-1000,0)*.10;
    map[key]=map[key]||{agent:u?u.name:'Unassigned',mode,orders:0,total:0,online:0,cod:0,incentive:0};
    map[key].orders++;map[key].total+=Number(o.total||0);map[key].online+=online;map[key].cod+=cod;map[key].incentive+=inc;
  });
  return {ok:true,rows:Object.values(map).map(x=>Object.assign(x,{online15:round_(x.online*.15),cod10:round_(x.cod*.10),incentive:round_(x.incentive)}))};
}
function dealerReport_(orders){
  const m={};
  orders.filter(o=>o.dealerCode).forEach(o=>{
    const k=o.dealerCode;m[k]=m[k]||{dealer:k,orders:0,qty:0,total:0,delivered:0};
    m[k].orders++;m[k].qty+=Number(o.qty||0);m[k].total+=Number(o.total||0);
    if(['Delivered','GPO Delivered'].includes(o.status))m[k].delivered++;
  });
  return Object.values(m);
}

/* ========================= USERS ========================= */

function users_(p) {
  const me=auth_(p.token);
  if(!['SUPER_ADMIN','MANAGER'].includes(me.role))throw new Error('Access denied.');
  return {ok:true,rows:sheetData_('Users').map(safeUser_)};
}
function saveUser_(p) {
  const me=auth_(p.token);
  if(!['SUPER_ADMIN','MANAGER'].includes(me.role))throw new Error('Access denied.');
  const u=Object.assign({},p.user||{});
  if(!u.name||!u.username||!u.password||!u.role)throw new Error('Name, username, password and role are required.');
  u.permissions=JSON.stringify(u.permissions||{});
  u.active=u.active!==false;
  if(!u.id){u.id=uid_('USR');u.createdAt=new Date();appendRowObject_('Users',u);}
  else updateById_('Users',u.id,u);
  audit_(me.username,'user.save',u.username,'Saved user');
  return {ok:true,user:safeUser_(u)};
}

/* ========================= SETTINGS ========================= */

function settings_(p) {
  const me=auth_(p.token);
  if(me.role!=='SUPER_ADMIN')throw new Error('SUPER_ADMIN only.');
  return {ok:true,settings:settingsObject_()};
}
function saveSettings_(p) {
  const me=auth_(p.token);
  if(me.role!=='SUPER_ADMIN')throw new Error('SUPER_ADMIN only.');
  Object.keys(p.settings||{}).forEach(k=>setSetting_(k,p.settings[k]));
  audit_(me.username,'settings.update','SETTINGS','Updated settings');
  return {ok:true,settings:settingsObject_()};
}
function settingsObject_(){
  const o={};sheetData_('Settings').forEach(x=>o[x.key]=parseJson_(x.value,x.value));return o;
}
function setSetting_(key,value){
  const rows=sheetData_('Settings');const found=rows.find(x=>x.key===key);
  if(found)updateById_('Settings',key,{value:JSON.stringify(value)});
  else appendRowObject_('Settings',{key,value:JSON.stringify(value)});
}

/* ========================= AUDIT / CALLS / NOTES / HEALTH ========================= */

function auditList_(p) {
  auth_(p.token);
  let rows=sheetData_('AuditLogs');
  if(p.user)rows=rows.filter(x=>String(x.user).toLowerCase().includes(String(p.user).toLowerCase()));
  if(p.action)rows=rows.filter(x=>String(x.action).toLowerCase().includes(String(p.action).toLowerCase()));
  if(p.module)rows=rows.filter(x=>String(x.module).toLowerCase()===String(p.module).toLowerCase());
  return {ok:true,rows:rows.reverse().slice(0,1000)};
}
function calls_(p) {
  auth_(p.token);
  let rows=sheetData_('Calls');
  if(p.agent)rows=rows.filter(x=>x.agent===p.agent);
  return {ok:true,rows:rows.reverse().slice(0,1000)};
}
function saveNote_(p) {
  const user=auth_(p.token);
  if(!p.orderId||!p.text)throw new Error('Order and note are required.');
  const o=sheetData_('Orders').find(x=>String(x.id)===String(p.orderId));
  appendRowObject_('Notes',{id:uid_('NOTE'),orderId:p.orderId,phone:o?o.phone:'',text:p.text,user:user.name,createdAt:new Date()});
  audit_(user.username,'note.add',p.orderId,'Shared note added');
  return {ok:true};
}
function health_(p) {
  auth_(p.token);
  const ss=SpreadsheetApp.getActive();
  const names=Object.keys(CRM.SHEETS).map(k=>CRM.SHEETS[k]);
  const counts={};names.forEach(n=>counts[n]=ss.getSheetByName(n)?Math.max(ss.getSheetByName(n).getLastRow()-1,0):0);
  return {ok:true,time:new Date(),spreadsheet:ss.getName(),counts};
}

/* ========================= HELPERS ========================= */

function sheetData_(name){
  const sh=SpreadsheetApp.getActive().getSheetByName(name);
  if(!sh||sh.getLastRow()<2)return [];
  const values=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getValues();
  const heads=values.shift();
  return values.filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>{
    const o={};heads.forEach((h,i)=>o[h]=r[i]);return o;
  });
}
function appendRowObject_(name,obj){
  const sh=SpreadsheetApp.getActive().getSheetByName(name);
  const heads=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.appendRow(heads.map(h=>obj[h]===undefined?'':obj[h]));
}
function updateById_(name,id,patch){
  const sh=SpreadsheetApp.getActive().getSheetByName(name);
  if(!sh)return;
  const values=sh.getDataRange().getValues();const heads=values.shift();
  const idIdx=heads.indexOf('id');if(idIdx<0)throw new Error('No id column in '+name);
  for(let i=0;i<values.length;i++){
    if(String(values[i][idIdx])===String(id)){
      Object.keys(patch).forEach(k=>{
        const c=heads.indexOf(k);if(c>=0)sh.getRange(i+2,c+1).setValue(patch[k]);
      });
      return;
    }
  }
}
function seedIfEmpty_(name,heads,rows){
  const sh=SpreadsheetApp.getActive().getSheetByName(name);
  if(sh.getLastRow()<=1)rows.forEach(r=>sh.appendRow(r));
}
function uid_(prefix){return prefix+'_'+Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase();}
function parseJson_(v,def){try{return typeof v==='object'?v:JSON.parse(String(v||''));}catch(e){return def;}}
function truthy_(v){return v===true||String(v).toLowerCase()==='true'||String(v)==='1';}
function round_(n){return Math.round(Number(n||0)*100)/100;}
function sum_(rows,key){return round_(rows.reduce((a,r)=>a+Number(r[key]||0),0));}
function groupCount_(rows,key){const o={};rows.forEach(r=>{const k=r[key]||'Unknown';o[k]=(o[k]||0)+1;});return o;}
function topGroup_(rows,key){return Object.entries(groupCount_(rows,key)).sort((a,b)=>b[1]-a[1]).slice(0,10).map(x=>({name:x[0],count:x[1]}));}
function isTerminal_(status){return ['Delivered','GPO Delivered','Cancelled','Final cancel','Lost','RTO','Double Cancel'].includes(String(status));}
function history_(orderId,field,oldValue,newValue,byUser,system){appendRowObject_('OrderHistory',{id:uid_('HIS'),orderId,field,oldValue,newValue,byUser,system,createdAt:new Date()});}
function audit_(user,action,target,details){try{appendRowObject_('AuditLogs',{id:uid_('AUD'),time:new Date(),user,module:String(action).split('.')[0]||'SYSTEM',action,target,details});}catch(e){}}
function dateRange_(range,from,to){
  const now=new Date();let a=startOfDay_(now),b=endOfDay_(now);
  if(range==='today'){a=startOfDay_(now);}
  else if(range==='yesterday'){a=startOfDay_(new Date(now-86400000));b=endOfDay_(new Date(now-86400000));}
  else if(range==='3days'){a=startOfDay_(new Date(now-2*86400000));}
  else if(range==='7days'){a=startOfDay_(new Date(now-6*86400000));}
  else if(range==='15days'){a=startOfDay_(new Date(now-14*86400000));}
  else if(range==='30days'){a=startOfDay_(new Date(now-29*86400000));}
  else if(range==='thismonth'){a=new Date(now.getFullYear(),now.getMonth(),1);}
  else if(range==='lastmonth'){a=new Date(now.getFullYear(),now.getMonth()-1,1);b=new Date(now.getFullYear(),now.getMonth(),0,23,59,59);}
  else if(range==='custom'){if(from)a=new Date(from);if(to)b=endOfDay_(new Date(to));}
  else {a=new Date(0);b=endOfDay_(now);}
  return {from:a,to:b};
}
function startOfDay_(d){return new Date(d.getFullYear(),d.getMonth(),d.getDate());}
function endOfDay_(d){return new Date(d.getFullYear(),d.getMonth(),d.getDate(),23,59,59,999);}
function sameDay_(a,b){return a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate();}
function inRange_(v,a,b){const d=new Date(v);return !isNaN(d)&&d>=a&&d<=b;}
function json_(x){return JSON.stringify(x);}
