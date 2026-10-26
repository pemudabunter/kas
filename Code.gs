/** Code.gs — Backend Kas, Aset & Inventaris Pemuda Bunter (Google Sheets + Drive)
 *  Jalankan fungsi setup() SEKALI dari editor Apps Script (aman diulang), lalu Deploy > Web app.
 *  Sheet yang dibuat: Users, Transaksi, Tahun, Kategori, Pengaturan, Log, Aset. */
const KAWAL='Kas awal kepemudaan',KBELI='Pengadaan/Pembelian Aset',KJUAL='Hasil Penjualan Aset';
const CFG={SESSION_SEC:1800,DEFAULT_YEAR:2026,FOLDER:'Bukti Kas Pemuda Bunter',MAX_B64:4000000};
const H={
  Users:['username','nama','role','aktif','password'],
  Transaksi:['id','tgl','jenis','kategori','nominal','pihak','ket','metode','tunai','bukti','status','voidAlasan','voidOleh','voidWaktu','dibuatOleh'],
  Tahun:['tahun','ditutup'],
  Kategori:['id','jenis','nama','anggaran'],
  Pengaturan:['kunci','nilai'],
  Log:['waktu','user','aksi'],
  Aset:['id','nama','qty','satuan','harga','kategori','ket','dibuatOleh'],
  Pengajuan:['id','tahun','kategori','keperluan','jumlah','tgl','metode','pihak','bukti','status','alasan','diajukanOleh','diputusOleh','trxId','waktu']};
const ROLES=['superadmin','admin','ketua','anggota'],WR=['superadmin','admin'],SA=['superadmin'],APR=['ketua','superadmin']; // APR = penyetuju pengajuan dana (Ketua/Wakil Ketua dan Super Admin)
const RLBL={superadmin:'Super Admin',admin:'Admin (Bendahara)',ketua:'Ketua / Wakil Ketua',anggota:'Anggota'};
const DEMO={superadmin:'admin123',bendahara:'admin123',ketua:'ketua123',anggota:'anggota123'}; // akun bawaan setup()

/* ===== SETUP SEKALI JALAN ===== */
function setup(){
  const ss=SpreadsheetApp.getActive();
  Object.keys(H).forEach(n=>{
    const s=ss.getSheetByName(n)||ss.insertSheet(n);
    if(s.getLastRow()===0){s.getRange(1,1,1,H[n].length).setValues([H[n]]).setFontWeight('bold').setBackground('#1e3a6e').setFontColor('#ffffff');s.setFrozenRows(1);}
  });
  const t=sh('Transaksi');t.getRange('B:B').setNumberFormat('@');t.getRange('N:N').setNumberFormat('@');sh('Log').getRange('A:A').setNumberFormat('@');sh('Pengajuan').getRange('F:F').setNumberFormat('@');
  const P=PropertiesService.getScriptProperties();if(!P.getProperty('SALT'))P.setProperty('SALT',Utilities.getUuid());
  if(!rows('Users').length)[['superadmin','Super Admin','superadmin','admin123'],['ketua','Ketua Pemuda','ketua','ketua123'],['bendahara','Bendahara','admin','admin123'],['anggota','Anggota','anggota','anggota123']]
    .forEach(u=>put('Users',{username:u[0],nama:u[1],role:u[2],aktif:true,password:hash(u[3])}));
  if(!rows('Kategori').length)seedKategori();
  if(!rows('Aset').length)seedAset();
  if(!rows('Pengaturan').length)put('Pengaturan',{kunci:'tahunBuku',nilai:CFG.DEFAULT_YEAR});
  Logger.log('Setup selesai. Ganti password bawaan, lalu isi saldo awal di Master Data > Setting Tahun Buku.');
}

function seedKategori(){
  [['Masuk','Iuran anggota pemuda',0],['Masuk','Dana kegiatan kepemudaan',0],['Masuk','Dana kegiatan olahraga',0],['Masuk','Sumbangan / donasi',0],
   ['Masuk','Dana kerja sama dengan pihak lain',0],['Masuk','Pendapatan kegiatan lainnya',0],['Masuk',KJUAL,0],
   ['Keluar','Dana kegiatan kepemudaan',2000000],['Keluar','Dana kegiatan olahraga',1500000],['Keluar','Sumbangan / donasi',500000],
   ['Keluar','Dana kerja sama dengan pihak lain',0],['Keluar','Dana pemeliharaan fasilitas',1000000],['Keluar',KBELI,0]]
  .forEach(k=>put('Kategori',{id:nextId('Kategori'),jenis:k[0],nama:k[1],anggaran:k[2]}));
}
function seedAset(){
  [['Kambing',4,'ekor',2500000,'Ternak'],['Sapi',1,'ekor',12000000,'Ternak'],['Kursi',50,'unit',85000,'Perabot'],['Jersey',2,'set',900000,'Perlengkapan Olahraga'],
   ['Sarung tangan kiper',3,'pasang',250000,'Perlengkapan Olahraga'],['Peralatan olahraga',1,'set',1500000,'Perlengkapan Olahraga'],
   ['Peralatan kegiatan',1,'set',2000000,'Perlengkapan Kegiatan'],['Inventaris lainnya',1,'set',500000,'Lainnya']]
  .forEach(a=>put('Aset',{id:nextId('Aset'),nama:a[0],qty:a[1],satuan:a[2],harga:a[3],kategori:a[4],ket:'',dibuatOleh:'sistem'}));
}
/** Jalankan SEKALI jika sheet lama masih memakai kategori/istilah lama: ganti daftar kategori & ubah 'Saldo Awal' jadi 'Kas awal kepemudaan'. */
function migrasi(){
  inv();const k=sh('Kategori');if(k.getLastRow()>1)k.deleteRows(2,k.getLastRow()-1);seedKategori();
  inv();const s=sh('Transaksi'),n=s.getLastRow();
  if(n>1){const r=s.getRange(2,4,n-1,1);r.setValues(r.getValues().map(x=>[x[0]==='Saldo Awal'?KAWAL:x[0]]));}
  Logger.log('Migrasi selesai. Transaksi lama dengan kategori lama perlu disesuaikan manual atau direset.');
}

/* ===== ENTRY POINT ===== */
function doGet(){return json({ok:true,msg:'API Kas Pemuda Bunter aktif'});} // sengaja tanpa data
const READ=['login','dashboard','master','bootstrap']; // aksi baca: tanpa kunci supaya tidak antre
function doPost(e){
  let lock=null,out;
  try{
    const p=JSON.parse(e.postData.contents);
    if(READ.indexOf(p.action)<0){lock=LockService.getScriptLock();lock.waitLock(20000);} // hanya aksi tulis yang diserialkan
    out=route(p);
  }
  catch(err){out={ok:false,msg:String(err.message||err)};}
  finally{try{if(lock)lock.releaseLock();}catch(x){}}
  return json(out);
}
const json=o=>ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
function route(p){
  switch(p.action){
    case 'login':return login(p);
    case 'dashboard':return dashboard(p);
    case 'bootstrap':return bootstrap(p);
    case 'master':auth(p);return master();
    case 'transaction':return transaction(p);
    case 'pengajuan':return pengajuanOp(p);
    case 'category':return category(p);
    case 'settings':return settings(p);
    case 'user':return userOp(p);
    case 'aset':return asetOp(p);
    case 'resetAll':return resetAll(p);
    default:throw new Error('Aksi tidak dikenal');
  }
}

/* ===== UTIL SHEET ===== */
const _ready={};
function sh(n){
  const ss=SpreadsheetApp.getActive();
  if(!ss)throw new Error('Script belum terhubung ke Google Sheets. Buat script lewat menu Extensions > Apps Script DI DALAM Spreadsheet.');
  let s=ss.getSheetByName(n);
  if(!s&&H[n])s=ss.insertSheet(n);
  if(s&&H[n]&&!_ready[n]){ // pastikan header ada (sekali per eksekusi)
    if(s.getLastRow()===0){s.getRange(1,1,1,H[n].length).setValues([H[n]]).setFontWeight('bold').setBackground('#1e3a6e').setFontColor('#ffffff');s.setFrozenRows(1);}
    _ready[n]=true;
  }
  return s;
}
/** Dipanggil saat login: bila sheet/akun awal/kunci belum ada (setup belum dijalankan), siapkan otomatis. */
function ensureInit(){
  const ss=SpreadsheetApp.getActive();
  if(!ss)throw new Error('Script belum terhubung ke Google Sheets. Buat script lewat menu Extensions > Apps Script DI DALAM Spreadsheet.');
  const need=()=>!PropertiesService.getScriptProperties().getProperty('SALT')||Object.keys(H).some(n=>!ss.getSheetByName(n))||!rows('Users').length;
  if(!need())return;
  const L=LockService.getScriptLock();L.waitLock(20000);
  try{inv();if(need())setup();}finally{L.releaseLock();}
}
const _rc={}; // cache isi sheet selama satu eksekusi (dibersihkan tiap ada penulisan)
const inv=n=>{if(n)delete _rc[n];else Object.keys(_rc).forEach(k=>delete _rc[k]);};
function rows(n){
  let c=_rc[n];if(!c){const v=sh(n).getDataRange().getValues();c=_rc[n]={h:v.shift(),v};}
  const v=c.v,h=c.h;
  return v.map((r,i)=>{const o={_r:i+2};h.forEach((k,j)=>o[k]=r[j]);return o;}).filter(o=>H[n].some(k=>o[k]!==''));}
const put=(n,o)=>{inv(n);sh(n).appendRow(H[n].map(k=>o[k]==null?'':o[k]));};
const upd=(n,o)=>{inv(n);sh(n).getRange(o._r,1,1,H[n].length).setValues([H[n].map(k=>o[k]==null?'':o[k])]);};
const ds=v=>v instanceof Date?Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd'):String(v==null?'':v);
const bool=v=>v===true||String(v).toUpperCase()==='TRUE';
const now=()=>Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm:ss');
const rp=n=>'Rp '+String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g,'.');
const log=(u,a)=>sh('Log').appendRow([now(),u,a]);
const nextId=n=>rows(n).reduce((m,r)=>Math.max(m,+r.id||0),0)+1;
let _salt=null;const salt=()=>_salt||(_salt=PropertiesService.getScriptProperties().getProperty('SALT'));
const hash=pw=>Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,salt()+pw));
/** true selama akun masih memakai username + password bawaan demo */
const isDemo=r=>DEMO[r.username]!==undefined&&r.password===hash(DEMO[r.username]);
const isClosed=y=>rows('Tahun').some(r=>String(r.tahun)===String(y)&&bool(r.ditutup));
function tx(o){return{id:+o.id,tgl:ds(o.tgl),jenis:o.jenis,kategori:o.kategori,nominal:+o.nominal,pihak:o.pihak,ket:o.ket,
  metode:o.metode==='Transfer'?'Transfer':'Tunai',tunai:o.tunai===''?undefined:+o.tunai,bukti:o.bukti,status:o.status||'Aktif',
  voidAlasan:o.voidAlasan,voidOleh:o.voidOleh,voidWaktu:ds(o.voidWaktu)};}
function bal(y,ex){y=String(y);const b={Tunai:0,Transfer:0};
  rows('Transaksi').map(tx).forEach(t=>{
    if(t.status==='Void'||t.id===ex||t.tgl.slice(0,4)!==y)return;
    if(t.kategori===KAWAL){const u=t.tunai!=null?t.tunai:(t.metode==='Tunai'?t.nominal:0);b.Tunai+=u;b.Transfer+=t.nominal-u;}
    else b[t.metode]+=(t.jenis==='Masuk'?t.nominal:-t.nominal);});
  return b;}

/* ===== AUTH & ROLE (SERVER-SIDE) ===== */
function login(p){
  ensureInit();
  const C=CacheService.getScriptCache(),k='fail_'+String(p.username).toLowerCase();
  if(+C.get(k)>=5)throw new Error('Terlalu banyak percobaan gagal. Coba lagi 10 menit lagi.');
  const u=rows('Users').find(r=>String(r.username).toLowerCase()===String(p.username).toLowerCase());
  if(!u||u.password!==hash(p.password||'')){C.put(k,String((+C.get(k)||0)+1),600);log(p.username||'?','Login gagal');return{ok:false,msg:'Username atau password salah'};}
  if(!bool(u.aktif))return{ok:false,msg:'Akun dinonaktifkan'};
  C.remove(k);const token=Utilities.getUuid()+Utilities.getUuid();C.put(token,sesVal(u),CFG.SESSION_SEC);log(u.username,'Login');
  return{ok:true,user:{username:u.username,nama:u.nama,role:u.role,token}};
}
const sesVal=u=>u.username+'|'+String(u.password).slice(-16);
/** Peran SELALU dibaca ulang dari sheet Users, bukan dari kiriman klien. */
function auth(p,roles){
  const C=CacheService.getScriptCache(),sv=C.get(p.token||'-');
  if(!sv)throw new Error('Sesi berakhir, silakan login ulang');
  const u=rows('Users').find(r=>sesVal(r)===sv); // username/password berubah atau user dihapus => sesi lama tidak berlaku
  if(!u||!bool(u.aktif))throw new Error('Sesi berakhir, silakan login ulang');
  if(roles&&roles.indexOf(u.role)<0)throw new Error('Akses ditolak: aksi ini hanya untuk '+roles.map(r=>RLBL[r]||r).join(' / ')+'. Peran Anda: '+(RLBL[u.role]||u.role));
  C.put(p.token,sv,CFG.SESSION_SEC); // sesi geser (sliding)
  return u;
}

/* ===== BACA DATA ===== */
function dashboard(p){
  const u=auth(p),sa=u.role==='superadmin',closed={};
  rows('Tahun').forEach(r=>{if(bool(r.ditutup))closed[r.tahun]=true;});
  return{ok:true,closed,
    aset:rows('Aset').map(as),pengajuan:u.role==='anggota'?[]:rows('Pengajuan').map(pj),
    trx:rows('Transaksi').map(tx).map(t=>{if(u.role==='anggota')t.bukti='';return t;}),
    logs:sa?recentLogs(100):[],
    users:sa?rows('Users').map(r=>({username:r.username,nama:r.nama,role:r.role,aktif:bool(r.aktif),demo:isDemo(r)})):[]};
}
function recentLogs(n){
  const s=sh('Log'),l=s.getLastRow();if(l<2)return[];
  const st=Math.max(2,l-n+1);
  return s.getRange(st,1,l-st+1,3).getValues().reverse().filter(r=>r[0]!==''||r[1]!==''||r[2]!=='').map(r=>({waktu:ds(r[0]),user:r[1],aksi:r[2]}));
}
/** Dashboard + master dalam SATU permintaan (lebih cepat daripada dua panggilan terpisah). */
function bootstrap(p){
  const d=dashboard(p),m=master();
  return Object.assign(d,{categories:m.categories,settings:m.settings});
}
function master(){
  const s=rows('Pengaturan').find(r=>r.kunci==='tahunBuku');
  return{ok:true,categories:rows('Kategori').map(r=>({id:+r.id,jenis:r.jenis,nama:r.nama,anggaran:+r.anggaran||0})),
    settings:{tahunBuku:s?+s.nilai:CFG.DEFAULT_YEAR}};
}

/* ===== TRANSAKSI ===== */
function transaction(p){
  if(p.op==='closeYear')return closeYear(auth(p,SA),+p.year);
  if(p.op==='void')return voidTx(auth(p,WR),p);
  return saveTx(auth(p,WR),p.data||{});
}
function saveTx(u,d,auto){
  const old=d.id?rows('Transaksi').find(r=>+r.id===+d.id):null,o=old?tx(old):null;
  if(d.id&&!old)throw new Error('Transaksi tidak ditemukan');
  if(o&&o.status==='Void')throw new Error('Transaksi yang sudah di-void tidak dapat diubah');
  if(o&&o.kategori===KAWAL)throw new Error('Kas awal tidak dapat diubah');
  if(o&&(o.kategori===KBELI||o.kategori===KJUAL))throw new Error('Transaksi otomatis dari modul Aset tidak dapat diedit, gunakan void');
  const n={jenis:d.jenis,kategori:String(d.kategori||''),nominal:Math.round(+d.nominal),tgl:String(d.tgl||''),
    pihak:String(d.pihak||'').trim().slice(0,100),ket:String(d.ket||'').trim().slice(0,200),metode:d.metode==='Transfer'?'Transfer':'Tunai'};
  if(['Masuk','Keluar'].indexOf(n.jenis)<0)throw new Error('Jenis transaksi tidak valid');
  if(!(n.nominal>0&&n.nominal<=1e11))throw new Error('Nominal tidak valid');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(n.tgl)||isNaN(new Date(n.tgl+'T00:00:00')))throw new Error('Format tanggal tidak valid');
  if(!n.pihak)throw new Error('Sumber/penerima wajib diisi');
  if(!rows('Kategori').some(c=>c.jenis===n.jenis&&c.nama===n.kategori))throw new Error('Kategori tidak valid');
  if(!auto&&(n.kategori===KBELI||n.kategori===KJUAL))throw new Error('Kategori otomatis hanya dibuat dari modul Aset');
  const y=n.tgl.slice(0,4);
  if(isClosed(y))throw new Error('Tahun buku '+y+' sudah ditutup (terkunci)');
  if(o&&o.tgl.slice(0,4)!==y)throw new Error('Tanggal tidak boleh dipindah ke tahun buku lain');
  const b=bal(y,o?o.id:-1);   // saldo guard: saldo metode terkait tidak boleh negatif
  if(b[n.metode]+(n.jenis==='Masuk'?n.nominal:-n.nominal)<0)throw new Error('Saldo '+n.metode+' tidak cukup. Tersedia '+rp(b[n.metode]));
  if(o&&o.jenis==='Masuk'&&o.metode!==n.metode&&b[o.metode]<0)throw new Error('Mengubah metode akan membuat saldo '+o.metode+' negatif');
  let bukti=o?o.bukti:(d.buktiUrl||'');if(d.buktiFile)bukti=saveFile(d.buktiFile,y);
  const row=old?Object.assign({},old):{id:nextId('Transaksi'),status:'Aktif',dibuatOleh:u.username,tunai:''};
  Object.assign(row,n,{bukti});
  if(old){const df=['tgl','jenis','kategori','nominal','pihak','ket','metode'].filter(k=>String(o[k])!==String(n[k])).map(k=>k+': '+o[k]+' → '+n[k]);
    upd('Transaksi',row);log(u.username,'Edit transaksi #'+row.id+(df.length?' ['+df.join('; ')+']':'')+(d.buktiFile?' [bukti diganti]':''));}
  else{put('Transaksi',row);log(u.username,'Tambah transaksi #'+row.id+' '+n.jenis+' '+rp(n.nominal)+' ('+n.metode+')');}
  return{ok:true};
}
function voidTx(u,p){
  const r=rows('Transaksi').find(x=>+x.id===+p.id);if(!r)throw new Error('Transaksi tidak ditemukan');
  const t=tx(r),y=t.tgl.slice(0,4),alasan=String(p.alasan||'').trim();
  if(t.status==='Void')throw new Error('Transaksi sudah di-void');
  if(t.kategori===KAWAL)throw new Error('Saldo Awal tidak dapat di-void');
  if(isClosed(y))throw new Error('Tahun buku '+y+' sudah ditutup (terkunci)');
  if(alasan.length<3)throw new Error('Alasan void wajib diisi');
  if(t.jenis==='Masuk'&&bal(y,t.id)[t.metode]<0)throw new Error('Void ditolak: saldo '+t.metode+' akan menjadi negatif');
  Object.assign(r,{status:'Void',voidAlasan:alasan.slice(0,200),voidOleh:u.username,voidWaktu:now()});
  upd('Transaksi',r);log(u.username,'VOID transaksi #'+t.id+' ('+t.jenis+' '+rp(t.nominal)+', '+t.kategori+') alasan: '+alasan);
  return{ok:true};
}
function closeYear(u,y){
  if(!(y>=2000))throw new Error('Tahun tidak valid');
  if(isClosed(y))throw new Error('Tahun buku '+y+' sudah ditutup');
  const b=bal(y),ny=String(y+1);
  const ex=rows('Transaksi').find(r=>ds(r.tgl).slice(0,4)===ny&&r.kategori===KAWAL&&r.status!=='Void');
  const row={id:ex?+ex.id:nextId('Transaksi'),tgl:ny+'-01-01',jenis:'Masuk',kategori:KAWAL,nominal:b.Tunai+b.Transfer,pihak:'Saldo awal',
    ket:'Saldo Awal Kas '+ny,metode:'Tunai',tunai:b.Tunai,bukti:'',status:'Aktif',dibuatOleh:u.username};
  if(ex){row._r=ex._r;upd('Transaksi',row);}else put('Transaksi',row);
  const th=rows('Tahun').find(r=>String(r.tahun)===String(y));
  if(th){th.ditutup=true;upd('Tahun',th);}else put('Tahun',{tahun:y,ditutup:true});
  log(u.username,'Tutup tahun buku '+y+' (saldo akhir '+rp(row.nominal)+')');
  return{ok:true};
}
function saveFile(f,y){
  const m=/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(f.dataUrl||'');
  if(!m||m[2].length>CFG.MAX_B64)throw new Error('File bukti tidak valid atau terlalu besar');
  const P=PropertiesService.getScriptProperties();let fo;
  try{fo=DriveApp.getFolderById(P.getProperty('FOLDER_ID'));}catch(e){fo=DriveApp.createFolder(CFG.FOLDER);P.setProperty('FOLDER_ID',fo.getId());}
  const name=y+'_'+Date.now()+'.'+m[1].split('/')[1];
  const file=fo.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]),m[1],name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  return file.getUrl();
}

/* ===== KATEGORI, SETTING, USER ===== */
function category(p){
  const u=auth(p,WR),C=rows('Kategori');
  if(p.op==='delete'){
    const c=C.find(r=>+r.id===+p.id);if(!c)throw new Error('Kategori tidak ditemukan');
    if(c.nama===KBELI||c.nama===KJUAL)throw new Error('Kategori otomatis tidak dapat dihapus');
    if(rows('Transaksi').some(t=>t.kategori===c.nama&&t.jenis===c.jenis))throw new Error('Kategori masih dipakai transaksi, tidak bisa dihapus');
    if(rows('Pengajuan').some(a=>a.kategori===c.nama))throw new Error('Kategori masih punya data pengajuan dana, tidak bisa dihapus');
    inv('Kategori');sh('Kategori').deleteRow(c._r);log(u.username,'Hapus kategori '+c.nama);return{ok:true};
  }
  const d=p.data||{},nama=String(d.nama||'').trim().slice(0,50),jenis=d.jenis;
  if(!nama||['Masuk','Keluar'].indexOf(jenis)<0)throw new Error('Data kategori tidak valid');
  if([KAWAL,KBELI,KJUAL].indexOf(nama)>=0)throw new Error('Nama kategori ini dicadangkan sistem');
  const ang=jenis==='Keluar'?Math.max(0,Math.round(+d.anggaran||0)):0;
  if(C.some(r=>r.jenis===jenis&&String(r.nama).toLowerCase()===nama.toLowerCase()&&+r.id!==+d.id))throw new Error('Kategori sudah ada');
  if(d.id){
    const c=C.find(r=>+r.id===+d.id);if(!c)throw new Error('Kategori tidak ditemukan');
    if(c.nama===KBELI||c.nama===KJUAL)throw new Error('Kategori otomatis tidak dapat diubah');
    if(c.nama!==nama||c.jenis!==jenis){
      const used=rows('Transaksi').filter(t=>t.kategori===c.nama&&t.jenis===c.jenis);
      if(c.jenis!==jenis&&used.length)throw new Error('Jenis tidak dapat diubah karena kategori sudah dipakai transaksi');
      inv('Transaksi');used.forEach(t=>sh('Transaksi').getRange(t._r,4).setValue(nama));
      inv('Pengajuan');rows('Pengajuan').filter(a=>a.kategori===c.nama).forEach(a=>{a.kategori=nama;upd('Pengajuan',a);});
    }
    Object.assign(c,{nama,jenis,anggaran:ang});upd('Kategori',c);log(u.username,'Edit kategori '+nama+' (anggaran '+rp(ang)+')');
  }else{put('Kategori',{id:nextId('Kategori'),jenis,nama,anggaran:ang});log(u.username,'Tambah kategori '+nama);}
  return{ok:true};
}
function settings(p){
  const u=auth(p,WR),d=p.data||{},y=Math.round(+d.tahunBuku),ys=String(y);
  if(!(y>=2000&&y<=2100))throw new Error('Tahun buku tidak valid');
  if(d.awalTunai!=null||d.awalBank!=null){ // saldo awal manual: hanya tahun pertama & belum ditutup
    const at=Math.round(+d.awalTunai||0),ab=Math.round(+d.awalBank||0);
    if(at<0||ab<0||at+ab>1e11)throw new Error('Saldo awal tidak valid');
    if(isClosed(y)||rows('Transaksi').some(t=>ds(t.tgl).slice(0,4)<ys))throw new Error('Saldo awal tahun ini otomatis dari saldo akhir tahun sebelumnya');
    const ex=rows('Transaksi').find(t=>ds(t.tgl).slice(0,4)===ys&&t.kategori===KAWAL&&t.status!=='Void'),b=bal(y,ex?+ex.id:-1);
    if(b.Tunai+at<0||b.Transfer+ab<0)throw new Error('Saldo awal terlalu kecil: saldo kas akan negatif');
    const row={id:ex?+ex.id:nextId('Transaksi'),tgl:ys+'-01-01',jenis:'Masuk',kategori:KAWAL,nominal:at+ab,pihak:'Saldo awal',
      ket:'Saldo Awal Kas '+ys,metode:'Tunai',tunai:at,bukti:'',status:'Aktif',dibuatOleh:u.username};
    if(ex){row._r=ex._r;upd('Transaksi',row);}else put('Transaksi',row);
    log(u.username,'Set saldo awal '+ys+': Tunai '+rp(at)+', Transfer '+rp(ab));
  }
  const r=rows('Pengaturan').find(x=>x.kunci==='tahunBuku');
  if(r){r.nilai=y;upd('Pengaturan',r);}else put('Pengaturan',{kunci:'tahunBuku',nilai:y});
  log(u.username,'Tahun buku aktif: '+y);return{ok:true};
}
function userOp(p){
  const u=auth(p,SA),U=rows('Users'),low=s=>String(s).toLowerCase();
  const lainSA=ex=>U.some(r=>r.role==='superadmin'&&bool(r.aktif)&&r.username!==ex.username);
  if(p.op==='delete'){
    const t=U.find(r=>r.username===String(p.username||'').trim());
    if(!t)throw new Error('User tidak ditemukan');
    if(t.username===u.username)throw new Error('Akun yang sedang Anda pakai tidak dapat dihapus');
    if(t.role==='superadmin'&&bool(t.aktif)&&!lainSA(t))throw new Error('Minimal harus ada satu Super Admin aktif');
    inv('Users');sh('Users').deleteRow(t._r);
    log(u.username,'Hapus user '+t.username+' ('+t.role+')');return{ok:true};
  }
  const d=p.data||{},un=String(d.username||'').trim(),old=String(d.oldUsername||un).trim();
  if(!/^[A-Za-z0-9_.-]{3,30}$/.test(un))throw new Error('Username 3-30 karakter (huruf, angka, . _ -)');
  if(ROLES.indexOf(d.role)<0)throw new Error('Role tidak valid');
  if(d.password&&String(d.password).length<4)throw new Error('Password minimal 4 karakter');
  const aktif=d.aktif!==false,ex=U.find(r=>low(r.username)===low(old));
  if(d.oldUsername&&!ex)throw new Error('User yang diedit tidak ditemukan');
  if(ex){
    if(U.some(r=>r.username!==ex.username&&low(r.username)===low(un)))throw new Error('Username sudah dipakai');
    if(ex.role==='superadmin'&&(d.role!=='superadmin'||!aktif)&&!lainSA(ex))throw new Error('Minimal harus ada satu Super Admin aktif');
    const info=[];
    if(ex.username!==un)info.push('username: '+ex.username+' → '+un);
    if(ex.role!==d.role)info.push('role: '+ex.role+' → '+d.role);
    if(bool(ex.aktif)!==aktif)info.push(aktif?'diaktifkan':'dinonaktifkan');
    if(d.password)info.push('password diganti');
    Object.assign(ex,{username:un,nama:String(d.nama||ex.nama).slice(0,60),role:d.role,aktif});
    if(d.password)ex.password=hash(d.password);
    upd('Users',ex);
    log(u.username,'Edit user '+un+(info.length?' ['+info.join('; ')+']':''));
  }else{
    if(!d.password)throw new Error('Password awal wajib diisi (minimal 4 karakter)');
    put('Users',{username:un,nama:String(d.nama||un).slice(0,60),role:d.role,aktif,password:hash(d.password)});
    log(u.username,'Tambah user '+un+' ('+d.role+')');
  }
  return{ok:true};
}

/** Hapus semua transaksi, status tahun, dan log (untuk uji ulang). Hanya Super Admin. */
function resetAll(p){
  const u=auth(p,SA);
  inv();['Transaksi','Tahun','Log','Aset','Pengajuan'].forEach(n=>{const s=sh(n),l=s.getLastRow();if(l>1)s.deleteRows(2,l-1);});
  const r=rows('Pengaturan').find(x=>x.kunci==='tahunBuku');
  if(r){r.nilai=CFG.DEFAULT_YEAR;upd('Pengaturan',r);}
  seedAset();log(u.username,'RESET semua data transaksi (aset dikembalikan ke contoh awal)');
  return{ok:true};
}

/* ===== ASET & INVENTARIS (terintegrasi dengan kas) ===== */
const as=o=>({id:+o.id,nama:o.nama,qty:+o.qty,satuan:o.satuan,harga:+o.harga,kategori:o.kategori,ket:o.ket});
function asetOp(p){
  const u=auth(p,WR); // anggota hanya boleh membaca lewat dashboard
  if(p.op==='delete'){const a=findAset(p.id);inv('Aset');sh('Aset').deleteRow(a._r);log(u.username,'Hapus aset #'+a.id+' '+a.nama);return{ok:true};}
  if(p.op==='sell')return sellAset(u,p);
  if(p.op==='addUnit')return addUnit(u,p);
  return saveAset(u,p.data||{});
}
function findAset(id){const a=rows('Aset').find(r=>+r.id===+id);if(!a)throw new Error('Aset tidak ditemukan');return a;}
function cleanAset(d){
  const n={nama:String(d.nama||'').trim().slice(0,80),kategori:String(d.kategori||'').trim().slice(0,40),ket:String(d.ket||'').trim().slice(0,200),
    satuan:String(d.satuan||'unit').trim().slice(0,15),qty:Math.round(+d.qty),harga:Math.round(+d.harga)};
  if(!n.nama)throw new Error('Nama aset wajib diisi');
  if(!(n.qty>=0&&n.qty<=1e6)||!(n.harga>=0&&n.harga<=1e11))throw new Error('Jumlah atau harga tidak valid');
  return n;
}
/** Catatan otomatis untuk aset yang berasal dari hadiah turnamen (tidak menyentuh kas). */
function hadiahNote(x){
  const t=String(x.turnamen||'').trim().slice(0,60),j=String(x.juara||'').trim().slice(0,30),tg=String(x.tglHadiah||'').trim();
  if(!t)throw new Error('Nama turnamen / penyelenggara wajib diisi untuk aset hadiah');
  if(!j)throw new Error('Peringkat (misal: Juara 1) wajib diisi untuk aset hadiah');
  if(tg&&(!/^\d{4}-\d{2}-\d{2}$/.test(tg)||isNaN(new Date(tg+'T00:00:00'))))throw new Error('Tanggal hadiah tidak valid');
  return 'Hadiah '+j+' — '+t+(tg?' ('+tg+')':'');
}
function gabungKet(lama,note){
  let k=(String(lama||'').trim()?String(lama).trim()+' · ':'')+note;
  while(k.length>200&&k.indexOf(' · ')>=0)k=k.slice(k.indexOf(' · ')+3); // buang catatan terlama bila kepanjangan
  return k.slice(0,200);
}
function saveAset(u,d){
  const n=cleanAset(d);
  if(d.id){const a=findAset(d.id);Object.assign(a,n);upd('Aset',a);log(u.username,'Edit aset #'+a.id+' '+n.nama);return{ok:true};}
  if(n.qty<1)throw new Error('Jumlah harus lebih dari 0');
  const hd=d.sumber==='hadiah'?hadiahNote(d):'';
  if(hd)n.ket=gabungKet(n.ket,hd);
  if(d.sumber==='kas') // potong kas: validasi tahun terkunci + saldo guard dijalankan oleh saveTx
    saveTx(u,{jenis:'Keluar',kategori:KBELI,nominal:n.qty*n.harga,tgl:d.tgl,pihak:d.pihak||'Pengadaan aset',metode:d.metode,buktiFile:d.buktiFile,
      ket:'Beli aset: '+n.nama+' '+n.qty+' '+n.satuan+' @ '+rp(n.harga)},true);
  put('Aset',Object.assign({id:nextId('Aset'),dibuatOleh:u.username},n));
  log(u.username,'Tambah aset '+n.nama+' '+n.qty+' '+n.satuan+' ('+(d.sumber==='kas'?'beli kas':d.sumber==='donasi'?'donasi':hd?hd:'aset awal')+')');
  return{ok:true};
}
function addUnit(u,p){
  const a=findAset(p.id),q=Math.round(+p.qty),h=Math.round(+p.harga);
  if(!(q>=1&&q<=1e6)||!(h>=0&&h<=1e11))throw new Error('Jumlah atau harga tidak valid');
  if(p.sumber==='kas')
    saveTx(u,{jenis:'Keluar',kategori:KBELI,nominal:q*h,tgl:p.tgl,pihak:p.pihak||'Pengadaan aset',metode:p.metode,buktiFile:p.buktiFile,
      ket:'Tambah unit aset: '+a.nama+' '+q+' '+a.satuan+' @ '+rp(h)},true);
  const hd=p.sumber==='hadiah'?hadiahNote(p):'';
  if(hd)a.ket=gabungKet(a.ket,hd);
  const nq=+a.qty+q;a.harga=Math.round((+a.qty*+a.harga+q*h)/nq);a.qty=nq;upd('Aset',a);
  log(u.username,'Tambah unit aset '+a.nama+' +'+q+' ('+(p.sumber==='kas'?'beli kas':p.sumber==='donasi'?'donasi':hd?hd:'penyesuaian')+')');
  return{ok:true};
}
function sellAset(u,p){
  const a=findAset(p.id),q=Math.round(+p.qty),h=Math.round(+p.hargaJual);
  if(!(q>=1&&q<=+a.qty))throw new Error('Jumlah jual melebihi stok ('+a.qty+')');
  if(!(h>0&&h<=1e11))throw new Error('Harga jual tidak valid');
  saveTx(u,{jenis:'Masuk',kategori:KJUAL,nominal:q*h,tgl:p.tgl,pihak:String(p.pembeli||'Pembeli'),metode:p.metode,
    ket:'Jual aset: '+a.nama+' '+q+' '+a.satuan+' @ '+rp(h)},true);
  a.qty=+a.qty-q;upd('Aset',a);log(u.username,'Jual aset '+a.nama+' '+q+' '+a.satuan+' = '+rp(q*h));
  return{ok:true};
}

/* ===== PENGAJUAN DANA: Bendahara mengajukan -> Ketua/Wakil Ketua approve -> kas OTOMATIS terpotong =====
 * Jumlah berbeda / ditolak? Bendahara mengedit lalu mengajukan lagi. Kas hanya terpotong saat approve (saldo guard & tahun terkunci berlaku).
 * Status: Diajukan -> Disetujui (transaksi Keluar tercatat) | Ditolak (bisa diedit & diajukan ulang) | Batal. Approve: Ketua/Wakil Ketua & Super Admin. */
const pj=o=>({id:+o.id,tahun:String(o.tahun),kategori:o.kategori,keperluan:o.keperluan,jumlah:+o.jumlah||0,tgl:ds(o.tgl),metode:o.metode,pihak:o.pihak,
  bukti:o.bukti,status:o.status,alasan:o.alasan,diajukanOleh:o.diajukanOleh,diputusOleh:o.diputusOleh,trxId:o.trxId===''?null:+o.trxId,waktu:ds(o.waktu)});
function pengajuanOp(p){
  const dec=p.op==='approve'||p.op==='reject',u=auth(p,dec?APR:WR),alasan=String(p.alasan||'').trim().slice(0,200);
  let r=null;
  if(p.op!=='save'||p.id){r=rows('Pengajuan').find(x=>+x.id===+p.id);if(!r)throw new Error('Pengajuan tidak ditemukan');}
  const buka=y=>{if(isClosed(y))throw new Error('Tahun buku '+y+' sudah ditutup (terkunci)');};
  const simpan=(st,extra,msg)=>{Object.assign(r,{status:st},extra);upd('Pengajuan',r);log(u.username,msg);return{ok:true};};
  if(p.op==='save'){
    const d=p.data||{},k=String(d.kategori||''),j=Math.round(+d.jumlah),kep=String(d.keperluan||'').trim().slice(0,150),tgl=String(d.tgl||''),metode=d.metode==='Transfer'?'Transfer':'Tunai';
    if(k===KBELI||!rows('Kategori').some(c=>c.jenis==='Keluar'&&c.nama===k))throw new Error('Kategori pengeluaran tidak valid');
    if(!kep)throw new Error('Keperluan wajib diisi');
    if(!(j>0&&j<=1e11))throw new Error('Jumlah tidak valid');
    if(!/^\d{4}-\d{2}-\d{2}$/.test(tgl)||isNaN(new Date(tgl+'T00:00:00')))throw new Error('Format tanggal tidak valid');
    const y=tgl.slice(0,4);buka(y);
    if(r&&r.status!=='Diajukan'&&r.status!=='Ditolak')throw new Error('Hanya pengajuan berstatus Diajukan atau Ditolak yang dapat diubah');
    const sd=bal(y)[metode];if(sd<j)throw new Error('Saldo '+metode+' tidak cukup. Tersedia '+rp(sd));
    const v={kategori:k,keperluan:kep,jumlah:j,tgl,metode,pihak:String(d.pihak||'').trim().slice(0,100),tahun:y,status:'Diajukan',alasan:''};
    const bukti=d.buktiFile?saveFile(d.buktiFile,y):(r?r.bukti:'');
    if(r){const ulang=r.status==='Ditolak';Object.assign(r,v,{bukti});upd('Pengajuan',r);log(u.username,(ulang?'Ajukan ULANG':'Ubah')+' pengajuan #'+r.id+' '+k+' '+rp(j));}
    else{const id=nextId('Pengajuan');put('Pengajuan',Object.assign({id,bukti,diajukanOleh:u.username,diputusOleh:'',trxId:'',waktu:now()},v));log(u.username,'Ajukan dana #'+id+' '+k+' '+rp(j)+' — '+kep);}
    return{ok:true};
  }
  if(p.op==='cancel'){
    if(r.status!=='Diajukan'&&r.status!=='Ditolak')throw new Error('Pengajuan ini tidak dapat dibatalkan');
    if(alasan.length<3)throw new Error('Alasan pembatalan wajib diisi');
    return simpan('Batal',{alasan:'Batal: '+alasan},'Batalkan pengajuan #'+r.id+' alasan: '+alasan);
  }
  if(dec){
    if(r.status!=='Diajukan')throw new Error('Pengajuan sudah diputuskan atau tidak lagi menunggu');
    if(p.op==='reject'){
      if(alasan.length<3)throw new Error('Alasan penolakan wajib diisi');
      return simpan('Ditolak',{alasan:'Ditolak: '+alasan,diputusOleh:u.username},'TOLAK pengajuan #'+r.id+' alasan: '+alasan);
    }
    const tid=nextId('Transaksi');
    saveTx(u,{jenis:'Keluar',kategori:r.kategori,nominal:+r.jumlah,tgl:ds(r.tgl),pihak:String(r.pihak||'').trim()||r.keperluan,metode:r.metode,buktiUrl:r.bukti,
      ket:'Pengajuan #'+r.id+' ('+r.diajukanOleh+'): '+r.keperluan}); // kas terpotong di sini; gagal bila saldo kurang / tahun terkunci
    return simpan('Disetujui',{alasan:'',diputusOleh:u.username,trxId:tid},'SETUJUI pengajuan #'+r.id+' '+r.kategori+' '+rp(r.jumlah)+' → kas terpotong, transaksi #'+tid);
  }
  throw new Error('Operasi pengajuan tidak dikenal');
}
