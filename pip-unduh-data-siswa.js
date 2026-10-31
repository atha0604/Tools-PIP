// Unduh Data Siswa PIP v1.5 (versi bookmarklet: tanpa Tampermonkey). Dibuat dari pip-unduh-data-siswa.user.js
(function () {
  if (window.__pipUnduh) { var p = document.getElementById('pipUnPanel'); if (p) p.classList.remove('mini'); return; }
  window.__pipUnduh = 1;
  var muat = function (u, ada) { return ada() ? Promise.resolve() : new Promise(function (ok, gagal) {
    var s = document.createElement('script'); s.src = u; s.onload = ok; s.onerror = function () { gagal(new Error('Gagal memuat ' + u)); }; document.head.appendChild(s); }); };
  Promise.all([
    muat('https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js', function () { return window.XLSX; }),
    muat('https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js', function () { return window.JSZip; })
  ]).then(function () {
(function () {
  'use strict';
  const W = window, VERSI = '1.5';
  const CAP = { pdUrl: null, headers: {}, sekolah: [], jenjang: null, tebak: false };
  const KEC = {};            // kode sekolah -> kecamatan
  let HASIL = {};            // kode sekolah -> { s, rows }
  let GAGAL = [], JALAN = false, STOP = false;
  const log = m => { console.log('[PIP Unduh]', m); const d = document.getElementById('pipUnLog'); if (d) { const b = document.createElement('div'); b.textContent = m; d.prepend(b); while (d.childElementCount > 120) d.lastElementChild.remove(); } };
  const bersih = v => String(v ?? '').replace(/^'/, '').trim();
  const amanNama = t => String(t || '').replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim();

  // timer lewat Web Worker agar tidak diperlambat Chrome saat tab di belakang
  let sleep = ms => new Promise(r => setTimeout(r, ms));
  try {
    const wk = new Worker(URL.createObjectURL(new Blob(['onmessage=e=>setTimeout(()=>postMessage(e.data.id),e.data.ms)'], { type: 'text/javascript' })));
    const tunggu = {}; let id = 0;
    wk.onmessage = e => { tunggu[e.data]?.(); delete tunggu[e.data]; };
    sleep = ms => new Promise(r => { tunggu[++id] = r; wk.postMessage({ id, ms }); });
  } catch (e) { /* pakai setTimeout biasa */ }

  // ---------- sadap request situs: alamat data siswa, header login, daftar sekolah ----------
  function catat(url, headers, teks) {
    try {
      const u = new URL(url, location.href);
      if (/\/pd$/.test(u.pathname) && u.searchParams.get('npsn')) {
        Object.assign(CAP, { pdUrl: u.origin + u.pathname, headers: headers || {}, tebak: false });
        if (u.searchParams.get('jenjang')) CAP.jenjang = u.searchParams.get('jenjang');
      }
      const j = teks && teks[0] === '{' ? JSON.parse(teks) : null, r0 = j?.data?.rows?.[0];
      if (r0 && /^[A-Za-z0-9]{8}$/.test(String(r0.npsn || '').trim()) && ('init_full' in r0 || 'init_half' in r0)) {
        if (r0.sub_jenjang && !CAP.jenjang) CAP.jenjang = r0.sub_jenjang;
        if (/\/sekolah$/.test(u.pathname) && !CAP.pdUrl) Object.assign(CAP, { pdUrl: u.origin + u.pathname.replace(/\/sekolah$/, '/pd'), headers: headers || {}, tebak: true });
        CAP.sekolah = j.data.rows;
      }
      status();
    } catch (e) { /* abaikan */ }
  }
  const _fetch = W.fetch;
  W.fetch = async function (input, init = {}) {
    const url = typeof input === 'string' ? input : input.url, h = {};
    new Headers(init.headers || (input instanceof Request ? input.headers : {})).forEach((v, k) => h[k] = v);
    const res = await _fetch.apply(this, arguments);
    res.clone().text().then(t => catat(url, h, t)).catch(() => {});
    return res;
  };
  const XO = XMLHttpRequest.prototype.open, XS = XMLHttpRequest.prototype.send, XH = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.open = function (m, url) { this._unUrl = url; this._unH = {}; return XO.apply(this, arguments); };
  XMLHttpRequest.prototype.setRequestHeader = function (k, v) { if (this._unH) this._unH[k] = v; return XH.apply(this, arguments); };
  XMLHttpRequest.prototype.send = function () {
    this.addEventListener('load', () => { let t = null; try { t = this.responseType === '' || this.responseType === 'text' ? this.responseText : JSON.stringify(this.response); } catch (e) {} catat(this._unUrl, this._unH, t); });
    return XS.apply(this, arguments);
  };

  // ---------- kolom ----------
  const KOLOM = [
    ['npsn', 'NPSN'], ['sekolah', 'Sekolah'], ['kecamatan', 'Kecamatan'], ['siswa', 'Nama Siswa'], ['nik', 'NIK'], ['nisn', 'NISN'],
    ['jenis_kelamin', 'JK'], ['tempat_lahir', 'Tempat Lahir'], ['tanggal_lahir', 'Tgl Lahir'], ['usia', 'Umur'], ['kelas', 'Kelas'],
    ['dtsen_desil', 'Desil'], ['nominal_id', 'Nominal (1=Awal,2=Berjalan)'], ['valid_dukcapil', 'Valid Dukcapil'], ['valid_nisn', 'Valid NISN'],
    ['nama_ibu', 'Ibu'], ['pekerjaan_ibu', 'Pekerjaan Ibu'], ['penghasilan_ibu', 'Penghasilan Ibu'],
    ['nama_ayah', 'Ayah'], ['pekerjaan_ayah', 'Pekerjaan Ayah'], ['penghasilan_ayah', 'Penghasilan Ayah'],
    ['nama_wali', 'Wali'], ['pekerjaan_wali', 'Pekerjaan Wali'], ['penghasilan_wali', 'Penghasilan Wali'], ['no_telepon', 'Telepon'],
    ['bank', 'Bank'], ['status_tabunganku', 'Status Tabunganku'], ['sk', 'Status SK'], ['sk_atas_nama', 'SK Atas Nama'],
    ['flag_drop', 'Flag Drop'], ['keterangan_flag', 'Ket Flag'], ['fase_id', 'Fase ID'], ['fase_name', 'Fase Usulan'], ['role', 'Diusulkan Oleh'],
    ['parent_name', 'Pengusul Induk'], ['child_name', 'Pengusul Anak'], ['created_at', 'Diusulkan Pada'], ['peserta_didik_id', 'PDID'],
    ['bentuk', 'Bentuk'], ['jenjang', 'Jenjang'], ['rombel', 'Rombel'], ['semester_id', 'Semester ID'], ['sekolah_id', 'SID'],
    ['pro', 'Kode Provinsi'], ['provinsi', 'Provinsi'], ['kab', 'Kode Kabupaten'], ['kabupaten', 'Kabupaten'], ['kec', 'Kode Kecamatan'],
    ['parent_id', 'UPID'], ['parent_username', 'Username Pengusul Induk'], ['child_id', 'UCID'], ['child_username', 'Username Pengusul Anak']
  ];
  const CADANGAN = { parent_username: 'userId_up', child_username: 'userId_uc' };
  const rapikan = r => {
    const o = {};
    KOLOM.forEach(([k, label]) => {
      let v = r[k] ?? r[CADANGAN[k]];
      if (typeof v === 'string') v = v.trim();
      if (k === 'tanggal_lahir' && v) v = v.slice(0, 10);
      o[label] = v ?? '';
    });
    return o;
  };
  // kolom file per sekolah: sama dengan tombol unduh di pop-up sekolah pada situs PIP
  const KOLOM_SITUS = [['peserta_didik_id', 'PDID'], ['siswa', 'Nama Murid'], ['nik', 'NIK'], ['valid_dukcapil', 'Valid Dukcapil'],
    ['nisn', 'NISN'], ['valid_nisn', 'Valid NISN'], ['tempat_lahir', 'Tempat Lahir'], ['tanggal_lahir', 'Tanggal Lahir'], ['usia', 'Umur'],
    ['jenis_kelamin', 'Jenis Kelamin'], ['nama_ibu', 'Nama Ibu'], ['pekerjaan_ibu', 'Pekerjaan Ibu'], ['penghasilan_ibu', 'Penghasilan Ibu'],
    ['nama_ayah', 'Nama Ayah'], ['pekerjaan_ayah', 'Pekerjaan Ayah'], ['penghasilan_ayah', 'Penghasilan Ayah'], ['nama_wali', 'Nama Wali'],
    ['pekerjaan_wali', 'Pekerjaan Wali'], ['penghasilan_wali', 'Penghasilan Wali'], ['no_telepon', 'Telefon'], ['bentuk', 'Bentuk'],
    ['jenjang', 'Jenjang'], ['kelas', 'Kelas'], ['rombel', 'Rombel'], ['semester_id', 'Semester ID'], ['sekolah_id', 'SID'], ['npsn', 'NPSN'],
    ['sekolah', 'Nama Sekolah'], ['pro', 'pro'], ['provinsi', 'provinsi'], ['kab', 'kab'], ['kabupaten', 'kabupaten'], ['kec', 'kec'],
    ['kecamatan', 'kecamatan'], ['nominal_id', 'Nominal ID'], ['dtsen_desil', 'Desil'], ['bank', 'Bank'], ['status_tabunganku', 'Status Tabunganku'],
    ['flag_drop', 'flag_drop'], ['keterangan_flag', 'keterangan_flag'], ['parent_id', 'UPID'], ['parent_name', 'Nama Pengusul Induk'],
    ['userId_up', 'Username Pengusul Induk'], ['child_id', 'UCID'], ['child_name', 'Nama Pengusul Anak'], ['userId_uc', 'Username Pengusul Anak'],
    ['fase_id', 'Fase ID'], ['fase_name', 'Nama Fase'], ['role', 'Role'], ['created_at', 'Diusulkan Pada'], ['sk', 'Status SK'], ['sk_atas_nama', 'SK Atas Nama']];

  // ---------- panel ----------
  const CSS = `
  #pipUnPanel{position:fixed;bottom:20px;right:340px;z-index:2147483646;width:290px;background:#fff;color:#1e293b;border-radius:12px;
    box-shadow:0 10px 30px rgba(15,23,42,.25);font:13px/1.4 'Segoe UI',Roboto,sans-serif;border:1px solid #e2e8f0;overflow:hidden}
  #pipUnPanel .ph{background:#0f766e;color:#fff;padding:9px 12px;display:flex;gap:8px;align-items:center;cursor:move}
  #pipUnPanel .ph b{flex:1}#pipUnPanel .min{background:none;border:0;color:#fff;font-size:18px;cursor:pointer}
  #pipUnPanel .pb{padding:10px 12px}#pipUnPanel.mini .pb{display:none}
  #pipUnSt{font-size:12px;color:#475569;margin-bottom:8px}
  .pipUnBtn{width:100%;border:0;border-radius:8px;padding:8px;margin-bottom:6px;font-weight:600;cursor:pointer;background:#ccfbf1;color:#115e59}
  .pipUnBtn.u{background:#0f766e;color:#fff}.pipUnBtn.s{background:#fee2e2;color:#b91c1c}.pipUnBtn:disabled{opacity:.5;cursor:not-allowed}
  #pipUnBar{height:6px;background:#e2e8f0;border-radius:4px;overflow:hidden;margin:4px 0 2px}#pipUnBar i{display:block;height:100%;width:0;background:#14b8a6}
  #pipUnPt{font-size:11px;color:#64748b;text-align:right}
  #pipUnLog{max-height:110px;overflow:auto;font:11px/1.5 Consolas,monospace;background:#0f172a;color:#cbd5e1;border-radius:6px;padding:6px 8px;margin-top:6px}`;
  function pasang() {
    if (!document.body || document.getElementById('pipUnPanel')) return;
    const st = document.createElement('style'); st.textContent = CSS; document.head.appendChild(st);
    const p = document.createElement('div'); p.id = 'pipUnPanel';
    p.innerHTML = `<div class="ph"><b>⬇ Unduh Data Siswa</b><span>v${VERSI}</span><button class="min">–</button></div>
      <div class="pb"><div id="pipUnSt">Menunggu daftar sekolah…</div>
      <button class="pipUnBtn u" id="pipUnExcel">Excel gabungan semua sekolah</button>
      <button class="pipUnBtn u" id="pipUnZip">ZIP file Excel per sekolah</button>
      <button class="pipUnBtn" id="pipUnSimpan" disabled>Simpan data yang sudah terkumpul</button>
      <button class="pipUnBtn" id="pipUnUlang" disabled>Ulangi sekolah yang gagal</button>
      <button class="pipUnBtn s" id="pipUnStop" disabled>⏹ Stop</button>
      <div id="pipUnBar"><i></i></div><div id="pipUnPt">0 / 0</div><div id="pipUnLog"></div></div>`;
    document.body.appendChild(p);
    p.querySelector('.min').onclick = () => p.classList.toggle('mini');
    const h = p.querySelector('.ph');
    h.onmousedown = e => { if (e.target.tagName === 'BUTTON') return; const r = p.getBoundingClientRect(), dx = e.clientX - r.left, dy = e.clientY - r.top;
      const mv = ev => { p.style.left = ev.clientX - dx + 'px'; p.style.top = ev.clientY - dy + 'px'; p.style.right = p.style.bottom = 'auto'; };
      document.addEventListener('mousemove', mv); document.addEventListener('mouseup', () => document.removeEventListener('mousemove', mv), { once: true }); };
    document.getElementById('pipUnExcel').onclick = () => mulai('excel');
    document.getElementById('pipUnZip').onclick = () => mulai('zip');
    document.getElementById('pipUnSimpan').onclick = () => simpanHasil(ModeTerakhir);
    document.getElementById('pipUnUlang').onclick = ulangi;
    document.getElementById('pipUnStop').onclick = () => { STOP = true; log('⏹ Stop diminta, menunggu langkah yang berjalan selesai…'); };
    log(`Script v${VERSI} aktif.`);
    status();
  }
  setInterval(pasang, 1000);
  function status() {
    const s = document.getElementById('pipUnSt'); if (!s) return;
    const n = Object.values(HASIL).reduce((a, x) => a + x.rows.length, 0);
    s.innerHTML = `Sekolah: <b>${CAP.sekolah.length || 'belum tertangkap'}</b> · Jenjang: <b>${CAP.jenjang || '-'}</b>` + (n ? `<br>Terkumpul: <b>${n.toLocaleString('id')}</b> siswa dari <b>${Object.keys(HASIL).length}</b> sekolah` : '');
    document.getElementById('pipUnSimpan').disabled = !n || JALAN;
    document.getElementById('pipUnUlang').disabled = !GAGAL.length || JALAN;
    document.getElementById('pipUnStop').disabled = !JALAN;
    document.getElementById('pipUnExcel').disabled = document.getElementById('pipUnZip').disabled = JALAN;
  }
  function progres(i, n) {
    const b = document.querySelector('#pipUnBar i'), t = document.getElementById('pipUnPt');
    if (b) b.style.width = (n ? i / n * 100 : 0) + '%'; if (t) t.textContent = `${i} / ${n}`;
  }

  // ---------- daftar sekolah & akses ----------
  const terlihat = el => !!el && el.isConnected && getComputedStyle(el).display !== 'none' && el.getClientRects().length > 0;
  const tungguSampai = async (fn, ms) => { const t = Date.now(); while (Date.now() - t < ms) { const v = fn(); if (v) return v; await sleep(150); } return null; };
  const kodeSek = (np, jj, ganda) => ganda ? `${np}·${jj}` : np;
  async function daftarSekolah() {
    const jj = r => r.sub_jenjang || r.jenjang || CAP.jenjang || 'SD';
    let rows = CAP.sekolah;
    if (!rows.length) {
      // cadangan: baca dari layar (tampilkan semua baris dulu)
      const sel = [...document.querySelectorAll('select')].filter(x => !x.closest('.modal') && [...x.options].some(o => o.text.trim() === 'All')).pop();
      if (sel) { Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(sel, [...sel.options].find(o => o.text.trim() === 'All').value);
        sel.dispatchEvent(new Event('change', { bubbles: true })); await sleep(1500); }
      rows = [...document.querySelectorAll('tbody tr td')].filter(td => !td.closest('.modal')).map(td => {
        const m = (td.querySelector('.ps-4 .col-12 > span')?.textContent || '').trim().match(/(?:^|\s-\s)([A-Za-z0-9]{8})\s*-\s*(.*?)\s*-\s*\(\d+\)$/);
        const kotak = (td.querySelector('.card .fs-5')?.textContent || '').trim();
        return m && { npsn: m[1], sekolah: m[2], sub_jenjang: /^[A-Z]{2,8}$/.test(kotak) ? kotak : (CAP.jenjang || 'SD') };
      }).filter(Boolean);
      if (rows.length) CAP.sekolah = rows;
    }
    const hit = {}; rows.forEach(r => { const n = bersih(r.npsn); hit[n] = (hit[n] || 0) + 1; });
    return rows.map(r => ({ npsn: bersih(r.npsn), nama: r.sekolah, jenjang: jj(r), kec: r.kecamatan || r.nama_kecamatan || '', kode: kodeSek(bersih(r.npsn), jj(r), hit[bersih(r.npsn)] > 1) }));
  }
  async function aktifkanAkses() {
    log('Mengaktifkan akses: membuka 1 sekolah…');
    const cari = () => [...document.querySelectorAll('tbody tr td')].find(t => !t.closest('.modal') && /(^|\s-\s)[A-Za-z0-9]{8}\s-\s/.test(t.querySelector('.ps-4 .col-12 > span')?.textContent || ''));
    let td = cari();
    if (!td) { const k = [...document.querySelectorAll('button')].find(b => !b.closest('.modal') && /^Keseluruhan\s*:/.test(b.textContent.trim())); if (k) k.click(); td = await tungguSampai(cari, 8000); }
    if (!td) throw new Error('Daftar sekolah tidak tampil. Buka halaman Pengusulan Dinas dulu.');
    CAP.tebak = true; td.click();
    const setuju = await tungguSampai(() => [...document.querySelectorAll('button')].find(b => terlihat(b) && b.textContent.trim() === 'Setuju'), 8000);
    if (setuju) setuju.click();
    const ok = await tungguSampai(() => CAP.pdUrl && !CAP.tebak, 15000);
    for (let k = 0; k < 4; k++) { const m = [...document.querySelectorAll('.modal')].find(terlihat); if (!m) break;
      const x = m.querySelector('.btn-close, [aria-label="Close"], .close'); if (x) x.click(); else document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true })); await sleep(600); }
    if (!ok) throw new Error('Akses tidak aktif. Klik 1 sekolah → Setuju secara manual, lalu ulangi.');
    log('Akses data siswa aktif ✔');
  }

  // ---------- ambil data ----------
  let ModeTerakhir = 'excel';
  async function ambil(daftar) {
    let i = 0, selesai = 0, gagalBeruntun = 0, jeda = false;
    const mulai = Date.now();
    async function pekerja() {
      while (!STOP && !jeda && i < daftar.length) {
        const s = daftar[i++];
        try {
          const res = await _fetch(`${CAP.pdUrl}?jenjang=${encodeURIComponent(s.jenjang)}&npsn=${s.npsn}`, { headers: CAP.headers, credentials: 'include' });
          if (!res.ok) throw new Error('HTTP ' + res.status + (res.status === 401 || res.status === 403 ? ' (sesi login habis?)' : ''));
          const j = await res.json();
          if (j.status && j.status !== 'OK') throw new Error(j.message || j.status);
          const mentah = j?.data?.rows || [];
          if (mentah[0]?.kecamatan) KEC[s.kode] = mentah[0].kecamatan;
          HASIL[s.kode] = { s, rows: mentah };
          gagalBeruntun = 0;
          log(`${++selesai}/${daftar.length} ${s.nama}: ${mentah.length} siswa`);
        } catch (e) {
          selesai++; GAGAL.push(s); log(`✖ ${s.npsn} ${s.nama}: ${e.message}`);
          if (++gagalBeruntun >= 5) { jeda = true; log('⚠ 5x gagal berturut-turut, proses dijeda. Cek login / internet, lalu tekan "Ulangi sekolah yang gagal".'); }
          await sleep(2000);
        }
        progres(selesai, daftar.length); if (selesai % 5 === 0) status();
      }
    }
    await Promise.all(Array.from({ length: Math.min(4, daftar.length) }, pekerja));
    if (i < daftar.length) { GAGAL.push(...daftar.slice(i)); log(`${STOP ? 'Dihentikan' : 'Dijeda'}. ${daftar.length - i} sekolah belum diambil (masuk daftar ulangi).`); }
    log(`Waktu ambil data: ${Math.round((Date.now() - mulai) / 1000)} dtk untuk ${selesai} sekolah.`);
  }
  async function mulai(mode) {
    if (JALAN) return;
    if (Object.keys(HASIL).length && !confirm('Data yang sudah terkumpul akan diganti. Ambil ulang dari awal?')) return simpanHasil(mode);
    try {
      // akses diaktifkan dulu: dari permintaan data siswa itu jenjang akun (SD/SMP/TK/Paket) diketahui, juga saat dijalankan lewat bookmarklet
      if (!CAP.pdUrl || CAP.tebak) await aktifkanAkses();
      const daftar = await daftarSekolah();
      if (!daftar.length) return log('⚠ Daftar sekolah belum ada. Buka halaman Pengusulan Dinas lalu tunggu daftar sekolah tampil.');
      HASIL = {}; GAGAL = []; STOP = false; JALAN = true; ModeTerakhir = mode; status();
      log(`Mulai ${daftar.length} sekolah (4 sekaligus)…`);
      await ambil(daftar);
    } catch (e) { log('⚠ ' + e.message); }
    JALAN = false; status();
    if (!GAGAL.length && Object.keys(HASIL).length) await simpanHasil(mode);
    else if (GAGAL.length) log(`Selesai sebagian: ${GAGAL.length} sekolah belum terambil. Tekan "Ulangi" atau "Simpan data yang sudah terkumpul".`);
  }
  async function ulangi() {
    if (JALAN || !GAGAL.length) return;
    try { if (!CAP.pdUrl || CAP.tebak) await aktifkanAkses(); } catch (e) { return log('⚠ ' + e.message); }
    const daftar = GAGAL.splice(0); STOP = false; JALAN = true; status();
    await ambil(daftar); JALAN = false; status();
    if (!GAGAL.length) { log('Semua sekolah berhasil. Menyimpan…'); await simpanHasil(ModeTerakhir); }
  }

  // ---------- rekap: jumlah siswa per kecamatan dan per sekolah ----------
  const KOL_REKAP = ['Jumlah Sekolah', 'Total Siswa', 'Valid', 'Tidak Valid', 'Valid Desil 1-4', 'Sudah SK', 'Sudah Diusulkan', 'Usulan Sekolah', 'Usulan Dinas', 'Usulan Pemangku (TA)', 'Kelas Berjalan', 'Kelas Awal'];
  // jenis pengusul dari kolom "Fase Usulan" (mis. "2 Sekolah 2026", "2 Dinas 2026", "2 TA 2026"); selain Sekolah/Dinas dihitung Pemangku (TA)
  const jenisUsul = r => {
    const t = (bersih(r.fase_name) || bersih(r.role)).toLowerCase();
    if (!t && !bersih(r.fase_id)) return '';
    return /sekolah/.test(t) ? 'Sekolah' : /dinas/.test(t) ? 'Dinas' : 'Pemangku';
  };
  function hitungSekolah(rows) {
    const o = { 'Total Siswa': rows.length, 'Valid': 0, 'Tidak Valid': 0, 'Valid Desil 1-4': 0, 'Sudah SK': 0, 'Sudah Diusulkan': 0, 'Usulan Sekolah': 0, 'Usulan Dinas': 0, 'Usulan Pemangku (TA)': 0, 'Kelas Berjalan': 0, 'Kelas Awal': 0 };
    rows.forEach(r => {
      const valid = String(r.valid_dukcapil) === '1' && String(r.valid_nisn) === '1';
      valid ? o.Valid++ : o['Tidak Valid']++;
      if (valid && ['1', '2', '3', '4'].includes(bersih(r.dtsen_desil))) o['Valid Desil 1-4']++;
      if (String(r.sk) === '1') o['Sudah SK']++;
      if (bersih(r.fase_name) || bersih(r.fase_id)) {
        o['Sudah Diusulkan']++;
        const j = jenisUsul(r);
        if (j === 'Sekolah') o['Usulan Sekolah']++; else if (j === 'Dinas') o['Usulan Dinas']++; else if (j) o['Usulan Pemangku (TA)']++;
      }
      if (Number(r.nominal_id) === 1) o['Kelas Awal']++; else if (Number(r.nominal_id) === 2) o['Kelas Berjalan']++;
    });
    return o;
  }
  function buatRekap(daftar) {
    const perSek = daftar.map(({ s, rows }) => ({ Kecamatan: rows[0]?.kecamatan || KEC[s.kode] || s.kec || 'Tanpa Kecamatan', NPSN: s.npsn, Sekolah: s.nama, Jenjang: s.jenjang, ...hitungSekolah(rows) }))
      .sort((a, b) => a.Kecamatan.localeCompare(b.Kecamatan) || String(a.Sekolah).localeCompare(String(b.Sekolah)));
    const jumlahkan = L => KOL_REKAP.reduce((o, k) => { o[k] = k === 'Jumlah Sekolah' ? L.length : L.reduce((a, x) => a + (x[k] || 0), 0); return o; }, {});
    const kec = {}; perSek.forEach(x => (kec[x.Kecamatan] = kec[x.Kecamatan] || []).push(x));
    const perKec = Object.keys(kec).sort().map(k => ({ Kecamatan: k, ...jumlahkan(kec[k]) }));
    perKec.push({ Kecamatan: 'TOTAL', ...jumlahkan(perSek) });
    perSek.push({ Kecamatan: 'TOTAL', NPSN: '', Sekolah: `${daftar.length} sekolah`, Jenjang: '', ...hitungSekolah(daftar.flatMap(x => x.rows)) });
    // total usulan per kecamatan menurut kolom "Fase Usulan": Sekolah, Dinas, Pemangku (TA)
    const totalUsulan = perKec.map(k => ({ Kecamatan: k.Kecamatan, 'Jumlah Sekolah': k['Jumlah Sekolah'], 'Usulan Sekolah': k['Usulan Sekolah'], 'Usulan Dinas': k['Usulan Dinas'],
      'Usulan Pemangku (TA)': k['Usulan Pemangku (TA)'], 'Total Usulan': k['Usulan Sekolah'] + k['Usulan Dinas'] + k['Usulan Pemangku (TA)'] }));
    return { perKec, perSek, totalUsulan };
  }
  const sheetRekap = (wb, nama, data, lebar) => { const ws = XLSX.utils.json_to_sheet(data); ws['!cols'] = lebar.map(w => ({ wch: w })); XLSX.utils.book_append_sheet(wb, ws, nama); };
  // tiga sheet rekap: Rekap per Kecamatan, Total Usulan, Rekap per Sekolah
  function tambahRekap(wb, daftar) {
    const { perKec, perSek, totalUsulan } = buatRekap(daftar);
    sheetRekap(wb, 'Rekap per Kecamatan', perKec, [24, 9, 9, 9, 9, 12, 9, 12, 12, 12, 20, 12, 12]);
    sheetRekap(wb, 'Total Usulan', totalUsulan, [24, 9, 14, 14, 20, 14]);
    sheetRekap(wb, 'Rekap per Sekolah', perSek, [22, 10, 36, 10, 9, 9, 9, 12, 12, 12, 12, 12, 20, 12, 12]);
  }
  const TULIS = { compression: true }; // tanpa ini file Excel tidak dikompres (bisa 8x lebih besar)

  // ---------- simpan ----------
  async function simpanHasil(mode) {
    if (typeof XLSX === 'undefined') return log('⚠ Library Excel gagal dimuat. Refresh halaman lalu coba lagi.');
    const daftar = Object.values(HASIL);
    if (!daftar.length) return log('Belum ada data terkumpul.');
    const kab = amanNama((daftar.find(x => x.rows[0]?.kabupaten)?.rows[0].kabupaten || '').replace(/^Kab\.\s*|^Kota\s*/, '')) || 'PIP';
    const tgl = new Date().toISOString().slice(0, 10);
    if (mode === 'zip') {
      if (typeof JSZip === 'undefined') return log('⚠ Library ZIP gagal dimuat. Refresh halaman lalu coba lagi.');
      const zip = new JSZip();
      { const wbR = XLSX.utils.book_new(); tambahRekap(wbR, daftar);
        zip.file('Rekap per Kecamatan.xlsx', XLSX.write(wbR, { type: 'array', bookType: 'xlsx', ...TULIS })); }
      daftar.forEach(({ s, rows }) => {
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([KOLOM_SITUS.map(k => k[1]), ...rows.map(r => KOLOM_SITUS.map(([k]) => r[k] ?? ''))]), 'Data');
        const kec = amanNama(rows[0]?.kecamatan || KEC[s.kode] || s.kec || 'Tanpa Kecamatan');
        const sub = s.kode.includes('·') ? ` (${s.jenjang})` : '';
        zip.folder(kec).file(`${s.npsn} - ${amanNama(s.nama)}${sub}.xlsx`, XLSX.write(wb, { type: 'array', bookType: 'xlsx', ...TULIS }));
      });
      log('Menyusun ZIP…');
      const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `Data_Siswa_per_Sekolah_${kab}_${tgl}.zip`; a.click();
      log(`✔ ZIP tersimpan: ${daftar.length} file sekolah.`);
    } else {
      const semua = daftar.flatMap(({ s, rows }) => rows.map(r => ({ ...rapikan(r), 'Sub Jenjang': s.jenjang })));
      const ws = XLSX.utils.json_to_sheet(semua);
      ws['!cols'] = [...KOLOM.map(([, l]) => ({ wch: Math.max(10, Math.min(35, l.length + 4)) })), { wch: 12 }];
      const wb = XLSX.utils.book_new();
      tambahRekap(wb, daftar);
      XLSX.utils.book_append_sheet(wb, ws, 'Siswa');
      XLSX.writeFile(wb, `Data_Siswa_PIP_${kab}_${tgl}.xlsx`, TULIS);
      log(`✔ Excel tersimpan: ${semua.length.toLocaleString('id')} siswa dari ${daftar.length} sekolah.`);
    }
  }
})();

  }).catch(function (e) { window.__pipUnduh = 0; alert('Unduh Data PIP gagal dimulai: ' + e.message + '\nPeriksa koneksi internet lalu klik bookmark lagi.'); });
})();
