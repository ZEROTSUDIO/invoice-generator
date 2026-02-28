// ─── CONSTANTS ───────────────────────────────────────────
const LOGO_SRC = 'rapa-logo.png';
const CO_NAME = 'RAPA CEMENT &amp; GRC';
const CO_ADDR = 'Jl. Ngadiretno no. 33, Tamanagung, Muntilan, Magelang 56413';
const CO_TELP = 'Telp: 08112959125 / 082134567874';
const CO_EMAIL = 'rapastone33@gmail.com';

// ─── STATE ───────────────────────────────────────────────
let activeTab = 'nota';
let notaItems = [newItem()];
let sjItems = [newSJItem()];

function newItem() { return { type: 'pcs', nama: '', pcs: '', m2: '', harga: '' }; }
function newSJItem() { return { kode: '', nama: '', qty: '', satuan: '' }; }

// ─── SUPABASE PARAMS ───────────────────────────────────────
const supabaseUrl = 'https://yhhxbmbjzrgtfxjdrizu.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InloaHhibWJqenJndGZ4amRyaXp1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzIxNzQyNjgsImV4cCI6MjA4Nzc1MDI2OH0.2bX25UW6r_BcN_yHUPN7ap5wRHuFhZFawTAuZKFvLmo';
const supabaseClient = supabase.createClient(supabaseUrl, supabaseKey);

// ─── NOMOR AUTO (SUPABASE) ─────────────────────────────────
function getYYYYMM() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`;
}

async function initNomor(key, prefix, inputId) {
  const ym = getYYYYMM();
  const inputEl = document.getElementById(inputId);
  inputEl.value = 'Loading...';

  const { data, error } = await supabaseClient
    .from('document_sequences')
    .select('*')
    .eq('id', key)
    .single();

  if (error && error.code !== 'PGRST116') {
    console.error('Error fetching sequence:', error);
    inputEl.value = 'ERROR';
    return;
  }

  let nomor;
  if (data) {
    if (data.ym === ym && data.current_val) {
      nomor = data.current_val;
    } else {
      let nextSeq = data.ym === ym ? (data.seq || 0) : 0;
      nextSeq++;
      nomor = `${prefix}-${ym}-${String(nextSeq).padStart(3, '0')}`;
      await supabaseClient
        .from('document_sequences')
        .update({ ym: ym, seq: nextSeq, current_val: nomor })
        .eq('id', key);
    }
  } else {
    // Fallback if row does not exist, though it should be created by the user manually
    nomor = `${prefix}-${ym}-001`;
  }

  inputEl.value = nomor;
  updatePreview();
}

async function saveCurrentNomor(key, val) {
  await supabaseClient
    .from('document_sequences')
    .update({ current_val: val })
    .eq('id', key);
}

let debounceTimer = {};
function debounceSaveNomor(key, val) {
  clearTimeout(debounceTimer[key]);
  debounceTimer[key] = setTimeout(() => {
    saveCurrentNomor(key, val);
  }, 500);
}

async function generateNewNomor(key, prefix, inputId) {
  const inputEl = document.getElementById(inputId);
  inputEl.value = 'Generating...';

  const ym = getYYYYMM();
  const { data } = await supabaseClient
    .from('document_sequences')
    .select('*')
    .eq('id', key)
    .single();

  if (data) {
    let nextSeq = data.ym === ym ? (data.seq || 0) : 0;
    nextSeq++;
    const nomor = `${prefix}-${ym}-${String(nextSeq).padStart(3, '0')}`;
    await supabaseClient
      .from('document_sequences')
      .update({ ym: ym, seq: nextSeq, current_val: nomor })
      .eq('id', key);

    inputEl.value = nomor;
    updatePreview();
  }
}

// ─── NUMBER FORMAT ────────────────────────────────────────
function fmt(n) {
  const val = parseFloat(n) || 0;
  return val === 0 ? '' : val.toLocaleString('id-ID');
}

function parseFmt(s) { return parseFloat(String(s).replace(/\./g, '').replace(',', '.')) || 0; }

// ─── TAB SWITCH ──────────────────────────────────────────
function switchTab(tab) {
  activeTab = tab;
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  document.getElementById('form-nota').style.display = tab === 'nota' ? '' : 'none';
  document.getElementById('form-sj').style.display = tab === 'sj' ? '' : 'none';
  updatePreview();
}

// ─── ITEM ROW MANAGEMENT ──────────────────────────────────
function addNotaRow() {
  notaItems.push(newItem());
  renderNotaItemsForm();
  updatePreview();
}

function removeNotaRow(i) {
  if (notaItems.length > 1) { notaItems.splice(i, 1); renderNotaItemsForm(); updatePreview(); }
}

function addSJRow() {
  sjItems.push(newSJItem());
  renderSJItemsForm();
  updatePreview();
}

function removeSJRow(i) {
  if (sjItems.length > 1) { sjItems.splice(i, 1); renderSJItemsForm(); updatePreview(); }
}

function updateNotaFormJumlah(i) {
  const it = notaItems[i];
  if (!it) return;
  const jumlah = it.type === 'tegel'
    ? (parseFloat(it.m2) || 0) * (parseFloat(it.harga) || 0)
    : (parseFloat(it.pcs) || 0) * (parseFloat(it.harga) || 0);
  const el = document.getElementById(`nota-jumlah-${i}`);
  if (el) el.innerText = fmt(jumlah);
}

function renderNotaItemsForm() {
  const tb = document.getElementById('nota-items-tbody');
  tb.innerHTML = '';
  notaItems.forEach((it, i) => {
    const tr = document.createElement('tr');
    const jumlah = it.type === 'tegel'
      ? (parseFloat(it.m2) || 0) * (parseFloat(it.harga) || 0)
      : (parseFloat(it.pcs) || 0) * (parseFloat(it.harga) || 0);

    tr.innerHTML = `
      <td style="text-align:center;color:#888;font-weight:600">${i + 1}</td>
      <td>
        <select onchange="notaItems[${i}].type=this.value;updatePreview();updateNotaFormJumlah(${i})" style="width:75px;font-size:11px">
          <option value="pcs" ${it.type === 'pcs' ? 'selected' : ''}>Lainnya</option>
          <option value="tegel" ${it.type === 'tegel' ? 'selected' : ''}>Tegel</option>
        </select>
      </td>
      <td><input type="text" value="${it.nama}" placeholder="Nama barang" oninput="notaItems[${i}].nama=this.value;updatePreview()"></td>
      <td><input type="number" value="${it.pcs}" placeholder="0" style="width:52px" oninput="notaItems[${i}].pcs=this.value;updatePreview();updateNotaFormJumlah(${i})"></td>
      <td><input type="number" value="${it.m2}" placeholder="0" style="width:52px" oninput="notaItems[${i}].m2=this.value;updatePreview();updateNotaFormJumlah(${i})"></td>
      <td><input type="number" value="${it.harga}" placeholder="0" style="width:90px" oninput="notaItems[${i}].harga=this.value;updatePreview();updateNotaFormJumlah(${i})"></td>
      <td id="nota-jumlah-${i}" style="text-align:right;font-weight:600;font-size:11px;white-space:nowrap">${fmt(jumlah)}</td>
      <td><button class="btn-del-row" onclick="removeNotaRow(${i})">×</button></td>`;
    tb.appendChild(tr);
  });
}

function renderSJItemsForm() {
  const tb = document.getElementById('sj-items-tbody');
  tb.innerHTML = '';
  sjItems.forEach((it, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align:center;color:#888;font-weight:600">${i + 1}</td>
      <td><input type="text" value="${it.kode}" placeholder="Kode" style="width:60px" oninput="sjItems[${i}].kode=this.value;updatePreview()"></td>
      <td><input type="text" value="${it.nama}" placeholder="Nama barang" oninput="sjItems[${i}].nama=this.value;updatePreview()"></td>
      <td><input type="number" value="${it.qty}" placeholder="0" style="width:48px" oninput="sjItems[${i}].qty=this.value;updatePreview()"></td>
      <td><input type="text" value="${it.satuan}" placeholder="pcs" style="width:52px" oninput="sjItems[${i}].satuan=this.value;updatePreview()"></td>
      <td><button class="btn-del-row" onclick="removeSJRow(${i})">×</button></td>`;
    tb.appendChild(tr);
  });
}

// ─── GET FORM VALUES ──────────────────────────────────────
function getV(id) { const el = document.getElementById(id); return el ? el.value : ''; }

function getNotaSummary() {
  const total = notaItems.reduce((s, it) => {
    const sub = it.type === 'tegel'
      ? (parseFloat(it.m2) || 0) * (parseFloat(it.harga) || 0)
      : (parseFloat(it.pcs) || 0) * (parseFloat(it.harga) || 0);
    return s + sub;
  }, 0);
  const dp = parseFmt(getV('nota-dp'));
  const ongkir = parseFmt(getV('nota-ongkir'));
  const kekurangan = total - dp + ongkir;
  // update kekurangan display
  const kekEl = document.getElementById('nota-kekurangan');
  if (kekEl) kekEl.value = total ? fmt(kekurangan) : '';
  return { total, dp, ongkir, kekurangan };
}

// ─── PREVIEW ─────────────────────────────────────────────
function updatePreview() {
  const container = document.getElementById('preview-container');
  if (activeTab === 'nota') container.innerHTML = buildNotaHTML();
  else container.innerHTML = buildSJHTML();
}

function headerHTML() {
  return `
  <div class="doc-header">
    <div class="doc-logo"><img src="${LOGO_SRC}" alt="Logo"/></div>
    <div class="doc-info">
      <p class="co-name"><strong>${CO_NAME}</strong></p>
      <p>${CO_ADDR}</p>
      <p>${CO_TELP}</p>
      <p>Email: <a href="mailto:${CO_EMAIL}">${CO_EMAIL}</a></p>
    </div>`;
}

function buildNotaHTML() {
  const nomor = getV('nota-nomor');
  const tgl = getV('nota-tgl') ? new Date(getV('nota-tgl')).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }) : '';
  const po = getV('nota-po');
  const { total, dp, ongkir, kekurangan } = getNotaSummary();

  // Empty rows to at least 10
  const minRows = Math.max(notaItems.length, 10);

  let rows = '';
  for (let i = 0; i < minRows; i++) {
    const it = notaItems[i] || {};
    const jumlah = it.type === 'tegel'
      ? (parseFloat(it.m2) || 0) * (parseFloat(it.harga) || 0)
      : (parseFloat(it.pcs) || 0) * (parseFloat(it.harga) || 0);
    rows += `<tr>
      <td class="tc">${i + 1}</td>
      <td>${it.nama || ''}</td>
      <td class="tc">${it.pcs || ''}</td>
      <td class="tc">${it.m2 || ''}</td>
      <td class="tr">${it.harga ? fmt(parseFloat(it.harga)) : ''}</td>
      <td class="tr">${jumlah ? fmt(jumlah) : ''}</td>
    </tr>`;
  }

  return `<div class="doc-wrap">
    ${headerHTML()}
    <div class="doc-title-area">
      <h2>NOTA<br>PENJUALAN</h2>
      <div class="doc-number">${tgl ? `Magelang, ${tgl}` : 'Magelang'}</div>
    </div>
  </div>
  <div style="font-size:10px;font-family:Arial;padding:4px 0 2px;font-weight:bold">NO. PO: ${po}</div>
  <table class="doc-table">
    <thead><tr>
      <th style="width:30px">NO</th>
      <th>NAMA BARANG</th>
      <th style="width:38px">PCs</th>
      <th style="width:38px">M2</th>
      <th style="width:90px">HARGA</th>
      <th style="width:90px">JUMLAH</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table>
  <table class="nota-footer-table">
    <tr>
      <td rowspan="4" style="width:75.5%;vertical-align:middle;font-size:10px">
        <strong>Nb.</strong> Pembayaran via Bank/Giro/Cek sah bila uang sudah diterima Perusahaan<br>
        <strong>BCA A.N NURJAMAL a c :1040257477</strong>
      </td>
      <td class="lbl">Jumlah</td><td style="width:12.25%; text-align:right">${fmt(total)}</td>
    </tr>
    <tr><td class="lbl">DP</td><td style="width:12.25%; text-align:right">${fmt(dp)}</td></tr>
    <tr><td class="lbl">Ongkir</td><td style="width:12.25%; text-align:right">${fmt(ongkir)}</td></tr>
    <tr><td class="lbl">Kekurangan</td><td style="width:12.25%; text-align:right;font-weight:bold">${total ? fmt(kekurangan) : ''}</td></tr>
  </table>
  <div class="nota-sig">
    <div>CUSTOMER</div>
    <div>RAPA CAST STONE</div>
  </div>
</div>`;
}

function buildSJHTML() {
  const nomor = getV('sj-nomor');
  const noSurat = getV('sj-nosurat');
  const tgl = getV('sj-tgl') ? new Date(getV('sj-tgl')).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }) : '';
  const po = getV('sj-po');
  const kepada = getV('sj-kepada');
  const catatan = getV('sj-catatan');

  const minRows = Math.max(sjItems.length, 10);
  let rows = '';
  for (let i = 0; i < minRows; i++) {
    const it = sjItems[i] || {};
    rows += `<tr>
      <td class="tc">${i + 1}</td>
      <td class="tc">${it.kode || ''}</td>
      <td>${it.nama || ''}</td>
      <td class="tc">${it.qty || ''}</td>
      <td class="tc">${it.satuan || ''}</td>
      ${i === 0 ? `<td class="tc" rowspan="${minRows}" style="font-style:italic;color:#555;vertical-align:middle;font-size:10px">${catatan}</td>` : ''}
    </tr>`;
  }

  return `<div class="doc-wrap">
    ${headerHTML()}
    <div class="doc-title-area">
      <h2>SURAT JALAN</h2>
      <div class="doc-number">NO : ${nomor}</div>
    </div>
  </div>
  <div class="sj-meta-grid" style="font-family:Arial;font-size:10px;margin-bottom:6px">
    <div>
      <div class="doc-meta-row"><span class="doc-meta-label" style="font-weight:bold;min-width:70px">No Surat:</span><span class="doc-meta-value">${noSurat}</span></div>
      <div class="doc-meta-row"><span class="doc-meta-label" style="font-weight:bold;min-width:70px">Tanggal</span><span class="doc-meta-value">${tgl}</span></div>
      <div class="doc-meta-row"><span class="doc-meta-label" style="font-weight:bold;min-width:70px">No. Po</span><span class="doc-meta-value">${po}</span></div>
    </div>
    <div>
      <div class="doc-meta-row"><span class="doc-meta-label" style="font-weight:bold;min-width:85px">Kepada Yth:</span><span class="doc-meta-value">${kepada}</span></div>
    </div>
  </div>
  <table class="doc-table">
    <thead><tr>
      <th style="width:28px">No</th>
      <th style="width:55px">Kode</th>
      <th>Nama Barang</th>
      <th style="width:38px">Qty</th>
      <th style="width:50px">Satuan</th>
      <th style="width:150px">Keterangan</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table>
  <p class="sj-footer-note">Barang diterima dengan kondisi baik</p>
  <div class="sj-sig-grid">
    ${['Dibuat Oleh', 'Dikirim Oleh', 'Diterima Oleh'].map(label => `
    <div class="sj-sig-col">
      <div class="sig-title">${label}</div>
      <div class="sig-space"></div>
      <div class="sig-bracket"><span>(</span><span>)</span></div>
      <div class="sig-date"><span>Tanggal:</span><span class="line"></span></div>
    </div>`).join('')}
  </div>
</div>`;
}

// ─── RESET ────────────────────────────────────────────────
function resetForm() {
  if (activeTab === 'nota') {
    ['nota-tgl', 'nota-po', 'nota-dp', 'nota-ongkir', 'nota-kekurangan'].forEach(id => {
      const el = document.getElementById(id); if (el) el.value = '';
    });
    notaItems = [newItem()];
    renderNotaItemsForm();
  } else {
    ['sj-nosurat', 'sj-tgl', 'sj-po', 'sj-kepada', 'sj-catatan'].forEach(id => {
      const el = document.getElementById(id); if (el) el.value = '';
    });
    sjItems = [newSJItem()];
    renderSJItemsForm();
  }
  updatePreview();
}

// ─── INIT ─────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  initNomor('nota_counter', 'NP', 'nota-nomor');
  initNomor('sj_counter', 'SJ', 'sj-nomor');
  renderNotaItemsForm();
  renderSJItemsForm();
  switchTab('nota');

  // Save nomor on change (debounced for Supabase)
  document.getElementById('nota-nomor').addEventListener('input', e => debounceSaveNomor('nota_counter', e.target.value));
  document.getElementById('sj-nomor').addEventListener('input', e => debounceSaveNomor('sj_counter', e.target.value));

  // Live update all inputs
  document.querySelectorAll('#form-nota input, #form-nota textarea').forEach(el => {
    el.addEventListener('input', updatePreview);
  });
  document.querySelectorAll('#form-sj input, #form-sj textarea').forEach(el => {
    el.addEventListener('input', updatePreview);
  });
});
