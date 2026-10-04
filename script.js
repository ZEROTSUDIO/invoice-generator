// ─── CONSTANTS & CONFIGURATION PRESETS ─────────────────────────
const DEFAULT_RAPA_CONFIG = {
  id: 'rapa',
  logo: 'rapa-logo.png',
  name: 'RAPA CEMENT & GRC',
  shortName: 'Rapa Cement & GRC',
  addr: 'Jl. Ngadiretno no. 33, Tamanagung, Muntilan, Magelang 56413',
  telp: 'Telp: 08112959125 / 082134567874',
  email: 'rapastone33@gmail.com',
  city: 'Magelang',
  bank: 'BCA A.N NURJAMAL a c :1040257477',
  signatureName: 'RAPA CAST STONE'
};

const DEFAULT_BINTANG_CONFIG = {
  id: 'bintang',
  logo: 'random.png',
  name: 'BINTANG JAYA MATERIAL',
  shortName: 'Bintang Jaya Material',
  addr: 'Jl. Contoh Fiktif No. 99, Jakarta Selatan 12345',
  telp: 'Telp: 081234567890',
  email: 'hello@bintangmaterial.dummy',
  city: 'Jakarta Selatan',
  bank: 'MANDIRI A.N BINTANG JAYA : 1234567890',
  signatureName: 'BINTANG JAYA MATERIAL'
};

const BLANK_CONFIG = {
  id: 'blank',
  logo: '',
  name: 'NAMA PERUSAHAAN / TOKO',
  shortName: 'Nama Toko / Perusahaan',
  addr: 'Jl. Alamat Usaha No. 1, Kota',
  telp: 'Telp: 08123456789',
  email: 'email@usaha.com',
  city: 'Jakarta',
  bank: 'BCA 1234567890 A.N NAMA PEMILIK',
  signatureName: 'HORMAT KAMI'
};

let currentConfig = { ...DEFAULT_RAPA_CONFIG };

// ─── CONFIG STORAGE & MANAGEMENT ──────────────────────────
function loadSavedConfig() {
  const saved = localStorage.getItem('saved_company_config');
  if (saved) {
    try {
      currentConfig = Object.assign({}, DEFAULT_RAPA_CONFIG, JSON.parse(saved));
    } catch (e) {
      console.error('Error parsing saved company config:', e);
    }
  }
}

function saveCompanyConfig() {
  localStorage.setItem('saved_company_config', JSON.stringify(currentConfig));
}

function updateConfigField(field, value) {
  currentConfig[field] = value;
  currentConfig.id = 'custom';
  saveCompanyConfig();
  updateHeaderAndCardBranding();
  updatePreview();

  const sel = document.getElementById('preset-selector');
  if (sel && sel.value !== 'custom') {
    sel.value = 'custom';
  }
}

function loadPreset(presetKey) {
  if (presetKey === 'rapa') {
    currentConfig = { ...DEFAULT_RAPA_CONFIG };
  } else if (presetKey === 'bintang') {
    currentConfig = { ...DEFAULT_BINTANG_CONFIG };
  } else if (presetKey === 'blank') {
    currentConfig = { ...BLANK_CONFIG };
  } else {
    currentConfig.id = 'custom';
  }
  saveCompanyConfig();
  updateCompanyFormFromConfig();
  updateHeaderAndCardBranding();
  updatePreview();
}

function handleLogoUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    currentConfig.logo = e.target.result;
    currentConfig.id = 'custom';
    saveCompanyConfig();
    updateCompanyFormFromConfig();
    updateHeaderAndCardBranding();
    updatePreview();
  };
  reader.readAsDataURL(file);
}

function removeLogo() {
  currentConfig.logo = '';
  currentConfig.id = 'custom';
  saveCompanyConfig();
  updateCompanyFormFromConfig();
  updateHeaderAndCardBranding();
  updatePreview();
}

function updateCompanyFormFromConfig() {
  const setVal = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.value = val || '';
  };
  setVal('cfg-name', currentConfig.name);
  setVal('cfg-city', currentConfig.city);
  setVal('cfg-telp', currentConfig.telp);
  setVal('cfg-addr', currentConfig.addr);
  setVal('cfg-email', currentConfig.email);
  setVal('cfg-signature', currentConfig.signatureName);
  setVal('cfg-bank', currentConfig.bank);

  const sel = document.getElementById('preset-selector');
  if (sel && currentConfig.id) sel.value = currentConfig.id;
}

function updateHeaderAndCardBranding() {
  const appLogo = document.getElementById('app-logo');
  const appCompanyName = document.getElementById('app-company-name');
  const cardCompanyName = document.getElementById('card-company-name');

  if (appLogo) {
    if (currentConfig.logo) {
      appLogo.src = currentConfig.logo;
      appLogo.style.display = 'block';
    } else {
      appLogo.style.display = 'none';
    }
  }
  if (appCompanyName) {
    appCompanyName.innerText = currentConfig.shortName || currentConfig.name || '';
  }
  if (cardCompanyName) {
    cardCompanyName.innerText = currentConfig.name || 'Profil Usaha';
  }
}

function toggleCompanyProfile() {
  const body = document.getElementById('company-card-body');
  const arrow = document.getElementById('profile-arrow');
  const card = document.getElementById('company-card');
  if (!body) return;
  const isHidden = body.style.display === 'none';
  body.style.display = isHidden ? 'block' : 'none';
  if (arrow) arrow.classList.toggle('open', isHidden);
  if (card) card.classList.toggle('expanded', isHidden);
}

// ─── STATE ───────────────────────────────────────────────
let activeTab = 'nota';
let isTegelMode = false;
let notaItems = [newItem()];
let sjItems = [newSJItem()];
let isGuest = false;

function newItem() {
  return {
    type: 'pcs',
    nama: '',
    qty: '1',
    satuan: 'pcs',
    pcs: '1',
    m2: '',
    harga: ''
  };
}

function newSJItem() {
  return {
    kode: '',
    nama: '',
    qty: '1',
    satuan: 'pcs'
  };
}

function toggleTegelMode(checked) {
  isTegelMode = checked;
  renderNotaItemsForm();
  updatePreview();
}

// ─── SUPABASE PARAMS ───────────────────────────────────────
const supabaseUrl = 'https://yhhxbmbjzrgtfxjdrizu.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InloaHhibWJqenJndGZ4amRyaXp1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzIxNzQyNjgsImV4cCI6MjA4Nzc1MDI2OH0.2bX25UW6r_BcN_yHUPN7ap5wRHuFhZFawTAuZKFvLmo';
const supabaseClient = supabase.createClient(supabaseUrl, supabaseKey);

// ─── NOMOR AUTO (SUPABASE / LOCALSTORAGE) ───────────────────
function getYYYYMM() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function getNotaTitle() {
  const sel = document.getElementById('nota-title-select');
  if (!sel) return 'NOTA PENJUALAN';
  if (sel.value === 'CUSTOM') {
    const cust = document.getElementById('nota-title-custom');
    return cust && cust.value.trim() ? cust.value.trim() : 'DOKUMEN';
  }
  return sel.value;
}

function getNotaPrefix() {
  const title = getNotaTitle().toUpperCase();
  if (title.includes('INVOICE')) return 'INV';
  if (title.includes('FAKTUR')) return 'FP';
  return 'NP';
}

function handleNotaTitleChange() {
  const sel = document.getElementById('nota-title-select');
  const wrap = document.getElementById('nota-title-custom-wrap');
  if (!sel) return;
  if (sel.value === 'CUSTOM') {
    if (wrap) wrap.style.display = '';
  } else {
    if (wrap) wrap.style.display = 'none';
  }
  const label = document.getElementById('nota-nomor-label');
  if (label) label.innerText = `No. ${getNotaTitle()}`;
  updatePreview();
}

function getSJTitle() {
  const sel = document.getElementById('sj-title-select');
  if (!sel) return 'SURAT JALAN';
  if (sel.value === 'CUSTOM') {
    const cust = document.getElementById('sj-title-custom');
    return cust && cust.value.trim() ? cust.value.trim() : 'SURAT JALAN';
  }
  return sel.value;
}

function getSJPrefix() {
  const title = getSJTitle().toUpperCase();
  if (title.includes('DELIVERY')) return 'DO';
  if (title.includes('PENGANTAR')) return 'SP';
  return 'SJ';
}

function handleSJTitleChange() {
  const sel = document.getElementById('sj-title-select');
  const wrap = document.getElementById('sj-title-custom-wrap');
  if (!sel) return;
  if (sel.value === 'CUSTOM') {
    if (wrap) wrap.style.display = '';
  } else {
    if (wrap) wrap.style.display = 'none';
  }
  updatePreview();
}

async function initNomor(key, prefix, inputId) {
  const ym = getYYYYMM();
  const inputEl = document.getElementById(inputId);
  if (!inputEl) return;
  inputEl.value = 'Loading...';

  let data, error;

  if (isGuest) {
    const localData = localStorage.getItem(`seq_${key}`);
    data = localData ? JSON.parse(localData) : null;
  } else {
    const res = await supabaseClient
      .from('document_sequences')
      .select('*')
      .eq('id', key)
      .single();
    data = res.data;
    error = res.error;
  }

  if (error && error.code !== 'PGRST116') {
    console.error('Error fetching sequence:', error);
    inputEl.value = `${prefix}-${ym}-001`;
    return;
  }

  let nomor;
  if (data) {
    if (data.ym === ym && data.current_val) {
      nomor = data.current_val;
    } else {
      let nextSeq = data.ym === ym ? data.seq || 0 : 0;
      nextSeq++;
      nomor = `${prefix}-${ym}-${String(nextSeq).padStart(3, '0')}`;

      if (isGuest) {
        localStorage.setItem(
          `seq_${key}`,
          JSON.stringify({ ym: ym, seq: nextSeq, current_val: nomor })
        );
      } else {
        await supabaseClient
          .from('document_sequences')
          .update({ ym: ym, seq: nextSeq, current_val: nomor })
          .eq('id', key);
      }
    }
  } else {
    nomor = `${prefix}-${ym}-001`;
    if (isGuest) {
      localStorage.setItem(
        `seq_${key}`,
        JSON.stringify({ ym: ym, seq: 1, current_val: nomor })
      );
    }
  }

  inputEl.value = nomor;
  updatePreview();
}

async function saveCurrentNomor(key, val) {
  const parts = val.split('-');
  const seqPart = parts[parts.length - 1];
  const seqNum = parseInt(seqPart) || 0;
  const ym = parts.length > 1 ? parts[parts.length - 2] : getYYYYMM();

  const updateData = { current_val: val };
  if (seqNum > 0) {
    updateData.seq = seqNum;
    updateData.ym = ym;
  }

  if (isGuest) {
    localStorage.setItem(
      `seq_${key}`,
      JSON.stringify({ ym: ym, seq: seqNum, current_val: val })
    );
  } else {
    await supabaseClient
      .from('document_sequences')
      .update(updateData)
      .eq('id', key);
  }
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
  if (!inputEl) return;
  inputEl.value = 'Generating...';

  const ym = getYYYYMM();
  let data;

  if (isGuest) {
    const localData = localStorage.getItem(`seq_${key}`);
    data = localData ? JSON.parse(localData) : { ym: ym, seq: 0 };
  } else {
    const res = await supabaseClient
      .from('document_sequences')
      .select('*')
      .eq('id', key)
      .single();
    data = res.data;
  }

  let nextSeq = data && data.ym === ym ? data.seq || 0 : 0;
  nextSeq++;
  const nomor = `${prefix}-${ym}-${String(nextSeq).padStart(3, '0')}`;

  if (isGuest) {
    localStorage.setItem(
      `seq_${key}`,
      JSON.stringify({ ym: ym, seq: nextSeq, current_val: nomor })
    );
  } else {
    await supabaseClient
      .from('document_sequences')
      .update({ ym: ym, seq: nextSeq, current_val: nomor })
      .eq('id', key);
  }

  inputEl.value = nomor;
  updatePreview();
}

async function resetSequence(key, prefix, inputId) {
  if (!confirm('Yakin ingin reset nomor ke 001 untuk bulan ini?')) return;

  const inputEl = document.getElementById(inputId);
  if (!inputEl) return;
  inputEl.value = 'Resetting...';

  const ym = getYYYYMM();
  const nomor = `${prefix}-${ym}-001`;

  if (isGuest) {
    localStorage.setItem(
      `seq_${key}`,
      JSON.stringify({ ym: ym, seq: 1, current_val: nomor })
    );
    inputEl.value = nomor;
    updatePreview();
  } else {
    const { error } = await supabaseClient
      .from('document_sequences')
      .update({ ym: ym, seq: 1, current_val: nomor })
      .eq('id', key);

    if (error) {
      alert('Gagal reset: ' + error.message);
      initNomor(key, prefix, inputId);
    } else {
      inputEl.value = nomor;
      updatePreview();
    }
  }
}

async function decrementSequence(key, prefix, inputId) {
  const inputEl = document.getElementById(inputId);
  if (!inputEl) return;
  const currentVal = inputEl.value;

  const parts = currentVal.split('-');
  let seqNum = parseInt(parts[parts.length - 1]) || 0;

  if (seqNum <= 1) {
    alert('Sudah di nomor paling awal (001).');
    return;
  }

  seqNum--;
  const ym = getYYYYMM();
  const nomor = `${prefix}-${ym}-${String(seqNum).padStart(3, '0')}`;

  inputEl.value = 'Rolling back...';

  if (isGuest) {
    localStorage.setItem(
      `seq_${key}`,
      JSON.stringify({ ym: ym, seq: seqNum, current_val: nomor })
    );
    inputEl.value = nomor;
    updatePreview();
  } else {
    const { error } = await supabaseClient
      .from('document_sequences')
      .update({ ym: ym, seq: seqNum, current_val: nomor })
      .eq('id', key);

    if (error) {
      alert('Gagal undo: ' + error.message);
      initNomor(key, prefix, inputId);
    } else {
      inputEl.value = nomor;
      updatePreview();
    }
  }
}

// ─── COPY FROM INVOICE TO SURAT JALAN ──────────────────────
function copyFromInvoice() {
  const notaPo = getV('nota-po');
  const notaKepada = getV('nota-kepada');
  const notaTgl = getV('nota-tgl');
  const notaAlamat = getV('nota-alamat-customer');

  const sjPo = document.getElementById('sj-po');
  const sjKepada = document.getElementById('sj-kepada');
  const sjTgl = document.getElementById('sj-tgl');
  const sjAlamat = document.getElementById('sj-alamat');

  if (sjPo && notaPo) sjPo.value = notaPo;
  if (sjKepada && notaKepada) sjKepada.value = notaKepada;
  if (sjTgl && notaTgl) sjTgl.value = notaTgl;
  if (sjAlamat && notaAlamat) sjAlamat.value = notaAlamat;

  if (notaItems && notaItems.length > 0) {
    sjItems = notaItems.map((it, idx) => {
      const qtyVal = isTegelMode
        ? (it.type === 'tegel' ? (it.m2 || it.pcs || '1') : (it.pcs || '1'))
        : (it.qty || it.pcs || '1');
      const satVal = isTegelMode
        ? (it.type === 'tegel' ? 'm²' : 'pcs')
        : (it.satuan || 'pcs');
      return {
        kode: (idx + 1).toString().padStart(3, '0'),
        nama: it.nama || '',
        qty: qtyVal,
        satuan: satVal
      };
    });
    renderSJItemsForm();
  }

  updatePreview();
  alert('Data dari Nota/Invoice berhasil disalin ke Surat Jalan!');
}

// ─── NUMBER FORMAT ────────────────────────────────────────
function fmt(n) {
  const val = parseFloat(n) || 0;
  return val === 0 ? '' : val.toLocaleString('id-ID');
}

function parseFmt(s) {
  return parseFloat(String(s).replace(/\./g, '').replace(',', '.')) || 0;
}

// ─── TAB SWITCH ──────────────────────────────────────────
function switchTab(tab) {
  activeTab = tab;
  document
    .querySelectorAll('.tab-btn')
    .forEach((b) => b.classList.toggle('active', b.dataset.tab === tab));
  document.getElementById('form-nota').style.display =
    tab === 'nota' ? '' : 'none';
  document.getElementById('form-sj').style.display =
    tab === 'sj' ? '' : 'none';
  updatePreview();
}

// ─── ITEM ROW MANAGEMENT ──────────────────────────────────
function addNotaRow() {
  notaItems.push(newItem());
  renderNotaItemsForm();
  updatePreview();
}

function removeNotaRow(i) {
  if (notaItems.length > 1) {
    notaItems.splice(i, 1);
    renderNotaItemsForm();
    updatePreview();
  }
}

function addSJRow() {
  sjItems.push(newSJItem());
  renderSJItemsForm();
  updatePreview();
}

function removeSJRow(i) {
  if (sjItems.length > 1) {
    sjItems.splice(i, 1);
    renderSJItemsForm();
    updatePreview();
  }
}

function updateNotaFormJumlah(i) {
  const it = notaItems[i];
  if (!it) return;
  const jumlah = isTegelMode
    ? it.type === 'tegel'
      ? (parseFloat(it.m2) || 0) * (parseFloat(it.harga) || 0)
      : (parseFloat(it.pcs) || 0) * (parseFloat(it.harga) || 0)
    : (parseFloat(it.qty) || 0) * (parseFloat(it.harga) || 0);

  const el = document.getElementById(`nota-jumlah-${i}`);
  if (el) el.innerText = fmt(jumlah);
  getNotaSummary();
}

function renderNotaItemsForm() {
  const thead = document.getElementById('nota-items-thead');
  const tbody = document.getElementById('nota-items-tbody');
  if (!tbody) return;

  if (thead) {
    if (isTegelMode) {
      thead.innerHTML = `
        <tr>
          <th style="width: 22px">#</th>
          <th style="width: 75px">Tipe</th>
          <th>Nama Barang</th>
          <th style="width: 50px">PCs</th>
          <th style="width: 50px">M2</th>
          <th style="width: 85px">Harga (Rp)</th>
          <th style="width: 75px">Jumlah</th>
          <th style="width: 24px"></th>
        </tr>`;
    } else {
      thead.innerHTML = `
        <tr>
          <th style="width: 22px">#</th>
          <th>Nama Barang / Jasa</th>
          <th style="width: 52px">Qty</th>
          <th style="width: 56px">Satuan</th>
          <th style="width: 85px">Harga (Rp)</th>
          <th style="width: 75px">Jumlah</th>
          <th style="width: 24px"></th>
        </tr>`;
    }
  }

  tbody.innerHTML = '';
  notaItems.forEach((it, i) => {
    const tr = document.createElement('tr');
    const jumlah = isTegelMode
      ? it.type === 'tegel'
        ? (parseFloat(it.m2) || 0) * (parseFloat(it.harga) || 0)
        : (parseFloat(it.pcs) || 0) * (parseFloat(it.harga) || 0)
      : (parseFloat(it.qty) || 0) * (parseFloat(it.harga) || 0);

    if (isTegelMode) {
      tr.innerHTML = `
        <td style="text-align:center;color:#888;font-weight:600">${i + 1}</td>
        <td>
          <select onchange="notaItems[${i}].type=this.value;updatePreview();updateNotaFormJumlah(${i})" style="width:72px;font-size:11px">
            <option value="pcs" ${it.type === 'pcs' ? 'selected' : ''}>Lainnya</option>
            <option value="tegel" ${it.type === 'tegel' ? 'selected' : ''}>Tegel</option>
          </select>
        </td>
        <td><input type="text" value="${it.nama || ''}" placeholder="Nama barang" oninput="notaItems[${i}].nama=this.value;updatePreview()"></td>
        <td><input type="number" value="${it.pcs || ''}" placeholder="0" style="width:50px" oninput="notaItems[${i}].pcs=this.value;updatePreview();updateNotaFormJumlah(${i})"></td>
        <td><input type="number" value="${it.m2 || ''}" placeholder="0" style="width:50px" oninput="notaItems[${i}].m2=this.value;updatePreview();updateNotaFormJumlah(${i})"></td>
        <td><input type="number" value="${it.harga || ''}" placeholder="0" style="width:85px" oninput="notaItems[${i}].harga=this.value;updatePreview();updateNotaFormJumlah(${i})"></td>
        <td id="nota-jumlah-${i}" style="text-align:right;font-weight:600;font-size:11px;white-space:nowrap">${fmt(jumlah)}</td>
        <td><button class="btn-del-row" onclick="removeNotaRow(${i})">×</button></td>`;
    } else {
      tr.innerHTML = `
        <td style="text-align:center;color:#888;font-weight:600">${i + 1}</td>
        <td><input type="text" value="${it.nama || ''}" placeholder="Nama barang / jasa" oninput="notaItems[${i}].nama=this.value;updatePreview()"></td>
        <td><input type="number" value="${it.qty || ''}" placeholder="1" style="width:52px" oninput="notaItems[${i}].qty=this.value;updatePreview();updateNotaFormJumlah(${i})"></td>
        <td><input type="text" list="satuan-presets" value="${it.satuan || 'pcs'}" placeholder="pcs" style="width:56px" oninput="notaItems[${i}].satuan=this.value;updatePreview()"></td>
        <td><input type="number" value="${it.harga || ''}" placeholder="0" style="width:85px" oninput="notaItems[${i}].harga=this.value;updatePreview();updateNotaFormJumlah(${i})"></td>
        <td id="nota-jumlah-${i}" style="text-align:right;font-weight:600;font-size:11px;white-space:nowrap">${fmt(jumlah)}</td>
        <td><button class="btn-del-row" onclick="removeNotaRow(${i})">×</button></td>`;
    }
    tbody.appendChild(tr);
  });
}

function renderSJItemsForm() {
  const tb = document.getElementById('sj-items-tbody');
  if (!tb) return;
  tb.innerHTML = '';
  sjItems.forEach((it, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align:center;color:#888;font-weight:600">${i + 1}</td>
      <td><input type="text" value="${it.kode || ''}" placeholder="Kode" style="width:60px" oninput="sjItems[${i}].kode=this.value;updatePreview()"></td>
      <td><input type="text" value="${it.nama || ''}" placeholder="Nama barang" oninput="sjItems[${i}].nama=this.value;updatePreview()"></td>
      <td><input type="number" value="${it.qty || ''}" placeholder="0" style="width:48px" oninput="sjItems[${i}].qty=this.value;updatePreview()"></td>
      <td><input type="text" list="satuan-presets" value="${it.satuan || 'pcs'}" placeholder="pcs" style="width:52px" oninput="sjItems[${i}].satuan=this.value;updatePreview()"></td>
      <td><button class="btn-del-row" onclick="removeSJRow(${i})">×</button></td>`;
    tb.appendChild(tr);
  });
}

// ─── GET FORM VALUES ──────────────────────────────────────
function getV(id) {
  const el = document.getElementById(id);
  return el ? el.value : '';
}

function getNotaSummary() {
  const subtotal = notaItems.reduce((s, it) => {
    const sub = isTegelMode
      ? it.type === 'tegel'
        ? (parseFloat(it.m2) || 0) * (parseFloat(it.harga) || 0)
        : (parseFloat(it.pcs) || 0) * (parseFloat(it.harga) || 0)
      : (parseFloat(it.qty) || 0) * (parseFloat(it.harga) || 0);
    return s + sub;
  }, 0);

  const diskon = parseFmt(getV('nota-diskon'));
  const dp = parseFmt(getV('nota-dp'));
  const ongkir = parseFmt(getV('nota-ongkir'));
  const kekurangan = subtotal - diskon + ongkir - dp;

  const subEl = document.getElementById('nota-subtotal');
  if (subEl) subEl.value = subtotal ? fmt(subtotal) : '0';

  const kekEl = document.getElementById('nota-kekurangan');
  if (kekEl) kekEl.value = subtotal ? fmt(kekurangan) : '';

  return { subtotal, diskon, dp, ongkir, kekurangan };
}

// ─── PREVIEW BUILDERS ─────────────────────────────────────
function updatePreview() {
  const container = document.getElementById('preview-container');
  if (!container) return;
  if (activeTab === 'nota') container.innerHTML = buildNotaHTML();
  else container.innerHTML = buildSJHTML();
}

function headerHTML() {
  const hasLogo = !!currentConfig.logo;
  return `
  <div class="doc-header ${!hasLogo ? 'no-logo' : ''}">
    ${hasLogo ? `<div class="doc-logo"><img src="${currentConfig.logo}" alt="Logo"/></div>` : ''}
    <div class="doc-info">
      <p class="co-name"><strong>${currentConfig.name || ''}</strong></p>
      ${currentConfig.addr ? `<p>${currentConfig.addr}</p>` : ''}
      ${currentConfig.telp ? `<p>${currentConfig.telp}</p>` : ''}
      ${currentConfig.email ? `<p>Email: <a href="mailto:${currentConfig.email}">${currentConfig.email}</a></p>` : ''}
    </div>`;
}

function buildNotaHTML() {
  const docTitle = getNotaTitle();
  const nomor = getV('nota-nomor');
  const tglRaw = getV('nota-tgl');
  const tgl = tglRaw
    ? new Date(tglRaw).toLocaleDateString('id-ID', {
        day: '2-digit',
        month: 'long',
        year: 'numeric'
      })
    : '';
  const city = currentConfig.city || '';
  const dateWithCity = tgl ? (city ? `${city}, ${tgl}` : tgl) : city;
  const po = getV('nota-po');
  const kepada = getV('nota-kepada');
  const alamatCust = getV('nota-alamat-customer');
  const { subtotal, diskon, dp, ongkir, kekurangan } = getNotaSummary();

  const minRows = Math.max(notaItems.length, 10);
  let rows = '';
  for (let i = 0; i < minRows; i++) {
    const it = notaItems[i] || {};
    if (isTegelMode) {
      const jumlah =
        it.type === 'tegel'
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
    } else {
      const qtyVal = parseFloat(it.qty) || 0;
      const hargaVal = parseFloat(it.harga) || 0;
      const jumlah = qtyVal * hargaVal;
      rows += `<tr>
        <td class="tc">${i + 1}</td>
        <td>${it.nama || ''}</td>
        <td class="tc">${it.qty || ''}</td>
        <td class="tc">${it.satuan || ''}</td>
        <td class="tr">${it.harga ? fmt(hargaVal) : ''}</td>
        <td class="tr">${jumlah ? fmt(jumlah) : ''}</td>
      </tr>`;
    }
  }

  const tableThead = isTegelMode
    ? `<thead><tr>
        <th style="width:30px">NO</th>
        <th>NAMA BARANG</th>
        <th style="width:38px">PCs</th>
        <th style="width:38px">M2</th>
        <th style="width:90px">HARGA</th>
        <th style="width:90px">JUMLAH</th>
      </tr></thead>`
    : `<thead><tr>
        <th style="width:30px">NO</th>
        <th>NAMA BARANG / DESKRIPSI</th>
        <th style="width:40px">QTY</th>
        <th style="width:50px">SATUAN</th>
        <th style="width:90px">HARGA</th>
        <th style="width:90px">JUMLAH</th>
      </tr></thead>`;

  return `<div class="doc-wrap">
    ${headerHTML()}
      <div class="doc-title-area">
        <h2>${docTitle}</h2>
        <div class="doc-number">${dateWithCity}</div>
        ${nomor ? `<div style="font-size:11px;font-weight:bold;margin-top:2px">NO: ${nomor}</div>` : ''}
      </div>
    </div>
    <div class="doc-customer-bar">
      <div class="customer-left">
        <div class="doc-meta-row">
          <span class="doc-meta-label">Kepada Yth:</span>
          <span class="doc-meta-value"><strong>${kepada || '—'}</strong>${alamatCust ? `<br>${alamatCust}` : ''}</span>
        </div>
      </div>
      <div class="customer-right">
        <div class="doc-meta-row" style="justify-content: flex-end">
          <span class="doc-meta-label" style="min-width: 60px">NO. PO:</span>
          <span class="doc-meta-value" style="flex: unset; min-width: 100px; text-align: right"><strong>${po || '—'}</strong></span>
        </div>
      </div>
    </div>
    <table class="doc-table">
      ${tableThead}
      <tbody>${rows}</tbody>
    </table>
    <table class="nota-footer-table">
      <tr>
        <td rowspan="${diskon > 0 ? '5' : '4'}" style="vertical-align:middle;font-size:10px">
          <strong>Nb.</strong> Pembayaran via Bank/Giro/Cek sah bila uang sudah diterima Perusahaan<br>
          <strong>${currentConfig.bank || ''}</strong>
        </td>
        <td class="lbl" style="width:90px">Subtotal</td><td style="width:90px; text-align:right">${fmt(subtotal)}</td>
      </tr>
      ${diskon > 0 ? `<tr><td class="lbl">Diskon</td><td style="text-align:right;color:#d92d20">- ${fmt(diskon)}</td></tr>` : ''}
      <tr><td class="lbl">DP</td><td style="text-align:right">${fmt(dp)}</td></tr>
      <tr><td class="lbl">Ongkir</td><td style="text-align:right">${fmt(ongkir)}</td></tr>
      <tr><td class="lbl">Kekurangan</td><td style="text-align:right;font-weight:bold">${subtotal ? fmt(kekurangan) : ''}</td></tr>
    </table>
    <div class="nota-sig">
      <div>CUSTOMER</div>
      <div>${currentConfig.signatureName || currentConfig.name || 'HORMAT KAMI'}</div>
    </div>
  </div>`;
}

function buildSJHTML() {
  const sjTitle = getSJTitle();
  const nomor = getV('sj-nomor');
  const noSurat = getV('sj-nosurat');
  const tglRaw = getV('sj-tgl');
  const tgl = tglRaw
    ? new Date(tglRaw).toLocaleDateString('id-ID', {
        day: '2-digit',
        month: 'long',
        year: 'numeric'
      })
    : '';
  const city = currentConfig.city || '';
  const dateWithCity = tgl ? (city ? `${city}, ${tgl}` : tgl) : city;
  const po = getV('sj-po');
  const kepada = getV('sj-kepada');
  const alamatKirim = getV('sj-alamat');
  const catatan = getV('sj-catatan');
  const noteBawah =
    getV('sj-note-bawah') || 'Barang diterima dengan kondisi baik';

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
        <h2>${sjTitle}</h2>
        <div class="doc-number">NO : ${nomor}</div>
      </div>
    </div>
    <div class="sj-meta-grid" style="font-family:Arial;font-size:10px;margin-bottom:6px">
      <div>
        <div class="doc-meta-row"><span class="doc-meta-label" style="font-weight:bold;min-width:70px">No Surat:</span><span class="doc-meta-value">${noSurat || '—'}</span></div>
        <div class="doc-meta-row"><span class="doc-meta-label" style="font-weight:bold;min-width:70px">Tanggal:</span><span class="doc-meta-value">${dateWithCity || '—'}</span></div>
        <div class="doc-meta-row"><span class="doc-meta-label" style="font-weight:bold;min-width:70px">No. PO:</span><span class="doc-meta-value">${po || '—'}</span></div>
      </div>
      <div>
        <div class="doc-meta-row"><span class="doc-meta-label" style="font-weight:bold;min-width:85px">Kepada Yth:</span><span class="doc-meta-value"><strong>${kepada || '—'}</strong></span></div>
        ${alamatKirim ? `<div class="doc-meta-row"><span class="doc-meta-label" style="font-weight:bold;min-width:85px">Alamat Kirim:</span><span class="doc-meta-value">${alamatKirim}</span></div>` : ''}
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
    <p class="sj-footer-note">${noteBawah}</p>
    <div class="sj-sig-grid">
      ${['Dibuat Oleh', 'Dikirim Oleh', 'Diterima Oleh']
        .map(
          (label) => `
      <div class="sj-sig-col">
        <div class="sig-title">${label}</div>
        <div class="sig-space"></div>
        <div class="sig-bracket"><span>(</span><span>)</span></div>
        <div class="sig-date"><span>Tanggal:</span><span class="line"></span></div>
      </div>`
        )
        .join('')}
    </div>
  </div>`;
}

// ─── RESET FORM ───────────────────────────────────────────
function resetForm() {
  if (activeTab === 'nota') {
    [
      'nota-tgl',
      'nota-po',
      'nota-kepada',
      'nota-alamat-customer',
      'nota-diskon',
      'nota-dp',
      'nota-ongkir',
      'nota-kekurangan'
    ].forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
    notaItems = [newItem()];
    renderNotaItemsForm();
  } else {
    [
      'sj-nosurat',
      'sj-tgl',
      'sj-po',
      'sj-kepada',
      'sj-alamat',
      'sj-catatan'
    ].forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.value = '';
    });
    sjItems = [newSJItem()];
    renderSJItemsForm();
  }
  updatePreview();
}

// ─── INIT ─────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  loadSavedConfig();
  updateCompanyFormFromConfig();
  updateHeaderAndCardBranding();

  initNomor('nota_counter', getNotaPrefix(), 'nota-nomor');
  initNomor('sj_counter', getSJPrefix(), 'sj-nomor');
  renderNotaItemsForm();
  renderSJItemsForm();
  switchTab('nota');

  // Debounced save for document numbers
  document
    .getElementById('nota-nomor')
    .addEventListener('input', (e) =>
      debounceSaveNomor('nota_counter', e.target.value)
    );
  document
    .getElementById('sj-nomor')
    .addEventListener('input', (e) =>
      debounceSaveNomor('sj_counter', e.target.value)
    );

  // Live preview update on all input elements
  document
    .querySelectorAll('#form-nota input, #form-nota textarea, #form-nota select')
    .forEach((el) => {
      el.addEventListener('input', updatePreview);
    });
  document
    .querySelectorAll('#form-sj input, #form-sj textarea, #form-sj select')
    .forEach((el) => {
      el.addEventListener('input', updatePreview);
    });

  // Auth Logic
  const authOverlay = document.getElementById('auth-overlay');
  const loginForm = document.getElementById('login-form');
  const loginError = document.getElementById('login-error');

  async function checkSession() {
    const {
      data: { session }
    } = await supabaseClient.auth.getSession();
    if (session) {
      setUIMode('admin');
    } else {
      if (sessionStorage.getItem('guestMode') === 'true') {
        setUIMode('guest');
      } else {
        showAuthOverlay(true);
      }
    }
  }

  function setUIMode(mode) {
    const guestBadge = document.getElementById('guest-badge');
    const logoutBtn = document.getElementById('logout-btn');

    if (mode === 'admin') {
      isGuest = false;
      showAuthOverlay(false);
      if (guestBadge) guestBadge.style.display = 'none';
      if (logoutBtn) logoutBtn.innerHTML = '🚪 Logout Admin';
      sessionStorage.removeItem('guestMode');
    } else if (mode === 'guest') {
      isGuest = true;
      showAuthOverlay(false);
      if (guestBadge) guestBadge.style.display = 'block';
      if (logoutBtn) logoutBtn.innerHTML = '🚪 Keluar Mode Bebas';
      sessionStorage.setItem('guestMode', 'true');
    } else {
      showAuthOverlay(true);
    }

    updateHeaderAndCardBranding();
    initNomor('nota_counter', getNotaPrefix(), 'nota-nomor');
    initNomor('sj_counter', getSJPrefix(), 'sj-nomor');
    updatePreview();
  }

  function showAuthOverlay(show) {
    if (authOverlay) {
      authOverlay.style.display = show ? 'flex' : 'none';
      document.body.style.overflow = show ? 'hidden' : '';
    }
  }

  window.continueAsGuest = function () {
    setUIMode('guest');
  };

  supabaseClient.auth.onAuthStateChange((event, session) => {
    if (event === 'SIGNED_IN') {
      setUIMode('admin');
    } else if (event === 'SIGNED_OUT') {
      showAuthOverlay(true);
      sessionStorage.removeItem('guestMode');
    }
  });

  loginForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    loginError.style.display = 'none';
    const email = document.getElementById('login-email').value;
    const password = document.getElementById('login-pwd').value;

    const { error } = await supabaseClient.auth.signInWithPassword({
      email,
      password
    });
    if (error) {
      loginError.innerText = error.message;
      loginError.style.display = 'block';
    }
  });

  window.handleLogout = async function () {
    if (isGuest) {
      sessionStorage.removeItem('guestMode');
      location.reload();
    } else {
      await supabaseClient.auth.signOut();
    }
  };

  checkSession();
});
