// ============================================
// SERVER.JS — MPK MAN 1 BEKASI (VERCEL READY)
// ============================================

const express = require('express');
const app = express();
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');

// ============ MIDDLEWARE ============
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Nge-serve file statis dari folder public secara aman
app.use(express.static(path.join(__dirname, 'public')));

// ============ FOLDER & DATA SETUP (Aman buat Vercel Serverless) ============
// Vercel hanya mengizinkan penulisan file di folder /tmp bawaan serverless
const DATA_FILE = path.join('/tmp', 'data.json');
const EXCEL_FILE = path.join('/tmp', 'pelanggaran.xlsx');

// Pastikan folder penyimpanan lokal temporary di Vercel sudah siap
if (!fs.existsSync('/tmp/uploads')) {
  fs.mkdirSync('/tmp/uploads', { recursive: true });
}

// ============ MULTER — KONFIGURASI UPLOAD GAMBAR ============
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, '/tmp/uploads/');
  },
  filename: function (req, file, cb) {
    const safeName = file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_');
    cb(null, Date.now() + '-' + file.fieldname + '-' + safeName);
  }
});

const upload = multer({
  storage: storage,
  limits: { fileSize: 5 * 1024 * 1024 }, // Maks 5MB
  fileFilter: function (req, file, cb) {
    const allowedTypes = /jpeg|jpg|png|gif|webp|svg/;
    const extOk = allowedTypes.test(path.extname(file.originalname).toLowerCase());
    const mimeOk = allowedTypes.test(file.mimetype) || file.mimetype === 'image/svg+xml';
    if (extOk || mimeOk) return cb(null, true);
    cb(new Error('Hanya file gambar yang diperbolehkan (jpg, png, gif, webp, svg)'));
  }
});

// Upload khusus Excel (untuk data pelanggaran)
const uploadExcelStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, '/tmp/'),
  filename: (req, file, cb) => cb(null, file.fieldname + path.extname(file.originalname))
});
const uploadExcel = multer({ storage: uploadExcelStorage });

// ============ INISIALISASI DATA AWAL ============
function initData() {
  if (!fs.existsSync(DATA_FILE)) {
    const defaultData = {
      berita: [
        {
          id: 1,
          judul: "Selamat Datang di Website MPK MAN 1 Bekasi",
          isi: "Website resmi Majelis Perwakilan Kelas MAN 1 Bekasi telah resmi diluncurkan. Kami hadir untuk menjadi jembatan antara siswa dan sekolah.",
          tanggal: new Date().toISOString().split('T')[0],
          kategori: "pengumuman",
          gambar: ""
        }
      ],
      kegiatan: [
        {
          id: 1,
          nama: "Pelantikan MPK 2025/2026",
          tanggal: "2025-08-15",
          waktu: "08:00",
          tempat: "Aula MAN 1 Bekasi",
          deskripsi: "Pelantikan resmi pengurus MPK periode 2025/2026 yang akan dihadiri oleh seluruh sivitas akademika MAN 1 Bekasi.",
          status: "upcoming",
          gambar: ""
        },
        {
          id: 2,
          nama: "Masa Ta'aruf Siswa Baru",
          tanggal: "2025-07-14",
          waktu: "07:00",
          tempat: "Lapangan MAN 1 Bekasi",
          deskripsi: "Kegiatan pengenalan lingkungan sekolah bagi siswa baru tahun ajaran 2025/2026.",
          status: "upcoming",
          gambar: ""
        }
      ],
      pengurus: [
        { id: 1, nama: "Ahmad Fauzan", jabatan: "Ketua MPK", kelas: "XII IPA 1", foto: "", instagram: "" },
        { id: 2, nama: "Siti Nurhaliza", jabatan: "Wakil Ketua MPK", kelas: "XII IPS 2", foto: "", instagram: "" },
        { id: 3, nama: "Muhammad Rizky", jabatan: "Sekretaris", kelas: "XI IPA 3", foto: "", instagram: "" },
        { id: 4, nama: "Dewi Rahmawati", jabatan: "Bendahara", kelas: "XI IPS 1", foto: "", instagram: "" },
        { id: 5, nama: "Fajar Maulana", jabatan: "Ketua Bidang Keamanan", kelas: "XII IPA 2", foto: "", instagram: "" },
        { id: 6, nama: "Nabila Azzahra", jabatan: "Ketua Bidang Kebersihan", kelas: "XI IPA 1", foto: "", instagram: "" },
        { id: 7, nama: "Ridho Pratama", jabatan: "Anggota", kelas: "X IPA 4", foto: "", instagram: "" },
        { id: 8, nama: "Zahra Fitria", jabatan: "Anggota", kelas: "X IPS 2", foto: "", instagram: "" }
      ],
      pengaturan: {
        namaSekolah: "MAN 1 Bekasi",
        namaMPK: "Majelis Perwakilan Kelas",
        tahunAjaran: "2024/2025",
        visi: "Mewujudkan MPK yang aktif, aspiratif, dan berintegritas dalam membangun ekosistem sekolah yang harmonis.",
        misi: "Menampung dan menyalurkan aspirasi siswa, mengawal pelaksanaan tata tertib, dan menjadi mitra strategis OSIS.",
        logoPath: "",
        adminPassword: "mpkman1bekasi2024"
      }
    };
    fs.writeFileSync(DATA_FILE, JSON.stringify(defaultData, null, 2));
  }

  if (!fs.existsSync(EXCEL_FILE)) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ['Nama', 'Kelas', 'Tanggal', 'Pelanggaran', 'Poin', 'Keterangan'],
      ['Ahmad Rifai', 'X IPA 1', '2024-01-15', 'Terlambat masuk kelas', 5, 'Terlambat 15 menit'],
      ['Budi Santoso', 'XI IPS 2', '2024-01-20', 'Tidak memakai seragam lengkap', 10, 'Tidak memakai dasi'],
      ['Citra Dewi', 'XII IPA 3', '2024-02-01', 'Tidak mengikuti upacara', 15, 'Tanpa keterangan'],
      ['Dian Putri', 'X IPS 1', '2024-02-10', 'Terlambat masuk kelas', 5, 'Terlambat 20 menit'],
      ['Ahmad Rifai', 'X IPA 1', '2024-03-05', 'Membawa HP tanpa izin', 20, 'HP disita sementara']
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'Pelanggaran');
    XLSX.writeFile(wb, EXCEL_FILE);
  }
}

initData();

function readData() {
  try {
    return JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8'));
  } catch {
    return {};
  }
}

function saveData(data) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));
}

function deleteOldFile(filePath) {
  if (!filePath) return;
  const fullPath = path.join('/tmp', filePath);
  if (fs.existsSync(fullPath)) {
    try { fs.unlinkSync(fullPath); } catch(e) { /* abaikan error */ }
  }
}

// ============ PUBLIC API ROUTES ============
app.get('/api/data', (req, res) => {
  const data = readData();
  const { pengaturan, ...publicData } = data;
  const { adminPassword, ...publicPengaturan } = pengaturan || {};
  res.json({ ...publicData, pengaturan: publicPengaturan });
});

app.get('/api/berita', (req, res) => {
  const data = readData();
  res.json(data.berita || []);
});

app.get('/api/kegiatan', (req, res) => {
  const data = readData();
  res.json(data.kegiatan || []);
});

app.get('/api/pengurus', (req, res) => {
  const data = readData();
  res.json(data.pengurus || []);
});

app.get('/api/pengaturan', (req, res) => {
  const data = readData();
  const { adminPassword, ...pub } = data.pengaturan || {};
  res.json(pub);
});

app.post('/api/pelanggaran', (req, res) => {
  const { nama, kelas } = req.body;
  if (!nama || !kelas) return res.status(400).json({ error: 'Nama dan kelas wajib diisi' });
  try {
    if (!fs.existsSync(EXCEL_FILE)) return res.json({ pelanggaran: [], totalPoin: 0 });
    const wb = XLSX.readFile(EXCEL_FILE);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws);
    const hasil = rows.filter(row => {
      const rowNama = (row['Nama'] || '').toLowerCase().trim();
      const rowKelas = (row['Kelas'] || '').toLowerCase().replace(/\s/g, '');
      const cariNama = nama.toLowerCase().trim();
      const cariKelas = kelas.toLowerCase().replace(/\s/g, '');
      return (rowNama.includes(cariNama) || cariNama.includes(rowNama)) && rowKelas === cariKelas;
    });
    const totalPoin = hasil.reduce((sum, r) => sum + (parseInt(r['Poin']) || 0), 0);
    res.json({ pelanggaran: hasil, totalPoin });
  } catch (err) {
    res.status(500).json({ error: 'Gagal membaca data pelanggaran' });
  }
});

// ============ ADMIN: LOGIN & AUTH ============
app.post('/api/admin/login', (req, res) => {
  const { password } = req.body;
  const data = readData();
  if (password === data.pengaturan.adminPassword) {
    res.json({ success: true, token: Buffer.from(password + ':mpk2024').toString('base64') });
  } else {
    res.status(401).json({ error: 'Password salah' });
  }
});

function authMiddleware(req, res, next) {
  const token = req.headers['x-auth-token'];
  const data = readData();
  const expected = Buffer.from(data.pengaturan.adminPassword + ':mpk2024').toString('base64');
  if (token === expected) return next();
  res.status(401).json({ error: 'Unauthorized' });
}

app.get('/api/admin/data', authMiddleware, (req, res) => {
  res.json(readData());
});

// ============ ADMIN: BERITA CRUD ============
app.post('/api/admin/berita', authMiddleware, upload.single('foto_berita'), (req, res) => {
  const data = readData();
  const gambar = req.file ? '/uploads/' + req.file.filename : (req.body.gambar || '');
  const newItem = {
    id: Date.now(),
    judul: req.body.judul,
    kategori: req.body.kategori || 'berita',
    tanggal: req.body.tanggal || new Date().toISOString().split('T')[0],
    isi: req.body.isi,
    gambar: gambar
  };
  data.berita.unshift(newItem);
  saveData(data);
  res.json(newItem);
});

app.put('/api/admin/berita/:id', authMiddleware, upload.single('foto_berita'), (req, res) => {
  const data = readData();
  const idx = data.berita.findIndex(b => b.id == req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Tidak ditemukan' });

  let gambar = data.berita[idx].gambar;
  if (req.file) {
    deleteOldFile(gambar);
    gambar = '/uploads/' + req.file.filename;
  } else if (req.body.hapus_gambar === 'true') {
    deleteOldFile(gambar);
    gambar = '';
  } else if (req.body.gambar !== undefined) {
    gambar = req.body.gambar;
  }

  data.berita[idx] = {
    ...data.berita[idx],
    judul: req.body.judul || data.berita[idx].judul,
    kategori: req.body.kategori || data.berita[idx].kategori,
    tanggal: req.body.tanggal || data.berita[idx].tanggal,
    isi: req.body.isi || data.berita[idx].isi,
    gambar: gambar
  };
  saveData(data);
  res.json(data.berita[idx]);
});

app.delete('/api/admin/berita/:id', authMiddleware, (req, res) => {
  const data = readData();
  const berita = data.berita.find(b => b.id == req.params.id);
  if (berita && berita.gambar) deleteOldFile(berita.gambar);
  data.berita = data.berita.filter(b => b.id != req.params.id);
  saveData(data);
  res.json({ success: true });
});

// ============ ADMIN: KEGIATAN CRUD ============
app.post('/api/admin/kegiatan', authMiddleware, upload.single('foto_kegiatan'), (req, res) => {
  const data = readData();
  const gambar = req.file ? '/uploads/' + req.file.filename : (req.body.gambar || '');
  const newItem = {
    id: Date.now(),
    nama: req.body.nama,
    tanggal: req.body.tanggal,
    waktu: req.body.waktu || '',
    tempat: req.body.tempat || '',
    deskripsi: req.body.deskripsi || '',
    status: req.body.status || 'upcoming',
    gambar: gambar
  };
  data.kegiatan.push(newItem);
  saveData(data);
  res.json(newItem);
});

app.put('/api/admin/kegiatan/:id', authMiddleware, upload.single('foto_kegiatan'), (req, res) => {
  const data = readData();
  const idx = data.kegiatan.findIndex(k => k.id == req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Tidak ditemukan' });

  let gambar = data.kegiatan[idx].gambar;
  if (req.file) {
    deleteOldFile(gambar);
    gambar = '/uploads/' + req.file.filename;
  } else if (req.body.hapus_gambar === 'true') {
    deleteOldFile(gambar);
    gambar = '';
  } else if (req.body.gambar !== undefined) {
    gambar = req.body.gambar;
  }

  data.kegiatan[idx] = {
    ...data.kegiatan[idx],
    nama: req.body.nama || data.kegiatan[idx].nama,
    tanggal: req.body.tanggal || data.kegiatan[idx].tanggal,
    waktu: req.body.waktu !== undefined ? req.body.waktu : data.kegiatan[idx].waktu,
    tempat: req.body.tempat !== undefined ? req.body.tempat : data.kegiatan[idx].tempat,
    deskripsi: req.body.deskripsi !== undefined ? req.body.deskripsi : data.kegiatan[idx].deskripsi,
    status: req.body.status || data.kegiatan[idx].status,
    gambar: gambar
  };
  saveData(data);
  res.json(data.kegiatan[idx]);
});

app.delete('/api/admin/kegiatan/:id', authMiddleware, (req, res) => {
  const data = readData();
  const kegiatan = data.kegiatan.find(k => k.id == req.params.id);
  if (kegiatan && kegiatan.gambar) deleteOldFile(kegiatan.gambar);
  data.kegiatan = data.kegiatan.filter(k => k.id != req.params.id);
  saveData(data);
  res.json({ success: true });
});

// ============ ADMIN: PENGURUS CRUD ============
app.post('/api/admin/pengurus', authMiddleware, upload.single('foto_pengurus'), (req, res) => {
  const data = readData();
  const foto = req.file ? '/uploads/' + req.file.filename : (req.body.foto || '');
  const newItem = {
    id: Date.now(),
    nama: req.body.nama,
    jabatan: req.body.jabatan,
    kelas: req.body.kelas || '',
    instagram: req.body.instagram || '',
    foto: foto
  };
  data.pengurus.push(newItem);
  saveData(data);
  res.json(newItem);
});

app.put('/api/admin/pengurus/:id', authMiddleware, upload.single('foto_pengurus'), (req, res) => {
  const data = readData();
  const idx = data.pengurus.findIndex(p => p.id == req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Tidak ditemukan' });

  let foto = data.pengurus[idx].foto;
  if (req.file) {
    deleteOldFile(foto);
    foto = '/uploads/' + req.file.filename;
  } else if (req.body.hapus_foto === 'true') {
    deleteOldFile(foto);
    foto = '';
  } else if (req.body.foto !== undefined) {
    foto = req.body.foto;
  }

  data.pengurus[idx] = {
    ...data.pengurus[idx],
    nama: req.body.nama || data.pengurus[idx].nama,
    jabatan: req.body.jabatan || data.pengurus[idx].jabatan,
    kelas: req.body.kelas !== undefined ? req.body.kelas : data.pengurus[idx].kelas,
    instagram: req.body.instagram !== undefined ? req.body.instagram : data.pengurus[idx].instagram,
    foto: foto
  };
  saveData(data);
  res.json(data.pengurus[idx]);
});

app.delete('/api/admin/pengurus/:id', authMiddleware, (req, res) => {
  const data = readData();
  const pengurus = data.pengurus.find(p => p.id == req.params.id);
  if (pengurus && pengurus.foto) deleteOldFile(pengurus.foto);
  data.pengurus = data.pengurus.filter(p => p.id != req.params.id);
  saveData(data);
  res.json({ success: true });
});

// ============ ADMIN: PENGATURAN ============
app.put('/api/admin/pengaturan', authMiddleware, (req, res) => {
  const data = readData();
  data.pengaturan = { ...data.pengaturan, ...req.body };
  saveData(data);
  const { adminPassword, ...pub } = data.pengaturan;
  res.json(pub);
});

// ============ ADMIN: UPLOAD EXCEL ============
app.post('/api/admin/upload-excel', authMiddleware, uploadExcel.single('excel'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'File tidak ditemukan' });
  res.json({ success: true, message: 'File Excel berhasil diupload ke server temporary' });
});

// ============ ADMIN: UPLOAD LOGO ============
app.post('/api/admin/upload-logo', authMiddleware, upload.single('logo'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'File tidak ditemukan' });
  const ext = path.extname(req.file.originalname);
  const data = readData();
  data.pengaturan.logoPath = '/uploads/' + req.file.filename;
  saveData(data);
  res.json({ success: true, logoPath: data.pengaturan.logoPath });
});

// ============ DOWNLOAD TEMPLATE EXCEL ============
app.get('/api/admin/template-excel', authMiddleware, (req, res) => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['Nama', 'Kelas', 'Tanggal', 'Pelanggaran', 'Poin', 'Keterangan'],
    ['Contoh Nama', 'X IPA 1', '2024-01-15', 'Terlambat masuk kelas', 5, 'Keterangan opsional']
  ]);
  XLSX.utils.book_append_sheet(wb, ws, 'Pelanggaran');
  const tmpFile = path.join('/tmp', 'template_pelanggaran.xlsx');
  XLSX.writeFile(wb, tmpFile);
  res.download(tmpFile, 'template_pelanggaran.xlsx');
});

// ============ ROUTING HTML ============
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
app.get('/admin', (req, res) => res.sendFile(path.join(__dirname, 'public', 'admin.html')));

// ============ EXPORT APP FOR VERCEL (Paling Penting!) ============
module.exports = app;
