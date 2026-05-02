// ===============================
// INVENTORY CONTROL SYSTEM
// Backend - Google Apps Script
// ===============================

const SPREADSHEET_ID = '1o1tg9db454ARo2kMnbTLjgAJO2_Pf9tL8NYFAHJxHp8';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('app')
    .setTitle('Inventory Control System');
}

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ===============================
// AMBIL DAFTAR BARANG
// ===============================
function getDaftarBarang() {
  const ss = getSS();
  const sheet = ss.getSheetByName('daftar barang');
  if (!sheet) throw new Error('Sheet daftar barang tidak ditemukan');

  const data = sheet.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < data.length; i++) {
    const kode = data[i][0];
    const nama = data[i][1];
    const stok = Number(data[i][2]);
    const saldo = Number(data[i][3]);

    if (kode && nama && saldo >= 0) { // tampilkan semua barang dengan saldo >=0
      result.push({
        kodeBarang: kode,
        namaBarang: nama,
        stok,
        sisaSaldo: saldo
      });
    }
  }
  return result;
}

// ===============================
// SIMPAN TRANSAKSI
// ===============================
function simpanTransaksi(payload) {
  const ss = getSS();
  const logSheet = ss.getSheetByName('LOG BARANG');
  const barangSheet = ss.getSheetByName('daftar barang');

  if (!logSheet || !barangSheet) {
    throw new Error('Sheet LOG BARANG atau daftar barang tidak ditemukan');
  }

  const tanggal = new Date();
  const dataBarang = barangSheet.getDataRange().getValues();

  payload.items.forEach(item => {
    // cari indeks barang di daftar barang
    let idx = dataBarang.findIndex(r => r[0] === item.kodeBarang);
    if (idx === -1) throw new Error(`Barang ${item.namaBarang} tidak ditemukan`);

    let stok = Number(dataBarang[idx][2]);
    let saldo = Number(dataBarang[idx][3]);
    const qty = Number(item.jumlah);

    // validasi saldo untuk transaksi keluar
    if (payload.jenisTransaksi === 'Keluar' && qty > saldo) {
      throw new Error(`Saldo ${item.namaBarang} tidak cukup!`);
    }

    // update stok & saldo
    if (payload.jenisTransaksi === 'Masuk') {
      saldo += qty;
      stok += qty;
    } else {
      saldo -= qty;
      stok -= qty;
    }

    // tulis kembali ke sheet
    barangSheet.getRange(idx + 1, 3).setValue(stok);
    barangSheet.getRange(idx + 1, 4).setValue(saldo);

    // catat ke log
    logSheet.appendRow([
      tanggal,
      payload.namaPegawai,
      payload.jenisTransaksi,
      item.kodeBarang,
      item.namaBarang,
      qty
    ]);
  });

  return true;
}
