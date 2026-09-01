// ID Spreadsheet db-damsurstore
const SPREADSHEET_ID = '1Vj_fmOaROpi8DYrlo_EaQIDn1I4nDX4381EYN3T5-Os';

function doGet() {
  var htmlOutput = HtmlService.createTemplateFromFile('Index').evaluate();
  htmlOutput.setTitle('DAMSUR STORE');
  htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return htmlOutput;
}

// Fungsi Ambil Data Toko (Aman dari Serialization Error)
function getStoreData() {
  var response = {
    success: true,
    produk: [],
    profil: {},
    terlaris: {}
  };

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. BACA PROFIL
    var sheetProfil = ss.getSheetByName('Profil');
    if (sheetProfil) {
      var dataProfil = sheetProfil.getDataRange().getDisplayValues(); // Gunakan getDisplayValues agar murni String
      if (dataProfil.length > 1) {
        var headersP = dataProfil[0];
        var rowP = dataProfil[1];
        headersP.forEach(function(h, i) {
          if (h) response.profil[h.trim()] = rowP[i] || '';
        });
      }
    }

    // 2. BACA PRODUK
    var sheetProduk = ss.getSheetByName('Produk');
    if (sheetProduk) {
      var dataProduk = sheetProduk.getDataRange().getDisplayValues(); // Gunakan getDisplayValues
      if (dataProduk.length > 1) {
        var headersPr = dataProduk[0];
        for (var r = 1; r < dataProduk.length; r++) {
          var row = dataProduk[r];
          var obj = {};
          var isValid = false;
          headersPr.forEach(function(h, i) {
            var key = h.trim();
            var val = row[i] ? row[i].toString().trim() : '';
            obj[key] = val;
            if (key === 'nama_produk' && val !== '') isValid = true;
          });
          if (isValid) {
            response.produk.push(obj);
          }
        }
      }
    }

    // 3. BACA RIWAYAT DetailOrder UNTUK MENGHITUNG PRODUK TERLARIS
    // Struktur DetailOrder yang digunakan saveOrderToSheet:
    // [orderId, nama_produk, qty, harga, subtotal]
    var sheetDetail = ss.getSheetByName('DetailOrder');
    if (sheetDetail) {
      var dataDetail = sheetDetail.getDataRange().getDisplayValues();
      if (dataDetail.length > 1) {
        for (var d = 1; d < dataDetail.length; d++) {
          var namaDetail = (dataDetail[d][1] || '').toString().trim();
          var qtyDetail = parseNumber_(dataDetail[d][2]);
          if (namaDetail && qtyDetail > 0) {
            response.terlaris[namaDetail] = (response.terlaris[namaDetail] || 0) + qtyDetail;
          }
        }
      }
    }

  } catch (err) {
    Logger.log("Error Buka Sheet: " + err.toString());
  }

  // Tempel jumlah terjual ke setiap produk agar frontend dapat mengurutkannya.
  response.produk.forEach(function(prod) {
    prod.jumlah_terjual = response.terlaris[prod.nama_produk] || 0;
  });

  // JIKA SPREADSHEET KOSONG/GAGAL, GUNAKAN DATA DUMMY AGAR TAMPILAN TIDAK INFINITE LOADING
  if (response.produk.length === 0) {
    response.produk = [
      {
        id: "P001",
        nama_produk: "Beras Premium Super 5kg",
        satuan: "Karung",
        deskripsi_produk: "Beras pulen kualita super bebas pemutih",
        harga_normal: "75000",
        harga_promo: "68000",
        kategori: "Sembako",
        jumlah_terjual: 0,
        link_image: "https://images.unsplash.com/photo-1586201375761-83865001e31c?w=400"
      },
      {
        id: "P002",
        nama_produk: "Minyak Goreng Sawit 2 Litri",
        satuan: "Pouch",
        deskripsi_produk: "Minyak goreng kelapa sawit murni jernih",
        harga_normal: "38000",
        harga_promo: "34000",
        kategori: "Sembako",
        jumlah_terjual: 0,
        link_image: "https://images.unsplash.com/photo-1474979266404-7eaacbcd87c5?w=400"
      },
      {
        id: "P003",
        nama_produk: "Gula Pasir Putih 1kg",
        satuan: "Kg",
        deskripsi_produk: "Gula murni manis alami konsumsi harian",
        harga_normal: "17500",
        harga_promo: "16000",
        kategori: "Sembako",
        jumlah_terjual: 0,
        link_image: "https://images.unsplash.com/photo-1622484210800-8851b576f9d2?w=400"
      },
      {
        id: "P004",
        nama_produk: "Kopi Hitam Bubuk 200g",
        satuan: "Pack",
        deskripsi_produk: "Kopi olahan khas aroma wangi mantap",
        harga_normal: "25000",
        harga_promo: "22000",
        kategori: "Minuman",
        jumlah_terjual: 0,
        link_image: "https://images.unsplash.com/photo-1559056199-641a0ac8b55e?w=400"
      }
    ];
  }

  if (!response.profil.nama_toko) {
    response.profil = {
      nama_toko: "DAMSUR STORE",
      motto_slogan: "Toko Online Modern Berbasis Checkout Whatsapp Lebih Praktis dan Mudah",
      whatsapp: "081917061753"
    };
  }

  // Return sebagai JSON String agar 100% lolos serialisasi RPC
  return JSON.stringify(response);
}

// Mengubah nilai angka dari Spreadsheet menjadi Number.
// Mendukung format seperti 100000, 100.000, Rp100.000, dll.
function parseNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  var text = value.toString().trim().replace(/[^0-9,-]/g, '');
  text = text.replace(/,/g, '.');
  var num = parseFloat(text);
  return isNaN(num) ? 0 : num;
}

// Fungsi Simpan Pesanan ke Spreadsheet
function saveOrderToSheet(jsonPayload) {
  try {
    var data = JSON.parse(jsonPayload);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var orderId = 'ORD-' + Date.now();
    var timestamp = new Date().toLocaleString('id-ID');

    // 1. Simpan Order
    var sheetOrder = ss.getSheetByName('Order');
    if (sheetOrder) {
      sheetOrder.appendRow([
        orderId,
        timestamp,
        data.nama,
        data.alamat,
        data.no_hp,
        data.total
      ]);
    }

    // 2. Simpan DetailOrder
    var sheetDetail = ss.getSheetByName('DetailOrder');
    if (sheetDetail) {
      data.items.forEach(function(item) {
        sheetDetail.appendRow([
          orderId,
          item.nama,
          item.qty,
          item.harga,
          item.subtotal
        ]);
      });
    }

    return JSON.stringify({ success: true, orderId: orderId });
  } catch (err) {
    return JSON.stringify({ success: false, error: err.toString() });
  }
}