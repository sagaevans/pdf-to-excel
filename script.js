// Konfigurasi PDF.js untuk versi 2.16.105
if (window["pdfjsLib"]) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js";
} else {
  console.error("PDF.js gagal dimuat.");
}

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const fileInfo = document.getElementById("fileInfo");
const convertBtn = document.getElementById("convertBtn");
const statusEl = document.getElementById("status");

let currentFile = null;
let pdfArrayBuffer = null;

function setStatus(text) {
  statusEl.textContent = text;
}

// ====== PILIH FILE ======

function handleFile(file) {
  if (!file) return;

  if (file.type !== "application/pdf") {
    setStatus("❌ File bukan PDF. Silakan pilih file .pdf");
    fileInfo.textContent = "Belum ada file yang valid.";
    convertBtn.disabled = true;
    return;
  }

  currentFile = file;
  fileInfo.textContent = `File terpilih: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  setStatus("Membaca file...");

  const reader = new FileReader();
  reader.onload = function (e) {
    pdfArrayBuffer = e.target.result;
    convertBtn.disabled = false;
    setStatus("File siap dikonversi. Klik Convert ke Excel.");
  };
  reader.onerror = function () {
    setStatus("❌ Gagal membaca file.");
  };
  reader.readAsArrayBuffer(file);
}

dropzone.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", (e) => handleFile(e.target.files[0]));

dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropzone.style.background = "#e0edff";
});
dropzone.addEventListener("dragleave", () => {
  dropzone.style.background = "#f9fbff";
});
dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropzone.style.background = "#f9fbff";
  handleFile(e.dataTransfer.files[0]);
});

// ====== FUNGSI: UBAH STRING ANGKA → NUMBER ======

function toNumberIfNumeric(str) {
  if (typeof str !== "string") return str;
  const s = str.replace(/\s/g, ""); // buang spasi

  if (!s) return str;

  // Jangan ubah kalau dia ID yang punya nol di depan (misal teller id 0374620)
  if (/^0\d+$/.test(s)) return str;

  // Pola angka US: 12,345,678.90
  const usPattern = /^\d{1,3}(,\d{3})*(\.\d+)?$/;
  // Pola angka EU/ID: 12.345.678,90
  const euPattern = /^\d{1,3}(\.\d{3})*(,\d+)?$/;
  // Pola sederhana: 1234,56 atau 1234.56
  const simplePattern = /^-?\d+([.,]\d+)?$/;

  let normalized;
  let num;

  if (usPattern.test(s)) {
    normalized = s.replace(/,/g, ""); // buang pemisah ribuan
    num = Number(normalized);
    return isNaN(num) ? str : num;
  }

  if (euPattern.test(s)) {
    normalized = s.replace(/\./g, "").replace(",", "."); // 12.345,67 -> 12345.67
    num = Number(normalized);
    return isNaN(num) ? str : num;
  }

  if (simplePattern.test(s)) {
    normalized = s.replace(",", "."); // 1234,56 -> 1234.56
    num = Number(normalized);
    return isNaN(num) ? str : num;
  }

  return str;
}

// ====== KONVERSI PDF → EXCEL ======

convertBtn.addEventListener("click", async () => {
  if (!pdfArrayBuffer || !currentFile) {
    setStatus("Tidak ada file untuk dikonversi.");
    return;
  }

  try {
    convertBtn.disabled = true;
    setStatus("Memuat PDF...");

    const loadingTask = pdfjsLib.getDocument({
      data: pdfArrayBuffer
    });
    const pdf = await loadingTask.promise;

    let allRows = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      setStatus(`Memproses halaman ${pageNum} dari ${pdf.numPages}...`);

      const page = await pdf.getPage(pageNum);
      const content = await page.getTextContent();
      const items = content.items;

      // Grup per baris berdasarkan koordinat Y
      let lines = [];

      items.forEach((item) => {
        const x = item.transform[4];
        const y = item.transform[5];

        let line = lines.find((l) => Math.abs(l.y - y) < 2);
        if (!line) {
          line = { y, cells: [] };
          lines.push(line);
        }
        line.cells.push({ x, str: item.str });
      });

      // Urutkan baris: atas → bawah
      lines.sort((a, b) => b.y - a.y);

      lines.forEach((line) => {
        // Urutkan teks dalam baris: kiri → kanan
        line.cells.sort((a, b) => a.x - b.x);

        // Gabung dengan 1 spasi
        const joined = line.cells.map((c) => c.str.trim()).join(" ");

        // Pecah kolom dengan 2+ spasi (jadi "DATE  TIME  REMARK  DEBET" dst)
        const cols = joined
          .split(/\s{2,}/)
          .map((c) => c.trim())
          .filter((c, idx, arr) => !(c === "" && idx === arr.length - 1));

        if (cols.some((c) => c)) allRows.push(cols);
      });
    }

    if (allRows.length === 0) {
      setStatus("❌ Tidak ada teks yang bisa diambil. Mungkin PDF berupa scan/gambar.");
      convertBtn.disabled = false;
      return;
    }

    // Samakan jumlah kolom di semua baris
    const maxCols = Math.max(...allRows.map((r) => r.length));
    allRows = allRows.map((r) => {
      while (r.length < maxCols) r.push("");
      return r;
    });

    // UBAH TEXT ANGKA → NUMBER (supaya di Excel muncul format lokal, misal 12.295.574.073,33)
    allRows = allRows.map((row) => row.map(toNumberIfNumeric));

    // Buat workbook Excel
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(allRows);
    XLSX.utils.book_append_sheet(wb, ws, "Hasil");

    const outName = currentFile.name.replace(/\.pdf$/i, "") + "_converted.xlsx";
    XLSX.writeFile(wb, outName);

    setStatus("✅ Selesai! File Excel telah diunduh (angka sudah dalam format numerik).");
  } catch (err) {
    console.error(err);
    setStatus("❌ Gagal mengonversi: " + (err.message || err));
  } finally {
    convertBtn.disabled = false;
  }
});
