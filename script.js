// Atur worker untuk PDF.js
if (window["pdfjsLib"]) {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.0.379/pdf.worker.min.js";
}

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const fileInfo = document.getElementById("fileInfo");
const convertBtn = document.getElementById("convertBtn");
const statusEl = document.getElementById("status");

let currentFile = null;
let pdfArrayBuffer = null;

// === Helper: tampilkan status ===
function setStatus(msg) {
  statusEl.textContent = msg;
}

// === Handler pilih file ===
function handleFile(file) {
  if (!file) return;

  if (file.type !== "application/pdf") {
    setStatus("File bukan PDF. Silakan pilih file dengan ekstensi .pdf");
    fileInfo.textContent = "Belum ada file yang valid.";
    convertBtn.disabled = false; // user bisa ganti file
    return;
  }

  currentFile = file;
  fileInfo.textContent = `File terpilih: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  setStatus("Membaca file PDF...");

  const reader = new FileReader();
  reader.onload = function (ev) {
    pdfArrayBuffer = ev.target.result;
    convertBtn.disabled = false;
    setStatus('Siap dikonversi. Klik "Convert ke Excel".');
  };
  reader.onerror = function () {
    setStatus("Gagal membaca file. Coba ulangi.");
  };
  reader.readAsArrayBuffer(file);
}

// === Klik dropzone untuk buka file explorer ===
dropzone.addEventListener("click", () => {
  fileInput.click();
});

fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  handleFile(file);
});

// === Drag & drop ===
dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropzone.classList.add("dragover");
});

dropzone.addEventListener("dragleave", (e) => {
  e.preventDefault();
  dropzone.classList.remove("dragover");
});

dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropzone.classList.remove("dragover");
  const file = e.dataTransfer.files[0];
  handleFile(file);
});

// === Convert ke Excel ===
convertBtn.addEventListener("click", async () => {
  if (!pdfArrayBuffer || !currentFile) {
    setStatus("Belum ada file PDF yang siap dikonversi.");
    return;
  }

  convertBtn.disabled = true;
  setStatus("Memuat dan memproses PDF...");

  try {
    const loadingTask = pdfjsLib.getDocument({ data: pdfArrayBuffer });
    const pdf = await loadingTask.promise;

    let allRows = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      setStatus(`Memproses halaman ${pageNum} dari ${pdf.numPages}...`);

      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();
      const items = textContent.items;

      // Kumpulkan teks per baris (kelompok berdasarkan koordinat Y)
      let lines = [];

      items.forEach((item) => {
        const tx = item.transform;
        const x = tx[4];
        const y = tx[5];

        // cari baris existing dengan y yang hampir sama
        let line = lines.find((l) => Math.abs(l.y - y) < 2);
        if (!line) {
          line = { y, cells: [] };
          lines.push(line);
        }
        line.cells.push({ x, str: item.str });
      });

      // Urutkan baris dari atas ke bawah (y besar ke kecil)
      lines.sort((a, b) => b.y - a.y);

      // Setiap baris: urutkan teks berdasar X (kiri ke kanan), gabung, lalu split jadi kolom
      lines.forEach((line) => {
        line.cells.sort((a, b) => a.x - b.x);
        const joined = line.cells.map((c) => c.str.trim()).join(" ");

        // Pisah kolom berdasarkan 2+ spasi (heuristic)
        const cols = joined.split(/\s{2,}/).map((c) => c.trim());

        if (cols.some((c) => c !== "")) {
          allRows.push(cols);
        }
      });
    }

    if (allRows.length === 0) {
      setStatus("Tidak berhasil mengambil teks dari PDF. Mungkin ini PDF scan / gambar.");
      convertBtn.disabled = false;
      return;
    }

    // Samakan jumlah kolom (pad dengan string kosong)
    const maxCols = allRows.reduce((max, row) => Math.max(max, row.length), 0);
    allRows = allRows.map((row) => {
      while (row.length < maxCols) row.push("");
      return row;
    });

    setStatus("Menyusun file Excel...");

    // Buat workbook Excel
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(allRows);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    // Nama file: namaPDF.xlsx
    const baseName = currentFile.name.replace(/\.pdf$/i, "");
    const outName = baseName + "_konversi.xlsx";

    XLSX.writeFile(wb, outName);
    setStatus("Selesai! File Excel telah diunduh.");
  } catch (err) {
    console.error(err);
    setStatus("Terjadi kesalahan saat mengonversi PDF. Detail di console.");
  } finally {
    convertBtn.disabled = false;
  }
});
