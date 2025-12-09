// =======================
// KONFIG PDF.JS
// =======================

// Pastikan PDF.js sudah ada
if (!window["pdfjsLib"]) {
  alert("PDF.js gagal dimuat. Coba refresh halaman.");
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

// =======================
// HANDLE FILE
// =======================

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
  setStatus("Membaca file PDF...");

  const reader = new FileReader();
  reader.onload = function (ev) {
    pdfArrayBuffer = ev.target.result;
    convertBtn.disabled = false;
    setStatus("File siap dikonversi. Klik tombol Convert ke Excel.");
  };
  reader.onerror = function () {
    setStatus("❌ Gagal membaca file. Coba ulangi.");
  };
  reader.readAsArrayBuffer(file);
}

// klik dropzone → buka file explorer
dropzone.addEventListener("click", () => fileInput.click());

// pilih file manual
fileInput.addEventListener("change", (e) => {
  const file = e.target.files[0];
  handleFile(file);
});

// drag & drop
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
  const file = e.dataTransfer.files[0];
  handleFile(file);
});

// =======================
// KONVERSI PDF → EXCEL
// =======================

convertBtn.addEventListener("click", async () => {
  if (!pdfArrayBuffer || !currentFile) {
    setStatus("Tidak ada file PDF yang siap dikonversi.");
    return;
  }

  if (!window["pdfjsLib"]) {
    setStatus("❌ PDF.js tidak tersedia di halaman ini.");
    return;
  }

  convertBtn.disabled = true;
  setStatus("Memuat dan memproses PDF... (bisa agak lama untuk file besar)");

  try {
    // PAKSA TANPA WORKER → menghindari error di GitHub Pages
    const loadingTask = pdfjsLib.getDocument({
      data: pdfArrayBuffer,
      disableWorker: true
    });

    const pdf = await loadingTask.promise;
    let allRows = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      setStatus(`Memproses halaman ${pageNum} dari ${pdf.numPages}...`);

      const page = await pdf.getPage(pageNum);
      const textContent = await page.getTextContent();
      const items = textContent.items;

      // Kumpulkan teks per garis Y
      let lines = [];

      items.forEach((item) => {
        const tx = item.transform;
        const x = tx[4];
        const y = tx[5];

        let line = lines.find((l) => Math.abs(l.y - y) < 2);
        if (!line) {
          line = { y, cells: [] };
          lines.push(line);
        }
        line.cells.push({ x, str: item.str });
      });

      // urut baris dari atas ke bawah
      lines.sort((a, b) => b.y - a.y);

      // setiap baris → urut kiri-kanan dan split jadi kolom
      lines.forEach((line) => {
        line.cells.sort((a, b) => a.x - b.x);
        const joined = line.cells.map((c) => c.str.trim()).join(" ");

        // heuristic: 2+ spasi dianggap pemisah kolom
        const cols = joined.split(/\s{2,}/).map((c) => c.trim());

        if (cols.some((c) => c !== "")) {
          allRows.push(cols);
        }
      });
    }

    if (allRows.length === 0) {
      setStatus("❌ Tidak ada teks yang bisa diambil. Mungkin PDF berupa scan/gambar.");
      convertBtn.disabled = false;
      return;
    }

    // samakan jumlah kolom
    const maxCols = allRows.reduce((max, row) => Math.max(max, row.length), 0);
    allRows = allRows.map((row) => {
      while (row.length < maxCols) row.push("");
      return row;
    });

    setStatus("Menyusun file Excel...");

    // buat workbook Excel
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(allRows);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    const baseName = currentFile.name.replace(/\.pdf$/i, "");
    const outName = baseName + "_konversi.xlsx";

    XLSX.writeFile(wb, outName);
    setStatus("✅ Selesai! File Excel telah diunduh.");
  } catch (err) {
    console.error(err);
    setStatus("❌ Gagal mengonversi: " + (err.message || err.toString()));
  } finally {
    convertBtn.disabled = false;
  }
});
