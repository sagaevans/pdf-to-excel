// Atur worker PDF.js (versi stabil 2.16)
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

function handleFile(file) {
  if (!file) return;
  if (file.type !== "application/pdf") {
    setStatus("File bukan PDF!");
    return;
  }
  currentFile = file;
  fileInfo.textContent = `File terpilih: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  setStatus("Membaca file...");
  const reader = new FileReader();
  reader.onload = function (e) {
    pdfArrayBuffer = e.target.result;
    convertBtn.disabled = false;
    setStatus("File siap dikonversi. Klik tombol Convert ke Excel.");
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

convertBtn.addEventListener("click", async () => {
  if (!pdfArrayBuffer) return setStatus("Tidak ada file PDF untuk diproses.");
  convertBtn.disabled = true;
  setStatus("Sedang memproses PDF...");

  try {
    const pdf = await pdfjsLib.getDocument({ data: pdfArrayBuffer }).promise;
    let allRows = [];

    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const lines = {};

      content.items.forEach((item) => {
        const y = Math.round(item.transform[5]);
        if (!lines[y]) lines[y] = [];
        lines[y].push({ x: item.transform[4], str: item.str });
      });

      const sortedY = Object.keys(lines).sort((a, b) => b - a);
      sortedY.forEach((y) => {
        const line = lines[y].sort((a, b) => a.x - b.x);
        const joined = line.map((c) => c.str.trim()).join(" ");
        const cols = joined.split(/\s{2,}/).map((c) => c.trim());
        if (cols.some((c) => c)) allRows.push(cols);
      });
    }

    if (allRows.length === 0) throw new Error("PDF kosong atau berupa gambar.");

    const maxCols = Math.max(...allRows.map((r) => r.length));
    allRows = allRows.map((r) => {
      while (r.length < maxCols) r.push("");
      return r;
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(allRows);
    XLSX.utils.book_append_sheet(wb, ws, "Hasil");
    const outName = currentFile.name.replace(/\.pdf$/i, "") + "_converted.xlsx";
    XLSX.writeFile(wb, outName);

    setStatus("✅ Selesai! File Excel telah diunduh.");
  } catch (err) {
    console.error(err);
    setStatus("❌ Gagal mengonversi. Mungkin file berupa gambar atau format tidak dikenali.");
  } finally {
    convertBtn.disabled = false;
  }
});
