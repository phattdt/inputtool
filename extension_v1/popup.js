import * as pdfjsLib from "./pdfjs/pdf.mjs";
import { parsePdfTextToPayload } from "./parse.js";

const fileInput = document.getElementById("file");
const log = document.getElementById("log");
const runBtn = document.getElementById("run");

const progressContainer = document.getElementById("progressContainer");
const progressBar = document.getElementById("progressBar");
const progressLabel = document.getElementById("progressLabel");

pdfjsLib.GlobalWorkerOptions.workerSrc = chrome.runtime.getURL("pdfjs/pdf.worker.mjs");

function println(msg) {
  log.value += msg + "\n";
  log.scrollTop = log.scrollHeight;
}

function updateProgress(percent, msg) {
  if (progressContainer) {
    progressContainer.style.display = "block";
    if (progressBar) progressBar.value = percent;
    if (progressLabel && msg) progressLabel.innerText = msg;
  }
}

async function extractTextFromPDF(file) {
  const data = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({ data });
  const pdf = await loadingTask.promise;

  let fullText = "";
  let useOCR = false;

  // First pass: Check if page 1 has text
  const page1 = await pdf.getPage(1);
  const content1 = await page1.getTextContent();
  const text1 = content1.items.map((it) => it.str).join(" ").trim();

  if (text1.length < 10) {
    println("⚠️ No text layer found (Scanned PDF?). Switching to OCR mode...");
    useOCR = true;
  }

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    updateProgress(Math.floor(((pageNum - 1) / pdf.numPages) * 100), `Processing Page ${pageNum}/${pdf.numPages}...`);

    const page = await pdf.getPage(pageNum);
    let pageText = "";

    if (useOCR) {
      // OCR Mode: Render to Canvas -> Tesseract
      const viewport = page.getViewport({ scale: 1.5 }); // Scale 1.5 for better OCR
      const canvas = document.createElement("canvas");
      const context = canvas.getContext("2d");
      canvas.height = viewport.height;
      canvas.width = viewport.width;

      await page.render({ canvasContext: context, viewport: viewport }).promise;

      const imageBlob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));

      println(`   Creating OCR worker for page ${pageNum}...`);

      // Initialize Tesseract Worker
      // MV3 requires disabling Blob URL generation so it uses "new Worker('path')" directly
      const worker = await Tesseract.createWorker("eng", 1, {
        workerPath: chrome.runtime.getURL("ocr_lib/worker.min.js"),
        corePath: chrome.runtime.getURL("ocr_lib/tesseract-core.wasm.js"),
        workerBlobURL: false,
        langPath: chrome.runtime.getURL("ocr_lib/"),
        gzip: false,
      });

      // Note: Getting core offline is hard without downloading it.
      // For now, let's assume online allowed or standard loading.
      // Or just standard usage:
      // const worker = await Tesseract.createWorker('eng');

      // Configure for Table/Text preservation (PSM 6 = Assume a single uniform block of text)
      // This forces reading line-by-line across the page, preventing column splitting.
      await worker.setParameters({
        tessedit_pageseg_mode: 6,
      });

      const ret = await worker.recognize(imageBlob);
      pageText = ret.data.text;
      await worker.terminate();

    } else {
      // Standard Text Mode
      const content = await page.getTextContent();
      pageText = content.items.map((it) => it.str).join(" ");
    }

    fullText += `\n\n--- PAGE ${pageNum} ---\n`;
    fullText += pageText;
  }

  updateProgress(100, "Done!");
  setTimeout(() => { if (progressContainer) progressContainer.style.display = "none"; }, 1000);

  return fullText;
}

async function getActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) throw new Error("No active tab.");
  return tab;
}

runBtn.addEventListener("click", async () => {
  const EXTENSION_START = performance.now();
  console.log('🟧 [PERF] Extension processing started');
  log.value = "";

  const file = fileInput.files?.[0];
  if (!file) {
    println("❌ Select a PDF first.");
    return;
  }

  try {
    println("1) Extracting PDF text locally…");
    const PDF_START = performance.now();
    const text = await extractTextFromPDF(file);
    const PDF_END = performance.now();
    println(`✅ Extracted ${text.length} chars in ${(PDF_END - PDF_START).toFixed(2)}ms`);

    // DEBUG: Dump text to log to check OCR quality
    println("--- START RAW TEXT ---");
    println(text.substring(0, 1000) + (text.length > 1000 ? "... (truncated)" : ""));
    println("--- END RAW TEXT ---");

    console.log(`📄 [PERF] PDF extraction: ${(PDF_END - PDF_START).toFixed(2)}ms`);

    println("2) Parsing text into payload JSON…");
    const PARSE_START = performance.now();
    const payload = parsePdfTextToPayload(text);
    const PARSE_END = performance.now();
    println(`✅ Parsed parts: ${payload.parts?.length ?? 0}`);
    println(`✅ Parsed details keys: ${Object.keys(payload.details ?? {}).length}`);
    println(`   Parse time: ${(PARSE_END - PARSE_START).toFixed(2)}ms`);
    console.log(`🔄 [PERF] Text parsing: ${(PARSE_END - PARSE_START).toFixed(2)}ms`);

    println("3) Injecting fill logic into the current tab…");
    const tab = await getActiveTab();

    // Inject the fill logic into MAIN world so the page can call it
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      world: "MAIN",
      files: ["injectFillLogic.js"]
    });

    println("4) Sending payload to the page (localStorage + call fill)…");
    const FILL_START = performance.now();

    // Send payload into the page and execute it
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      world: "MAIN",
      func: (payloadObj) => {
        // Save for the page’s existing hook
        localStorage.setItem("gproc_excel_data", JSON.stringify(payloadObj));

        // Fill immediately if logic exists
        if (window.gprocFillLogic) {
          window.gprocFillLogic(payloadObj);
        }

        // If there’s exactly one part, you can auto-open detail
        const first = payloadObj?.parts?.[0]?.drawingNumber;
      },
      args: [payload]
    });

    const FILL_END = performance.now();
    const EXTENSION_END = performance.now();
    const totalTime = (EXTENSION_END - EXTENSION_START).toFixed(2);

    println(`✅ Done. Check the Demo-Tool page.`);
    println(`📊 TOTAL TIME: ${totalTime}ms`);
    console.log(`✅ [PERF] Extension TOTAL TIME: ${totalTime}ms`);
    console.log('   Breakdown:');
    console.log(`     PDF extraction: ${(PDF_END - PDF_START).toFixed(2)}ms`);
    console.log(`     Text parsing: ${(PARSE_END - PARSE_START).toFixed(2)}ms`);
    console.log(`     Fill execution: ${(FILL_END - FILL_START).toFixed(2)}ms`);
  } catch (e) {
    console.error(e);
    println("❌ Error: " + (e?.message ?? String(e)));
  }
});

// Download log button handler
const downloadBtn = document.getElementById("downloadLog");
if (downloadBtn) {
  downloadBtn.addEventListener("click", () => {
    if (!log.value || log.value.trim() === "") {
      alert("⚠️ No log data to download. Run the extension first!");
      return;
    }

    // Build txt content
    let content = '='.repeat(80) + '\n';
    content += '  📊 EXTENSION PERFORMANCE LOG\n';
    content += '='.repeat(80) + '\n\n';
    content += `Generated: ${new Date().toLocaleString()}\n\n`;
    content += '-'.repeat(80) + '\n';
    content += '  Processing Log:\n';
    content += '-'.repeat(80) + '\n\n';
    content += log.value;
    content += '\n\n' + '='.repeat(80) + '\n';

    // Create and download file
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `extension_performance_log_${new Date().toISOString().replace(/[:.]/g, '-')}.txt`;
    a.click();
    URL.revokeObjectURL(url);

    println("📥 Performance log downloaded!");
  });
}
