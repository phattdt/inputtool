
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs';
import fs from 'fs';

// Load the PDF file
const pdfPath = './pdf_template.pdf';
const data = new Uint8Array(fs.readFileSync(pdfPath));

async function extractText() {
    const loadingTask = pdfjsLib.getDocument({ data: data });
    const pdf = await loadingTask.promise;
    console.log(`PDF loaded. Pages: ${pdf.numPages}`);

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        const strings = content.items.map(function (item) {
            return item.str;
        });
        console.log(`\n--- Page ${i} Content Start ---`);
        console.log(strings.join(" "));
        console.log(`--- Page ${i} Content End ---\n`);
    }
}

extractText().catch(console.error);
