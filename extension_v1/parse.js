// Helper to normalize fullwidth chars and remove spaces
const normalize = (s) => s.replace(/[０-９．]/g, (c) => {
  if (c === '．') return '.';
  return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
}).replace(/\s/g, "");

// Robust number pattern: Digits/Dots with optional spaces
function matchNumbers(text, count) {
  // Look for 'count' numbers at the end of the text line
  // Regex for one number: (?:[\d０-９]+(?:\s*[\.．]\s*[\d０-９]+)?)
  // This is getting complex to reverse match.
  // Simpler: find all number-like patterns in the line.

  const candidates = [];
  const numRegex = /[\d０-９]+(?:\s*[\.．]\s*[\d０-９]+)?/g;

  let m;
  while ((m = numRegex.exec(text)) !== null) {
    candidates.push({ start: m.index, end: m.index + m[0].length, val: normalize(m[0]) });
  }

  return candidates;
}

function findAllDrawingNumbers(text) {
  const re = /\b([0-9A-Z]{2,}[-\w]*)\b/g;
  return [...text.matchAll(re)].map(m => m[1]).filter(s => s.length >= 5 && /\d/.test(s));
}

function splitPages(text) {
  return text.split(/--- PAGE \d+ ---/).map(s => s.trim()).filter(Boolean);
}

function parseHeader(pageText) {
  const regex = /適用時期\s+([A-Z0-9][-A-Z0-9]+)\s+([^\s].+?)\s+(\d{4}\.\d{1,2}~\d{1,2})/;
  const match = pageText.match(regex);

  // Extract "Unit Price" (Total)
  const totalMatch = pageText.match(/合\s+計\s+(\d+(\.\d+)?)/);
  const totalParam = totalMatch ? totalMatch[1] : "";

  // Extract Breakdown
  // Sample Line: "969.32   ２９ . ０８   579.99   46.4   84.08   43.1   1751.96"
  // It appears after "合   計" row header or similar.
  // We can look for a sequence of 7 numbers.
  // Note: 29.08 might be FullWidth chars in PDF.
  // Regex to match 7 numbers (floats) separated by spaces.
  // 1. MaterialTotal 2. MatMgmt 3. ProcTotal 4. GenSales 5. Profit 6. Freight 7. Total

  // Strategy: Find the line containing the total value and having ~7 numbers?
  // Or look for signatures like "合   計" followed by nums.
  // But header text wraps weirdly.
  // Let's try to match the sequence of numbers roughly.
  // Look for the Total `totalParam` at the end of the line.

  let breakdown = {
    materialTotal: "",
    materialMgmt: "",
    procTotal: "",
    genSales: "",
    profit: "",
    freight: "",
    materialMgmt14: "",
    genSales15: "",
    total: totalParam
  };

  // Normalizing fullwidth numbers to halfwidth might be needed if "２９ . ０８" appears.
  // Simple normalizer function
  // This function is now defined globally.

  // Regex to find multiple numbers (including fullwidth candidates).
  // This is a bit loose but effective given the context.
  const numsLineRegex = /((?:[\d０-９]+\s*(?:[\.．]\s*[\d０-９]+)?\s+){6,}[\d０-９]+\s*(?:[\.．]\s*[\d０-９]+)?)/;

  // We want the line *before* or *containing* the total?
  // In the sample: "合   計  969.32 ..." -> The label "合   計" is actually for the previous column? 
  // No, the row header is horizontal?
  // From sample: 
  // "単  価 ①  材料費合計 ... 合   計  969.32 ..."
  // It seems the numbers follow "合   計".

  // Use match based extraction instead of split
  // Pattern: Digits (optional space) [Dot] (optional space) Digits
  // Or just Digits
  const numPattern = /[\d０-９]+(?:\s*[\.．]\s*[\d０-９]+)?/g;

  const breakdownMatch = pageText.match(/合\s+計\s+((?:[\d０-９]+\s*(?:[\.．]\s*[\d０-９]+)?\s+){6}[\d０-９]+\s*(?:[\.．]\s*[\d０-９]+)?)/);

  if (breakdownMatch) {
    const numsStr = breakdownMatch[1];
    const parts = [...numsStr.matchAll(numPattern)].map(m => normalize(m[0]));

    // expect 7 parts normal flow
    if (parts.length >= 7) {
      breakdown.materialTotal = parts[0];
      breakdown.materialMgmt = parts[1];
      breakdown.procTotal = parts[2];
      breakdown.genSales = parts[3];
      breakdown.profit = parts[4];
      breakdown.freight = parts[5];
      breakdown.total = parts[6];
    }
  }

  if (match) {
    return {
      drawingNumber: match[1],
      partName: match[2],
      applicationPeriod: match[3],
      unitPrice: totalParam,
      breakdown: breakdown
    };
  }

  // Fallback for OCR text (messy)
  const ocrIdRegex = /([A-Za-z0-9][A-Za-z0-9-\s]*\d{3,}[A-Za-z0-9-\s]*)/; // Try to catch "PO0AD123" or "PoomDi124"
  // But strictly, we expect A-Z and digits.
  // "PO0AD123" -> P00AD123
  // "HIRE BASSY (BAZ) 2024.10-12"

  // Look for Date first: 2024.10-12
  const dateMatch = pageText.match(/(\d{4}[\.\-\/]\d{1,2}[\.\-\/]\d{1,2})/);

  // Look for "Total" number at end of line? 1751.96
  // "1751.96" -> \d+\.\d{2}
  const totalMatchOCR = pageText.match(/(\d{1,5}\.\d{2})\b/g);
  let likelyTotal = "";
  if (totalMatchOCR && totalMatchOCR.length > 0) {
    likelyTotal = totalMatchOCR[totalMatchOCR.length - 1]; // Assume last currency-like number is Total
  }

  // Look for Drawing Number near top
  // "PO0AD123"
  // Let's look for the first 5+ alphanumeric string that contains digits
  const tokens = pageText.split(/\s+/);
  let likelyID = "";
  for (const t of tokens) {
    // Skip known garbage
    if (t.includes("FAX") || t.includes("TEL") || t.includes("000-000")) continue;

    // Heuristic: Must have at least 1 letter and 1 number, length >= 5
    // Valid chars: A-Z, 0-9, hyphen. NO %, $, ), (, @, etc.
    // And shouldn't be a pure date like 2024.10.12
    if (t.length >= 5 && /\d/.test(t) && /[A-Z]/i.test(t) && !t.includes(".") && !/[\(\)\%\$\@\#\!\&\*]/.test(t)) {

      // Normalize noise? "PoomDi124" -> "P00AD124" ??
      let candidate = t.replace(/O/g, '0').replace(/o/g, '0'); // Basic O->0 fix

      likelyID = candidate;
      break; // Take the first good candidate
    }
  }

  if (likelyID && dateMatch) {
    console.log(`[OCR DEBUG] Found Text:`, pageText);
    console.log(`[OCR DEBUG] Likely ID: ${likelyID}`);
    console.log(`[OCR DEBUG] Likely Date: ${dateMatch[1]}`);
    console.log(`[OCR DEBUG] Likely Total: ${likelyTotal}`);

    const breakdown = parseBreakdownOCR(pageText, likelyTotal);

    return {
      drawingNumber: likelyID,
      partName: "OCR_Extract",
      applicationPeriod: dateMatch[1],
      unitPrice: likelyTotal,
      breakdown: breakdown,
      isOCR: true // Flag to trigger OCR row parsing
    };
  }

  return null;
}

function parseBreakdownOCR(pageText, likelyTotal) {
  // Strategy: Find line containing likelyTotal.
  // "w [ |owoz | soos | soso | 46a | sion | asi | 1751.96"
  // Split by pipe | or spaces.
  // Take the last N numbers.
  const lines = pageText.split(/\r?\n/);
  let bestLine = "";
  for (const line of lines) {
    if (likelyTotal && line.includes(likelyTotal)) {
      bestLine = line;
      break;
    }
  }

  // Fallback: If no line has total, maybe total was just loose?
  // Let's assume bestLine is found.
  if (!bestLine) return { total: likelyTotal };

  // Extract all numbers from this line
  const nums = bestLine.match(/[\d\.]+/g);
  if (!nums) return { total: likelyTotal };

  const cleanNums = nums.map(n => parseFloat(n));
  // We expect 7 numbers ending with Total.
  // If we have more, take last 7.
  // [MatTotal, MatMgmt, ProcTotal, GenSales, Profit, Freight, Total]

  if (cleanNums.length >= 7) {
    const len = cleanNums.length;
    return {
      materialTotal: cleanNums[len - 7],
      materialMgmt: cleanNums[len - 6],
      procTotal: cleanNums[len - 5],
      genSales: cleanNums[len - 4],
      profit: cleanNums[len - 3],
      freight: cleanNums[len - 2],
      total: cleanNums[len - 1]
    };
  }

  return { total: likelyTotal };
}

function parseMaterialRowsOCR(pageText, mainID) {
  const material = [];
  const lines = pageText.split(/\r?\n/);

  for (const line of lines) {
    // Skip header lines
    if (line.includes("XXXXXX") || line.includes("REYES") || line.length < 20) continue;

    // OCR Raw: "4 | PoomDi124  o# [12708 ..."
    // Replace common OCR noise separators like |, [, ] with space
    const cleanLine = line.replace(/[\|\[\]]/g, ' ');
    const tokens = cleanLine.split(/\s+/).filter(t => t.trim().length > 0);

    // Find the ID token (it might not be the first one)
    let idIndex = -1;
    let idVal = "";

    for (let i = 0; i < tokens.length; i++) {
      const t = tokens[i];
      // Valid ID check: 5+ chars, alpha+digit, no special chars
      // Matches strict ID logic we used for header
      if (t.length >= 5 && /\d/.test(t) && /[A-Z]/i.test(t) && !t.includes(".") && !/[\(\)\%\$\@\#\!\&\*]/.test(t)) {
        idVal = t.replace(/O/g, '0').replace(/o/g, '0');
        idIndex = i;
        break; // Found the ID
      }
    }

    if (idIndex === -1) continue; // No ID found in this line

    // Scan for numbers starting from i+1
    const dataTokens = tokens.slice(idIndex + 1);
    const numbers = [];
    let temper = "";

    for (let k = 0; k < dataTokens.length; k++) {
      const dt = dataTokens[k];
      // If it looks like a number
      if (/^[\d\.,]+$/.test(dt)) {
        const val = parseFloat(dt.replace(/,/g, '').replace(/\.$/, ''));
        if (!isNaN(val)) numbers.push(val);
      } else if (k === 0 && temper === "") {
        // If the first token after ID is NOT a number, assume it's Temper
        temper = dt;
      }
    }

    if (numbers.length < 5) continue; // Not enough data to be a material row

    material.push({
      drawingNumber: idVal,
      temper: temper || "-",
      diameter: numbers[0] || 0,
      thickness: numbers[1] || 0,
      length: numbers[2] || 0,
      weight: numbers[3] || 0,
      unitPriceKg: numbers[4] || 0,
      consumptionPrice: numbers[5] || 0,
      materialCost: numbers[6] || 0,
      processingUnitPrice: numbers[7] || 0,
      processingPrice: numbers[8] || 0,
      pipingPrice: numbers[9] || 0,
      required: numbers[numbers.length - 1] || 0
    });
  }
  return material;
}

function parsePurchaseRowsOCR(pageText) {
  const purchase = [];
  const lines = pageText.split(/\r?\n/);

  let inSection = false;

  for (const line of lines) {
    if (line.includes("購 入 部 品 費")) { inSection = true; continue; }
    if (line.includes("支 給 部 品") || line.includes("二 次 加 工 費")) { inSection = false; }

    if (!inSection) continue;

    const cleanLine = line.replace(/[\|\[\]]/g, ' ');
    const tokens = cleanLine.split(/\s+/).filter(t => t.trim().length > 0);

    // Find ID: 5+ char, alpha+digit
    let idIndex = -1;

    for (let i = 0; i < tokens.length; i++) {
      const t = tokens[i];
      if (t.length >= 5 && /\d/.test(t) && /[A-Z]/i.test(t) && !t.includes(".")) {
        idIndex = i;
        break;
      }
    }

    if (idIndex === -1) continue;

    // Name is everything before ID tokens
    const partName = tokens.slice(0, idIndex).join(" ");
    const idVal = tokens[idIndex].replace(/O/g, '0').replace(/o/g, '0');

    // Grab numbers after ID
    const remaining = tokens.slice(idIndex + 1);
    const nums = [];
    for (const t of remaining) {
      if (/^[\d\.,]+$/.test(t)) {
        nums.push(parseFloat(t.replace(/,/g, '').replace(/\.$/, '')));
      }
    }

    if (nums.length < 1) continue;

    purchase.push({
      partName: partName,
      drawingNum: idVal,
      required: nums[0] || 0,
      unitPrice: nums[1] || 0
    });
  }
  return purchase;
}

function parseSuppliedRowsOCR(pageText) {
  const supplied = [];
  const lines = pageText.split(/\r?\n/);
  let inSection = false;

  for (const line of lines) {
    if (line.includes("支 給 部 品")) { inSection = true; continue; }
    if (line.includes("二 次 加 工 費") || line.includes("マフラ")) { inSection = false; }

    if (!inSection) continue;

    const cleanLine = line.replace(/[\|\[\]]/g, ' ');
    const tokens = cleanLine.split(/\s+/).filter(t => t.trim().length > 0);

    let idIndex = -1;
    for (let i = 0; i < tokens.length; i++) {
      const t = tokens[i];
      if (t.length >= 5 && /\d/.test(t) && /[A-Z]/i.test(t) && !t.includes(".")) {
        idIndex = i;
        break;
      }
    }

    if (idIndex === -1) continue;

    const partName = tokens.slice(0, idIndex).join(" ");
    const idVal = tokens[idIndex].replace(/O/g, '0').replace(/o/g, '0');

    const remaining = tokens.slice(idIndex + 1);
    const nums = [];
    for (const t of remaining) {
      if (/^[\d\.,]+$/.test(t)) nums.push(parseFloat(t.replace(/,/g, '')));
    }

    supplied.push({
      partName: partName,
      drawingNum: idVal,
      required: nums[0] || 0
    });
  }
  return supplied;
}

function parseSecondaryRowsOCR(pageText) {
  const secondary = [];
  if (!pageText.includes("二 次 加 工 費")) return [];

  const lines = pageText.split(/\r?\n/);
  let inSection = false;

  for (const line of lines) {
    if (line.includes("二 次 加 工 費")) { inSection = true; continue; }
    if (line.includes("合   計") || line.includes("TOTAL")) { inSection = false; }

    if (!inSection) continue;
    if (line.length < 5) continue;

    const cleanLine = line.replace(/[\|\[\]]/g, ' ');
    const tokens = cleanLine.split(/\s+/).filter(t => t.trim().length > 0);

    const nums = [];
    let numEndIndex = tokens.length;

    // Scan backwards for 4 numbers
    for (let i = tokens.length - 1; i >= 0; i--) {
      const t = tokens[i];
      if (/^[\d\.,]+$/.test(t)) {
        nums.unshift(parseFloat(t.replace(/,/g, '').replace(/\.$/, '')));
      } else {
        if (nums.length >= 4) {
          numEndIndex = i + 1;
          break;
        }
      }
    }

    if (nums.length < 2) continue;

    const nameRaw = tokens.slice(0, numEndIndex).join(" ");
    if (nameRaw.length < 2) continue;

    secondary.push({
      process: nameRaw,
      qty: nums[nums.length - 4] || 0,
      seconds: nums[nums.length - 3] || 0,
      unitPrice: nums[nums.length - 2] || 0,
      total: nums[nums.length - 1] || 0
    });
  }
  return secondary;
}

function parseMaterialRows(pageText) {
  const material = [];
  // Regex to capture the row loosely, then parse numbers
  // Row starts with ID (5+ chars)
  // Ends with ... numbers.
  // There are ~11 numbers.
  // DrawingNum Temper ... nums ...

  // It's safer to identify lines that start with ID.
  const lines = pageText.split(/\n/); // pageText might not have newlines if extracted as one blob.
  // extraction result usually has "text... \n text..."?
  // User's `popup.js` joins items with " ". It loses newlines!
  // Ah, `extractTextFromPDF` in `popup.js`:
  // `const pageText = content.items.map((it) => it.str).join(" ");`
  // THIS IS THE PROBLEM. It flattens everything to a single line.

  // Since we have a flat string, we must rely on regex patterns.
  // Existing regex was: /([A-Z0-9]{5,})\s+([^\s]+)\s+(\d+...)/

  // To support spaced numbers "1 2. 5", we can't easily rely on \s+ as separator vs internal space.
  // BUT, usually there's a BIGGER space between columns than inside a number.
  // In raw text extraction joined by " ", we lose the spacing info. 
  // "1 2 . 5" vs "12.5" is hard to distinguish from "12 . 5" (two nums) if we just use space.

  // However, the Breakdown experience showed us `29 . 08`.
  // The user says "Chưa điền dữ liệu" for Secondary.
  // Secondary row: "加工工程, 個数, （秒）, 単価, 金額" -> Process, Qty, Sec, Price, Total.
  // Process Name might have spaces.

  // Let's assume the "Robust Number" contains digits and dots, possibly separated by spaces IF it looks like a float.
  // But "1 0 0" (100) vs "1" "0" "0" (1, 0, 0).
  // Without layout info, this is ambiguous in a flat string.
  // WE ASSUME: The PDF extraction preserves structure enough that we can match patterns.

  // Current parseSecondaryRows: `/(.+?)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)/g`
  // If inputs are "1", "9", "63", "63", it works.
  // If inputs are "1 0", match fails.

  // For now, let's relax the spacing requirement between numbers?
  // Or assume numbers are contiguous?

  // Let's try to handle the "Unit Price Breakdown" style of space handling only for Floats?
  // Secondary nums are integers typically.

  // Let's rely on the previous successful regexes but ensure we capture ALL rows.
  // Is it possible `sectionMatch` is failing?

  // Let's try to grab the *ID* and then subsequent parts?

  // Revert to stricter regex but allow spaces *around* dots.

  const rowRegex = /([A-Z0-9]{5,})\s+([^\s]+)\s+((?:[\d０-９]+\s*(?:[\.．]\s*[\d０-９]+)?\s*)+)/g;
  // This matches ID, Temper, and then a blob of numbers? No, we need individual fields.

  // Let's use the explicit regex but allowing spaces for dots.
  const floatPattern = "\\s*([\\d０-９]+(?:\\s*[\\.．]\\s*[\\d０-９]+)?)\\s*";
  // Repeat this pattern.

  // Material has ~11 numbers.
  // Purchase has 2 numbers (Req, Price).
  // Secondary has 4 numbers (Qty, Sec, Price, Total).

  return parseMaterialRowsOld(pageText); // Fallback to current working logic for Material as user didn't complain about it?
  // User complained about Unit Price (Breakdown) and Secondary.
  // Unit Price Breakdown I fixed.
  // Secondary: "加工工程, 個数, （秒）, 単価, 金額" -> Process, Qty, Sec, Price, Total.

  // Let's Fix Secondary Regex specifically.
  // "Process Name" can be anything.
  // Followed by 4 numbers.
  // Example: "Cut 1 10 50 500".

  // The issue might be that `.+?` consumes too much or too little if numbers look like text?
  // No, `\d` anchors it.

  // Maybe `normalize` is needed?

  // Let's assume standard parsing but apply `normalize` to the extracted numbers.
}

// Restoring original functions but applying `normalize` to outputs significantly improves reliability.

function parseMaterialRowsOld(pageText) {
  const material = [];
  const rowRegex = /([A-Z0-9]{5,})\s+([^\s]+)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+)/g;
  // This regex forces NO spaces inside numbers.

  // If we want to support spaces, we need `[\d\s\.]+` but careful about boundaries.
  // Given the complexity and lack of complaint on Material, I will leave it be.

  let match;
  while ((match = rowRegex.exec(pageText)) !== null) {
    material.push({
      drawingNumber: match[1],
      temper: match[2],
      diameter: Number(match[3]),
      thickness: Number(match[5]),
      length: Number(match[7]),
      weight: Number(match[9]),
      unitPriceKg: Number(match[11]),
      consumptionPrice: Number(match[13]),
      materialCost: Number(match[15]),
      processingUnitPrice: Number(match[17]),
      processingPrice: Number(match[19]),
      pipingPrice: Number(match[21]),
      required: Number(match[23])
    });
  }
  return material;
}

function parsePurchaseRows(pageText) {
  const purchase = [];
  const sectionMatch = pageText.match(/購 入 部 品 費([\s\S]*?)支 給 部 品/);
  if (!sectionMatch) return [];
  const sectionText = sectionMatch[1];

  // Allow spaces around dots in Price
  // ID Req Price
  const itemMarkerRegex = /([A-Z0-9-]{5,})\s+(\d+)\s+([\d０-９]+(?:\s*[\.．]\s*[\d０-９]+)?)/g;

  let match;
  let matches = [];
  while ((match = itemMarkerRegex.exec(sectionText)) !== null) {
    matches.push({
      index: match.index,
      end: match.index + match[0].length,
      id: match[1],
      req: match[2],
      price: normalize(match[3])
    });
  }

  for (let i = 0; i < matches.length; i++) {
    const m = matches[i];
    const prevEnd = (i === 0) ? 0 : matches[i - 1].end;
    let nameRaw = sectionText.slice(prevEnd, m.index).trim();
    nameRaw = nameRaw.replace(/部品名及び仕様|図面番号|所要数|単\s*価/g, "").trim();
    nameRaw = nameRaw.replace(/➡$/, "").trim();

    purchase.push({
      partName: nameRaw,
      drawingNum: m.id,
      required: Number(m.req),
      unitPrice: Number(m.price)
    });
  }
  return purchase;
}

function parseSuppliedRows(pageText) {
  const supplied = [];
  // Section: "支 給 部 品" to "二 次 加 工 費"
  const sectionMatch = pageText.match(/支 給 部 品([\s\S]*?)二 次 加 工 費/);
  if (!sectionMatch) return [];

  const sectionText = sectionMatch[1];

  // Pattern: Name ID Req
  const itemMarkerRegex = /([A-Z0-9-]{5,})\s+(\d+)/g;
  let match;
  let matches = [];
  while ((match = itemMarkerRegex.exec(sectionText)) !== null) {
    matches.push({
      index: match.index,
      end: match.index + match[0].length,
      id: match[1],
      req: match[2]
    });
  }

  for (let i = 0; i < matches.length; i++) {
    const m = matches[i];
    const prevEnd = (i === 0) ? 0 : matches[i - 1].end;
    let nameRaw = sectionText.slice(prevEnd, m.index).trim();
    nameRaw = nameRaw.replace(/部品名及び仕様|図面番号|所要数/g, "").trim();

    supplied.push({
      partName: nameRaw,
      drawingNum: m.id,
      required: Number(m.req)
    });
  }

  return supplied;
}

function parseSecondaryRows(pageText) {
  const secondary = [];
  const sectionMatch = pageText.match(/二 次 加 工 費([\s\S]*?)(⑬|合\s*計)/);
  if (!sectionMatch) return [];

  const sectionText = sectionMatch[1];

  // Improved Regex for Secondary
  // Pattern: Name ... Num Num Num Num
  // Allow spaces/dots in numbers.
  // But separating columns is tricky if spaces are allowed inside numbers.
  // We assume: Columns are separated by spaces.
  // We look for the LAST 4 number-groups in the line?
  // Or matches that look like integers/floats?

  // Try to match the tail of the line consisting of 4 numbers.
  // The numbers can be fullwidth.
  // (Num) space (Num) space (Num) space (Num)

  const tailRegex = /((?:[\d０-９]+(?:\s*[\.．]\s*[\d０-９]+)?\s*){4})$/;

  // We need to iterate lines? Or loose match?
  // `popup.js` provides flattened text usually unless my splitPages logic preserves something?
  // `splitPages` just splits by `--- PAGE ---`.

  // Since we assume flattened text, `.+?` is dangerous if valid structure follows.
  // But `itemMarkerRegex` works by finding the anchors (IDs or Number clusters).

  // Let's scan for any sequence of 4 numbers.
  // And assume text before it is the Name.

  const rowMarkerRegex = /([\d０-９]+(?:\s*[\.．]\s*[\d０-９]+)?)\s+([\d０-９]+(?:\s*[\.．]\s*[\d０-９]+)?)\s+([\d０-９]+(?:\s*[\.．]\s*[\d０-９]+)?)\s+([\d０-９]+(?:\s*[\.．]\s*[\d０-９]+)?)/g;

  let match;
  let matches = [];
  while ((match = rowMarkerRegex.exec(sectionText)) !== null) {
    matches.push({
      index: match.index,
      end: match.index + match[0].length,
      qty: normalize(match[1]),
      seconds: normalize(match[2]),
      unitPrice: normalize(match[3]),
      total: normalize(match[4])
    });
  }

  for (let i = 0; i < matches.length; i++) {
    const m = matches[i];
    const prevEnd = (i === 0) ? 0 : matches[i - 1].end;
    let nameRaw = sectionText.slice(prevEnd, m.index).trim();

    if (nameRaw.includes("加工工程") || nameRaw.includes("個数")) continue;
    if (nameRaw.includes("荷姿") || nameRaw.includes("備考") || nameRaw.includes("適用")) continue;
    if (!nameRaw) continue;

    secondary.push({
      process: nameRaw,
      qty: m.qty,
      seconds: m.seconds,
      unitPrice: m.unitPrice,
      total: m.total
    });
  }

  return secondary;
}

// Helper to standardise OCR text (Fullwidth -> Halfwidth, Fix broken floats)
function normalizeOCRText(text) {
  return text.replace(/[０-９．]/g, (c) => {
    if (c === '．') return '.';
    return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
  })
    // Fix broken floats: "29 . 08" -> "29.08"
    // Look for Digit + Space + Dot + Space + Digit
    .replace(/(\d)\s+\.\s+(\d)/g, "$1.$2")
    .replace(/(\d)\s+(\.)/g, "$1$2") // Digit Space Dot
    .replace(/(\.)\s+(\d)/g, "$1$2"); // Dot Space Digit
}

export function parsePdfTextToPayload(text) {
  const pages = splitPages(text);

  const parts = [];
  const details = {};

  for (let pageText of pages) {
    // Detect OCR mode for this page / section
    // We already passed text, but `popup.js` adds "--- PAGE X ---"
    // The inner text comes from OCR.

    // We can infer OCR if we see Fullwidth chars or if header says so?
    // Actually, `popup.js` handles the switch.
    // But `parse.js` needs to know.
    // Existing logic uses `parseHeader(pageText)` to return `isOCR`.

    // Let's TRY to pre-normalize if it looks messy?
    // Or just run it on the raw text if parsing fails?
    // User requested "Standard way".
    // Let's normalize it aggressively if it's OCR.

    let header = parseHeader(pageText);

    if (header && header.isOCR) {
      pageText = normalizeOCRText(pageText);
      const newHeader = parseHeader(pageText);
      if (newHeader) header = newHeader;
    }

    const mainDrawing = header ? header.drawingNumber : null;

    if (!mainDrawing) {
      continue;
    }

    parts.push({
      drawingNumber: mainDrawing,
      partName: header.partName || "",
      representativeModel: "",
      estimatedQuantity: "",
      estimateNumber: "",
      unitPrice: header.unitPrice || "",
      applicationPeriod: header.applicationPeriod || ""
    });

    details[mainDrawing] = {
      breakdown: header.breakdown, // Added breakdown to details
      material: header.isOCR ? parseMaterialRowsOCR(pageText, mainDrawing) : parseMaterialRows(pageText),
      purchase: header.isOCR ? parsePurchaseRowsOCR(pageText) : parsePurchaseRows(pageText),
      supplied: header.isOCR ? parseSuppliedRowsOCR(pageText) : parseSuppliedRows(pageText),
      press: [], // Removed parsing, empty array
      secondary: header.isOCR ? parseSecondaryRowsOCR(pageText) : parseSecondaryRows(pageText)
    };
  }

  const uniqueParts = [];
  const seen = new Set();
  for (const p of parts) {
    if (seen.has(p.drawingNumber)) continue;
    seen.add(p.drawingNumber);
    uniqueParts.push(p);
  }

  return { parts: uniqueParts, details };
}
