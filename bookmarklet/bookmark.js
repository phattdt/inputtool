javascript: (function () {
	var lib = document.createElement('script');
	lib.src = 'https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js';
	document.head.appendChild(lib);
	lib.onload = function () {
		function clean(s) {
			return String(s || "").replace(/\s/g, "");
		}
		function injectStyles() {
			if (document.getElementById('gproc-modal-style')) return;
			const style = document.createElement('style');
			style.id = 'gproc-modal-style';
			style.innerHTML = `
				.gproc-modal-overlay {
					position: fixed; top: 0; left: 0; width: 100%; height: 100%;
					background: rgba(0, 0, 0, 0.5); z-index: 9999;
					display: flex; justify-content: center; align-items: center;
					font-family: sans-serif;
				}
				.gproc-modal {
					background: white; padding: 20px; border-radius: 8px;
					width: 800px; max-height: 90vh; overflow-y: auto;
					box-shadow: 0 4px 6px rgba(0,0,0,0.1);
				}
				.gproc-modal h2 { margin-top: 0; color: #333; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
				.gproc-section-title {
					background: #f8f9fa; padding: 8px; margin: 15px 0 5px; 
					font-weight: bold; border-left: 4px solid #3498db;
				}
				.gproc-field-row {
					display: flex; justify-content: space-between; align-items: center;
					margin-bottom: 5px; padding: 5px 0; border-bottom: 1px solid #eee;
				}
				.gproc-field-label { font-size: 13px; color: #555; width: 40%; }
				.gproc-select { padding: 5px; width: 55%; border: 1px solid #ccc; border-radius: 4px; }
				.gproc-btn-group { margin-top: 20px; text-align: right; position: sticky; bottom: 0; background: white; padding-top: 10px; }
				.gproc-btn {
					padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;
					margin-left: 10px; font-weight: bold;
				}
				.gproc-btn-cancel { background: #ccc; color: #333; }
				.gproc-btn-confirm { background: #3498db; color: white; }
				.gproc-btn-confirm:hover { background: #2980b9; }
			`;
			document.head.appendChild(style);
		}
		function createMappingModal(excelHeaders, onConfirm) {
			injectStyles();
			const sections = [
				{
					title: "部品リスト (Main Parts List)",
					sectionKey: "ヘッダー",
					fields: [
						{ key: "parts_drawingNumber", label: "図面番号", autoMatch: ["図面番号"] },
						{ key: "parts_partName", label: "部品名称", autoMatch: ["部品名称"] },
						{ key: "parts_model", label: "代表機種", autoMatch: ["代表機種"] },
						{ key: "parts_estimatedQty", label: "見積り数", autoMatch: ["見積り数", "見積数"] },
						{ key: "parts_estimateNum", label: "見積書＃", autoMatch: ["見積書＃", "見積り"] },
						{ key: "parts_unitPrice", label: "単価 (リスト用)", autoMatch: ["単価"] },
						{ key: "parts_appPeriod", label: "適用時期", autoMatch: ["適用時期", "運用時期"] }
					]
				},
				{
					title: "材料費 (Material Cost)",
					sectionKey: "材料費",
					fields: [
						{ key: "mat_drawing_number", label: "図面番号 (材料)", autoMatch: ["図面番号"] },
						{ key: "mat_temper", label: "調質", autoMatch: ["調質", "素材"] },
						{ key: "mat_diameter", label: "φ (外径)", autoMatch: ["φ", "外径", "寸法"] },
						{ key: "mat_thickness", label: "t (肉厚)", autoMatch: ["t", "肉厚", "ｔ"] },
						{ key: "mat_length", label: "l (長さ)", autoMatch: ["l", "長さ", "ｌ"] },
						{ key: "mat_weight", label: "重量(g)", autoMatch: ["重量", "単重量"] },
						{ key: "mat_unit_price_kg", label: "単価/Kg", autoMatch: ["単価/Kg", "単価"] },
						{ key: "mat_consumption_price", label: "消費材料費", autoMatch: ["消費材料費", "価格"] },
						{ key: "mat_material_cost", label: "材料費", autoMatch: ["材料費"] },
						{ key: "mat_proc_unit_price", label: "加工単価", autoMatch: ["加工単価", "単価"] },
						{ key: "mat_proc_price", label: "加工価格", autoMatch: ["加工価格", "価格"] },
						{ key: "mat_piping_price", label: "配管価格", autoMatch: ["配管価格"] },
						{ key: "mat_required", label: "所要数", autoMatch: ["所要数"] }
					]
				},
				{
					title: "購入部品費 (Purchase Parts)",
					sectionKey: "購入部品費",
					fields: [
						{ key: "pur_part_name", label: "部品名及び仕様", autoMatch: ["部品名及び仕様", "部品名"] },
						{ key: "pur_drawing_num", label: "図面番号", autoMatch: ["図面番号"] },
						{ key: "pur_required", label: "所要数", autoMatch: ["所要数"] },
						{ key: "pur_unit_price", label: "単価", autoMatch: ["単価"] }
					]
				},
				{
					title: "支給部品 (Supplied Parts)",
					sectionKey: "支給部品",
					fields: [
						{ key: "sup_part_name", label: "部品名及び仕様", autoMatch: ["部品名及び仕様", "部品名"] },
						{ key: "sup_drawing_num", label: "図面番号", autoMatch: ["図面番号"] },
						{ key: "sup_required", label: "所要数", autoMatch: ["所要数"] }
					]
				},
				{
					title: "二次加工費 (Secondary Processing)",
					sectionKey: "二次加工費",
					fields: [
						{ key: "sec_process", label: "加工工程", autoMatch: ["加工工程"] },
						{ key: "sec_qty", label: "個数", autoMatch: ["個数", "人員"] },
						{ key: "sec_seconds", label: "（秒）/ サイクル", autoMatch: ["（秒）", "サイクル", "サイクルタイム"] },
						{ key: "sec_unit_price", label: "単価/チャージ", autoMatch: ["単価", "チャージ"] },
						{ key: "sec_total", label: "金額/加工費", autoMatch: ["金額", "加工費"] }
					]
				},
				{
					title: "単価内訳 (Breakdown)",
					sectionKey: "単価",
					fields: [
						{ key: "bd_mat_total", label: "材料費合計", autoMatch: ["材料費合計", "1台当り材料費合計"] },
						{ key: "bd_mat_mgmt", label: "材料管理費", autoMatch: ["材料管理費"] },
						{ key: "bd_proc_total", label: "加工費合計", autoMatch: ["加工費合計", "1台当り加工費合計"] },
						{ key: "bd_gen_sales", label: "一般販売管理費", autoMatch: ["一般販売管理費"] },
						{ key: "bd_profit", label: "利益", autoMatch: ["利益"] },
						{ key: "bd_freight", label: "運賃・家賃", autoMatch: ["運賃・家賃", "運賃"] },
						{ key: "bd_total", label: "合計", autoMatch: ["合計"] }
					]
				}
			];
			const defaultMapping = {};
			sections.forEach(sec => {
				const sectionHeaders = excelHeaders.filter(h => h.section === sec.sectionKey);
				sec.fields.forEach(f => {
					const match = sectionHeaders.find(h => f.autoMatch.some(m => clean(h.text).includes(clean(m))));
					if (match) defaultMapping[f.key] = JSON.stringify(match);
				});
			});
			const overlay = document.createElement('div');
			overlay.className = 'gproc-modal-overlay';
			const modal = document.createElement('div');
			modal.className = 'gproc-modal';
			const titleHeader = document.createElement('h2');
			titleHeader.textContent = 'Excel列のマッピング設定';
			modal.appendChild(titleHeader);
			const desc = document.createElement('p');
			desc.textContent = 'Excelの列名とシステムの項目を紐付けてください。名称が部分一致する項目は自動選択されています。';
			desc.style.color = '#666';
			modal.appendChild(desc);
			const form = document.createElement('div');
			sections.forEach(sec => {
				const secTitle = document.createElement('div');
				secTitle.className = 'gproc-section-title';
				secTitle.textContent = sec.title;
				form.appendChild(secTitle);
				sec.fields.forEach(f => {
					const row = document.createElement('div');
					row.className = 'gproc-field-row';
					const label = document.createElement('div');
					label.className = 'gproc-field-label';
					label.textContent = f.label;
					const select = document.createElement('select');
					select.className = 'gproc-select';
					select.dataset.key = f.key;
					const defaultOpt = document.createElement('option');
					defaultOpt.value = "";
					defaultOpt.textContent = "-- 指定なし (空欄) --";
					select.appendChild(defaultOpt);
					let sectionHeaders = excelHeaders.filter(h => h.section === sec.sectionKey);
					if (sectionHeaders.length === 0) {
						sectionHeaders = excelHeaders;
					}
					sectionHeaders.forEach(h => {
						const opt = document.createElement('option');
						opt.value = JSON.stringify(h);
						opt.textContent = sectionHeaders === excelHeaders ? h.display : h.text;
						if (defaultMapping[f.key] === JSON.stringify(h)) opt.selected = true;
						select.appendChild(opt);
					});
					row.appendChild(label);
					row.appendChild(select);
					form.appendChild(row);
				});
			});
			modal.appendChild(form);
			const btnGroup = document.createElement('div');
			btnGroup.className = 'gproc-btn-group';
			const cancelBtn = document.createElement('button');
			cancelBtn.className = 'gproc-btn gproc-btn-cancel';
			cancelBtn.textContent = 'キャンセル';
			cancelBtn.onclick = () => document.body.removeChild(overlay);
			const confirmBtn = document.createElement('button');
			confirmBtn.className = 'gproc-btn gproc-btn-confirm';
			confirmBtn.textContent = '反映する';
			confirmBtn.onclick = () => {
				const mapping = {};
				form.querySelectorAll('select').forEach(s => {
					if (s.value) {
						try {
							const parsed = JSON.parse(s.value);
							mapping[s.dataset.key] = { header: parsed.text, colIdx: parsed.colIdx };
						} catch (e) {
							mapping[s.dataset.key] = s.value;
						}
					}
				});
				document.body.removeChild(overlay);
				onConfirm(mapping);
			};
			btnGroup.appendChild(cancelBtn);
			btnGroup.appendChild(confirmBtn);
			modal.appendChild(btnGroup);
			overlay.appendChild(modal);
			document.body.appendChild(overlay);
		}
		function findRowByLabel(rows, label) {
			const target = clean(label);
			for (let r = 0; r < rows.length; r++) {
				if (clean(rows[r][0]).includes(target)) return r;
			}
			return -1;
		}
		function parseKeyValue(cellValue) {
			if (!cellValue || typeof cellValue !== 'string') return null;
			const delimiters = ['\n', '\r\n', ':', '：', '＝', '='];
			for (const d of delimiters) {
				const idx = cellValue.indexOf(d);
				if (idx > 0 && idx < cellValue.length - 1) {
					return { key: cellValue.substring(0, idx).trim(), value: cellValue.substring(idx + 1).trim() };
				}
			}
			return null;
		}
		function getVal(row, headerRow, mapKey, mapping) {
			const mappedItem = mapping[mapKey];
			if (!mappedItem) return "";
			let cellValue = "";
			if (typeof mappedItem === 'object' && mappedItem.colIdx !== undefined) {
				cellValue = row[mappedItem.colIdx] || "";
			} else {
				const colIndex = headerRow.findIndex(h => h === mappedItem);
				if (colIndex === -1) return "";
				cellValue = row[colIndex] || "";
			}
			const kv = parseKeyValue(String(cellValue));
			if (kv) return kv.value;
			return cellValue;
		}
		function getValNum(row, headerRow, mapKey, mapping) {
			const val = getVal(row, headerRow, mapKey, mapping);
			if (typeof val === 'number') return val;
			if (!val) return "";
			const cleanVal = String(val).replace(/,/g, '');
			const num = parseFloat(cleanVal);
			return isNaN(num) ? "" : num;
		}
		function parseNum(val) {
			if (typeof val === 'number') return val;
			if (!val) return "";
			const cleanVal = String(val).replace(/,/g, '');
			const num = parseFloat(cleanVal);
			return isNaN(num) ? "" : num;
		}
		window.gprocFillLogic = function (sheetsData, mapping) {
			if (!mapping) {
				try {
					const saved = localStorage.getItem('gproc_excel_mapping');
					mapping = saved ? JSON.parse(saved) : {};
				} catch (e) { mapping = {}; }
			}
			const PERF_START = performance.now();
			console.log('⏱️  [PERF] Bookmarklet parsing started');
			const detailForm = document.getElementById('detailForm');
			const isDetailView = detailForm && !detailForm.classList.contains('hidden');
			const sheetNames = Object.keys(sheetsData);

			function findHeaderIdx(rows, keyword) {
				for (let i = 0; i < Math.min(rows.length, 50); i++) {
					if (rows[i] && rows[i].some(c => clean(c).includes(keyword))) return i;
				}
				return -1;
			}
			if (!isDetailView) {
				sheetNames.forEach((name) => {
					var rows = sheetsData[name];
					if (!rows || rows.length === 0) return;

					let headerIdx = findHeaderIdx(rows, "図面番号");
					if (headerIdx === -1) headerIdx = 0;
					const headerRow = rows[headerIdx];

					function getValSmart(mapKey) {
						const mappedItem = mapping[mapKey];
						if (!mappedItem || typeof mappedItem !== 'object') return "";
						const colIdx = mappedItem.colIdx;
						if (colIdx === undefined) return "";
						const headerCellValue = String(headerRow[colIdx] || "");
						const kv = parseKeyValue(headerCellValue);
						if (kv && kv.value) return kv.value;
						if (rows.length > headerIdx + 1) {
							const dataRow = rows[headerIdx + 1];
							return dataRow[colIdx] || "";
						}
						return "";
					}
					var record = {};
					record["図面番号"] = getValSmart("parts_drawingNumber");
					record["部品名称"] = getValSmart("parts_partName");
					record["代表機種"] = getValSmart("parts_model");
					record["見積り数"] = getValSmart("parts_estimatedQty");
					record["見積書＃"] = getValSmart("parts_estimateNum");
					record["適用時期"] = getValSmart("parts_appPeriod");
					if (record["図面番号"] == undefined && record["適用機種名"]) {
						record["図面番号"] = record["適用機種名"];
					}
					var targetRow = Array.from(document.querySelectorAll('#partsTableBody tr')).find(r => !r.querySelector('input').value);
					if (!targetRow && window.addPartRow) {
						window.addPartRow();
						targetRow = document.querySelector('#partsTableBody tr:last-child');
					}
					if (targetRow) {
						var ins = targetRow.querySelectorAll('input');
						ins[0].value = record["図面番号"] || "";
						ins[1].value = record["部品名称"] || "";
						ins[2].value = record["代表機種"] || "";
						ins[3].value = record["見積り数"] || "";
						ins[4].value = record["見積書＃"] || "";
						const up = getVal((rows.length > headerIdx + 1 ? rows[headerIdx + 1] : []), headerRow, "parts_unitPrice", mapping);
						ins[5].value = up || "";
						ins[6].value = record["適用時期"] || "";
						ins.forEach(inp => inp.dispatchEvent(new Event('change', { bubbles: true })));
					}
				});
			} else {
				const headerEl = document.getElementById('drawingNumberHeader');
				const currentDN = clean(headerEl ? headerEl.textContent.replace('図面番号:', '') : "");
				const targetSheetName = sheetNames.find(name => sheetsData[name].some(r => r.some(c => clean(c).includes(currentDN)))) || sheetNames[0];
				var rows = sheetsData[targetSheetName];
				const drawingNumber = window.currentDrawingNumber;
				if (!drawingNumber || !window.detailData || !window.detailData[drawingNumber]) {
					console.error('No detailData found');
					return;
				}
				let mainHeaderIdx = findHeaderIdx(rows, "図面番号");
				if (mainHeaderIdx === -1) mainHeaderIdx = 0;
				const mainHeaderRow = rows[mainHeaderIdx];
				function extractSectionRows(label, offset, maxRows) {
					const targetLabel = label;
					const rIdx = findRowByLabel(rows, targetLabel);
					if (rIdx === -1) return [];
					const result = [];
					for (let i = 0; i < maxRows; i++) {
						let exR = rows[rIdx + offset + i];
						if (!exR) break;
						if (i > 0 && exR[0] && clean(exR[0]) !== "") break;
						result.push(exR);
					}
					return result;
				}
				function getLocalHeader(label, expectedKey, defaultHeader) {
					const rIdx = findRowByLabel(rows, label);
					if (rIdx === -1) return defaultHeader;
					if (rows[rIdx] && rows[rIdx].some(c => clean(c).includes(expectedKey))) return rows[rIdx];
					if (rows[rIdx + 1]) return rows[rIdx + 1];
					return defaultHeader;
				}
				const materialRows = extractSectionRows("材料費", 2, 20);
				const matHeader = getLocalHeader("材料費", "調質", mainHeaderRow);
				materialRows.forEach((exR, i) => {
					while (window.detailData[drawingNumber].material.length <= i) {
						window.detailData[drawingNumber].material.push({ id: 'id_' + Math.random().toString(36).substr(2, 9), drawingNumber: '', temper: '', diameter: '', thickness: '', length: '', weight: '', unitPriceKg: '', consumptionPrice: '', materialCost: '', processingUnitPrice: '', processingPrice: '', pipingPrice: '', required: '' });
					}
					const m = window.detailData[drawingNumber].material[i];
					const h = matHeader;
					m.drawingNumber = getVal(exR, h, "mat_drawing_number", mapping);
					m.temper = getVal(exR, h, "mat_temper", mapping);
					m.diameter = getValNum(exR, h, "mat_diameter", mapping);
					m.thickness = getValNum(exR, h, "mat_thickness", mapping);
					m.length = getValNum(exR, h, "mat_length", mapping);
					m.weight = getValNum(exR, h, "mat_weight", mapping);
					m.unitPriceKg = getValNum(exR, h, "mat_unit_price_kg", mapping);
					m.consumptionPrice = getValNum(exR, h, "mat_consumption_price", mapping);
					m.materialCost = getValNum(exR, h, "mat_material_cost", mapping);
					m.processingUnitPrice = getValNum(exR, h, "mat_proc_unit_price", mapping);
					m.processingPrice = getValNum(exR, h, "mat_proc_price", mapping);
					m.pipingPrice = getValNum(exR, h, "mat_piping_price", mapping);
					m.required = getValNum(exR, h, "mat_required", mapping);
				});
				const purchaseRows = extractSectionRows("購入部品費", 1, 20);
				const purHeader = getLocalHeader("購入部品費", "部品名", mainHeaderRow);
				let pIdx = 0;
				purchaseRows.forEach((exR) => {
					if (!exR || exR.length === 0) return;
					while (window.detailData[drawingNumber].purchase.length <= pIdx) {
						window.detailData[drawingNumber].purchase.push({ id: 'id_' + Math.random().toString(36).substr(2, 9), partName: '', drawingNum: '', required: '', unitPrice: '' });
					}
					const p = window.detailData[drawingNumber].purchase[pIdx];
					const h = purHeader;
					p.partName = getVal(exR, h, "pur_part_name", mapping);
					p.drawingNum = getVal(exR, h, "pur_drawing_num", mapping);
					p.required = getValNum(exR, h, "pur_required", mapping);
					p.unitPrice = getValNum(exR, h, "pur_unit_price", mapping);
					if (p.partName || p.drawingNum) pIdx++;
				});

				const suppliedRows = extractSectionRows("支給部品", 1, 20);
				const supHeader = getLocalHeader("支給部品", "部品名", mainHeaderRow);
				let sIdx = 0;
				suppliedRows.forEach((exR) => {
					while (window.detailData[drawingNumber].supplied.length <= sIdx) {
						window.detailData[drawingNumber].supplied.push({ id: 'id_' + Math.random().toString(36).substr(2, 9), partName: '', drawingNum: '', required: '' });
					}
					const sp = window.detailData[drawingNumber].supplied[sIdx];
					const h = supHeader;
					sp.partName = getVal(exR, h, "sup_part_name", mapping);
					sp.drawingNum = getVal(exR, h, "sup_drawing_num", mapping);
					sp.required = getValNum(exR, h, "sup_required", mapping);
					if (sp.partName || sp.drawingNum) sIdx++;
				});

				const secondaryRows = extractSectionRows("二次加工費", 1, 20);
				const secHeader = getLocalHeader("二次加工費", "加工工程", mainHeaderRow);
				secondaryRows.forEach((exR, i) => {
					while (window.detailData[drawingNumber].secondary.length <= i) {
						window.detailData[drawingNumber].secondary.push({ id: 'id_' + Math.random().toString(36).substr(2, 9), process: '', qty: '', seconds: '', unitPrice: '', total: '' });
					}
					const s = window.detailData[drawingNumber].secondary[i];
					const h = secHeader;
					s.process = getVal(exR, h, "sec_process", mapping);
					s.qty = getValNum(exR, h, "sec_qty", mapping);
					s.seconds = getValNum(exR, h, "sec_seconds", mapping);
					s.unitPrice = getValNum(exR, h, "sec_unit_price", mapping);
					s.total = getValNum(exR, h, "sec_total", mapping);
				});
				const breakdownLabel = "単価";
				const breakdownIdx = findRowByLabel(rows, breakdownLabel);
				if (breakdownIdx !== -1) {
					if (!window.detailData[drawingNumber].breakdown) {
						window.detailData[drawingNumber].breakdown = {};
					}
					const bd = window.detailData[drawingNumber].breakdown;
					const valuesRow = rows[breakdownIdx + 3];
					const bdHeader = (breakdownIdx !== -1 && rows[breakdownIdx + 1]) ? rows[breakdownIdx + 1] : mainHeaderRow;
					const h = bdHeader;
					if (valuesRow) {
						bd.materialTotal = getValNum(valuesRow, h, "bd_mat_total", mapping);
						bd.materialMgmt = getValNum(valuesRow, h, "bd_mat_mgmt", mapping);
						bd.procTotal = getValNum(valuesRow, h, "bd_proc_total", mapping);
						bd.genSales = getValNum(valuesRow, h, "bd_gen_sales", mapping);
						bd.profit = getValNum(valuesRow, h, "bd_profit", mapping);
						bd.freight = getValNum(valuesRow, h, "bd_freight", mapping);
						bd.materialMgmt14 = "";
						bd.genSales15 = "";
						bd.total = getValNum(valuesRow, h, "bd_total", mapping);
					}
				}
				if (window.renderDetailTables) {
					window.renderDetailTables();
				}
				const PERF_END = performance.now();
				console.log(`⏱️  [PERF] Detail parsing completed in ${(PERF_END - PERF_START).toFixed(2)}ms`);
			}
			const PERF_TOTAL = performance.now();
			console.log(`⏱️  [PERF] Total bookmarklet execution: ${(PERF_TOTAL - PERF_START).toFixed(2)}ms`);
		};
		const detailForm = document.getElementById('detailForm');
		const isDetailView = detailForm && !detailForm.classList.contains('hidden');
		const savedData = localStorage.getItem('gproc_excel_data');
		const savedMapping = localStorage.getItem('gproc_excel_mapping');
		function initFileInput() {
			var input = document.createElement('input');
			input.type = 'file';
			input.accept = '.xlsx, .xls';
			input.style.display = 'none';
			document.body.appendChild(input);
			input.onchange = function (e) {
				var file = e.target.files[0];
				var reader = new FileReader();
				reader.onload = function (ev) {
					var data = new Uint8Array(ev.target.result);
					var workbook = XLSX.read(data, { type: 'array', cellDates: true });
					var storageObj = {};
					workbook.SheetNames.forEach(name => {
						storageObj[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "" });
					});
					localStorage.setItem('gproc_excel_data', JSON.stringify(storageObj));
					let headerItems = [];
					if (workbook.SheetNames.length > 0) {
						const firstSheet = storageObj[workbook.SheetNames[0]];
						if (firstSheet && firstSheet.length > 0) {
							const seen = new Set();
							const sectionLabels = ["部品リスト", "材料費", "購入部品費", "支給部品", "二次加工費", "単価"];
							const headerKeywords = ["図面番号", "部品名", "機種", "見積", "適用時期", "調質", "外径", "肉厚", "長さ", "重量", "単価", "価格", "所要数", "加工工程", "個数", "サイクル", "チャージ", "金額", "合計", "管理費", "利益", "運賃"];
							let currentSection = "ヘッダー";
							let isHeaderRow = false;
							firstSheet.slice(0, 50).forEach((row, rowIdx) => {
								if (Array.isArray(row)) {
									const col0 = clean(row[0] || "");
									const foundSection = sectionLabels.find(s => col0.includes(s));
									if (foundSection) currentSection = foundSection;
									isHeaderRow = row.some(cell => {
										const cellText = clean(cell || "");
										return headerKeywords.some(kw => cellText.includes(kw));
									});
									if (isHeaderRow) {
										row.forEach((cell, colIdx) => {
											let text = '';
											if (cell && typeof cell === 'string') text = clean(cell);
											else if (cell) text = String(cell).trim();
											const kvDelimiters = ['\n', '\r\n', ':', '：', '＝', '='];
											let keyPart = text;
											let isKeyValueCell = false;
											for (const d of kvDelimiters) {
												const idx = text.indexOf(d);
												if (idx > 0 && idx < text.length - 1) {
													keyPart = text.substring(0, idx).trim();
													isKeyValueCell = true;
													break;
												}
											}
											if (keyPart && !seen.has(keyPart + '_' + currentSection)) {
												seen.add(keyPart + '_' + currentSection);
												headerItems.push({ text: keyPart, rowIdx: rowIdx, colIdx: colIdx, section: currentSection, isKeyValueCell: isKeyValueCell, display: keyPart + ' (' + currentSection + ')' });
											}
										});
									}
								}
							});
						}
					}
					createMappingModal(headerItems, (mapping) => {
						console.log("Selected Mapping:", mapping);
						localStorage.setItem('gproc_excel_mapping', JSON.stringify(mapping));
						window.gprocFillLogic(storageObj, mapping);
						alert("アップロードとマッピングが完了しました！");
					});
				};
				reader.readAsArrayBuffer(file);
			};
			input.click();
		}
		if (isDetailView && savedData) {
			const mapping = savedMapping ? JSON.parse(savedMapping) : {};
			window.gprocFillLogic(JSON.parse(savedData), mapping);
		} else {
			initFileInput();
		}
	};
})();