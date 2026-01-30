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
						{ key: "mat_drawing_number", label: "図面番号", autoMatch: ["図面番号"] },
						{ key: "mat_material", label: "素材", autoMatch: ["素材", "C1220"] },
						{ key: "mat_temper", label: "調質", autoMatch: ["調質"] },
						{ key: "mat_diameter", label: "外径", autoMatch: ["外径", "φ"] },
						{ key: "mat_thickness", label: "肉厚", autoMatch: ["肉厚", "t"] },
						{ key: "mat_length", label: "長さ", autoMatch: ["長さ", "l"] },
						{ key: "mat_quantity", label: "取り数", autoMatch: ["取り数"] },
						{ key: "mat_weight", label: "単重量(消費)", autoMatch: ["単重量"] },
						{ key: "mat_unit_price", label: "単価(消費)", autoMatch: ["単価"] },
						{ key: "mat_price", label: "価格(消費)", autoMatch: ["価格"] },
						{ key: "mat_proc_weight", label: "単重量(加工)", autoMatch: ["単重量"] },
						{ key: "mat_proc_unit_price", label: "単価(加工)", autoMatch: ["単価"] },
						{ key: "mat_proc_price", label: "価格(加工)", autoMatch: ["価格"] },
						{ key: "mat_material_cost", label: "材料費", autoMatch: ["材料費"] },
						{ key: "mat_piping_price", label: "配管価格", autoMatch: ["配管価格"] }
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
					title: "プレス・配管加工費 (Press/Piping)",
					sectionKey: "プレス・配管加工費",
					fields: [
						{ key: "prs_part_name", label: "部品名(金型費)", autoMatch: ["部品名", "金型費"] },
						{ key: "prs_process", label: "加工工程", autoMatch: ["加工工程"] },
						{ key: "prs_equipment", label: "使用設備", autoMatch: ["使用設備"] },
						{ key: "prs_cycle_setup", label: "段取り", autoMatch: ["段取り"] },
						{ key: "prs_cycle_proc", label: "加工", autoMatch: ["加工"] },
						{ key: "prs_cycle_total", label: "計", autoMatch: ["計"] },
						{ key: "prs_charge", label: "チャージ(円/分)", autoMatch: ["チャージ", "円/分"] },
						{ key: "prs_cost", label: "加工費", autoMatch: ["加工費"] },
						{ key: "prs_required", label: "所要数", autoMatch: ["所要数"] }
					]
				},
				{
					title: "二次加工費 (Secondary Processing)",
					sectionKey: "二次加工費",
					fields: [
						{ key: "sec_process", label: "加工工程", autoMatch: ["加工工程"] },
						{ key: "sec_personnel", label: "人員", autoMatch: ["人員"] },
						{ key: "sec_cycle_time", label: "サイクルタイム", autoMatch: ["サイクルタイム", "サイクル"] },
						{ key: "sec_charge", label: "チャージ(円/分)", autoMatch: ["チャージ", "円/分"] },
						{ key: "sec_cost", label: "加工費", autoMatch: ["加工費"] }
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
						window.detailData[drawingNumber].material.push({ id: 'id_' + Math.random().toString(36).substr(2, 9), drawingNumber: '', material: '', temper: '', diameter: '', thickness: '', length: '', quantity: '', weight: '', unitPrice: '', price: '', procWeight: '', procUnitPrice: '', procPrice: '', materialCost: '', pipingPrice: '' });
					}
					const m = window.detailData[drawingNumber].material[i];
					const h = matHeader;
					m.drawingNumber = getVal(exR, h, "mat_drawing_number", mapping);
					m.material = getVal(exR, h, "mat_material", mapping);
					m.temper = getVal(exR, h, "mat_temper", mapping);
					m.diameter = getValNum(exR, h, "mat_diameter", mapping);
					m.thickness = getValNum(exR, h, "mat_thickness", mapping);
					m.length = getValNum(exR, h, "mat_length", mapping);
					m.quantity = getValNum(exR, h, "mat_quantity", mapping);
					m.weight = getValNum(exR, h, "mat_weight", mapping);
					m.unitPrice = getValNum(exR, h, "mat_unit_price", mapping);
					m.price = getValNum(exR, h, "mat_price", mapping);
					m.procWeight = getValNum(exR, h, "mat_proc_weight", mapping);
					m.procUnitPrice = getValNum(exR, h, "mat_proc_unit_price", mapping);
					m.procPrice = getValNum(exR, h, "mat_proc_price", mapping);
					m.materialCost = getValNum(exR, h, "mat_material_cost", mapping);
					m.pipingPrice = getValNum(exR, h, "mat_piping_price", mapping);
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

				const pressRows = extractSectionRows("プレス・配管加工費", 3, 20);
				const prsHeader = getLocalHeader("プレス・配管加工費", "部品名", mainHeaderRow);
				let prsIdx = 0;
				pressRows.forEach((exR) => {
					while (window.detailData[drawingNumber].press.length <= prsIdx) {
						window.detailData[drawingNumber].press.push({ id: 'id_' + Math.random().toString(36).substr(2, 9), partName: '', process: '', equipment: '', cycleSetup: '', cycleProc: '', cycleTotal: '', charge: '', cost: '', required: '' });
					}
					const pr = window.detailData[drawingNumber].press[prsIdx];
					const h = prsHeader;
					pr.partName = getVal(exR, h, "prs_part_name", mapping);
					pr.process = getVal(exR, h, "prs_process", mapping);
					pr.equipment = getVal(exR, h, "prs_equipment", mapping);
					pr.cycleSetup = getValNum(exR, h, "prs_cycle_setup", mapping);
					pr.cycleProc = getValNum(exR, h, "prs_cycle_proc", mapping);
					pr.cycleTotal = getValNum(exR, h, "prs_cycle_total", mapping);
					pr.charge = getValNum(exR, h, "prs_charge", mapping);
					pr.cost = getValNum(exR, h, "prs_cost", mapping);
					pr.required = getValNum(exR, h, "prs_required", mapping);
					if (pr.partName || pr.process) prsIdx++;
				});

				const secondaryRows = extractSectionRows("二次加工費", 1, 20);
				const secHeader = getLocalHeader("二次加工費", "加工工程", mainHeaderRow);
				secondaryRows.forEach((exR, i) => {
					while (window.detailData[drawingNumber].secondary.length <= i) {
						window.detailData[drawingNumber].secondary.push({ id: 'id_' + Math.random().toString(36).substr(2, 9), process: '', personnel: '', cycleTime: '', charge: '', cost: '' });
					}
					const s = window.detailData[drawingNumber].secondary[i];
					const h = secHeader;
					s.process = getVal(exR, h, "sec_process", mapping);
					s.personnel = getValNum(exR, h, "sec_personnel", mapping);
					s.cycleTime = getValNum(exR, h, "sec_cycle_time", mapping);
					s.charge = getValNum(exR, h, "sec_charge", mapping);
					s.cost = getValNum(exR, h, "sec_cost", mapping);
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
							const sectionLabels = ["部品リスト", "材料費", "購入部品費", "支給部品", "二次加工費", "プレス・配管加工費", "単価"];
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
			showModeSelection();
		}
		function showModeSelection() {
			injectStyles();
			var input = document.createElement('input');
			input.type = 'file';
			input.accept = '.xlsx, .xls';
			input.style.display = 'none';
			document.body.appendChild(input);

			input.onchange = function (e) {
				var file = e.target.files[0];
				if (!file) return;

				var reader = new FileReader();
				reader.onload = function (ev) {
					try {
						var data = new Uint8Array(ev.target.result);
						var workbook = XLSX.read(data, { type: 'array', cellDates: true });
						var storageObj = {};
						workbook.SheetNames.forEach(function (name) {
							storageObj[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "" });
						});
						showExcelOptionsModal(file.name, storageObj);
					} catch (err) {
						alert('Excelファイルの読み込みに失敗しました: ' + err.message);
						console.error(err);
					}
				};
				reader.readAsArrayBuffer(file);
			};
			input.click();
		}

		function showExcelOptionsModal(fileName, sheetsData) {
			const overlay = document.createElement('div');
			overlay.className = 'gproc-modal-overlay';
			const modal = document.createElement('div');
			modal.className = 'gproc-modal';
			modal.style.maxWidth = '450px';

			const title = document.createElement('h2');
			title.textContent = 'Excel処理モード選択';
			modal.appendChild(title);

			const fileInfo = document.createElement('p');
			fileInfo.innerHTML = '<strong>📄 ファイル:</strong> ' + fileName;
			fileInfo.style.color = '#333';
			fileInfo.style.background = '#f0f0f0';
			fileInfo.style.padding = '10px';
			fileInfo.style.borderRadius = '5px';
			modal.appendChild(fileInfo);

			const desc = document.createElement('p');
			desc.textContent = '処理モードを選択してください';
			desc.style.color = '#666';
			modal.appendChild(desc);

			const btnGroup = document.createElement('div');
			btnGroup.style.display = 'flex';
			btnGroup.style.flexDirection = 'column';
			btnGroup.style.gap = '10px';
			btnGroup.style.marginTop = '20px';

			const btn1 = document.createElement('button');
			btn1.className = 'gproc-btn gproc-btn-confirm';
			btn1.textContent = '📥 G-PROCに直接入力';
			btn1.onclick = function () {
				document.body.removeChild(overlay);
				processExcelToGproc(sheetsData);
			};

			const btn2 = document.createElement('button');
			btn2.className = 'gproc-btn gproc-btn-confirm';
			btn2.textContent = '📤 TXTにエクスポート';
			btn2.onclick = function () {
				document.body.removeChild(overlay);
				processExcelToTxt(fileName, sheetsData);
			};

			const cancelBtn = document.createElement('button');
			cancelBtn.className = 'gproc-btn gproc-btn-cancel';
			cancelBtn.textContent = 'キャンセル';
			cancelBtn.onclick = function () {
				document.body.removeChild(overlay);
			};

			btnGroup.appendChild(btn1);
			btnGroup.appendChild(btn2);
			btnGroup.appendChild(cancelBtn);
			modal.appendChild(btnGroup);
			overlay.appendChild(modal);
			document.body.appendChild(overlay);
		}

		function processExcelToGproc(sheetsData) {
			localStorage.setItem('gproc_excel_data', JSON.stringify(sheetsData));
			let headerItems = [];
			var sheetNames = Object.keys(sheetsData);
			if (sheetNames.length > 0) {
				const firstSheet = sheetsData[sheetNames[0]];
				if (firstSheet && firstSheet.length > 0) {
					const seen = new Set();
					const sectionLabels = ["部品リスト", "材料費", "購入部品費", "支給部品", "二次加工費", "プレス・配管加工費", "単価"];
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
				localStorage.setItem('gproc_excel_mapping', JSON.stringify(mapping));
				window.gprocFillLogic(sheetsData, mapping);
				alert("G-PROCへのインポートが完了しました！");
			});
		}

		function processExcelToTxt(fileName, sheetsData) {
			var txtContent = convertExcelToTxt(sheetsData);
			if (!txtContent) {
				alert('Excelファイルからデータを抽出できませんでした。');
				return;
			}
			var blob = new Blob([txtContent], { type: 'text/plain;charset=utf-8' });
			var url = URL.createObjectURL(blob);
			var a = document.createElement('a');
			a.href = url;
			a.download = 'QuoteDat_' + fileName.replace(/\.[^/.]+$/, "") + '.txt';
			document.body.appendChild(a);
			a.click();
			document.body.removeChild(a);
			URL.revokeObjectURL(url);
			alert('TXTファイルをエクスポートしました！');
		}

		function initExportToTxt() {
			var input = document.createElement('input');
			input.type = 'file';
			input.accept = '.xlsx, .xls';
			input.style.display = 'none';
			document.body.appendChild(input);
			input.onchange = function (e) {
				var file = e.target.files[0];
				if (!file) return;
				var reader = new FileReader();
				reader.onload = function (ev) {
					try {
						var data = new Uint8Array(ev.target.result);
						var workbook = XLSX.read(data, { type: 'array', cellDates: true });
						var sheetsData = {};
						workbook.SheetNames.forEach(function (name) {
							sheetsData[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "" });
						});
						var txtContent = convertExcelToTxt(sheetsData);
						if (!txtContent) {
							alert('Excelファイルからデータを抽出できませんでした。');
							return;
						}
						var blob = new Blob([txtContent], { type: 'text/plain;charset=utf-8' });
						var url = URL.createObjectURL(blob);
						var a = document.createElement('a');
						a.href = url;
						a.download = 'QuoteDat_' + (file.name.replace(/\.[^/.]+$/, "")) + '.txt';
						document.body.appendChild(a);
						a.click();
						document.body.removeChild(a);
						URL.revokeObjectURL(url);
						alert('TXTファイルをエクスポートしました！');
					} catch (err) {
						alert('Excelファイルの解析に失敗しました: ' + err.message);
						console.error(err);
					}
				};
				reader.readAsArrayBuffer(file);
			};
			input.click();
		}

		function convertExcelToTxt(sheetsData) {
			var lines = [];
			var sheetNames = Object.keys(sheetsData);
			if (sheetNames.length === 0) return null;
			var firstSheet = sheetsData[sheetNames[0]];
			if (!firstSheet || firstSheet.length === 0) return null;
			function findCellValue(rows, keyword) {
				for (var i = 0; i < rows.length; i++) {
					var row = rows[i];
					if (!row) continue;
					for (var j = 0; j < row.length; j++) {
						var cell = String(row[j] || '').trim();
						if (cell.includes(keyword)) {
							if (i + 1 < rows.length && rows[i + 1]) {
								var nextRowVal = String(rows[i + 1][j] || '').trim();
								if (nextRowVal && !nextRowVal.includes(keyword)) {
									return nextRowVal;
								}
							}
							for (var k = j + 1; k < row.length; k++) {
								var val = String(row[k] || '').trim();
								if (val) return val;
							}
						}
					}
				}
				return '';
			}

			function findRowWithKeyword(rows, keyword, startFrom) {
				startFrom = startFrom || 0;
				for (var i = startFrom; i < rows.length; i++) {
					var row = rows[i];
					if (!row) continue;
					for (var j = 0; j < row.length; j++) {
						var cell = String(row[j] || '').trim();
						if (cell.includes(keyword)) {
							return i;
						}
					}
				}
				return -1;
			}

			function findColWithKeyword(row, keyword) {
				if (!row) return -1;
				for (var j = 0; j < row.length; j++) {
					var cell = String(row[j] || '').trim();
					if (cell.includes(keyword)) {
						return j;
					}
				}
				return -1;
			}

			var drawingNumber = findCellValue(firstSheet, '図面番号') || 'UNKNOWN';
			var productName = findCellValue(firstSheet, '部品名称') || '';
			var quoteNumber = findCellValue(firstSheet, '見積り数') || '';

			lines.push(formatLine('QUOTE', [
				quoteNumber, '', drawingNumber, '', '', productName, ''
			]));

			var materialSectionRow = findRowWithKeyword(firstSheet, '材料費');
			var sourceCurrency = '', targetCurrency = '', weightUnit = '', yieldRate = '';
			var recycledWeight = '', recoveredWeight = '', recoveredPrice = '', displayPattern = '';
			if (materialSectionRow >= 0) {
				for (var r = materialSectionRow; r < Math.min(materialSectionRow + 5, firstSheet.length); r++) {
					var row = firstSheet[r];
					if (!row) continue;
					var colTanka = findColWithKeyword(row, '単価');
					var colKakaku = findColWithKeyword(row, '価格');
					var colTanjuryo = findColWithKeyword(row, '単重量');
					var colTorisu = findColWithKeyword(row, '取り数');

					if (colTanka >= 0 || colKakaku >= 0) {
						for (var dataR = r + 1; dataR < Math.min(r + 10, firstSheet.length); dataR++) {
							var dataRow = firstSheet[dataR];
							if (!dataRow) continue;
							var firstCell = String(dataRow[0] || '').trim();
							if (!firstCell || firstCell.includes('合計') || firstCell.includes('小計')) continue;

							var hasNumeric = dataRow.some(function (c) {
								var v = String(c || '').trim();
								return v && !isNaN(parseFloat(v));
							});

							if (hasNumeric) {
								sourceCurrency = colTanka >= 0 ? String(dataRow[colTanka] || '') : '';
								targetCurrency = colKakaku >= 0 ? String(dataRow[colKakaku] || '') : '';
								weightUnit = colTanjuryo >= 0 ? String(dataRow[colTanjuryo] || '') : '';
								yieldRate = colTorisu >= 0 ? String(dataRow[colTorisu] || '') : '';
								recycledWeight = weightUnit;
								recoveredWeight = targetCurrency;
								recoveredPrice = sourceCurrency;
								break;
							}
						}
						break;
					}
				}
			}

			lines.push(formatLine('MAT06', [
				'1', '1', '1.000000',
				sourceCurrency || 'JPY', targetCurrency || 'JPY', weightUnit || 'KG',
				'', '', yieldRate, recycledWeight, recoveredWeight, recoveredPrice,
				'', '', '', displayPattern, ''
			]));

			var materialRows = [];
			if (materialSectionRow >= 0) {
				var headerRowIdx = -1;
				var colZumenBango = -1;
				var colZairyohi = -1;

				for (var r = materialSectionRow; r < Math.min(materialSectionRow + 5, firstSheet.length); r++) {
					var row = firstSheet[r];
					if (!row) continue;
					var idx = findColWithKeyword(row, '図面番号');
					if (idx >= 0) {
						headerRowIdx = r;
						colZumenBango = idx;
						colZairyohi = findColWithKeyword(row, '価格');
						if (colZairyohi < 0) colZairyohi = findColWithKeyword(row, '材料費');
						break;
					}
				}

				if (headerRowIdx >= 0) {
					for (var i = headerRowIdx + 1; i < Math.min(headerRowIdx + 20, firstSheet.length); i++) {
						var row = firstSheet[i];
						if (!row) continue;
						var firstCell = String(row[0] || '').trim();
						var secondCell = colZumenBango >= 0 ? String(row[colZumenBango] || '').trim() : '';

						if (firstCell.includes('合計') || firstCell.includes('小計') || firstCell.includes('購入')) break;
						if (secondCell && secondCell.length > 0) {
							materialRows.push({
								code: secondCell,
								price: colZairyohi >= 0 ? String(row[colZairyohi] || '') : ''
							});
						}
					}
				}
			}

			materialRows.forEach(function (item, idx) {
				lines.push(formatLine('MATD06_1', [
					String(idx + 1), '04', item.code, '01', '', '', '', '', '', '', '', item.price, String(idx + 1), ''
				]));
			});

			lines.push(formatLine('PROC03', [
				'1', '射出成形', '', '', '1.000000', 'JPY', 'JPY', '',
				'', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''
			]));
			lines.push(formatLine('PROC02', [
				'1', '加工工程', '', '1.000000', 'JPY', 'JPY', '', '', '', '', '', '', '', '', ''
			]));
			lines.push(formatLine('PART', [
				'1', '1.000000', 'JPY', 'JPY', '', '', '', '', '', '', '', '', '', '', '', ''
			]));
			lines.push(formatLine('PAC01', [
				'1', '1', '1.000000', 'JPY', 'JPY', '', '', '', '', '', '', '', '', ''
			]));
			lines.push(formatLine('LOGI01', [
				'1', '1.000000', 'JPY', 'JPY', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'KG', 'MM', ''
			]));
			lines.push(formatLine('OTHER', [
				'1', '', '1.000000', 'JPY', 'JPY', '', ''
			]));
			lines.push(formatLine('TOOL', [
				'1', '1.000000', 'JPY', 'JPY', '', '', '', '', '', '', '', ''
			]));
			return lines.join('\r\n');
		}
		function formatLine(section, values) {
			var parts = ['"' + section + '"'];
			values.forEach(function (v) {
				parts.push('"' + (v || '') + '"');
			});
			return parts.join('\t');
		}
		function initImportFromTxt() {
			var input = document.createElement('input');
			input.type = 'file';
			input.accept = '.txt';
			input.style.display = 'none';
			document.body.appendChild(input);
			input.onchange = function (e) {
				var file = e.target.files[0];
				var reader = new FileReader();
				reader.onload = function (ev) {
					try {
						var content = ev.target.result;
						var lines = content.split('\n').filter(function (l) { return l.trim(); });
						var data = parseTxtContent(lines);
						localStorage.setItem('gproc_excel_data', JSON.stringify(data));
						gprocFillLogic(data);
						alert('TXTファイルからインポートしました！詳細ボタンをクリックして確認してください。');
					} catch (err) {
						alert('TXTファイルの解析に失敗しました: ' + err.message);
						console.error(err);
					}
				};
				reader.readAsText(file, 'UTF-8');
			};
			input.click();
		}
		function parseTxtContent(lines) {
			var data = {
				header: {},
				materialResin: {},
				materialDetail: [],
				processing3: [],
				processing2: [],
				partsDetail: [],
				packaging: [],
				transport: {},
				otherCost: {},
				moldCost: {},
				management: {},
				profit: {}
			};
			lines.forEach(function (line) {
				var matches = line.match(/"([^"]*)"/g);
				if (!matches) return;
				var cells = matches.map(function (m) { return m.replace(/^"|"$/g, ''); });
				var section = cells[0];

				if (section === 'QUOTE') {
					data.header = {
						quoteNumber: cells[1],
						deptCode: cells[2],
						drawingNumber: cells[3],
						process: cells[4],
						progressive: cells[5],
						productName: cells[6],
						remarks: cells[7]
					};
				} else if (section === 'MAT06') {
					data.materialResin = {
						detailNo: cells[1], quoteNo: cells[2], conversionRate: cells[3], sourceCurrency: cells[4],
						targetCurrency: cells[5], weightUnit: cells[6], productWeight: cells[7], runnerWeight: cells[8],
						yieldRate: cells[9], recycledWeight: cells[10], recoveredWeight: cells[11], recoveredPrice: cells[12],
						productionLot: cells[13], wasteWeight: cells[14], calcPattern: cells[15], displayPattern: cells[16],
						remarks: cells[17]
					};
				} else if (section === 'MATD06_1') {
					data.materialDetail.push({
						id: 'id_' + Math.random().toString(36).substr(2, 9),
						detailNo: cells[1], procurementType: cells[2], materialCode: cells[3], deliveryForm: cells[4],
						supplier: cells[5], manufacturer: cells[6], color: cells[7], grade: cells[8],
						deliveryLot: cells[9], blendingRatio: cells[10], smallLotEx: cells[11], materialPrice: cells[12],
						detailType: cells[13], specification: cells[14]
					});
				} else if (section === 'PROC03') {
					data.processing3.push({
						id: 'id_' + Math.random().toString(36).substr(2, 9),
						quoteNo: cells[1], processName: cells[2], processId: cells[3], equipment: cells[4],
						conversionRate: cells[5], sourceCurrency: cells[6], targetCurrency: cells[7], location: cells[8],
						machineRate: cells[9], machineCT: cells[10], yieldRate: cells[11], changeoverTime: cells[12],
						outputQty: cells[13], calcPattern: cells[14], remarks: cells[15],
						itemName1: cells[16], itemName2: cells[17], itemName3: cells[18], itemName4: cells[19],
						itemName5: cells[20], itemName6: cells[21], value1: cells[22], value2: cells[23],
						value3: cells[24], value4: cells[25], value5: cells[26], value6: cells[27], lotCount: cells[28]
					});
				} else if (section === 'PROC02') {
					data.processing2.push({
						id: 'id_' + Math.random().toString(36).substr(2, 9),
						quoteNo: cells[1], processName: cells[2], processId: cells[3], conversionRate: cells[4],
						sourceCurrency: cells[5], targetCurrency: cells[6], location: cells[7], manRate: cells[8],
						manCT: cells[9], yieldRate: cells[10], changeoverTime: cells[11], outputQty: cells[12],
						personnel: cells[13], calcPattern: cells[14], remarks: cells[15], lotCount: cells[16]
					});
				} else if (section === 'PART') {
					data.partsDetail.push({
						id: 'id_' + Math.random().toString(36).substr(2, 9),
						quoteNo: cells[1], conversionRate: cells[2], sourceCurrency: cells[3], targetCurrency: cells[4],
						partName: cells[5], modelNumber: cells[6], processClass: cells[7], procurementType: cells[8],
						originCountry: cells[9], supplier: cells[10], manufacturer: cells[11], remarks: cells[12],
						quantity: cells[13], unitPrice: cells[14], daikinDrawingNo: cells[15], vendorDrawingNo: cells[16]
					});
				} else if (section === 'PAC01') {
					data.packaging.push({
						id: 'id_' + Math.random().toString(36).substr(2, 9),
						quoteNo: cells[1], detailType: cells[2], conversionRate: cells[3], sourceCurrency: cells[4],
						targetCurrency: cells[5], quantity: cells[6], unit: cells[7], unitPrice: cells[8],
						rotationCount: cells[9], supplier: cells[10], piecesPerBox: cells[11], materialName: cells[12],
						specification: cells[13], remarks: cells[14]
					});
				} else if (section === 'LOGI01') {
					data.transport = {
						quoteNo: cells[1], conversionRate: cells[2], sourceCurrency: cells[3], targetCurrency: cells[4],
						shippingOrigin: cells[5], deliveryDest: cells[6], distance: cells[7], boxHeight: cells[8],
						boxWidth: cells[9], boxDepth: cells[10], weight: cells[11], packageForm: cells[12],
						transportCategory: cells[13], transportMethod: cells[14], logisticsPrice: cells[15],
						boxQuantity: cells[16], transportQty: cells[17], handlingFee: cells[18], otherFee: cells[19],
						weightUnit: cells[20], lengthUnit: cells[21], remarks: cells[22]
					};
				} else if (section === 'OTHER') {
					data.otherCost = {
						quoteNo: cells[1], otherCost: cells[2], conversionRate: cells[3], sourceCurrency: cells[4],
						targetCurrency: cells[5], item: cells[6], remarks: cells[7]
					};
				} else if (section === 'TOOL') {
					data.moldCost = {
						quoteNo: cells[1], conversionRate: cells[2], sourceCurrency: cells[3], targetCurrency: cells[4],
						moldName: cells[5], storageLocation: cells[6], procurementType: cells[7], paymentMethod: cells[8],
						investmentAmount: cells[9], depreciationCount: cells[10], supplier: cells[11], remarks: cells[12]
					};
				}
			});
			return data;
		}
		function gprocFillLogic(data) {
			if (!data || !data.header || !data.header.drawingNumber) {
				console.error('No valid data found');
				return;
			}

			var drawingNumber = data.header.drawingNumber;
			window.currentDrawingNumber = drawingNumber;
			if (!window.detailData[drawingNumber]) {
				window.detailData[drawingNumber] = createEmptyDetailData();
			}
			if (!window.partsData) {
				window.partsData = [];
			}
			var partIdx = window.partsData.findIndex(function (p) { return p.drawingNumber === drawingNumber; });

			if (partIdx === -1) {
				partIdx = window.partsData.findIndex(function (p) { return !p.drawingNumber; });
				if (partIdx === -1) {
					if (window.addPartRow) {
						window.addPartRow();
					} else {
						window.partsData.push({
							id: 'id_' + Math.random().toString(36).substr(2, 9),
							drawingNumber: '',
							quoteNumber: '',
							deptCode: '',
							process: '',
							progressive: '',
							productName: '',
							remarks: ''
						});
					}
					partIdx = window.partsData.length - 1;
				}
			}
			var part = window.partsData[partIdx];
			part.drawingNumber = data.header.drawingNumber;
			part.quoteNumber = data.header.quoteNumber;
			part.deptCode = data.header.deptCode;
			part.process = data.header.process;
			part.progressive = data.header.progressive;
			part.productName = data.header.productName;
			part.remarks = data.header.remarks;
			if (window.renderPartsTable) window.renderPartsTable();
			var details = window.detailData[drawingNumber];
			if (data.materialResin) Object.assign(details.materialResin, data.materialResin);
			if (data.materialDetail && data.materialDetail.length > 0) {
				details.materialDetail = data.materialDetail;
			}
			if (data.processing3 && data.processing3.length > 0) {
				details.processing3 = data.processing3;
			}
			if (data.processing2 && data.processing2.length > 0) {
				details.processing2 = data.processing2;
			}
			if (data.partsDetail && data.partsDetail.length > 0) {
				details.partsDetail = data.partsDetail;
			}
			if (data.packaging && data.packaging.length > 0) {
				details.packaging = data.packaging;
			}
			if (data.transport) Object.assign(details.transport, data.transport);
			if (data.otherCost) Object.assign(details.otherCost, data.otherCost);
			if (data.moldCost) Object.assign(details.moldCost, data.moldCost);
			if (!document.getElementById('detailForm').classList.contains('hidden')) {
				window.viewDetail(drawingNumber);
			}
		}
		function createEmptyDetailData() {
			return {
				materialResin: {},
				materialDetail: [],
				processing3: [],
				processing2: [],
				partsDetail: [],
				packaging: [],
				transport: {},
				otherCost: {},
				moldCost: {},
				management: {},
				profit: {}
			};
		}
	};
})();