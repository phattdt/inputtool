javascript: (function () {
	var lib = document.createElement('script');
	lib.src = 'https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js';
	document.head.appendChild(lib);
	lib.onload = function () {
		function clean(s) {
			return String(s || "").replace(/\s/g, "");
		}
		function findRowByLabel(rows, label) {
			const target = clean(label);
			for (let r = 0; r < rows.length; r++) {
				if (clean(rows[r][0]).includes(target)) return r;
			}
			return -1;
		}
		window.gprocFillLogic = function (sheetsData) {
			const PERF_START = performance.now();
			console.log('⏱️  [PERF] Bookmarklet parsing started');
			const detailForm = document.getElementById('detailForm');
			const isDetailView = detailForm && !detailForm.classList.contains('hidden');
			const sheetNames = Object.keys(sheetsData);
			if (!isDetailView) {
				sheetNames.forEach((name) => {
					var rows = sheetsData[name];
					var record = {};
					["図面番号", "部品名称", "代表機種", "見積り数", "見積書＃", "適用時期", "適用機種名"].forEach(f => {
						for (let r = 0; r < 15; r++) {
							for (let c = 0; c < (rows[r]?.length || 0); c++) {
								if (clean(rows[r][c]) === clean(f)) {
									let val = rows[r + 1] ? rows[r + 1][c] : "";
									if (val !== "" && record[f] == null) {
										record[f] = val;
									}
								}
							}
						}
					});
					if (!record["図面番号"] && record["適用機種名"]) {
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
						ins[5].value = record["適用時期"] || "";
						ins.forEach(inp => inp.dispatchEvent(new Event('change', { bubbles: true })));
					}
				});
			} else {
				const headerEl = document.getElementById('drawingNumberHeader');
				const currentDN = clean(headerEl ? headerEl.textContent.replace('図面番号:', '') : "");
				const targetSheetName = sheetNames.find(name => sheetsData[name].some(r => r.some(c => clean(c).includes(currentDN)))) || sheetNames[0];
				const rows = sheetsData[targetSheetName];
				const drawingNumber = window.currentDrawingNumber;
				if (!drawingNumber || !window.detailData || !window.detailData[drawingNumber]) {
					console.error('No detailData found for', drawingNumber);
					return;
				}

				function extractSectionRows(label, offset, maxRows) {
					const rIdx = findRowByLabel(rows, label);
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
				const materialRows = extractSectionRows("材料費", 2, 20);
				materialRows.forEach((exR, i) => {
					while (window.detailData[drawingNumber].material.length <= i) {
						window.detailData[drawingNumber].material.push({
							id: 'id_' + Math.random().toString(36).substr(2, 9),
							drawingNumber: '', temper: '', diameter: '', thickness: '', length: '',
							weight: '', unitPriceKg: '', consumptionPrice: '', materialCost: '',
							processingUnitPrice: '', processingPrice: '', pipingPrice: '', required: ''
						});
					}
					const m = window.detailData[drawingNumber].material[i];
					m.drawingNumber = exR[1] || '';
					m.temper = exR[5] || '';
					m.diameter = exR[6] || '';
					m.thickness = exR[7] || '';
					m.length = exR[8] || '';
					m.weight = exR[9] || '';
					m.unitPriceKg = exR[10] || '';
					m.consumptionPrice = exR[11] || '';
					m.materialCost = exR[12] || '';
					m.processingUnitPrice = exR[13] || '';
					m.processingPrice = exR[14] || '';
					m.pipingPrice = exR[15] || '';
					m.required = exR[16] || '';
				});
				const purchaseRows = extractSectionRows("購入部品費", 1, 20);
				let purchaseIndex = 0;
				purchaseRows.forEach((exR, i) => {
					const leftPartName = exR[1] || '';
					const leftDrawingNum = exR[5] || '';
					const leftRequired = exR[8] || '';
					const leftUnitPrice = exR[9] || '';

					if (leftPartName || leftDrawingNum || leftRequired || leftUnitPrice) {
						while (window.detailData[drawingNumber].purchase.length <= purchaseIndex) {
							window.detailData[drawingNumber].purchase.push({
								id: 'id_' + Math.random().toString(36).substr(2, 9),
								partName: '', drawingNum: '', required: '', unitPrice: ''
							});
						}
						const p = window.detailData[drawingNumber].purchase[purchaseIndex];
						p.partName = leftPartName;
						p.drawingNum = leftDrawingNum;
						p.required = leftRequired;
						p.unitPrice = leftUnitPrice;
						purchaseIndex++;
					}
					const rightPartName = exR[10] || '';
					const rightDrawingNum = exR[12] || '';
					const rightRequired = exR[15] || '';
					const rightUnitPrice = exR[16] || '';

					if (rightPartName || rightDrawingNum || rightRequired || rightUnitPrice) {
						while (window.detailData[drawingNumber].purchase.length <= purchaseIndex) {
							window.detailData[drawingNumber].purchase.push({
								id: 'id_' + Math.random().toString(36).substr(2, 9),
								partName: '', drawingNum: '', required: '', unitPrice: ''
							});
						}
						const p = window.detailData[drawingNumber].purchase[purchaseIndex];
						p.partName = rightPartName;
						p.drawingNum = rightDrawingNum;
						p.required = rightRequired;
						p.unitPrice = rightUnitPrice;
						purchaseIndex++;
					}
				});
				const suppliedRows = extractSectionRows("支給部品", 1, 20);
				let suppliedIndex = 0;
				suppliedRows.forEach((exR, i) => {
					const leftPartName = exR[1] || '';
					const leftDrawingNum = exR[5] || '';
					const leftRequired = exR[8] || '';

					if (leftPartName || leftDrawingNum || leftRequired) {
						while (window.detailData[drawingNumber].supplied.length <= suppliedIndex) {
							window.detailData[drawingNumber].supplied.push({
								id: 'id_' + Math.random().toString(36).substr(2, 9),
								partName: '', drawingNum: '', required: ''
							});
						}
						const sp = window.detailData[drawingNumber].supplied[suppliedIndex];
						sp.partName = leftPartName;
						sp.drawingNum = leftDrawingNum;
						sp.required = leftRequired;
						suppliedIndex++;
					}
					const rightPartName = exR[10] || '';
					const rightDrawingNum = exR[12] || '';
					const rightRequired = exR[15] || '';

					if (rightPartName || rightDrawingNum || rightRequired) {
						while (window.detailData[drawingNumber].supplied.length <= suppliedIndex) {
							window.detailData[drawingNumber].supplied.push({
								id: 'id_' + Math.random().toString(36).substr(2, 9),
								partName: '', drawingNum: '', required: ''
							});
						}
						const sp = window.detailData[drawingNumber].supplied[suppliedIndex];
						sp.partName = rightPartName;
						sp.drawingNum = rightDrawingNum;
						sp.required = rightRequired;
						suppliedIndex++;
					}
				});
				const secondaryRows = extractSectionRows("二次加工費", 1, 20);
				secondaryRows.forEach((exR, i) => {
					while (window.detailData[drawingNumber].secondary.length <= i) {
						window.detailData[drawingNumber].secondary.push({
							id: 'id_' + Math.random().toString(36).substr(2, 9),
							process: '', qty: '', seconds: '', unitPrice: '', total: ''
						});
					}
					const s = window.detailData[drawingNumber].secondary[i];
					s.process = exR[1] || '';
					s.qty = exR[6] || '';
					s.seconds = exR[7] || '';
					s.unitPrice = exR[8] || '';
					s.total = exR[9] || '';
				});
				const breakdownIdx = findRowByLabel(rows, "単価");
				console.log('🔍 DEBUG: Looking for 単価 section...');
				console.log('  breakdownIdx:', breakdownIdx);
				if (breakdownIdx !== -1) {
					console.log('  Found 単価 at row:', breakdownIdx);
					console.log('  Row content:', rows[breakdownIdx]);
					console.log('  Row +1:', rows[breakdownIdx + 1]);
					console.log('  Row +2:', rows[breakdownIdx + 2]);
					console.log('  Row +3:', rows[breakdownIdx + 3]);
					if (!window.detailData[drawingNumber].breakdown) {
						window.detailData[drawingNumber].breakdown = {};
					}
					const bd = window.detailData[drawingNumber].breakdown;
					const valuesRow = rows[breakdownIdx + 3];
					if (valuesRow) {
						console.log('  Parsing breakdown from row +3:', valuesRow);
						bd.materialTotal = valuesRow[3] || "";
						bd.materialMgmt = valuesRow[5] || "";
						bd.procTotal = valuesRow[7] || "";
						bd.genSales = valuesRow[9] || "";
						bd.profit = valuesRow[11] || "";
						bd.freight = valuesRow[13] || "";
						bd.materialMgmt14 = "";
						bd.genSales15 = "";
						bd.total = valuesRow[15] || "";
						console.log('  ✅ Breakdown data:', bd);
					} else {
						console.log('  ❌ No valuesRow found at +3 offset');
					}
				} else {
					console.log('  ❌ 単価 section not found in Excel');
					console.log('  Searched in', rows.length, 'rows');
					console.log('  First 20 row labels:', rows.slice(0, 20).map((r, i) => `${i}: ${clean(r[0])}`));
				}

				if (window.renderDetailTables) {
					window.renderDetailTables();
				}
				const PERF_END = performance.now();
				const totalTime = (PERF_END - PERF_START).toFixed(2);
				console.log(`⏱️  [PERF] Detail parsing completed in ${totalTime}ms`);
			}
			const PERF_TOTAL = performance.now();
			console.log(`⏱️  [PERF] Total bookmarklet execution: ${(PERF_TOTAL - PERF_START).toFixed(2)}ms`);
		};
		const detailForm = document.getElementById('detailForm');
		const isDetailView = detailForm && !detailForm.classList.contains('hidden');
		const savedData = localStorage.getItem('gproc_excel_data');
		if (isDetailView) {
			if (savedData) {
				window.gprocFillLogic(JSON.parse(savedData));
			} else {
				var input = document.createElement('input');
				input.type = 'file';
				input.accept = '.xlsx, .xls';
				input.style.display = 'none';
				document.body.appendChild(input);
				input.onchange = function (e) {
					const FILE_START = performance.now();
					console.log('📁 [PERF] File selected, starting upload...');
					var file = e.target.files[0];
					var reader = new FileReader();
					reader.onload = function (ev) {
						const READ_END = performance.now();
						console.log(`📖 [PERF] File read completed in ${(READ_END - FILE_START).toFixed(2)}ms`);

						const PARSE_START = performance.now();
						var data = new Uint8Array(ev.target.result);
						var workbook = XLSX.read(data, { type: 'array', cellDates: true });
						var storageObj = {};
						workbook.SheetNames.forEach(name => {
							storageObj[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "" });
						});
						const PARSE_END = performance.now();
						console.log(`🔄 [PERF] Excel parsed in ${(PARSE_END - PARSE_START).toFixed(2)}ms`);

						localStorage.setItem('gproc_excel_data', JSON.stringify(storageObj));
						window.gprocFillLogic(storageObj);

						const TOTAL_END = performance.now();
						console.log(`✅ [PERF] TOTAL TIME (file select → fill complete): ${(TOTAL_END - FILE_START).toFixed(2)}ms`);
						alert("アップロード成功！");
					};
					reader.readAsArrayBuffer(file);
				};
				input.click();
			}
		} else {
			var input = document.createElement('input');
			input.type = 'file';
			input.accept = '.xlsx, .xls';
			input.style.display = 'none';
			document.body.appendChild(input);
			input.onchange = function (e) {
				const FILE_START = performance.now();
				console.log('📁 [PERF] File selected, starting upload...');
				var file = e.target.files[0];
				var reader = new FileReader();
				reader.onload = function (ev) {
					const READ_END = performance.now();
					console.log(`📖 [PERF] File read completed in ${(READ_END - FILE_START).toFixed(2)}ms`);

					const PARSE_START = performance.now();
					var data = new Uint8Array(ev.target.result);
					var workbook = XLSX.read(data, { type: 'array', cellDates: true });
					var storageObj = {};
					workbook.SheetNames.forEach(name => {
						storageObj[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "" });
					});
					const PARSE_END = performance.now();
					console.log(`🔄 [PERF] Excel parsed in ${(PARSE_END - PARSE_START).toFixed(2)}ms`);

					localStorage.setItem('gproc_excel_data', JSON.stringify(storageObj));
					window.gprocFillLogic(storageObj);

					const TOTAL_END = performance.now();
					console.log(`✅ [PERF] TOTAL TIME (file select → fill complete): ${(TOTAL_END - FILE_START).toFixed(2)}ms`);
					alert("アップロード成功！");
				};
				reader.readAsArrayBuffer(file);
			};
			input.click();
		}
	};
})();