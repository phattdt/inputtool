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
					title: "見積書リスト (Parts List)",
					sectionKey: "ヘッダー",
					fields: [
						{ key: "parts_drawingNumber", label: "図面番号", autoMatch: ["図面番号", "部品名"] },
						{ key: "parts_quoteNumber", label: "見積書番号", autoMatch: ["見積書番号", "見積書＃", "見積番号"] },
						{ key: "parts_deptCode", label: "事部コード", autoMatch: ["事部コード", "部門コード"] },
						{ key: "parts_process", label: "工程", autoMatch: ["工程"] },
						{ key: "parts_progressive", label: "累進", autoMatch: ["累進"] },
						{ key: "parts_productName", label: "品名仕様", autoMatch: ["品名仕様", "品名"] },
						{ key: "parts_remarks", label: "備考", autoMatch: ["備考"] }
					]
				},
				{
					title: "材料費（配管）(Material Piping)",
					sectionKey: "材料費",
					fields: [
						{ key: "mat_recordType", label: "レコード区分", autoMatch: ["レコード区分"] },
						{ key: "mat_detailNo", label: "明細番号", autoMatch: ["明細書番号"] },
						{ key: "mat_quoteNumber", label: "見積書番号", autoMatch: ["見積書番号", "見積書＃", "見積番号"] },
						{ key: "mat_conversionRate", label: "換算レート", autoMatch: ["換算レート"] },
						{ key: "mat_sourceCurrency", label: "換算元通貨", autoMatch: ["換算元通貨"] },
						{ key: "mat_targetCurrency", label: "換算先通貨", autoMatch: ["換算先通貨"] },
						{ key: "mat_weightUnit", label: "重量単位", autoMatch: ["重量単位"] },
						{ key: "mat_lengthUnit", label: "長さ単位", autoMatch: ["長さ単位"] },
						{ key: "mat_specificGravity", label: "比重", autoMatch: ["比重", "材質"] },
						{ key: "mat_quantity", label: "取り数", autoMatch: ["取り数", "取り数"] },
						{ key: "mat_outerDiameter", label: "外径", autoMatch: ["外径"] },
						{ key: "mat_thickness", label: "肉厚", autoMatch: ["肉厚"] },
						{ key: "mat_length", label: "長さ", autoMatch: ["長さ"] },
						{ key: "mat_productWeight", label: "製品重量", autoMatch: ["製品重量", "重量"] },
						{ key: "mat_lossWeight", label: "ロス重量", autoMatch: ["ロス重量"] },
						{ key: "mat_scpUnitPrice", label: "SCP単価", autoMatch: ["SCP単価"] },
						{ key: "mat_weightPerPiece", label: "1個あたり重量（直接入力）", autoMatch: ["1個あたり重量", "直接入力", "台 り 材 料 費合 計"] },
						{ key: "mat_calcPattern", label: "計算パターン", autoMatch: ["計算パターン"] },
						{ key: "mat_displayPattern", label: "表示パターン", autoMatch: ["表示パターン"] },
						{ key: "mat_remarks", label: "備考", autoMatch: ["備考"] }
					]
				},
				{
					title: "材料費（樹脂）明細 (Material Detail)",
					sectionKey: "材料費明細",
					fields: [
						{ key: "matd_recordType", label: "レコード区分", autoMatch: ["レコード区分"] },
						{ key: "matd_required", label: "必須", autoMatch: ["必須"] },
						{ key: "matd_priority", label: "優先", autoMatch: ["優先"] },
						{ key: "matd_appType", label: "アプリ区分", autoMatch: ["アプリ区分"] },
						{ key: "matd_specificGravity", label: "比重", autoMatch: ["比重"] },
						{ key: "matd_smallLot", label: "小ロット", autoMatch: ["小ロット"] },
						{ key: "matd_deliveryPrice", label: "納入価格", autoMatch: ["納入価格"] },
						{ key: "matd_codeName", label: "コード名称", autoMatch: ["コード名称"] },
						{ key: "matd_lotProduction", label: "ロット生産数", autoMatch: ["ロット生産数"] },
						{ key: "matd_remarks", label: "備考", autoMatch: ["備考"] }
					]
				},
				{
					title: "加工費（パターン３）(Processing Pattern 3)",
					sectionKey: "プレス・配管加工費",
					fields: [
						{ key: "proc3_processName", label: "加工工程", autoMatch: ["加工工程"] },
						{ key: "proc3_processId", label: "加工工程識別名", autoMatch: ["加工工程識別名", "識別名"] },
						{ key: "proc3_equipment", label: "使用設備", autoMatch: ["使用設備"] },
						{ key: "proc3_conversionRate", label: "換算レート", autoMatch: ["換算レート"] },
						{ key: "proc3_machineRate", label: "マシンレート", autoMatch: ["マシンレート"] },
						{ key: "proc3_machineCT", label: "マシンコストCT", autoMatch: ["マシンコストCT", "CT"] },
						{ key: "proc3_yieldRate", label: "歩留り", autoMatch: ["歩留り", "歩留"] },
						{ key: "proc3_changeoverTime", label: "段替時間", autoMatch: ["段替時間", "段取り"] },
						{ key: "proc3_outputQty", label: "取り数", autoMatch: ["取り数", "所要数"] },
						{ key: "proc3_calcPattern", label: "計算パターン", autoMatch: ["計算パターン"] },
						{ key: "proc3_cost", label: "加工費", autoMatch: ["加工費"] },
						{ key: "proc3_remarks", label: "備考", autoMatch: ["備考"] }
					]
				},
				{
					title: "加工費（パターン２）(Processing Pattern 2)",
					sectionKey: "二次加工費",
					fields: [
						{ key: "proc2_recordType", label: "レコード区分", autoMatch: ["レコード区分"] },
						{ key: "proc2_quoteNumber", label: "見積書番号", autoMatch: ["見積書番号"] },
						{ key: "proc2_processName", label: "加工工程", autoMatch: ["加工工程"] },
						{ key: "proc2_processId", label: "加工工程識別名", autoMatch: ["加工工程識別名", "識別名", "部品名（金型費）"] },
						{ key: "proc2_conversionRate", label: "換算レート", autoMatch: ["換算レート"] },
						{ key: "proc2_sourceCurrency", label: "換算元通貨", autoMatch: ["換算元通貨"] },
						{ key: "proc2_targetCurrency", label: "換算先通貨", autoMatch: ["換算先通貨"] },
						{ key: "proc2_location", label: "拠点", autoMatch: ["拠点"] },
						{ key: "proc2_manRate", label: "マンレート（@/Hr）", autoMatch: ["マンレート", "チャージ", "@/Hr"] },
						{ key: "proc2_manCT", label: "マンコストCT（s）", autoMatch: ["マンコストCT", "サイクルタイム", "s"] },
						{ key: "proc2_yieldRate", label: "歩留り(%)", autoMatch: ["歩留り", "歩留", "%"] },
						{ key: "proc2_changeoverTime", label: "段替時間（s）", autoMatch: ["段替時間", "段取り", "s"] },
						{ key: "proc2_outputQty", label: "取り数", autoMatch: ["取り数", "所要数"] },
						{ key: "proc2_personnel", label: "人数", autoMatch: ["人数", "人員"] },
						{ key: "proc2_calcPattern", label: "計算パターン", autoMatch: ["計算パターン"] },
						{ key: "proc2_remarks", label: "備考", autoMatch: ["備考"] },
						{ key: "proc2_lotCount", label: "ロット数", autoMatch: ["ロット数"] }
					]
				},
				{
					title: "部品費 (Parts Cost)",
					sectionKey: "購入部品費",
					fields: [
						{ key: "part_recordType", label: "レコード区分", autoMatch: ["レコード区分"] },
						{ key: "part_quoteNumber", label: "見積書番号", autoMatch: ["見積書番号"] },
						{ key: "part_conversionRate", label: "換算レート", autoMatch: ["換算レート"] },
						{ key: "part_sourceCurrency", label: "換算元通貨", autoMatch: ["換算元通貨"] },
						{ key: "part_targetCurrency", label: "換算先通貨", autoMatch: ["換算先通貨"] },
						{ key: "part_partName", label: "品名", autoMatch: ["品名", "部品名"] },
						{ key: "part_modelNumber", label: "型式番号", autoMatch: ["型式番号", "部品名及び仕様"] },
						{ key: "part_processClass", label: "工程分類", autoMatch: ["工程分類"] },
						{ key: "part_procurementType", label: "調達区分", autoMatch: ["調達区分"] },
						{ key: "part_originCountry", label: "原産国", autoMatch: ["原産国"] },
						{ key: "part_supplier", label: "購入元", autoMatch: ["購入元"] },
						{ key: "part_manufacturer", label: "メーカー", autoMatch: ["メーカー"] },
						{ key: "part_remarks", label: "備考", autoMatch: ["備考"] },
						{ key: "part_quantity", label: "所要量", autoMatch: ["所要量", "所要数"] },
						{ key: "part_unitPrice", label: "単価", autoMatch: ["単価"] },
						{ key: "part_daikinDrawingNo", label: "ダイキン図面番号", autoMatch: ["ダイキン図面番号", "明細書"] },
						{ key: "part_vendorDrawingNo", label: "取引先図面番号", autoMatch: ["取引先図面番号"] }
					]
				},
				{
					title: "梱包費 (Packaging)",
					sectionKey: "梱包費",
					fields: [
						{ key: "pac_detailType", label: "明細区分", autoMatch: ["明細区分"] },
						{ key: "pac_quantity", label: "所要量", autoMatch: ["所要量", "所要数"] },
						{ key: "pac_unit", label: "単位", autoMatch: ["単位"] },
						{ key: "pac_unitPrice", label: "単価", autoMatch: ["単価"] },
						{ key: "pac_rotationCount", label: "回転数", autoMatch: ["回転数"] },
						{ key: "pac_supplier", label: "購入元", autoMatch: ["購入元"] },
						{ key: "pac_piecesPerBox", label: "入り数", autoMatch: ["入り数"] },
						{ key: "pac_materialName", label: "材料名", autoMatch: ["材料名"] },
						{ key: "pac_specification", label: "仕様", autoMatch: ["仕様"] },
						{ key: "pac_remarks", label: "備考", autoMatch: ["備考"] }
					]
				},
				{
					title: "輸送費 (Transport)",
					sectionKey: "輸送費",
					fields: [
						{ key: "trans_shippingOrigin", label: "出荷元", autoMatch: ["出荷元"] },
						{ key: "trans_deliveryDest", label: "納入先", autoMatch: ["納入先"] },
						{ key: "trans_distance", label: "輸送距離", autoMatch: ["輸送距離"] },
						{ key: "trans_boxHeight", label: "箱サイズ（タテ）", autoMatch: ["タテ", "箱サイズ"] },
						{ key: "trans_boxWidth", label: "箱サイズ（ヨコ）", autoMatch: ["ヨコ"] },
						{ key: "trans_boxDepth", label: "箱サイズ（タカサ）", autoMatch: ["タカサ"] },
						{ key: "trans_weight", label: "重量", autoMatch: ["重量"] },
						{ key: "trans_packageForm", label: "荷姿", autoMatch: ["荷姿"] },
						{ key: "trans_transportCategory", label: "輸送カテゴリー", autoMatch: ["輸送カテゴリー"] },
						{ key: "trans_transportMethod", label: "輸送手段", autoMatch: ["輸送手段"] },
						{ key: "trans_logisticsPrice", label: "物流単価", autoMatch: ["物流単価"] },
						{ key: "trans_boxQuantity", label: "箱入り数", autoMatch: ["箱入り数"] },
						{ key: "trans_transportQty", label: "輸送数量", autoMatch: ["輸送数量"] },
						{ key: "trans_handlingFee", label: "荷役費", autoMatch: ["荷役費"] },
						{ key: "trans_otherFee", label: "その他費", autoMatch: ["その他費"] },
						{ key: "trans_remarks", label: "備考", autoMatch: ["備考"] }
					]
				},
				{
					title: "その他費 (Other Cost)",
					sectionKey: "その他費",
					fields: [
						{ key: "other_cost", label: "その他費", autoMatch: ["その他費"] },
						{ key: "other_item", label: "項目", autoMatch: ["項目"] },
						{ key: "other_remarks", label: "備考", autoMatch: ["備考"] }
					]
				},
				{
					title: "金型費 (Mold Cost)",
					sectionKey: "金型費",
					fields: [
						{ key: "mold_recordType", label: "レコード区分", autoMatch: ["レコード区分"] },
						{ key: "mold_quoteNumber", label: "見積書番号", autoMatch: ["見積書番号"] },
						{ key: "mold_conversionRate", label: "換算レート", autoMatch: ["換算レート"] },
						{ key: "mold_sourceCurrency", label: "換算元通貨", autoMatch: ["換算元通貨"] },
						{ key: "mold_targetCurrency", label: "換算先通貨", autoMatch: ["換算先通貨"] },
						{ key: "mold_moldName", label: "金型名称", autoMatch: ["金型名称", "金型費", "代表機種"] },
						{ key: "mold_storageLocation", label: "金型保管場所", autoMatch: ["保管場所"] },
						{ key: "mold_procurementType", label: "調達区分", autoMatch: ["調達区分"] },
						{ key: "mold_paymentMethod", label: "支払方法", autoMatch: ["支払方法"] },
						{ key: "mold_investmentAmount", label: "投資額", autoMatch: ["投資額"] },
						{ key: "mold_depreciationCount", label: "償却数", autoMatch: ["償却数", "尼崎パイプ製作所"] },
						{ key: "mold_supplier", label: "購入元", autoMatch: ["購入元"] },
						{ key: "mold_remarks", label: "備考", autoMatch: ["備考"] }
					]
				},
				{
					title: "管理費 (EXPENSE)",
					sectionKey: "管理費",
					fields: [
						{ key: "mgmt_recordType", label: "レコード区分", autoMatch: ["レコード区分"] },
						{ key: "mgmt_quoteNumber", label: "見積書番号", autoMatch: ["見積書番号"] },
						{ key: "mgmt_expenseCode", label: "費目コード", autoMatch: ["費目コード"] },
						{ key: "mgmt_managementFee", label: "管理費", autoMatch: ["管理費", "材料管理費", "一般販売管理費"] },
						{ key: "mgmt_ratio", label: "比率", autoMatch: ["比率"] }
					]
				},
				{
					title: "利益 (Profit)",
					sectionKey: "利益",
					fields: [
						{ key: "profit_recordType", label: "レコード区分", autoMatch: ["レコード区分"] },
						{ key: "profit_quoteNumber", label: "見積書番号", autoMatch: ["見積書番号"] },
						{ key: "profit_expenseCode", label: "費目コード", autoMatch: ["費目コード"] },
						{ key: "profit_profitAmount", label: "利益", autoMatch: ["利益"] },
						{ key: "profit_ratio", label: "比率", autoMatch: ["比率"] }
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
					record["drawingNumber"] = getValSmart("parts_drawingNumber");
					record["quoteNumber"] = getValSmart("parts_quoteNumber");
					record["deptCode"] = getValSmart("parts_deptCode");
					record["process"] = getValSmart("parts_process");
					record["progressive"] = getValSmart("parts_progressive");
					record["productName"] = getValSmart("parts_productName");
					record["remarks"] = getValSmart("parts_remarks");

					if (!record["drawingNumber"]) return;

					var targetRow = Array.from(document.querySelectorAll('#partsTableBody tr')).find(r => !r.querySelector('input').value);
					if (!targetRow && window.addPartRow) {
						window.addPartRow();
						targetRow = document.querySelector('#partsTableBody tr:last-child');
					}
					if (targetRow) {
						var ins = targetRow.querySelectorAll('input');
						ins[0].value = record["drawingNumber"] || "";
						ins[1].value = record["quoteNumber"] || "";
						ins[2].value = record["deptCode"] || "";
						ins[3].value = record["process"] || "";
						ins[4].value = record["progressive"] || "";
						ins[5].value = record["productName"] || "";
						ins[6].value = record["remarks"] || "";
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

				const dd = window.detailData[drawingNumber];
				const matResinRow = extractSectionRows("材料費", 1, 1)[0];
				if (matResinRow) {
					const h = getLocalHeader("材料費", "製品重量", mainHeaderRow);
					dd.materialResin.detailNo = getVal(matResinRow, h, "mat_detailNo", mapping);
					dd.materialResin.conversionRate = getValNum(matResinRow, h, "mat_conversionRate", mapping) || "1";
					dd.materialResin.sourceCurrency = getVal(matResinRow, h, "mat_sourceCurrency", mapping) || "JPY";
					dd.materialResin.targetCurrency = getVal(matResinRow, h, "mat_targetCurrency", mapping) || "JPY";
					dd.materialResin.weightUnit = getVal(matResinRow, h, "mat_weightUnit", mapping) || "KG";
					dd.materialResin.productWeight = getValNum(matResinRow, h, "mat_productWeight", mapping);
					dd.materialResin.runnerWeight = getValNum(matResinRow, h, "mat_runnerWeight", mapping);
					dd.materialResin.yieldRate = getValNum(matResinRow, h, "mat_yieldRate", mapping);
					dd.materialResin.recycledWeight = getValNum(matResinRow, h, "mat_recycledWeight", mapping);
					dd.materialResin.recoveredWeight = getValNum(matResinRow, h, "mat_recoveredWeight", mapping);
					dd.materialResin.recoveredPrice = getValNum(matResinRow, h, "mat_recoveredPrice", mapping);
					dd.materialResin.productionLot = getValNum(matResinRow, h, "mat_productionLot", mapping);
					dd.materialResin.wasteWeight = getValNum(matResinRow, h, "mat_wasteWeight", mapping);
					dd.materialResin.calcPattern = getVal(matResinRow, h, "mat_calcPattern", mapping);
					dd.materialResin.displayPattern = getVal(matResinRow, h, "mat_displayPattern", mapping);
					dd.materialResin.remarks = getVal(matResinRow, h, "mat_remarks", mapping);
				}

				const matDetailRows = extractSectionRows("材料費", 2, 20);
				const matDetailHeader = getLocalHeader("材料費", "調達区分", mainHeaderRow);
				matDetailRows.forEach((exR, i) => {
					while (dd.materialDetail.length <= i) {
						dd.materialDetail.push(window.createEmptyMaterialDetail ? window.createEmptyMaterialDetail() : { id: 'id_' + Math.random().toString(36).substr(2, 9) });
					}
					const m = dd.materialDetail[i];
					const h = matDetailHeader;
					m.procurementType = getVal(exR, h, "matd_procurementType", mapping);
					m.materialCode = getVal(exR, h, "matd_materialCode", mapping);
					m.deliveryForm = getVal(exR, h, "matd_deliveryForm", mapping);
					m.supplier = getVal(exR, h, "matd_supplier", mapping);
					m.manufacturer = getVal(exR, h, "matd_manufacturer", mapping);
					m.color = getVal(exR, h, "matd_color", mapping);
					m.grade = getVal(exR, h, "matd_grade", mapping);
					m.deliveryLot = getValNum(exR, h, "matd_deliveryLot", mapping);
					m.blendingRatio = getValNum(exR, h, "matd_blendingRatio", mapping);
					m.smallLotEx = getValNum(exR, h, "matd_smallLotEx", mapping);
					m.materialPrice = getValNum(exR, h, "matd_materialPrice", mapping);
					m.detailType = getVal(exR, h, "matd_detailType", mapping);
					m.specification = getVal(exR, h, "matd_specification", mapping);
				});

				const proc3Rows = extractSectionRows("プレス・配管加工費", 3, 20);
				const proc3Header = getLocalHeader("プレス・配管加工費", "加工工程", mainHeaderRow);
				let proc3Idx = 0;
				proc3Rows.forEach((exR) => {
					while (dd.processing3.length <= proc3Idx) {
						dd.processing3.push(window.createEmptyProcessing3 ? window.createEmptyProcessing3() : { id: 'id_' + Math.random().toString(36).substr(2, 9) });
					}
					const p = dd.processing3[proc3Idx];
					const h = proc3Header;
					p.processName = getVal(exR, h, "proc3_processName", mapping);
					p.processId = getVal(exR, h, "proc3_processId", mapping);
					p.equipment = getVal(exR, h, "proc3_equipment", mapping);
					p.conversionRate = getValNum(exR, h, "proc3_conversionRate", mapping) || "1";
					p.machineRate = getValNum(exR, h, "proc3_machineRate", mapping);
					p.machineCT = getValNum(exR, h, "proc3_machineCT", mapping);
					p.yieldRate = getValNum(exR, h, "proc3_yieldRate", mapping);
					p.changeoverTime = getValNum(exR, h, "proc3_changeoverTime", mapping);
					p.outputQty = getValNum(exR, h, "proc3_outputQty", mapping);
					p.calcPattern = getVal(exR, h, "proc3_calcPattern", mapping) || "3";
					p.value1 = getValNum(exR, h, "proc3_cost", mapping);
					p.remarks = getVal(exR, h, "proc3_remarks", mapping);
					if (p.processName || p.value1) proc3Idx++;
				});

				const proc2Rows = extractSectionRows("二次加工費", 1, 20);
				const proc2Header = getLocalHeader("二次加工費", "加工工程", mainHeaderRow);
				let proc2Idx = 0;
				proc2Rows.forEach((exR) => {
					while (dd.processing2.length <= proc2Idx) {
						dd.processing2.push(window.createEmptyProcessing2 ? window.createEmptyProcessing2() : { id: 'id_' + Math.random().toString(36).substr(2, 9) });
					}
					const p = dd.processing2[proc2Idx];
					const h = proc2Header;
					p.processName = getVal(exR, h, "proc2_processName", mapping);
					p.processId = getVal(exR, h, "proc2_processId", mapping);
					p.conversionRate = getValNum(exR, h, "proc2_conversionRate", mapping) || "1";
					p.manRate = getValNum(exR, h, "proc2_manRate", mapping);
					p.manCT = getValNum(exR, h, "proc2_manCT", mapping);
					p.yieldRate = getValNum(exR, h, "proc2_yieldRate", mapping);
					p.changeoverTime = getValNum(exR, h, "proc2_changeoverTime", mapping);
					p.outputQty = getValNum(exR, h, "proc2_outputQty", mapping);
					p.personnel = getValNum(exR, h, "proc2_personnel", mapping);
					p.calcPattern = getVal(exR, h, "proc2_calcPattern", mapping) || "2";
					p.remarks = getVal(exR, h, "proc2_remarks", mapping);
					p.lotCount = getValNum(exR, h, "proc2_lotCount", mapping);
					if (p.processName) proc2Idx++;
				});

				const partsRows = extractSectionRows("購入部品費", 1, 20);
				const partsHeader = getLocalHeader("購入部品費", "品名", mainHeaderRow);
				let partsIdx = 0;
				partsRows.forEach((exR) => {
					while (dd.partsDetail.length <= partsIdx) {
						dd.partsDetail.push(window.createEmptyPartsDetail ? window.createEmptyPartsDetail() : { id: 'id_' + Math.random().toString(36).substr(2, 9) });
					}
					const p = dd.partsDetail[partsIdx];
					const h = partsHeader;
					p.partName = getVal(exR, h, "part_partName", mapping);
					p.modelNumber = getVal(exR, h, "part_modelNumber", mapping);
					p.processClass = getVal(exR, h, "part_processClass", mapping);
					p.procurementType = getVal(exR, h, "part_procurementType", mapping);
					p.originCountry = getVal(exR, h, "part_originCountry", mapping);
					p.supplier = getVal(exR, h, "part_supplier", mapping);
					p.manufacturer = getVal(exR, h, "part_manufacturer", mapping);
					p.remarks = getVal(exR, h, "part_remarks", mapping);
					p.quantity = getValNum(exR, h, "part_quantity", mapping);
					p.unitPrice = getValNum(exR, h, "part_unitPrice", mapping);
					p.daikinDrawingNo = getVal(exR, h, "part_daikinDrawingNo", mapping);
					p.vendorDrawingNo = getVal(exR, h, "part_vendorDrawingNo", mapping);
					if (p.partName || p.unitPrice) partsIdx++;
				});

				const pacRows = extractSectionRows("梱包費", 1, 20);
				const pacHeader = getLocalHeader("梱包費", "所要量", mainHeaderRow);
				let pacIdx = 0;
				pacRows.forEach((exR) => {
					while (dd.packaging.length <= pacIdx) {
						dd.packaging.push(window.createEmptyPackaging ? window.createEmptyPackaging() : { id: 'id_' + Math.random().toString(36).substr(2, 9) });
					}
					const p = dd.packaging[pacIdx];
					const h = pacHeader;
					p.detailType = getVal(exR, h, "pac_detailType", mapping);
					p.quantity = getValNum(exR, h, "pac_quantity", mapping);
					p.unit = getVal(exR, h, "pac_unit", mapping);
					p.unitPrice = getValNum(exR, h, "pac_unitPrice", mapping);
					p.rotationCount = getValNum(exR, h, "pac_rotationCount", mapping);
					p.supplier = getVal(exR, h, "pac_supplier", mapping);
					p.piecesPerBox = getValNum(exR, h, "pac_piecesPerBox", mapping);
					p.materialName = getVal(exR, h, "pac_materialName", mapping);
					p.specification = getVal(exR, h, "pac_specification", mapping);
					p.remarks = getVal(exR, h, "pac_remarks", mapping);
					if (p.materialName || p.unitPrice) pacIdx++;
				});

				const transRow = extractSectionRows("輸送費", 1, 1)[0];
				if (transRow) {
					const h = getLocalHeader("輸送費", "出荷元", mainHeaderRow);
					dd.transport.shippingOrigin = getVal(transRow, h, "trans_shippingOrigin", mapping);
					dd.transport.deliveryDest = getVal(transRow, h, "trans_deliveryDest", mapping);
					dd.transport.distance = getValNum(transRow, h, "trans_distance", mapping);
					dd.transport.boxHeight = getValNum(transRow, h, "trans_boxHeight", mapping);
					dd.transport.boxWidth = getValNum(transRow, h, "trans_boxWidth", mapping);
					dd.transport.boxDepth = getValNum(transRow, h, "trans_boxDepth", mapping);
					dd.transport.weight = getValNum(transRow, h, "trans_weight", mapping);
					dd.transport.packageForm = getVal(transRow, h, "trans_packageForm", mapping);
					dd.transport.transportCategory = getVal(transRow, h, "trans_transportCategory", mapping);
					dd.transport.transportMethod = getVal(transRow, h, "trans_transportMethod", mapping);
					dd.transport.logisticsPrice = getValNum(transRow, h, "trans_logisticsPrice", mapping);
					dd.transport.boxQuantity = getValNum(transRow, h, "trans_boxQuantity", mapping);
					dd.transport.transportQty = getValNum(transRow, h, "trans_transportQty", mapping);
					dd.transport.handlingFee = getValNum(transRow, h, "trans_handlingFee", mapping);
					dd.transport.otherFee = getValNum(transRow, h, "trans_otherFee", mapping);
					dd.transport.remarks = getVal(transRow, h, "trans_remarks", mapping);
				}

				const otherRow = extractSectionRows("その他費", 1, 1)[0];
				if (otherRow) {
					const h = getLocalHeader("その他費", "その他費", mainHeaderRow);
					dd.otherCost.otherCost = getValNum(otherRow, h, "other_cost", mapping);
					dd.otherCost.item = getVal(otherRow, h, "other_item", mapping);
					dd.otherCost.remarks = getVal(otherRow, h, "other_remarks", mapping);
				}

				const moldRow = extractSectionRows("金型費", 1, 1)[0];
				if (moldRow) {
					const h = getLocalHeader("金型費", "金型名称", mainHeaderRow);
					dd.moldCost.moldName = getVal(moldRow, h, "mold_moldName", mapping);
					dd.moldCost.storageLocation = getVal(moldRow, h, "mold_storageLocation", mapping);
					dd.moldCost.procurementType = getVal(moldRow, h, "mold_procurementType", mapping);
					dd.moldCost.paymentMethod = getVal(moldRow, h, "mold_paymentMethod", mapping);
					dd.moldCost.investmentAmount = getValNum(moldRow, h, "mold_investmentAmount", mapping);
					dd.moldCost.depreciationCount = getValNum(moldRow, h, "mold_depreciationCount", mapping);
					dd.moldCost.supplier = getVal(moldRow, h, "mold_supplier", mapping);
					dd.moldCost.remarks = getVal(moldRow, h, "mold_remarks", mapping);
				}

				const mgmtRow = extractSectionRows("管理費", 1, 1)[0];
				if (mgmtRow) {
					const h = getLocalHeader("管理費", "費目コード", mainHeaderRow);
					dd.management.expenseCode = getVal(mgmtRow, h, "mgmt_expenseCode", mapping);
					dd.management.ratio = getValNum(mgmtRow, h, "mgmt_ratio", mapping);
				}

				const profitRow = extractSectionRows("利益", 1, 1)[0];
				if (profitRow) {
					const h = getLocalHeader("利益", "費目コード", mainHeaderRow);
					dd.profit.expenseCode = getVal(profitRow, h, "profit_expenseCode", mapping);
					dd.profit.ratio = getValNum(profitRow, h, "profit_ratio", mapping);
				}

				if (window.renderAllDetailSections) {
					window.renderAllDetailSections();
				} else if (window.renderDetailTables) {
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
				var k = keyword.replace(/\s/g, '');
				for (var j = 0; j < row.length; j++) {
					var cell = String(row[j] || '').replace(/\s/g, '');
					if (cell.includes(k)) {
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

			lines.push(formatLine('MAT02', [
				'1', '1', '1.000000',
				sourceCurrency || 'JPY', targetCurrency || 'JPY', weightUnit || 'KG',
				'', '', yieldRate, recycledWeight, recoveredWeight, recoveredPrice,
				'', '', '', displayPattern, ''
			]));

			var materialRows = [];
			if (materialSectionRow >= 0) {
				var headerRowIdx = -1;
				var colZumenBango = -1;
				var colTanka = -1;
				var colKakaku = -1;
				for (var r = materialSectionRow; r < Math.min(materialSectionRow + 8, firstSheet.length); r++) {
					var row = firstSheet[r];
					if (!row) continue;
					var idx = findColWithKeyword(row, '図面番号');
					if (idx < 0) idx = findColWithKeyword(row, '部品名称');
					if (idx < 0) idx = findColWithKeyword(row, '素材');

					if (idx >= 0) {
						headerRowIdx = r;
						colZumenBango = idx;
						break;
					}
				}
				if (headerRowIdx >= 0) {
					for (var r = headerRowIdx + 1; r < Math.min(headerRowIdx + 4, firstSheet.length); r++) {
						var row = firstSheet[r];
						if (!row) continue;
						var tankaIdx = findColWithKeyword(row, '単価');
						var kakakuIdx = findColWithKeyword(row, '価格');
						if (tankaIdx >= 0 || kakakuIdx >= 0) {
							colTanka = tankaIdx;
							colKakaku = kakakuIdx;
							break;
						}
					}
				}
				if (headerRowIdx >= 0) {
					var startRow = headerRowIdx + 2;
					for (var i = startRow; i < Math.min(startRow + 20, firstSheet.length); i++) {
						var row = firstSheet[i];
						if (!row) continue;
						var firstCell = String(row[0] || '').trim();
						var zumenCell = colZumenBango >= 0 ? String(row[colZumenBango] || '').trim() : '';
						if (firstCell.includes('合計') || firstCell.includes('小計') || firstCell.includes('購入')) break;
						if (zumenCell && zumenCell.length > 0 && /[A-Z0-9]/i.test(zumenCell) && !/[①-⑳⓪]/.test(zumenCell)) {
							var price = '';
							if (colTanka >= 0) {
								price = String(row[colTanka] || '').trim();
							}
							if (!price && colKakaku >= 0) {
								price = String(row[colKakaku] || '').trim();
							}
							materialRows.push({
								code: zumenCell,
								price: price
							});
						}
					}
				}
			}

			materialRows.forEach(function (item, idx) {
				var sectionName = 'MATD02_' + (idx + 1);
				lines.push(formatLine(sectionName, [
					String(idx + 1), '04', item.code, '01', '', '', '', '', '', '', '', item.price, String(idx + 1), ''
				]));
			});
			lines.push(formatLine('PROC02', [
				'1', '射出成形', '', '', '1.000000', 'JPY', 'JPY', '',
				'', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''
			]));
			lines.push(formatLine('PROC02', [
				'1', '加工工程', '', '1.000000', 'JPY', 'JPY', '', '', '', '', '', '', '', '', ''
			]));
			lines.push(formatLine('PART', [
				'1', '1.000000', 'JPY', 'JPY', '', '', '', '', '', '', '', '', '', '', '', ''
			]));
			lines.push(formatLine('TOOL', [
				'1', '1.000000', 'JPY', 'JPY', '', '1', '', '1', '1', '1', '1', ''
			]));
			lines.push(formatLine('EXPENSE', [
				'1', '', '', '', '1.000000'
			]));
			lines.push(formatLine('PROFIT', [
				'1', '', '', '', '1.000000'
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
				} else if (section === 'MAT02') {
					data.materialResin = {
						detailNo: cells[1], quoteNo: cells[2], conversionRate: cells[3], sourceCurrency: cells[4],
						targetCurrency: cells[5], weightUnit: cells[6], productWeight: cells[7], runnerWeight: cells[8],
						yieldRate: cells[9], recycledWeight: cells[10], recoveredWeight: cells[11], recoveredPrice: cells[12],
						productionLot: cells[13], wasteWeight: cells[14], calcPattern: cells[15], displayPattern: cells[16],
						remarks: cells[17]
					};
				} else if (section.indexOf('MATD02_') === 0) {
					data.materialDetail.push({
						id: 'id_' + Math.random().toString(36).substr(2, 9),
						detailNo: cells[1], procurementType: cells[2], materialCode: cells[3], deliveryForm: cells[4],
						supplier: cells[5], manufacturer: cells[6], color: cells[7], grade: cells[8],
						deliveryLot: cells[9], blendingRatio: cells[10], smallLotEx: cells[11], materialPrice: cells[12],
						detailType: cells[13], specification: cells[14]
					});
				} else if (section === 'PROC02') {
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
				} else if (section === 'PAC02') {
					data.packaging.push({
						id: 'id_' + Math.random().toString(36).substr(2, 9),
						quoteNo: cells[1], detailType: cells[2], conversionRate: cells[3], sourceCurrency: cells[4],
						targetCurrency: cells[5], quantity: cells[6], unit: cells[7], unitPrice: cells[8],
						rotationCount: cells[9], supplier: cells[10], piecesPerBox: cells[11], materialName: cells[12],
						specification: cells[13], remarks: cells[14]
					});
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
				} else if (section === 'EXPENSE') {
					data.management = {
						quoteNo: cells[1], expenseCode: cells[2], managementFee: cells[3], ratio: cells[4]
					};
				} else if (section === 'PROFIT') {
					data.profit = {
						quoteNo: cells[1], expenseCode: cells[2], profitAmount: cells[3], ratio: cells[4]
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
			if (data.management) Object.assign(details.management, data.management);
			if (data.profit) Object.assign(details.profit, data.profit);
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