(() => {
  if (window.gprocFillLogic) return;

  function ensurePartsRows(countNeeded) {
    while (document.querySelectorAll("#partsTableBody tr").length < countNeeded) {
      if (typeof window.addPartRow === "function") window.addPartRow();
      else break;
    }
  }


  function setParts(parts) {
    ensurePartsRows(parts.length);

    parts.forEach((p, idx) => {
      window.updatePart(idx, "drawingNumber", p.drawingNumber || "");
      window.updatePart(idx, "partName", p.partName || "");
      window.updatePart(idx, "representativeModel", p.representativeModel || "");
      window.updatePart(idx, "estimatedQuantity", p.estimatedQuantity || "");
      window.updatePart(idx, "estimateNumber", p.estimateNumber || "");
      window.updatePart(idx, "unitPrice", p.unitPrice || ""); // Added
      window.updatePart(idx, "applicationPeriod", p.applicationPeriod || "");
    });

    if (typeof window.renderPartsTable === "function") window.renderPartsTable();
  }


  function ensureDetailBucket(drawingNumber) {
    if (!window.detailData) window.detailData = {};
    if (!window.detailData[drawingNumber]) {
      const genId = window.generateId || (() => "id_" + Math.random().toString(36).slice(2));
      window.detailData[drawingNumber] = {
        material: Array(10).fill(null).map(() => ({
          id: genId(),
          drawingNumber: "", temper: "", diameter: "", thickness: "", length: "",
          weight: "", unitPriceKg: "", consumptionPrice: "", materialCost: "",
          processingUnitPrice: "", processingPrice: "", pipingPrice: "", required: ""
        })),
        purchase: Array(10).fill(null).map(() => ({
          id: genId(),
          partName: "", drawingNum: "", required: "", unitPrice: ""
        })),
        supplied: Array(10).fill(null).map(() => ({
          id: genId(),
          partName: "", drawingNum: "", required: ""
        })),
        // press: removed
        secondary: Array(10).fill(null).map(() => ({
          id: genId(),
          process: "", qty: "", seconds: "", unitPrice: "", total: ""
        })),
        // New: Unit Price Breakdown (Single Object)
        breakdown: {
          materialTotal: "",
          materialMgmt: "",
          procTotal: "",
          genSales: "",
          profit: "",
          freight: "",
          materialMgmt14: "",
          genSales15: "",
          total: ""
        }
      };
    }
  }

  function fillArray(targetArr, sourceArr, defaultsFactory) {
    targetArr.length = 0;
    (sourceArr || []).forEach((x) => targetArr.push(x));
    while (targetArr.length < 10) targetArr.push(defaultsFactory());
  }

  function setDetail(drawingNumber, detail) {
    ensureDetailBucket(drawingNumber);
    const bucket = window.detailData[drawingNumber];

    const genId = window.generateId || (() => "id_" + Math.random().toString(36).slice(2));

    fillArray(
      bucket.material,
      detail.material,
      () => ({
        id: genId(),
        drawingNumber: "", temper: "", diameter: "", thickness: "", length: "",
        weight: "", unitPriceKg: "", consumptionPrice: "", materialCost: "",
        processingUnitPrice: "", processingPrice: "", pipingPrice: "", required: ""
      })
    );

    fillArray(
      bucket.purchase,
      detail.purchase,
      () => ({
        id: genId(),
        partName: "", drawingNum: "", required: "", unitPrice: ""
      })
    );

    fillArray(
      bucket.supplied,
      detail.supplied,
      () => ({
        id: genId(),
        partName: "", drawingNum: "", required: ""
      })
    );

    // Press removed

    fillArray(
      bucket.secondary,
      detail.secondary,
      () => ({
        id: genId(),
        process: "", qty: "", seconds: "", unitPrice: "", total: ""
      })
    );

    // Fill Breakdown
    if (detail.breakdown) {
      bucket.breakdown = { ...detail.breakdown };
    }
  }

  // Custom logic to populate DOM for breakdown if this drawing is current
  function renderBreakdown(drawingNumber) {
    const bd = window.detailData[drawingNumber]?.breakdown;
    if (!bd) return;

    const fields = ['materialTotal', 'materialMgmt', 'procTotal', 'genSales', 'profit', 'freight', 'materialMgmt14', 'genSales15', 'total'];
    fields.forEach(f => {
      const el = document.getElementById('breakdown_' + f);
      if (el) el.value = bd[f] || "";
    });
  }

  window.gprocFillLogic = function (payload) {
    if (!payload) return;

    if (Array.isArray(payload.parts)) setParts(payload.parts);

    if (payload.details && typeof payload.details === "object") {
      for (const [drawingNumber, detail] of Object.entries(payload.details)) {
        setDetail(drawingNumber, detail);
      }
    }

    if (window.currentDrawingNumber && typeof window.renderDetailTables === "function") {
      window.renderDetailTables();
      // Also render breakdown
      renderBreakdown(window.currentDrawingNumber);
    }

    // Also patch viewDetail to call renderBreakdown
    const originalViewDetail = window.viewDetail;
    if (originalViewDetail && !originalViewDetail.isPatched) {
      window.viewDetail = function (dn) {
        originalViewDetail(dn);
        renderBreakdown(dn);
      };
      window.viewDetail.isPatched = true;
    }

    console.log("[Demo-Tool] Filled payload ✅", payload);
  };

  console.log("[Demo-Tool] gprocFillLogic injected ✅");
})();
