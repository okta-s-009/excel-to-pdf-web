/**
 * Excel2PDF Studio - Offline Client-Side Conversion Engine v5.0
 * ============================================================
 * Ultra-high-fidelity Excel-to-PDF conversion running 100% in the
 * browser and Android APK (SheetJS + jsPDF + AutoTable).
 *
 * Engineered for Indonesian Accounting, Contractor, & Government Spreadsheets:
 *   - Smart table boundary detection: eliminates ghost columns & stray scratch notes
 *   - Smart semantic column coloring: peach dates, mint income, rose expenses, sky quantities
 *   - 100% preservation of Excel explicit fill colors (e.g. bright yellow highlights)
 *   - Solid, continuous, crisp grid lines on every cell (never faint or broken)
 *   - Multi-tier header auto-merging: columns without subheaders merge cleanly vertically
 *   - Clean date formatting (no long GMT timezone strings)
 *   - Intelligent multi-table and multi-sheet separation
 *   - Robust Android tablet & phone PDF download & share delivery
 */

(function (window) {
  'use strict';

  // Configure PDF.js worker if available
  if (window.pdfjsLib) {
    const isStatic = window.location.pathname.startsWith('/static');
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = isStatic
      ? '/static/libs/pdf.worker.min.js'
      : 'libs/pdf.worker.min.js';
  }

  const PAPER_DIMENSIONS_MM = {
    A4: { width: 210, height: 297 },
    Letter: { width: 215.9, height: 279.4 },
    Legal: { width: 215.9, height: 355.6 },
    A3: { width: 297, height: 420 },
    Executive: { width: 184.1, height: 266.7 }
  };

  function cmToMm(cm) {
    return (parseFloat(cm) || 0) * 10;
  }

  function inchToMm(inch) {
    return (parseFloat(inch) || 0) * 25.4;
  }

  function formatBytes(bytes) {
    if (!bytes || bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  // --- SMART SEMANTIC COLUMN PALETTE ---
  const PALETTE = {
    tanggal:    { headBg: [254, 215, 170], headText: [124, 45, 18],  bodyBg: [255, 247, 237], halign: 'center' }, // Soft Peach
    keterangan: { headBg: [226, 232, 240], headText: [15, 23, 42],   bodyBg: [255, 255, 255], halign: 'left' },   // Clean Slate
    quantity:   { headBg: [224, 242, 254], headText: [7, 89, 133],   bodyBg: [248, 250, 252], halign: 'center' }, // Sky Blue
    masuk:      { headBg: [187, 247, 208], headText: [20, 83, 45],   bodyBg: [240, 253, 244], halign: 'right' },  // Mint Green
    keluar:     { headBg: [254, 205, 211], headText: [159, 18, 57],  bodyBg: [255, 245, 245], halign: 'right' },  // Soft Rose
    harga:      { headBg: [254, 240, 138], headText: [113, 63, 18],  bodyBg: [254, 252, 232], halign: 'right' },  // Honey Gold
    total:      { headBg: [224, 231, 255], headText: [49, 46, 129],  bodyBg: [238, 242, 255], halign: 'right' },  // Soft Indigo
    nomor:      { headBg: [226, 232, 240], headText: [15, 23, 42],   bodyBg: [248, 250, 252], halign: 'center' }, // Slate
    default:    { headBg: [226, 232, 240], headText: [15, 23, 42],   bodyBg: [255, 255, 255], halign: 'left' }    // Slate / White
  };

  function detectSemanticColType(headerText, subHeaderText) {
    const combined = ((headerText || '') + ' ' + (subHeaderText || '')).toLowerCase().trim();
    if (/\b(tanggal|tgl|date|waktu|hari)\b/.test(combined)) return 'tanggal';
    if (/\b(masuk|debit|penerimaan|setoran|uang masuk)\b/.test(combined)) return 'masuk';
    if (/\b(keluar|kredit|pengeluaran|biaya|belanja|bayar|potongan|kasbon)\b/.test(combined)) return 'keluar';
    if (/\b(total|sub\s*total|subtotal|grand\s*total|sisa\s*saldo|saldo)\b/.test(combined)) return 'total';
    if (/\b(harga|tarif|price|rate|cost|upah)\b/.test(combined)) return 'harga';
    if (/\b(quantity|kuantitas|volume|vol|satuan|sat|unit|qty|banyak)\b/.test(combined)) return 'quantity';
    if (/\b(keterangan|uraian|nama|kegiatan|item|deskripsi|pekerjaan)\b/.test(combined)) return 'keterangan';
    if (/^(no|no\.|nomor|kode)$/.test(combined.trim())) return 'nomor';
    return 'default';
  }

  // Standard Excel theme colors (Theme 0..9)
  const EXCEL_THEME_PALETTE = [
    [255, 255, 255], // 0: Light 1 (White)
    [0, 0, 0],       // 1: Dark 1 (Black)
    [231, 230, 230], // 2: Light 2 (Warm Gray)
    [68, 84, 106],   // 3: Dark 2 (Navy Slate)
    [68, 114, 196],  // 4: Accent 1 (Blue)
    [237, 125, 49],  // 5: Accent 2 (Orange)
    [165, 165, 165], // 6: Accent 3 (Gray)
    [255, 192, 0],   // 7: Accent 4 (Gold / Yellow)
    [91, 155, 213],  // 8: Accent 5 (Sky Blue)
    [112, 173, 71]   // 9: Accent 6 (Green)
  ];

  function applyExcelTint(rgb, tint) {
    if (!tint || typeof tint !== 'number') return rgb;
    let [r, g, b] = rgb;
    if (tint > 0) {
      r = Math.round(r * (1 - tint) + 255 * tint);
      g = Math.round(g * (1 - tint) + 255 * tint);
      b = Math.round(b * (1 - tint) + 255 * tint);
    } else if (tint < 0) {
      r = Math.round(r * (1 + tint));
      g = Math.round(g * (1 + tint));
      b = Math.round(b * (1 + tint));
    }
    return [Math.max(0, Math.min(255, r)), Math.max(0, Math.min(255, g)), Math.max(0, Math.min(255, b))];
  }

  function parseExcelCellColor(cell) {
    if (!cell || !cell.s) return null;
    const s = cell.s;
    const colorObj = (s.fill && s.fill.fgColor) || s.fgColor || s.bgColor || (s.fill && s.fill.bgColor) || null;
    if (!colorObj) return null;

    // 1. Direct RGB hex string
    if (colorObj.rgb) {
      let rgbStr = String(colorObj.rgb).trim().toUpperCase();
      if (rgbStr.length === 8) rgbStr = rgbStr.substring(2);
      if (rgbStr.length === 6 && !/^000000$|^FFFFFF$/i.test(rgbStr)) {
        const r = parseInt(rgbStr.substring(0, 2), 16);
        const g = parseInt(rgbStr.substring(2, 4), 16);
        const b = parseInt(rgbStr.substring(4, 6), 16);
        if (!isNaN(r) && !isNaN(g) && !isNaN(b)) {
          return colorObj.tint ? applyExcelTint([r, g, b], colorObj.tint) : [r, g, b];
        }
      }
    }

    // 2. Theme color index (0..9)
    if (typeof colorObj.theme === 'number' && colorObj.theme >= 0 && colorObj.theme < EXCEL_THEME_PALETTE.length) {
      if (colorObj.theme === 0 && (!colorObj.tint || colorObj.tint >= 0)) return null;
      if (colorObj.theme === 1 && (!colorObj.tint || colorObj.tint <= 0)) return null;
      const baseRgb = EXCEL_THEME_PALETTE[colorObj.theme];
      return colorObj.tint ? applyExcelTint(baseRgb, colorObj.tint) : baseRgb;
    }

    // 3. Common indexed colors
    if (typeof colorObj.indexed === 'number' && colorObj.indexed >= 2 && colorObj.indexed <= 63) {
      const idxColors = {
        2: [255, 0, 0], 3: [0, 255, 0], 4: [0, 0, 255], 5: [255, 255, 0],
        6: [255, 0, 255], 7: [0, 255, 255], 9: [128, 0, 0], 10: [0, 128, 0],
        11: [0, 0, 128], 12: [128, 128, 0], 13: [128, 0, 128], 14: [0, 128, 128],
        44: [237, 125, 49], 45: [255, 192, 0], 50: [112, 173, 71], 51: [255, 204, 153]
      };
      if (idxColors[colorObj.indexed]) return idxColors[colorObj.indexed];
    }

    return null;
  }

  function formatDateValue(val, cell) {
    if (cell && cell.w && !/GMT|Indochina|Western/i.test(cell.w)) {
      return String(cell.w).trim();
    }
    if (cell && (cell.t === 'd' || cell.v instanceof Date)) {
      const d = cell.v instanceof Date ? cell.v : new Date(cell.v);
      if (!isNaN(d.getTime())) {
        const utcHours = d.getUTCHours();
        let dCalc = d;
        if (utcHours >= 12) {
          dCalc = new Date(d.getTime() + 7 * 3600 * 1000);
        }
        const day = String(dCalc.getUTCDate()).padStart(2, '0');
        const month = String(dCalc.getUTCMonth() + 1).padStart(2, '0');
        const year = String(dCalc.getUTCFullYear()).slice(-2);
        return `${day}/${month}/${year}`;
      }
    }
    if (typeof val === 'string' && /^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+[A-Za-z]+\s+\d{1,2}\s+\d{4}/i.test(val)) {
      const d = new Date(val);
      if (!isNaN(d.getTime())) {
        const day = String(d.getDate()).padStart(2, '0');
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const year = String(d.getFullYear()).slice(-2);
        return `${day}/${month}/${year}`;
      }
    }
    return val;
  }

  const OfflineConverter = {
    /**
     * Checks if all required offline libraries are loaded.
     */
    isReady: function () {
      return Boolean(window.XLSX && (window.jspdf || window.jsPDF));
    },

    /**
     * Inspects an Excel file directly from File or ArrayBuffer.
     */
    inspectFile: async function (file) {
      if (!window.XLSX) {
        throw new Error("Pustaka SheetJS (XLSX) belum dimuat.");
      }

      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(buffer, { type: 'array', cellDates: true, cellStyles: true });
      const sheets = workbook.SheetNames || [];
      const details = [];

      sheets.forEach(name => {
        const ws = workbook.Sheets[name];
        let maxRow = 0;
        let maxCol = 0;
        if (ws && ws['!ref']) {
          const range = window.XLSX.utils.decode_range(ws['!ref']);
          maxRow = range.e.r - range.s.r + 1;
          maxCol = range.e.c - range.s.c + 1;
        }
        details.push({
          name: name,
          max_row: maxRow,
          max_column: maxCol
        });
      });

      return {
        success: true,
        filename: file.name,
        filesize: file.size,
        filesize_formatted: formatBytes(file.size),
        sheets: sheets,
        details: details,
        workbook: workbook
      };
    },

    /**
     * Converts Excel workbook to PDF client-side.
     */
    convertToPdf: async function (fileOrWorkbook, options = {}) {
      if (!this.isReady()) {
        throw new Error("Pustaka konversi offline belum siap (XLSX / jsPDF).");
      }

      const startTime = performance.now();

      // Resolve workbook
      let workbook = null;
      let filename = "dokumen.xlsx";

      if (options && options.cachedWorkbook) {
        workbook = options.cachedWorkbook;
        filename = (fileOrWorkbook && fileOrWorkbook.name) || options.filename || "dokumen.xlsx";
      } else if (fileOrWorkbook instanceof File || fileOrWorkbook instanceof Blob) {
        filename = fileOrWorkbook.name || "dokumen.xlsx";
        const buffer = await fileOrWorkbook.arrayBuffer();
        workbook = window.XLSX.read(buffer, { type: 'array', cellDates: true, cellStyles: true });
      } else if (fileOrWorkbook && fileOrWorkbook.SheetNames) {
        workbook = fileOrWorkbook;
        filename = options.filename || "dokumen.xlsx";
      } else {
        throw new Error("Sumber data Excel tidak valid.");
      }

      // Settings extraction
      const unit = (options.unit || 'cm').toLowerCase();
      const topMm = unit === 'inch' ? inchToMm(options.top_margin) : cmToMm(options.top_margin);
      const bottomMm = unit === 'inch' ? inchToMm(options.bottom_margin) : cmToMm(options.bottom_margin);
      const leftMm = unit === 'inch' ? inchToMm(options.left_margin) : cmToMm(options.left_margin);
      const rightMm = unit === 'inch' ? inchToMm(options.right_margin) : cmToMm(options.right_margin);

      let orientation = (options.orientation || 'portrait').toLowerCase();
      const paperSizeName = options.paper_size || 'A4';
      const paperDim = PAPER_DIMENSIONS_MM[paperSizeName] || PAPER_DIMENSIONS_MM.A4;
      const selectedSheet = options.sheet_name || 'all';

      // Sheets selection
      let sheetsToRender = [];
      if (selectedSheet !== 'all' && workbook.Sheets[selectedSheet]) {
        sheetsToRender.push(selectedSheet);
      } else {
        // Skip sheets with no valid content (e.g. empty sheet '   ')
        sheetsToRender = (workbook.SheetNames || []).filter(sn => {
          const ws = workbook.Sheets[sn];
          return ws && ws['!ref'];
        });
      }

      if (sheetsToRender.length === 0) {
        throw new Error("Tidak ada lembar kerja yang memiliki data.");
      }

      // Pre-scan sheets to find global max REAL columns for auto-landscape
      let globalMaxCols = 0;
      for (const sName of sheetsToRender) {
        const ws = workbook.Sheets[sName];
        if (!ws || !ws['!ref']) continue;
        const r0 = window.XLSX.utils.decode_range(ws['!ref']);
        const qDensity = {};
        for (let qr = r0.s.r; qr <= Math.min(r0.e.r, r0.s.r + 50); qr++) {
          for (let qc = r0.s.c; qc <= r0.e.c; qc++) {
            const addr = window.XLSX.utils.encode_cell({ r: qr, c: qc });
            const cell = ws[addr];
            if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
              qDensity[qc] = (qDensity[qc] || 0) + 1;
            }
          }
        }
        let realColCount = Object.keys(qDensity).length;
        if (realColCount > globalMaxCols) globalMaxCols = realColCount;
      }

      // Auto-switch to landscape if user selected portrait but spreadsheet is wide (>7 cols)
      const autoLandscapeTriggered = orientation === 'portrait' && globalMaxCols > 7;
      if (autoLandscapeTriggered) {
        orientation = 'landscape';
      }

      const pageWidth = orientation === 'landscape' ? paperDim.height : paperDim.width;
      const pageHeight = orientation === 'landscape' ? paperDim.width : paperDim.height;
      const printableWidth = Math.max(pageWidth - leftMm - rightMm, 40);

      // Initialize jsPDF
      const jspdfModule = window.jspdf || window;
      const { jsPDF } = jspdfModule;
      const doc = new jsPDF({
        orientation: orientation === 'landscape' ? 'landscape' : 'portrait',
        unit: 'mm',
        format: [pageWidth, pageHeight],
        compress: true
      });

      // Ensure AutoTable plugin is attached
      if (!doc.autoTable && window.jspdf && window.jspdf.plugin && window.jspdf.plugin.autotable) {
        window.jspdf.plugin.autotable.applyPlugin(jsPDF);
      }

      let isFirstTable = true;

      for (let sIdx = 0; sIdx < sheetsToRender.length; sIdx++) {
        const sName = sheetsToRender[sIdx];
        const ws = workbook.Sheets[sName];
        if (!ws || !ws['!ref']) continue;

        const range = window.XLSX.utils.decode_range(ws['!ref']);
        const totalRows = range.e.r - range.s.r + 1;
        if (totalRows <= 0) continue;

        // === SMART COLUMN BOUNDS: Eliminate ghost empty columns & isolated scratch notes ===
        const colDensity = {};
        for (let r = range.s.r; r <= range.e.r; r++) {
          for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = window.XLSX.utils.encode_cell({ r, c });
            const cell = ws[addr];
            if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
              colDensity[c] = (colDensity[c] || 0) + 1;
            }
          }
        }

        const merges = ws['!merges'] || [];
        const mergedCols = new Set();
        for (const m of merges) {
          for (let mc = m.s.c; mc <= m.e.c; mc++) mergedCols.add(mc);
        }

        let realStartCol = range.s.c;
        for (let c = range.s.c; c <= range.e.c; c++) {
          if ((colDensity[c] || 0) > 0) {
            realStartCol = c;
            break;
          }
        }

        const maxColDensity = Math.max(1, ...Object.values(colDensity));

        // Header-like rows near the top (>= 3 text cells). A column carrying a real text
        // label in one of these rows is a genuine table column even if its data is sparse.
        const headerLikeRows = [];
        for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 15); r++) {
          let textCount = 0;
          for (let c = range.s.c; c <= range.e.c; c++) {
            const cell = ws[window.XLSX.utils.encode_cell({ r, c })];
            if (cell && cell.t === 's' && String(cell.v).trim().length >= 2) textCount++;
          }
          if (textCount >= 3) headerLikeRows.push(r);
        }
        const hasHeaderLabel = (c) => headerLikeRows.some(r => {
          const cell = ws[window.XLSX.utils.encode_cell({ r, c })];
          return Boolean(cell && cell.t === 's' && /[a-zA-Z]{3,}/.test(String(cell.v)));
        });
        const isRealColumn = (c) => {
          const d = colDensity[c] || 0;
          return d > maxColDensity * 0.15 || mergedCols.has(c) || (d >= 2 && hasHeaderLabel(c));
        };
        const isScratchColumn = (c) => {
          // Only a handful of tiny values (e.g. "1", "T", "an") with no header label
          if ((colDensity[c] || 0) > 2 || mergedCols.has(c) || hasHeaderLabel(c)) return false;
          for (let r = range.s.r; r <= range.e.r; r++) {
            const cell = ws[window.XLSX.utils.encode_cell({ r, c })];
            if (cell && cell.v !== undefined && String(cell.v).trim().length > 3) return false;
          }
          return true;
        };

        let realEndCol = realStartCol;
        for (let c = realStartCol; c <= range.e.c; c++) {
          const d = colDensity[c] || 0;
          if (d === 0 && !mergedCols.has(c)) {
            // Blank gap: continue only if genuine table columns follow (otherwise stray notes)
            let nextHasRealTable = false;
            for (let nc = c + 1; nc <= Math.min(c + 3, range.e.c); nc++) {
              if (isRealColumn(nc)) { nextHasRealTable = true; break; }
            }
            if (!nextHasRealTable) break;
          } else {
            // Populated column directly attached to the table: ALWAYS keep it,
            // even when sparse (e.g. "Adm", "Adm transfer", status columns).
            realEndCol = c;
          }
        }
        // Trim only a trailing scratch column glued to the table edge
        while (realEndCol > realStartCol + 3 && isScratchColumn(realEndCol)) {
          realEndCol--;
          while (realEndCol > realStartCol && !(colDensity[realEndCol] || 0) && !mergedCols.has(realEndCol)) realEndCol--;
        }

        if (realEndCol < realStartCol) realEndCol = realStartCol;
        const totalCols = Math.max(realEndCol - realStartCol + 1, 1);

        // Build merge map strictly within real bounds
        const mergeMap = {};
        for (const m of merges) {
          const rs = m.e.r - m.s.r + 1;
          const cs = m.e.c - m.s.c + 1;
          for (let r = m.s.r; r <= m.e.r; r++) {
            if (!mergeMap[r]) mergeMap[r] = {};
            for (let c = m.s.c; c <= m.e.c; c++) {
              if (r === m.s.r && c === m.s.c) {
                mergeMap[r][c] = { rs, cs, origin: true };
              } else {
                mergeMap[r][c] = { skip: true, originR: m.s.r, originC: m.s.c };
              }
            }
          }
        }

        // Separate Title rows from Table rows
        const titleLines = [];
        let tableStartR = range.s.r;

        for (let r = range.s.r; r <= range.e.r; r++) {
          const filled = [];
          let isWideMerge = false;

          for (let c = realStartCol; c <= realEndCol; c++) {
            const m = (mergeMap[r] && mergeMap[r][c]) || null;
            if (m && m.origin && m.cs >= Math.floor(totalCols * 0.5)) {
              isWideMerge = true;
            }
            if (m && m.skip) continue;

            const addr = window.XLSX.utils.encode_cell({ r, c });
            const cell = ws[addr];
            if (cell && cell.v !== undefined && String(cell.v).trim() !== '') {
              filled.push(String(cell.w || cell.v).trim());
            }
          }

          if (filled.length === 0) {
            if (r < range.s.r + 8) continue;
            tableStartR = r;
            break;
          }

          const isTitleRow = (filled.length <= 2 && (isWideMerge || totalCols >= 4));
          if (r < range.s.r + 8 && isTitleRow) {
            const titleStr = filled.join(' - ').trim();
            // Filter out isolated stray 1-3 letter scratch notes (e.g. "an")
            if (titleStr.length >= 4 && !/^[a-zA-Z0-9\/\-\.]{1,3}$/.test(titleStr)) {
              titleLines.push(titleStr);
            }
          } else {
            tableStartR = r;
            break;
          }
        }

        // Determine header rows count (multi-tier headers)
        let headerRowsCount = 1;
        for (let c = realStartCol; c <= realEndCol; c++) {
          const m = (mergeMap[tableStartR] && mergeMap[tableStartR][c]);
          if (m && m.origin && m.rs > 1 && m.rs <= 4) {
            if (m.rs > headerRowsCount) headerRowsCount = m.rs;
          }
          if (m && m.origin && m.cs > 1 && tableStartR + 1 <= range.e.r) {
            headerRowsCount = Math.max(headerRowsCount, 2);
          }
        }

        // Heuristic: check if row below tableStartR contains subheader keywords
        if (headerRowsCount === 1 && tableStartR + 1 <= range.e.r) {
          let subScore = 0;
          for (let c = realStartCol; c <= realEndCol; c++) {
            const addr = window.XLSX.utils.encode_cell({ r: tableStartR + 1, c });
            const cell = ws[addr];
            if (cell && cell.v !== undefined) {
              const v = String(cell.v).trim().toLowerCase();
              if (/^(vol|volume|sat|satuan|unit|masuk|keluar|debit|kredit|jumlah|qty|harga|list)$/i.test(v)) {
                subScore++;
              }
            }
          }
          if (subScore >= 2) headerRowsCount = 2;
        }

        // A vertical merge that starts in a lower header row but continues below the header
        // means that row is really DATA (e.g. "batu split | 200 | m3" spanning several rows).
        for (let hr = tableStartR + 1; hr < tableStartR + headerRowsCount; hr++) {
          let spillsIntoBody = false;
          for (let c = realStartCol; c <= realEndCol; c++) {
            const m = mergeMap[hr] && mergeMap[hr][c];
            if (m && m.origin && hr + m.rs - 1 >= tableStartR + headerRowsCount) {
              spillsIntoBody = true;
              break;
            }
          }
          if (spillsIntoBody) {
            headerRowsCount = hr - tableStartR;
            break;
          }
        }

        // Multi-tier header auto-merging:
        // When headerRowsCount === 2, columns in tableStartR that have no subheaders below
        // should automatically merge vertically across both header rows!
        if (headerRowsCount === 2) {
          for (let c = realStartCol; c <= realEndCol; c++) {
            const m1 = (mergeMap[tableStartR] && mergeMap[tableStartR][c]);
            const nextR = tableStartR + 1;
            const addr1 = window.XLSX.utils.encode_cell({ r: tableStartR, c });
            const addr2 = window.XLSX.utils.encode_cell({ r: nextR, c });
            const cell1 = ws[addr1];
            const cell2 = ws[addr2];
            const hasContentAbove = Boolean(cell1 && cell1.v !== undefined && String(cell1.v).trim() !== '');
            const hasContentBelow = Boolean(cell2 && cell2.v !== undefined && String(cell2.v).trim() !== '');

            if (!m1) {
              if (!hasContentAbove && hasContentBelow) {
                // Orphaned subheader (e.g. C4 empty, C5 "KETERANGAN"): lift to row 1 & span both
                ws[addr1] = { ...cell2 };
                if (!mergeMap[tableStartR]) mergeMap[tableStartR] = {};
                mergeMap[tableStartR][c] = { rs: 2, cs: 1, origin: true, autoHeaderMerge: true };
                if (!mergeMap[nextR]) mergeMap[nextR] = {};
                mergeMap[nextR][c] = { skip: true, originR: tableStartR, originC: c };
              } else if (!hasContentBelow) {
                // Top header has no subheader below: span both rows.
                // If top header is also blank but column contains dates, auto-label "Tanggal"
                if (!hasContentAbove) {
                  let dateMatches = 0;
                  for (let r = tableStartR + headerRowsCount; r <= Math.min(range.e.r, tableStartR + 35); r++) {
                    const cl = ws[window.XLSX.utils.encode_cell({ r, c })];
                    if (cl && (cl.t === 'd' || (cl.t === 'n' && cl.v >= 30000 && cl.v <= 60000))) dateMatches++;
                  }
                  if (dateMatches >= 3) {
                    ws[addr1] = { t: 's', v: 'Tanggal', w: 'Tanggal' };
                  }
                }
                if (!mergeMap[tableStartR]) mergeMap[tableStartR] = {};
                mergeMap[tableStartR][c] = { rs: 2, cs: 1, origin: true, autoHeaderMerge: true };
                if (!mergeMap[nextR]) mergeMap[nextR] = {};
                mergeMap[nextR][c] = { skip: true, originR: tableStartR, originC: c };
              }
            }
          }
        }

        // Determine semantic column types for smart coloring
        const colTypes = {};
        for (let c = realStartCol; c <= realEndCol; c++) {
          let h1 = '', h2 = '';
          const c1 = ws[window.XLSX.utils.encode_cell({ r: tableStartR, c })];
          if (c1 && c1.v !== undefined) h1 = String(c1.w || c1.v).trim();
          if (headerRowsCount > 1) {
            const c2 = ws[window.XLSX.utils.encode_cell({ r: tableStartR + 1, c })];
            if (c2 && c2.v !== undefined) h2 = String(c2.w || c2.v).trim();
          }
          // If h1 is blank (merged from left), find origin text
          if (!h1) {
            for (let sc = c - 1; sc >= realStartCol; sc--) {
              const oc = ws[window.XLSX.utils.encode_cell({ r: tableStartR, c: sc })];
              if (oc && oc.v !== undefined && String(oc.w || oc.v).trim() !== '') {
                h1 = String(oc.w || oc.v).trim();
                break;
              }
            }
          }
          colTypes[c] = detectSemanticColType(h1, h2);
        }

        // Proportional column widths from Excel metadata
        const wsCols = ws['!cols'] || [];
        const colWidthChars = [];
        for (let c = realStartCol; c <= realEndCol; c++) {
          const colInfo = wsCols[c];
          if (colInfo && colInfo.wch) {
            colWidthChars.push(Math.max(colInfo.wch, 4));
          } else if (colInfo && colInfo.wpx) {
            colWidthChars.push(Math.max(colInfo.wpx / 7, 4));
          } else {
            let maxLen = 4;
            for (let r = tableStartR; r <= Math.min(range.e.r, tableStartR + 35); r++) {
              const cell = ws[window.XLSX.utils.encode_cell({ r, c })];
              if (cell) {
                const v = String(cell.w || cell.v || '');
                if (v.length > maxLen) maxLen = Math.min(v.length, 45);
              }
            }
            colWidthChars.push(maxLen);
          }
        }

        // ---- PASS 1: decide which rows are rendered and which block each belongs to ----
        // Block 'H' = header rows, 'B0', 'B1', ... = body table sections (split by 2+ blank rows)
        const rowHasData = (r) => {
          for (let c = realStartCol; c <= realEndCol; c++) {
            const m = mergeMap[r] && mergeMap[r][c];
            if (m && m.skip) continue;
            const cell = ws[window.XLSX.utils.encode_cell({ r, c })];
            if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') return true;
          }
          return false;
        };
        const rowBlock = {};
        const bodyBlockIds = [];
        let secCounter = 0;
        let blankRun = 0;
        let secRowCount = headerRowsCount;
        for (let r = tableStartR; r <= range.e.r; r++) {
          if (r < tableStartR + headerRowsCount) { rowBlock[r] = 'H'; continue; }
          if (rowHasData(r)) {
            if (blankRun >= 2 && secRowCount >= 2) { secCounter++; secRowCount = 0; }
            blankRun = 0;
            const blk = 'B' + secCounter;
            if (!bodyBlockIds.includes(blk)) bodyBlockIds.push(blk);
            rowBlock[r] = blk;
            secRowCount++;
          } else {
            blankRun++;
          }
        }

        // ---- PASS 2: build cells. RowSpans are clipped to rendered rows of the same block, and
        // cells whose merge origin lives in another block are materialized, so every row keeps
        // exactly the right number of columns (no shifted / lost columns). ----
        const headRows = [];
        const bodyBlocks = {};
        const keptRows = Object.keys(rowBlock).map(Number).sort((a, b) => a - b);

        for (const r of keptRows) {
          const blk = rowBlock[r];
          const isHeaderRow = blk === 'H';
          const rowCells = [];
          let rowHasContent = false;

          for (let c = realStartCol; c <= realEndCol; c++) {
            const m = (mergeMap[r] && mergeMap[r][c]) || null;
            let materialColSpan = 0;

            if (m && m.skip) {
              const oR = m.originR, oC = m.originC;
              const originInfo = (mergeMap[oR] && mergeMap[oR][oC]) || { rs: 1, cs: 1 };
              const originRendered = oC >= realStartCol && (oR === r || rowBlock[oR] === blk);
              if (originRendered) continue; // covered by origin's rowSpan / colSpan
              const firstVisibleC = Math.max(oC, realStartCol);
              if (c !== firstVisibleC) continue; // covered by the materialized cell's colSpan
              materialColSpan = Math.min(oC + originInfo.cs - 1, realEndCol) - c + 1;
            }

            const addr = window.XLSX.utils.encode_cell({ r, c });
            const cell = materialColSpan ? null : ws[addr];
            let val = '';
            if (cell && cell.v !== undefined) {
              val = cell.w !== undefined ? String(cell.w).trim() : String(cell.v).trim();
              val = formatDateValue(val, cell);
              if (val) rowHasContent = true;
            }

            const cType = colTypes[c] || 'default';
            const pal = PALETTE[cType] || PALETTE.default;

            let halign = pal.halign;
            // Heuristic override for alignment if cell has specific format
            if (!isHeaderRow && halign === 'left') {
              if (cell && cell.t === 'n') halign = 'right';
              else if (/^[\(]?[\-]?\s*(Rp|IDR|\$|€|£)?\s*[\d.,]+\s*[\)]?$/.test(val) && /\d/.test(val)) halign = 'right';
              else if (val.length <= 5 && /^\d+$/.test(val)) halign = 'center';
            }
            if (isHeaderRow) halign = 'center';

            const cellObj = {
              content: val,
              styles: {
                halign: halign,
                valign: 'middle'
              },
              _c: c,
              _r: r,
              _colType: cType
            };

            // Capture custom Excel cell background color
            const customColor = parseExcelCellColor(cell);
            if (customColor) {
              cellObj._excelColor = customColor;
            }

            if (materialColSpan > 1) cellObj.colSpan = materialColSpan;

            if (m && m.origin) {
              // Effective rowSpan = rendered rows of the SAME block inside the merge range
              let rs = 0;
              for (let r2 = r; r2 <= r + m.rs - 1; r2++) {
                if (rowBlock[r2] === blk) rs++;
              }
              let cs = m.cs;
              if (c + cs - 1 > realEndCol) cs = realEndCol - c + 1;
              if (rs > 1) cellObj.rowSpan = rs;
              if (cs > 1) cellObj.colSpan = cs;
            }

            rowCells.push(cellObj);
          }

          // Check if summary / total row
          const isTotalRow = !isHeaderRow && rowHasContent && rowCells.some(c => {
            const txt = (c.content || '').toUpperCase();
            return /\b(TOTAL|JUMLAH|SUBTOTAL|SUB TOTAL|GRAND TOTAL|SISA SALDO|SALDO AKHIR)\b/.test(txt);
          });
          if (isTotalRow) {
            rowCells.forEach(c => { c._isTotal = true; });
          }

          if (isHeaderRow) {
            headRows.push(rowCells);
          } else {
            if (!bodyBlocks[blk]) bodyBlocks[blk] = [];
            bodyBlocks[blk].push(rowCells);
          }
        }

        // Assemble sections: [{ head, body }]
        const cloneRows = (rows) => rows.map(row => row.map(cl => ({ ...cl, styles: { ...cl.styles } })));
        const tableSections = [];
        if (bodyBlockIds.length === 0) {
          if (headRows.length > 0) tableSections.push({ head: headRows, body: [] });
        } else {
          bodyBlockIds.forEach((blk, idx) => {
            const body = bodyBlocks[blk] || [];
            if (body.length === 0) return;
            if (idx === 0) {
              tableSections.push({ head: headRows, body });
              return;
            }
            // Follow-on block: repeat the main header only if it spans most of the table width
            // (e.g. totals under the same columns); small side tables get no foreign header.
            let minC = Infinity, maxC = -Infinity;
            body.forEach(row => row.forEach(cl => {
              if (cl.content) {
                minC = Math.min(minC, cl._c);
                maxC = Math.max(maxC, cl._c + (cl.colSpan || 1) - 1);
              }
            }));
            const spanCols = maxC >= minC ? (maxC - minC + 1) : 0;
            const head = (headRows.length > 0 && spanCols >= totalCols * 0.5) ? cloneRows(headRows) : [];
            tableSections.push({ head, body });
          });
        }

        if (tableSections.length === 0) continue;

        // Render each table section
        for (let secIdx = 0; secIdx < tableSections.length; secIdx++) {
          const headPart = tableSections[secIdx].head;
          const bodyPart = tableSections[secIdx].body;
          if (headPart.length === 0 && bodyPart.length === 0) continue;

          if (!isFirstTable) {
            doc.addPage([pageWidth, pageHeight], orientation);
          }
          isFirstTable = false;

          // Render Document Title above Table
          let curY = Math.max(topMm, 8);
          if (secIdx === 0 && titleLines.length > 0) {
            for (let ti = 0; ti < titleLines.length; ti++) {
              const titleFontSize = ti === 0 ? 11 : 9;
              doc.setFontSize(titleFontSize);
              doc.setTextColor(15, 23, 42);
              doc.setFont('helvetica', 'bold');
              const centerX = leftMm + printableWidth / 2;
              doc.text(titleLines[ti], centerX, curY, {
                align: 'center',
                maxWidth: printableWidth - 4
              });
              curY += titleFontSize * 0.45 + 3.0;
            }
            curY += 2.0;
          } else if (sheetsToRender.length > 1 || tableSections.length > 1) {
            doc.setFontSize(10);
            doc.setTextColor(30, 41, 59);
            doc.setFont('helvetica', 'bold');
            const secLabel = tableSections.length > 1
              ? `Sheet: ${sName} (Tabel ${secIdx + 1})`
              : `Sheet: ${sName}`;
            doc.text(secLabel, leftMm, curY);
            curY += 6.0;
          }

          // Font sizing & padding based on column count
          let fontSize = 7.5;
          let cellPadding = 1.0;
          if (totalCols > 18) { fontSize = 5.2; cellPadding = 0.5; }
          else if (totalCols > 14) { fontSize = 6.0; cellPadding = 0.7; }
          else if (totalCols > 10) { fontSize = 6.8; cellPadding = 0.8; }
          else if (totalCols > 7) { fontSize = 7.5; cellPadding = 1.0; }
          else if (totalCols <= 6) { fontSize = 8.5; cellPadding = 1.3; }

          // Proportional column widths
          const columnStyles = {};
          if (colWidthChars.length > 0) {
            const totalCharW = colWidthChars.reduce((a, b) => a + b, 0);
            if (totalCharW > 0) {
              let totalWidthMm = 0;
              const widths = [];
              for (let ci = 0; ci < colWidthChars.length; ci++) {
                const pct = colWidthChars[ci] / totalCharW;
                const minColW = totalCols > 12 ? 4.5 : 6.0;
                const w = Math.max(pct * printableWidth, minColW);
                widths.push(w);
                totalWidthMm += w;
              }
              const scaleFactor = printableWidth / totalWidthMm;
              for (let ci = 0; ci < widths.length; ci++) {
                columnStyles[ci] = { cellWidth: widths[ci] * scaleFactor };
              }
            }
          }

          // AutoTable invocation
          doc.autoTable({
            startY: curY,
            head: headPart,
            body: bodyPart,
            margin: {
              left: leftMm,
              right: rightMm,
              top: curY,
              bottom: bottomMm
            },
            theme: 'grid',
            showHead: 'everyPage',
            tableWidth: printableWidth,
            columnStyles: Object.keys(columnStyles).length > 0 ? columnStyles : undefined,
            styles: {
              fontSize: fontSize,
              cellPadding: cellPadding,
              lineWidth: 0.18,
              lineColor: [148, 163, 184],
              textColor: [30, 41, 59],
              overflow: 'linebreak',
              valign: 'middle',
              font: 'helvetica',
              minCellHeight: fontSize * 0.35 + cellPadding * 2
            },
            headStyles: {
              lineWidth: 0.25,
              lineColor: [100, 116, 139],
              fontStyle: 'bold',
              valign: 'middle',
              minCellHeight: fontSize * 0.35 + cellPadding * 2 + 1.5
            },
            didParseCell: function (data) {
              // Guarantee visible, crisp, unbroken borders on all cells
              data.cell.styles.lineWidth = 0.18;
              data.cell.styles.lineColor = [148, 163, 184];

              const raw = data.cell.raw;

              // PRIORITY 1: Explicit custom fill color from Excel (e.g. bright yellow)
              if (raw && raw._excelColor) {
                data.cell.styles.fillColor = raw._excelColor;
                const [fr, fg, fb] = raw._excelColor;
                const lum = (0.299 * fr + 0.587 * fg + 0.114 * fb);
                data.cell.styles.textColor = lum < 128 ? [255, 255, 255] : [15, 23, 42];
                return;
              }

              // PRIORITY 2: Total / Summary Row
              if (raw && raw._isTotal) {
                data.cell.styles.fillColor = [224, 231, 255]; // Soft indigo/slate
                data.cell.styles.textColor = [30, 27, 75];
                data.cell.styles.fontStyle = 'bold';
                return;
              }

              // PRIORITY 3: Smart semantic column palette
              const cType = (raw && raw._colType) || 'default';
              const pal = PALETTE[cType] || PALETTE.default;

              if (data.section === 'head') {
                data.cell.styles.fillColor = pal.headBg;
                data.cell.styles.textColor = pal.headText;
                data.cell.styles.lineColor = [100, 116, 139];
              } else if (data.section === 'body') {
                data.cell.styles.fillColor = pal.bodyBg;
                data.cell.styles.textColor = [15, 23, 42];
              }
            }
          });
        }
      }

      // Add clean page numbers in footer
      const totalPages = doc.internal.getNumberOfPages();
      for (let p = 1; p <= totalPages; p++) {
        doc.setPage(p);
        doc.setFontSize(7.5);
        doc.setTextColor(100, 116, 139);
        doc.setFont('helvetica', 'normal');
        const footerText = `Hal ${p} / ${totalPages}`;
        const tw = doc.getTextWidth(footerText);
        doc.text(footerText, pageWidth - rightMm - tw, pageHeight - 5);
      }

      // Finalize PDF
      const pdfBlob = doc.output('blob');
      const pdfArrayBuffer = doc.output('arraybuffer');
      const pdfSize = pdfBlob.size;
      const durationSeconds = ((performance.now() - startTime) / 1000).toFixed(2);
      const downloadName = filename.replace(/\.[^/.]+$/, "") + ".pdf";

      return {
        success: true,
        blob: pdfBlob,
        arrayBuffer: pdfArrayBuffer,
        filename: downloadName,
        pdf_size: pdfSize,
        pdf_size_formatted: formatBytes(pdfSize),
        total_pages: totalPages,
        duration_seconds: durationSeconds,
        engine: "Client-Side Engine v5.0 (Smart Precision)",
        auto_landscape: autoLandscapeTriggered
      };
    },

    /**
     * Generates visual page preview thumbnails from PDF buffer using PDF.js.
     */
    generatePreviews: async function (pdfArrayBuffer, maxPages = 1) {
      if (!window.pdfjsLib || !pdfArrayBuffer) {
        return { total_pages: 1, previews: [] };
      }

      try {
        const bufferCopy = pdfArrayBuffer.slice ? pdfArrayBuffer.slice(0) : new Uint8Array(pdfArrayBuffer).buffer.slice(0);
        const loadingTask = window.pdfjsLib.getDocument({ data: new Uint8Array(bufferCopy) });
        const pdf = await loadingTask.promise;
        const totalPages = pdf.numPages;
        const pagesToRender = Math.min(totalPages, maxPages);
        const previews = [];

        for (let i = 1; i <= pagesToRender; i++) {
          const page = await pdf.getPage(i);
          const viewport = page.getViewport({ scale: 1.2 });

          const canvas = document.createElement('canvas');
          const context = canvas.getContext('2d');
          canvas.height = viewport.height;
          canvas.width = viewport.width;

          await page.render({ canvasContext: context, viewport: viewport }).promise;
          const dataUrl = canvas.toDataURL('image/jpeg', 0.85);

          previews.push({ page: i, url: dataUrl });
        }

        return { total_pages: totalPages, previews: previews };
      } catch (err) {
        console.warn("Pratinjau cepat PDF.js:", err);
        return { total_pages: 1, previews: [] };
      }
    },

    /**
     * Renders a single specific page on demand.
     */
    renderSinglePage: async function (pdfArrayBuffer, pageNumber) {
      if (!window.pdfjsLib || !pdfArrayBuffer) return null;
      try {
        const bufferCopy = pdfArrayBuffer.slice ? pdfArrayBuffer.slice(0) : new Uint8Array(pdfArrayBuffer).buffer.slice(0);
        const loadingTask = window.pdfjsLib.getDocument({ data: new Uint8Array(bufferCopy) });
        const pdf = await loadingTask.promise;
        if (pageNumber < 1 || pageNumber > pdf.numPages) return null;

        const page = await pdf.getPage(pageNumber);
        const viewport = page.getViewport({ scale: 1.2 });

        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        await page.render({ canvasContext: context, viewport: viewport }).promise;
        return {
          page: pageNumber,
          url: canvas.toDataURL('image/jpeg', 0.85)
        };
      } catch (e) {
        console.warn("Gagal render halaman tunggal:", e);
        return null;
      }
    },

    /**
     * Saves or shares the PDF based on platform (Native Android APK vs Mobile Tablet Browser vs Desktop)
     */
    deliverPdf: function (pdfBlob, filename) {
      const safeName = (filename && filename.trim()) ? filename : 'dokumen.pdf';

      // 1. Native Android APK Bridge
      if (window.AndroidBridge) {
        const reader = new FileReader();
        reader.onloadend = function () {
          const base64Data = reader.result.split(',')[1];
          if (window.AndroidBridge.savePdf) {
            window.AndroidBridge.savePdf(base64Data, safeName);
          } else if (window.AndroidBridge.sharePdf) {
            window.AndroidBridge.sharePdf(base64Data, safeName);
          }
        };
        reader.readAsDataURL(pdfBlob);
        return;
      }

      // 2. Mobile Browser (Android Tablet, iPad, Mobile Phone) & Desktop
      try {
        const blobUrl = URL.createObjectURL(pdfBlob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = blobUrl;
        a.download = safeName;
        a.target = '_blank';
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
          document.body.removeChild(a);
          setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);
        }, 300);
      } catch (err) {
        console.warn('Download fallback trigger:', err);
        const reader = new FileReader();
        reader.onloadend = function () {
          const a = document.createElement('a');
          a.href = reader.result;
          a.download = safeName;
          document.body.appendChild(a);
          a.click();
          setTimeout(() => document.body.removeChild(a), 200);
        };
        reader.readAsDataURL(pdfBlob);
      }
    },

    /**
     * Triggers native Android Share sheet if in APK or Web Share API
     */
    sharePdf: function (pdfBlob, filename) {
      const safeName = (filename && filename.trim()) ? filename : 'dokumen.pdf';

      // 1. Android APK Bridge
      if (window.AndroidBridge && window.AndroidBridge.sharePdf) {
        const reader = new FileReader();
        reader.onloadend = function () {
          const base64Data = reader.result.split(',')[1];
          window.AndroidBridge.sharePdf(base64Data, safeName);
        };
        reader.readAsDataURL(pdfBlob);
        return;
      }

      // 2. Web Share API (Android Tablet, iPad, Mobile Phone)
      if (navigator.share) {
        try {
          const file = new File([pdfBlob], safeName, { type: 'application/pdf' });
          if (navigator.canShare && navigator.canShare({ files: [file] })) {
            navigator.share({
              files: [file],
              title: safeName,
              text: 'PDF hasil konversi dari Excel2PDF Studio'
            }).catch(err => {
              if (err && err.name !== 'AbortError') {
                this.deliverPdf(pdfBlob, safeName);
              }
            });
            return;
          }
        } catch (e) {
          console.warn("Share file fallback to download:", e);
        }
      }

      // 3. Fallback to direct deliver
      this.deliverPdf(pdfBlob, safeName);
    }
  };

  window.OfflineConverter = OfflineConverter;
})(window);
