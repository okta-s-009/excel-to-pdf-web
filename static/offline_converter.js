/**
 * Excel2PDF Studio - Offline Client-Side Conversion Engine v4.0
 * ============================================================
 * High-performance, pixel-accurate Excel-to-PDF conversion running
 * 100% in the browser (and inside Android APK) using SheetJS + jsPDF + AutoTable.
 *
 * Specially engineered for wide government / contractor spreadsheets (e.g. RAB):
 *   - Auto-detects and extracts document title blocks above tables
 *   - Accurately maps SheetJS merged cells (colSpan & rowSpan)
 *   - Automatic multi-level header recognition
 *   - Auto-landscape switching for wide spreadsheets (>7 columns)
 *   - Dynamic scale-to-fit: fits all columns cleanly across page width
 *   - Number & currency right-alignment, index centering
 *   - Razor-sharp vector output with selectable text & pagination
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

  const SMART_COL_PALETTE = {
    total: { head: [187, 247, 208], body: [220, 252, 231], text: [20, 83, 45] },    // Mint green (#DCFCE7)
    harga: { head: [253, 230, 138], body: [254, 243, 199], text: [113, 63, 18] },   // Amber yellow (#FEF3C7)
    jumlah: { head: [254, 215, 170], body: [255, 237, 213], text: [124, 45, 18] },  // Warm peach (#FFEDD5)
    potongan: { head: [254, 205, 211], body: [254, 226, 226], text: [136, 19, 55] },// Rose red (#FEE2E2)
    neutral: { head: [226, 232, 240], body: [241, 245, 249], text: [30, 41, 59] }    // Slate neutral (#F1F5F9)
  };

  // Standard Excel theme colors (Theme 0..9)
  const EXCEL_THEME_PALETTE = [
    [255, 255, 255], // 0: Light 1 (White)
    [0, 0, 0],       // 1: Dark 1 (Black)
    [231, 230, 230], // 2: Light 2 (Warm Gray)
    [68, 84, 106],   // 3: Dark 2 (Navy Slate)
    [68, 114, 196],  // 4: Accent 1 (Blue)
    [237, 125, 49],  // 5: Accent 2 (Orange - prominent in Indonesian accounting)
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

  function detectColType(text) {
    if (!text) return null;
    const t = String(text).trim().toLowerCase();
    if (/\b(total|sub\s*total|subtotal|grand\s*total|total\s*harga|total\s*gaji|total\s*biaya|netto)\b/.test(t)) {
      return 'total';
    } else if (/\b(harga|tarif|price|rate|cost|upah)\b/.test(t)) {
      return 'harga';
    } else if (/\b(jumlah|volume|vol|qty|kuantitas|banyak)\b/.test(t)) {
      return 'jumlah';
    } else if (/\b(kasbon|potongan|minus|denda|pph|diskon|disc)\b/.test(t)) {
      return 'potongan';
    } else if (/\b(no|satuan|unit)\b/.test(t)) {
      return 'neutral';
    }
    return null;
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
    convertToPdf: async function (fileOrWorkbook, options) {
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
      const showGridlines = options.show_gridlines !== false;
      const selectedSheet = options.sheet_name || 'all';

      // Sheets selection
      let sheetsToRender = [];
      if (selectedSheet !== 'all' && workbook.Sheets[selectedSheet]) {
        sheetsToRender.push(selectedSheet);
      } else {
        sheetsToRender = workbook.SheetNames || [];
      }

      // Pre-scan all sheets to find global max REAL columns for auto-landscape
      // Uses content density to ignore ghost empty columns from Excel formatting artifacts
      let globalMaxCols = 0;
      for (const sName of sheetsToRender) {
        const ws = workbook.Sheets[sName];
        if (!ws || !ws['!ref']) continue;
        const r0 = window.XLSX.utils.decode_range(ws['!ref']);
        const qDensity = {};
        const scanEnd = Math.min(r0.e.r, r0.s.r + 60);
        for (let qr = r0.s.r; qr <= scanEnd; qr++) {
          for (let qc = r0.s.c; qc <= r0.e.c; qc++) {
            const addr = window.XLSX.utils.encode_cell({ r: qr, c: qc });
            const cell = ws[addr];
            if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
              qDensity[qc] = (qDensity[qc] || 0) + 1;
            }
          }
        }
        let qStart = r0.s.c, qEnd = r0.s.c;
        for (let qc = r0.s.c; qc <= r0.e.c; qc++) {
          if ((qDensity[qc] || 0) >= 1 && qStart === r0.s.c) qStart = qc;
          if ((qDensity[qc] || 0) >= 2) qEnd = qc;
        }
        const realCols = Math.max(qEnd - qStart + 1, 1);
        if (realCols > globalMaxCols) globalMaxCols = realCols;
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
        if (totalRows === 0) continue;

        // === SMART COLUMN BOUNDS: Eliminate ghost empty columns & stray scratch notes ===
        // 1. Calculate row-by-row density of every column
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

        // 2. Identify merged ranges and header region
        const merges = ws['!merges'] || [];
        const mergedCols = new Set();
        for (const m of merges) {
          for (let mc = m.s.c; mc <= m.e.c; mc++) {
            mergedCols.add(mc);
          }
        }

        let realStartCol = range.s.c;
        let foundStart = false;
        for (let c = range.s.c; c <= range.e.c; c++) {
          if ((colDensity[c] || 0) > 0) {
            realStartCol = c;
            foundStart = true;
            break;
          }
        }

        // 3. Find real end column by detecting true continuous table extent
        // Real table columns have continuous presence, headers, or substantial rows.
        // A gap of 2 or more consecutive empty columns marks the end of the table.
        let realEndCol = realStartCol;
        let emptyColGap = 0;

        for (let c = realStartCol; c <= range.e.c; c++) {
          const density = colDensity[c] || 0;
          const isMerged = mergedCols.has(c);

          if (density > 0 || isMerged) {
            // If we previously had a gap of 2+ empty columns and this column only has 1 or 2 stray cells:
            // This is an isolated scratch note outside the table! Stop the table.
            if (emptyColGap >= 2 && density <= 2 && !isMerged) {
              break;
            }
            realEndCol = c;
            emptyColGap = 0;
          } else {
            emptyColGap++;
            if (emptyColGap >= 3) {
              // 3 consecutive blank columns: definitely beyond the table
              break;
            }
          }
        }

        if (realEndCol < realStartCol) realEndCol = realStartCol;
        const totalCols = Math.max(realEndCol - realStartCol + 1, 1);

        // === EXTRACT EXCEL CELL FILL COLORS ===
        const excelFillColors = {};
        for (let r = range.s.r; r <= range.e.r; r++) {
          for (let c = realStartCol; c <= realEndCol; c++) {
            const addr = window.XLSX.utils.encode_cell({ r, c });
            const cell = ws[addr];
            if (cell) {
              const color = parseExcelCellColor(cell);
              if (color) {
                if (!excelFillColors[r]) excelFillColors[r] = {};
                excelFillColors[r][c] = color;
              }
            }
          }
        }

        // === PROPORTIONAL COLUMN WIDTHS from Excel metadata ===
        const wsCols = ws['!cols'] || [];
        const colWidthChars = [];
        for (let c = realStartCol; c <= realEndCol; c++) {
          const colInfo = wsCols[c];
          if (colInfo && colInfo.wch) {
            colWidthChars.push(Math.max(colInfo.wch, 3));
          } else if (colInfo && colInfo.wpx) {
            colWidthChars.push(Math.max(colInfo.wpx / 7, 3));
          } else {
            let maxLen = 4;
            for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 30); r++) {
              const addr = window.XLSX.utils.encode_cell({ r, c });
              const cell = ws[addr];
              if (cell) {
                const v = cell.w !== undefined ? String(cell.w) : String(cell.v || '');
                if (v.length > maxLen) maxLen = Math.min(v.length, 50);
              }
            }
            colWidthChars.push(maxLen);
          }
        }

        // 1. Build lookup of merged cells by absolute row (r) and col (c)
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
                mergeMap[r][c] = { skip: true };
              }
            }
          }
        }

        // 2. Separate Title rows from Table rows
        // Title rows: early rows with <= 3 filled cells in wide sheets
        const titleLines = [];
        let tableStartR = range.s.r;

        for (let r = range.s.r; r <= range.e.r; r++) {
          const filled = [];
          for (let c = realStartCol; c <= realEndCol; c++) {
            const addr = window.XLSX.utils.encode_cell({ r, c });
            const cell = ws[addr];
            if (cell && cell.w !== undefined && String(cell.w).trim() !== '') {
              filled.push(String(cell.w).trim());
            } else if (cell && cell.v !== undefined && String(cell.v).trim() !== '') {
              filled.push(String(cell.v).trim());
            }
          }

          if (filled.length === 0) {
            continue; // Skip blank spacing row
          }

          if (((filled.length === 1 && totalCols >= 2) || (filled.length <= 3 && totalCols > 4)) && r < range.s.r + 6) {
            // SMART FILTER: Only accept genuine title text (at least 4 chars, not short noise like "an", "x", "ok")
            const titleStr = filled.join(' - ').trim();
            if (titleStr.length >= 4 && !/^[a-zA-Z0-9\/\-\.]{1,3}$/.test(titleStr)) {
              titleLines.push(titleStr);
            }
          } else {
            tableStartR = r;
            break;
          }
        }

        // 3. Determine number of header rows (e.g. multi-tier headers)
        let headerRowsCount = 1;
        for (let c = realStartCol; c <= realEndCol; c++) {
          const m = (mergeMap[tableStartR] && mergeMap[tableStartR][c]) || null;
          if (m && m.origin && m.rs > headerRowsCount) {
            headerRowsCount = m.rs;
          }
        }

        // 4. Construct table rows with colSpan and rowSpan, supporting multiple tables per sheet
        const tableSections = [];
        let currentSectionRows = [];
        let blankCount = 0;

        for (let r = tableStartR; r <= range.e.r; r++) {
          const cells = [];
          for (let c = realStartCol; c <= realEndCol; c++) {
            const m = (mergeMap[r] && mergeMap[r][c]) || null;
            if (m && m.skip) {
              continue; // Skip cell covered by a colSpan or rowSpan
            }

            const addr = window.XLSX.utils.encode_cell({ r, c });
            const cell = ws[addr];
            let val = '';
            if (cell) {
              val = cell.w !== undefined ? String(cell.w).trim() : (cell.v !== undefined ? String(cell.v).trim() : '');
            }

            // Detect cell text alignment
            let halign = 'left';
            if (cell && cell.t === 'n') {
              halign = 'right';
            } else if (/^[\(]?[\-]?\s*(Rp|IDR|\$|€|£)?\s*[\d.,]+\s*[\)]?$/.test(val) && /\d/.test(val)) {
              halign = 'right';
            } else if (/^[\d.,]+\s*%$/.test(val)) {
              halign = 'right';
            } else if (val.length <= 5 && /^\d+$/.test(val)) {
              halign = 'center';
            } else if (/^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4}$/.test(val)) {
              halign = 'center';
            } else if (val.length <= 7 && /^(M3|Sak|Kg|Buah|Batang|Keping|Lembar|Unit|m'|m2|m3|Ltr|Pkt|Bln|Hari|Jam|Ls|Set)$/i.test(val)) {
              halign = 'center';
            }

            const cellObj = {
              content: val,
              styles: { halign: halign }
            };

            // === PRESERVE EXCEL FILL COLORS ===
            const rowColors = excelFillColors[r];
            if (rowColors) {
              const cellColor = rowColors[c];
              if (cellColor) {
                cellObj._excelFillColor = cellColor;
              } else {
                // Row-level fill: if majority of cells in this row are colored, apply to all
                const rowColorCount = Object.keys(rowColors).length;
                if (rowColorCount >= Math.floor(totalCols * 0.4)) {
                  cellObj._excelFillColor = Object.values(rowColors)[0];
                }
              }
            }

            if (m && m.origin) {
              if (m.rs > 1) cellObj.rowSpan = m.rs;
              if (m.cs > 1) cellObj.colSpan = m.cs;
            }
            cells.push(cellObj);
          }

          const hasContent = cells.some(c => c.content !== '');
          if (!hasContent) {
            blankCount++;
            if (blankCount >= 2 && currentSectionRows.length >= 2) {
              tableSections.push(currentSectionRows);
              currentSectionRows = [];
            }
          } else {
            // Intelligent split: Check if this row is a new table title or header following a blank row
            if (blankCount >= 1 && currentSectionRows.length >= 3) {
              const filledInRow = cells.filter(c => c.content !== '');
              const firstTxt = filledInRow.length > 0 ? filledInRow[0].content : '';
              const isNewTitle = (filledInRow.length === 1 && firstTxt.length >= 5) ||
                                 (filledInRow.length <= 3 && /^(LAPORAN|TABEL|REKAPITULASI|DAFTAR|SECTION|BAGIAN|HALAMAN|PAGE|KETERANGAN)\b/i.test(firstTxt));
              const isNewHeader = filledInRow.some(c => /^(No|Kode|Tanggal|Nama|Keterangan|Uraian|Item)\b/i.test(c.content)) && filledInRow.length >= Math.min(3, totalCols);
              if (isNewTitle || isNewHeader) {
                tableSections.push(currentSectionRows);
                currentSectionRows = [];
              }
            }
            blankCount = 0;
            const isTotalRow = cells.some(c => {
              const txt = (c.content || '').toUpperCase();
              return /\b(TOTAL|JUMLAH|SUBTOTAL|SUB TOTAL|GRAND TOTAL|TERBILANG)\b/.test(txt);
            });
            if (isTotalRow) {
              cells.forEach(c => {
                c.styles = c.styles || {};
                c.styles.fontStyle = 'bold';
                c.styles.fillColor = [241, 245, 249];
                c.styles.textColor = [15, 23, 42];
              });
            }
            currentSectionRows.push(cells);
          }
        }
        if (currentSectionRows.length > 0) {
          tableSections.push(currentSectionRows);
        }

        if (tableSections.length === 0) continue;

        for (let secIdx = 0; secIdx < tableSections.length; secIdx++) {
          const allTableRows = tableSections[secIdx];
          if (allTableRows.length === 0) continue;

          // Add page for subsequent sheets or subsequent table sections
          if (!isFirstTable) {
            doc.addPage([pageWidth, pageHeight], orientation);
          }
          isFirstTable = false;

          // 5. Render titles above table (for first section or if section has title)
          let curY = Math.max(topMm, 8);
          if (secIdx === 0 && titleLines.length > 0) {
            for (let ti = 0; ti < titleLines.length; ti++) {
              const titleFontSize = ti === 0 ? 10 : 8.5;
              doc.setFontSize(titleFontSize);
              doc.setTextColor(15, 23, 42);
              doc.setFont('helvetica', ti === 0 ? 'bold' : 'normal');

              const centerX = leftMm + printableWidth / 2;
              doc.text(titleLines[ti], centerX, curY, {
                align: 'center',
                maxWidth: printableWidth - 4
              });
              curY += titleFontSize * 0.45 + 2.5;
            }
            curY += 1.5;
          } else if (sheetsToRender.length > 1 || tableSections.length > 1) {
            doc.setFontSize(9);
            doc.setTextColor(51, 65, 85);
            doc.setFont('helvetica', 'bold');
            const secLabel = tableSections.length > 1 ? `Sheet: ${sName} (Tabel ${secIdx + 1})` : `Sheet: ${sName}`;
            doc.text(secLabel, leftMm, curY);
            curY += 5;
          }

          // 6. Split header vs body rows
          const headPart = allTableRows.slice(0, headerRowsCount);
          const bodyPart = allTableRows.slice(headerRowsCount);

          // Detect column types for smart column coloring
          const detectedColTypes = [];
          if (headPart.length > 0) {
            const lastHeaderRow = headPart[headPart.length - 1];
            for (let c = 0; c < lastHeaderRow.length; c++) {
              const hText = (lastHeaderRow[c] && lastHeaderRow[c].content) || '';
              detectedColTypes.push(detectColType(hText));
            }
          }

          // 7. Dynamic font size & cell padding based on column count
          const fitMode = options.fit_mode || 'auto';
          const isAutoFit = fitMode === 'auto';
          const useHorizontalBreak = isAutoFit && totalCols > 10;

          let fontSize = 7;
          let cellPadding = 1.0;
          if (useHorizontalBreak) {
            fontSize = 7.5;
            cellPadding = 1.0;
          } else if (totalCols > 25) {
            fontSize = 4.5;
            cellPadding = 0.55;
          } else if (totalCols > 16) {
            fontSize = 5.5;
            cellPadding = 0.75;
          } else if (totalCols > 10) {
            fontSize = 6.5;
            cellPadding = 0.9;
          } else if (totalCols <= 6) {
            fontSize = 8.5;
            cellPadding = 1.4;
          }

          // 8. Calculate proportional column widths from Excel metadata
          const columnStyles = {};
          if (!useHorizontalBreak && colWidthChars.length > 0) {
            const totalCharW = colWidthChars.reduce((a, b) => a + b, 0);
            if (totalCharW > 0) {
              for (let ci = 0; ci < colWidthChars.length; ci++) {
                const pct = colWidthChars[ci] / totalCharW;
                columnStyles[ci] = { cellWidth: Math.max(pct * printableWidth, 8) };
              }
            }
          }

          // 9. AutoTable invocation with smart styling
          const autoTableConfig = {
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
            tableWidth: useHorizontalBreak ? 'wrap' : printableWidth,
            columnStyles: Object.keys(columnStyles).length > 0 ? columnStyles : undefined,
            styles: {
              fontSize: fontSize,
              cellPadding: cellPadding,
              lineWidth: 0.15,
              lineColor: [148, 163, 184],
              textColor: [30, 41, 59],
              overflow: 'linebreak',
              valign: 'middle',
              font: 'helvetica'
            },
            headStyles: {
              fillColor: [241, 245, 249],
              textColor: [15, 23, 42],
              fontStyle: 'bold',
              halign: 'center',
              lineWidth: 0.2,
              lineColor: [100, 116, 139],
              valign: 'middle'
            },
            alternateRowStyles: {
              fillColor: [248, 250, 252]
            },
            didParseCell: function (data) {
              // Guarantee visible crisp borders on all cells
              data.cell.styles.lineWidth = 0.15;
              data.cell.styles.lineColor = [148, 163, 184];

              // PRIORITY 1: Excel original fill color (highest priority)
              const rawCell = data.cell.raw;
              if (rawCell && rawCell._excelFillColor) {
                data.cell.styles.fillColor = rawCell._excelFillColor;
                const [fr, fg, fb] = rawCell._excelFillColor;
                const luminance = (0.299 * fr + 0.587 * fg + 0.114 * fb);
                data.cell.styles.textColor = luminance < 128 ? [255, 255, 255] : [15, 23, 42];
                return; // Excel color takes full precedence
              }

              // PRIORITY 2: Smart column palette based on header keywords
              const colIdx = data.column.index;
              const cType = detectedColTypes[colIdx];
              if (cType && SMART_COL_PALETTE[cType]) {
                const pal = SMART_COL_PALETTE[cType];
                if (data.section === 'head') {
                  data.cell.styles.fillColor = pal.head;
                  data.cell.styles.textColor = pal.text;
                } else if (data.section === 'body') {
                  const isTotal = data.row.raw && Array.isArray(data.row.raw) && data.row.raw.some(c => (c.styles && c.styles.fontStyle === 'bold'));
                  if (!isTotal) {
                    data.cell.styles.fillColor = pal.body;
                    data.cell.styles.textColor = pal.text;
                  }
                }
              }
            }
          };

          if (useHorizontalBreak) {
            autoTableConfig.horizontalPageBreak = true;
            autoTableConfig.horizontalPageBreakRepeat = [0, 1];
            autoTableConfig.horizontalPageBreakBehaviour = 'afterAllRows';
          }

          doc.autoTable(autoTableConfig);
        }
      }

      // 9. Add professional page numbers in the footer of every page
      const totalPages = doc.internal.getNumberOfPages();
      for (let p = 1; p <= totalPages; p++) {
        doc.setPage(p);
        doc.setFontSize(7);
        doc.setTextColor(140, 150, 160);
        doc.setFont('helvetica', 'normal');
        const footerText = `Hal ${p} / ${totalPages}`;
        const textWidth = doc.getTextWidth(footerText);
        doc.text(footerText, pageWidth - rightMm - textWidth, pageHeight - Math.max(bottomMm / 2, 4));
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
        engine: "Client-Side Engine v4.0 (Smart Precision)",
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
        // ALWAYS pass a cloned copy as Uint8Array so PDF.js Web Worker never detaches original ArrayBuffer
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
     * Saves or shares the PDF based on platform (Native Android APK vs Browser)
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

      // 2. Mobile Browser (Android Tablet, iPad, Mobile Phone)
      const isMobile = /Android|iPhone|iPad|iPod/i.test(navigator.userAgent);
      if (isMobile) {
        try {
          const file = new File([pdfBlob], safeName, { type: 'application/pdf' });
          const url = URL.createObjectURL(file);
          const a = document.createElement('a');
          a.style.display = 'none';
          a.href = url;
          a.download = safeName;
          document.body.appendChild(a);
          a.click();
          setTimeout(() => {
            document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(url), 60000);
          }, 200);
          return;
        } catch (e) {
          console.warn("Mobile blob download trigger:", e);
        }
      }

      // 3. Desktop / General Browser Download
      try {
        const url = URL.createObjectURL(pdfBlob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = safeName;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
          document.body.removeChild(a);
          setTimeout(() => URL.revokeObjectURL(url), 60000);
        }, 200);
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

      // 3. Fallback
      this.deliverPdf(pdfBlob, safeName);
    }
  };

  window.OfflineConverter = OfflineConverter;
})(window);
