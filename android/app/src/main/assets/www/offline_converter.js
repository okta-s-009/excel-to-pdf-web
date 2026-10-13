/**
 * Excel2PDF Studio - Offline Client-Side Conversion Engine v3.5
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

      // Pre-scan all sheets to find global max columns for auto-landscape
      let globalMaxCols = 0;
      for (const sName of sheetsToRender) {
        const ws = workbook.Sheets[sName];
        if (!ws || !ws['!ref']) continue;
        const range = window.XLSX.utils.decode_range(ws['!ref']);
        const cols = range.e.c - range.s.c + 1;
        if (cols > globalMaxCols) globalMaxCols = cols;
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
        const totalCols = range.e.c - range.s.c + 1;
        const totalRows = range.e.r - range.s.r + 1;
        if (totalCols === 0 || totalRows === 0) continue;

        // 1. Build lookup of merged cells by absolute row (r) and col (c)
        const merges = ws['!merges'] || [];
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
          for (let c = range.s.c; c <= range.e.c; c++) {
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

          if (filled.length <= 3 && totalCols > 4 && r < range.s.r + 6) {
            titleLines.push(filled.join(' - '));
          } else {
            tableStartR = r;
            break;
          }
        }

        // 3. Determine number of header rows (e.g. multi-tier headers)
        let headerRowsCount = 1;
        for (let c = range.s.c; c <= range.e.c; c++) {
          const m = (mergeMap[tableStartR] && mergeMap[tableStartR][c]) || null;
          if (m && m.origin && m.rs > headerRowsCount) {
            headerRowsCount = m.rs;
          }
        }

        // 4. Construct table rows with colSpan and rowSpan
        const allTableRows = [];
        for (let r = tableStartR; r <= range.e.r; r++) {
          const cells = [];
          for (let c = range.s.c; c <= range.e.c; c++) {
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
            } else if (/^[\d.,Rp\s\-]+$/.test(val) && /\d/.test(val) && !/[a-zA-Z]/.test(val.replace(/Rp/gi, ''))) {
              halign = 'right';
            } else if (val.length <= 4 && /^\d+$/.test(val)) {
              halign = 'center';
            } else if (val.length <= 5 && /^(M3|Sak|Kg|Buah|Batang|Keping|Lembar|Unit|m'|m2|m3)$/i.test(val)) {
              halign = 'center';
            }

            const cellObj = {
              content: val,
              styles: { halign: halign }
            };

            if (m && m.origin) {
              if (m.rs > 1) cellObj.rowSpan = m.rs;
              if (m.cs > 1) cellObj.colSpan = m.cs;
            }
            cells.push(cellObj);
          }

          if (cells.some(c => c.content !== '')) {
            allTableRows.push(cells);
          }
        }

        if (allTableRows.length === 0) continue;

        // Add page for subsequent sheets
        if (!isFirstTable) {
          doc.addPage([pageWidth, pageHeight], orientation);
        }
        isFirstTable = false;

        // 5. Render titles above table
        let curY = Math.max(topMm, 8);
        if (titleLines.length > 0) {
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
        } else if (sheetsToRender.length > 1) {
          doc.setFontSize(9);
          doc.setTextColor(51, 65, 85);
          doc.setFont('helvetica', 'bold');
          doc.text(`Sheet: ${sName}`, leftMm, curY);
          curY += 5;
        }

        // 6. Split header vs body rows
        const headPart = allTableRows.slice(0, headerRowsCount);
        const bodyPart = allTableRows.slice(headerRowsCount);

        // 7. Dynamic font size & cell padding based on column count
        let fontSize = 7;
        let cellPadding = 1.0;
        if (totalCols > 25) {
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

        // 8. AutoTable invocation
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
          styles: {
            fontSize: fontSize,
            cellPadding: cellPadding,
            lineWidth: showGridlines ? 0.1 : 0.05,
            lineColor: [203, 213, 225],
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
            lineWidth: 0.15,
            lineColor: [148, 163, 184],
            valign: 'middle'
          },
          alternateRowStyles: {
            fillColor: [248, 250, 252]
          }
        });
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
        engine: "Client-Side Engine v3.5 (Precision AutoTable)",
        auto_landscape: autoLandscapeTriggered
      };
    },

    /**
     * Generates visual page preview thumbnails from PDF buffer using PDF.js.
     */
    generatePreviews: async function (pdfArrayBuffer, maxPages = 1) {
      if (!window.pdfjsLib) {
        return { total_pages: 1, previews: [] };
      }

      try {
        const loadingTask = window.pdfjsLib.getDocument({ data: pdfArrayBuffer });
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
      if (!window.pdfjsLib) return null;
      try {
        const loadingTask = window.pdfjsLib.getDocument({ data: pdfArrayBuffer });
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
      if (window.AndroidBridge) {
        const reader = new FileReader();
        reader.onloadend = function () {
          const base64Data = reader.result.split(',')[1];
          if (window.AndroidBridge.savePdf) {
            window.AndroidBridge.savePdf(base64Data, filename);
          } else if (window.AndroidBridge.sharePdf) {
            window.AndroidBridge.sharePdf(base64Data, filename);
          }
        };
        reader.readAsDataURL(pdfBlob);
        return;
      }

      const url = URL.createObjectURL(pdfBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 1000);
    },

    /**
     * Triggers native Android Share sheet if in APK or Web Share API
     */
    sharePdf: function (pdfBlob, filename) {
      if (window.AndroidBridge && window.AndroidBridge.sharePdf) {
        const reader = new FileReader();
        reader.onloadend = function () {
          const base64Data = reader.result.split(',')[1];
          window.AndroidBridge.sharePdf(base64Data, filename);
        };
        reader.readAsDataURL(pdfBlob);
        return;
      }

      if (navigator.canShare && navigator.canShare({ files: [new File([pdfBlob], filename, { type: 'application/pdf' })] })) {
        navigator.share({
          files: [new File([pdfBlob], filename, { type: 'application/pdf' })],
          title: filename,
          text: 'PDF hasil konversi dari Excel2PDF Studio'
        }).catch(() => {
          this.deliverPdf(pdfBlob, filename);
        });
      } else {
        this.deliverPdf(pdfBlob, filename);
      }
    }
  };

  window.OfflineConverter = OfflineConverter;
})(window);
