/**
 * Excel2PDF Studio - Offline Client-Side Conversion Engine
 * Runs 100% locally in the browser or inside an Android WebView APK.
 * Uses: SheetJS (xlsx), jsPDF, jsPDF-AutoTable, and PDF.js
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
      const workbook = window.XLSX.read(buffer, { type: 'array', cellDates: true });
      const sheets = workbook.SheetNames || [];
      const details = [];

      sheets.forEach(name => {
        const ws = workbook.Sheets[name];
        let maxRow = 0;
        let maxCol = 0;
        if (ws && ws['!ref']) {
          const range = window.XLSX.utils.decode_range(ws['!ref']);
          maxRow = range.e.r + 1;
          maxCol = range.e.c + 1;
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
        workbook: workbook // cached in memory for immediate conversion
      };
    },

    /**
     * Converts Excel workbook to PDF client-side.
     * Enhanced: Smart title detection, auto-landscape, content-weighted columns,
     * Excel-like header styling, and number/currency right-alignment.
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

      const fitMode = options.fit_mode || 'fit_width';
      const showGridlines = Boolean(options.show_gridlines);
      const showHeadings = Boolean(options.show_headings);
      const selectedSheet = options.sheet_name || 'all';

      // --- Pre-scan all sheets to detect max columns for auto-landscape ---
      let sheetsToRender = [];
      if (selectedSheet !== 'all' && workbook.Sheets[selectedSheet]) {
        sheetsToRender.push(selectedSheet);
      } else {
        sheetsToRender = workbook.SheetNames || [];
      }

      let globalMaxCols = 0;
      const sheetDataCache = {};
      for (const sName of sheetsToRender) {
        const ws = workbook.Sheets[sName];
        if (!ws || !ws['!ref']) continue;
        const rawData = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
        if (!rawData || rawData.length === 0) continue;
        let cleanRows = rawData.filter(row => row.some(cell => String(cell).trim() !== ''));
        if (cleanRows.length === 0) continue;
        let maxCols = 0;
        cleanRows.forEach(r => { if (r.length > maxCols) maxCols = r.length; });
        if (maxCols > globalMaxCols) globalMaxCols = maxCols;
        sheetDataCache[sName] = { cleanRows, maxCols };
      }

      // AUTO-LANDSCAPE: If user chose portrait but table has > 7 columns, auto-switch to landscape
      if (orientation === 'portrait' && globalMaxCols > 7) {
        orientation = 'landscape';
      }

      const pageWidth = orientation === 'landscape' ? paperDim.height : paperDim.width;
      const pageHeight = orientation === 'landscape' ? paperDim.width : paperDim.height;

      // Initialize jsPDF
      const { jsPDF } = window.jspdf || window;
      const doc = new jsPDF({
        orientation: orientation === 'landscape' ? 'landscape' : 'portrait',
        unit: 'mm',
        format: [pageWidth, pageHeight],
        compress: true
      });

      // Printable usable width & height in mm
      const printableWidth = Math.max(pageWidth - leftMm - rightMm, 20);

      let isFirstTable = true;

      for (let sIdx = 0; sIdx < sheetsToRender.length; sIdx++) {
        const sName = sheetsToRender[sIdx];
        const cached = sheetDataCache[sName];
        if (!cached) continue;

        let cleanRows = cached.cleanRows;
        let maxCols = cached.maxCols;

        // Ensure every row has maxCols entries
        cleanRows = cleanRows.map(row => {
          const filled = [...row];
          while (filled.length < maxCols) filled.push('');
          return filled;
        });

        // ======= SMART TITLE ROW DETECTION =======
        // Scan the first few rows: if a row has <= 2 filled cells and long text (>20 chars),
        // it is a title/subtitle row. Extract these and render as centered text above the table.
        const titleLines = [];
        let dataStartIdx = 0;

        for (let ri = 0; ri < Math.min(cleanRows.length, 5); ri++) {
          const row = cleanRows[ri];
          const filledCells = row.filter(c => String(c).trim() !== '');
          const filledCount = filledCells.length;
          const longestCell = filledCells.reduce((max, c) => Math.max(max, String(c).length), 0);

          // Title criteria: very few filled cells (1-2) AND at least one has long text
          // OR the row is completely empty (spacer row between title and data)
          if (filledCount === 0) {
            // Empty spacer row - skip it but keep scanning
            dataStartIdx = ri + 1;
            continue;
          }
          if (filledCount <= 3 && maxCols > 3) {
            // This is a title row
            const titleText = filledCells.join(' ').trim();
            titleLines.push(titleText);
            dataStartIdx = ri + 1;
          } else {
            // Real data starts here
            break;
          }
        }

        // Separate actual data rows (after title rows)
        const dataRows = cleanRows.slice(dataStartIdx);
        if (dataRows.length === 0) continue;

        // ======= HEADER & BODY SEPARATION =======
        let headerRow = null;
        let bodyRows = dataRows;
        let multiRowHeader = [];

        if (showHeadings) {
          const colLetters = ['#'];
          for (let c = 0; c < maxCols; c++) {
            colLetters.push(window.XLSX.utils.encode_col(c));
          }
          headerRow = colLetters;
          bodyRows = dataRows.map((r, rIdx) => [String(rIdx + 1), ...r]);
          maxCols += 1;
        } else {
          // Detect multi-row headers: consecutive rows where most cells contain short text
          // and no cells look like pure numbers
          let headerEndIdx = 0;
          for (let ri = 0; ri < Math.min(dataRows.length, 4); ri++) {
            const row = dataRows[ri];
            const filledCells = row.filter(c => String(c).trim() !== '');
            const filledRatio = filledCells.length / maxCols;
            // Header rows typically have > 40% filled cells and mostly short text
            const avgLen = filledCells.reduce((s, c) => s + String(c).length, 0) / (filledCells.length || 1);
            const hasNumbers = filledCells.some(c => /^\d+[\d.,]*$/.test(String(c).trim()) && String(c).trim().length > 2);
            
            if (ri === 0 && filledRatio > 0.3) {
              // First row is always treated as header if it has enough content
              headerEndIdx = ri + 1;
            } else if (ri > 0 && filledRatio > 0.3 && avgLen < 30 && !hasNumbers) {
              // Subsequent header row (sub-headers like "Volume" spanning columns)
              headerEndIdx = ri + 1;
            } else {
              break;
            }
          }

          if (headerEndIdx > 0 && dataRows.length > headerEndIdx) {
            if (headerEndIdx === 1) {
              headerRow = dataRows[0];
              bodyRows = dataRows.slice(1);
            } else {
              // Multi-row header
              multiRowHeader = dataRows.slice(0, headerEndIdx);
              bodyRows = dataRows.slice(headerEndIdx);
            }
          } else {
            headerRow = null;
            bodyRows = dataRows;
          }
        }

        // ======= COLUMN TYPE DETECTION (for alignment & width weighting) =======
        const colMeta = [];
        for (let c = 0; c < maxCols; c++) {
          let numCount = 0;
          let textCount = 0;
          let maxTextLen = 0;
          let totalLen = 0;
          let sampleCount = 0;
          let isCurrency = false;

          // Sample from body rows
          const samplesToCheck = Math.min(bodyRows.length, 30);
          for (let r = 0; r < samplesToCheck; r++) {
            const val = String(bodyRows[r][c] || '').trim();
            if (val === '') continue;
            sampleCount++;
            totalLen += val.length;
            if (val.length > maxTextLen) maxTextLen = val.length;

            // Check if it looks like a number or currency
            const cleanVal = val.replace(/[Rp.,\s]/g, '');
            if (/^-?\d+$/.test(cleanVal) && cleanVal.length > 0) {
              numCount++;
              if (/Rp|rupiah/i.test(val) || /^\d{1,3}([.,]\d{3})+/.test(val) || val.includes('.000')) {
                isCurrency = true;
              }
            } else {
              textCount++;
            }
          }

          const avgLen = sampleCount > 0 ? totalLen / sampleCount : 5;
          const isNumeric = sampleCount > 0 && numCount / sampleCount > 0.6;
          const isIndex = maxTextLen <= 4 && isNumeric;

          colMeta.push({
            isNumeric,
            isIndex,
            isCurrency,
            maxTextLen,
            avgLen,
            // Weight for proportional column width
            weight: isIndex ? 1 : (isNumeric ? Math.max(avgLen * 0.8, 3) : Math.max(avgLen, 4))
          });
        }

        // Also factor in header text lengths for column widths
        const headersToCheck = multiRowHeader.length > 0 ? multiRowHeader : (headerRow ? [headerRow] : []);
        for (const hRow of headersToCheck) {
          for (let c = 0; c < Math.min(hRow.length, maxCols); c++) {
            const hLen = String(hRow[c] || '').length;
            if (hLen > colMeta[c].weight) {
              colMeta[c].weight = Math.max(colMeta[c].weight, hLen * 0.7);
            }
          }
        }

        // ======= COMPUTE COLUMN STYLES (Intelligent sizing & Wide Table Protection) =======
        const isWideTable = maxCols > 12;
        const columnStyles = {};

        if (isWideTable) {
          // For wide tables (e.g. RAB with 15-30 columns):
          // Give every column its comfortable natural width so numbers and words never wrap awkwardly.
          // jsPDF AutoTable will cleanly split excess columns across pages using horizontalPageBreak.
          for (let c = 0; c < maxCols; c++) {
            const m = colMeta[c];
            let widthMm = 16; // default comfortable width
            if (m.isIndex) {
              widthMm = 10;
            } else if (m.isCurrency) {
              widthMm = 28; // Fits 'Rp 115.000,00' on a single clean line
            } else if (m.isNumeric) {
              widthMm = 16;
            } else if (c === 1 || m.maxTextLen > 25) {
              widthMm = 40; // Wide room for 'Uraian Pekerjaan' / long names
            } else {
              widthMm = Math.max(m.avgLen * 1.8, 18);
            }

            columnStyles[c] = {
              cellWidth: widthMm,
              halign: m.isNumeric ? 'right' : (m.isIndex ? 'center' : 'left')
            };
          }
        } else {
          // For standard tables (<= 12 columns):
          // Scale proportionally to fill printable page width perfectly
          const totalWeight = colMeta.reduce((s, m) => s + m.weight, 0);
          for (let c = 0; c < maxCols; c++) {
            const proportion = colMeta[c].weight / totalWeight;
            const colWidthMm = Math.max(proportion * printableWidth, 10);
            
            columnStyles[c] = {
              cellWidth: colWidthMm,
              halign: colMeta[c].isNumeric ? 'right' : (colMeta[c].isIndex ? 'center' : 'left')
            };
          }
        }

        // Add page if not the first table
        if (!isFirstTable) {
          doc.addPage([pageWidth, pageHeight], orientation);
        }
        isFirstTable = false;

        // ======= RENDER TITLE LINES ABOVE TABLE =======
        let currentY = topMm;
        if (titleLines.length > 0) {
          for (let ti = 0; ti < titleLines.length; ti++) {
            const titleText = titleLines[ti];
            const titleFontSize = ti === 0 ? 11 : 9;
            doc.setFontSize(titleFontSize);
            doc.setTextColor(15, 23, 42);
            doc.setFont('helvetica', ti === 0 ? 'bold' : 'normal');

            // Center the title text in the printable area
            const centerX = leftMm + printableWidth / 2;
            doc.text(titleText, centerX, currentY + 4, { align: 'center', maxWidth: printableWidth - 4 });
            currentY += titleFontSize * 0.5 + 3;
          }
          currentY += 2; // Small gap before table
        }

        // Sheet name label for multi-sheet documents
        if (sheetsToRender.length > 1 && titleLines.length === 0) {
          doc.setFontSize(9);
          doc.setTextColor(100, 116, 139);
          doc.setFont('helvetica', 'italic');
          doc.text(`Sheet: ${sName}`, leftMm, currentY + 3);
          currentY += 6;
        }

        // ======= DYNAMIC FONT SIZE =======
        let fontSize = 8;
        let cellPadding = 1.4;

        if (!isWideTable) {
          if (maxCols > 9) {
            fontSize = 7.5;
            cellPadding = 1.2;
          } else if (maxCols <= 6) {
            fontSize = 9;
            cellPadding = 1.8;
          }
        } else {
          // In wide tables with horizontalPageBreak, columns have ample room
          fontSize = 7;
          cellPadding = 1.0;
        }

        // ======= BUILD AUTOTABLE CONFIG =======
        const headConfig = multiRowHeader.length > 0 
          ? multiRowHeader 
          : (headerRow ? [headerRow] : []);

        const autoTableConfig = {
          head: headConfig,
          body: bodyRows,
          startY: currentY,
          margin: {
            top: topMm,
            bottom: bottomMm,
            left: leftMm,
            right: rightMm
          },
          tableWidth: isWideTable ? 'wrap' : printableWidth,
          horizontalPageBreak: isWideTable,
          horizontalPageBreakRepeat: isWideTable ? [0, 1] : null,
          theme: 'grid',
          columnStyles: columnStyles,
          styles: {
            fontSize: fontSize,
            cellPadding: cellPadding,
            overflow: 'linebreak',
            font: 'helvetica',
            lineWidth: 0.15,
            lineColor: [203, 213, 225],
            textColor: [30, 41, 59],
            valign: 'middle'
          },
          headStyles: {
            fillColor: [241, 245, 249],
            textColor: [15, 23, 42],
            fontStyle: 'bold',
            halign: 'center',
            lineWidth: 0.2,
            lineColor: [148, 163, 184],
            valign: 'middle',
            cellPadding: Math.max(cellPadding, 1.0)
          },
          alternateRowStyles: {
            fillColor: [248, 250, 252]
          },
          didDrawPage: function (data) {
            // Page number in footer margin
            const pageNum = doc.internal.getNumberOfPages();
            doc.setFontSize(7);
            doc.setTextColor(140, 150, 160);
            doc.setFont('helvetica', 'normal');
            const footerText = `Hal ${pageNum}`;
            const textWidth = doc.getTextWidth(footerText);
            doc.text(footerText, pageWidth - rightMm - textWidth, pageHeight - Math.max(bottomMm / 2, 4));

            // Re-draw sheet name on continuation pages for multi-sheet docs
            if (sheetsToRender.length > 1) {
              doc.text(`Sheet: ${sName}`, leftMm, Math.max(topMm - 2, 4));
            }
          }
        };

        // Render table
        if (doc.autoTable) {
          doc.autoTable(autoTableConfig);
        } else {
          throw new Error("Plugin jsPDF-AutoTable belum tersedia.");
        }
      }

      // Finalize PDF
      const pdfBlob = doc.output('blob');
      const pdfArrayBuffer = doc.output('arraybuffer');
      const totalPages = doc.internal.getNumberOfPages();
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
        engine: "Offline Client-Side Engine (SheetJS + jsPDF)",
        auto_landscape: orientation === 'landscape' && (options.orientation || 'portrait').toLowerCase() === 'portrait'
      };
    },

    /**
     * Generates visual page preview thumbnails from PDF buffer using PDF.js.
     * Uses optimized JPEG rendering for blazingly fast preview generation.
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

          const renderContext = {
            canvasContext: context,
            viewport: viewport
          };

          await page.render(renderContext).promise;
          const dataUrl = canvas.toDataURL('image/jpeg', 0.85);

          previews.push({
            page: i,
            url: dataUrl
          });
        }

        return {
          total_pages: totalPages,
          previews: previews
        };
      } catch (err) {
        console.warn("Pratinjau cepat PDF.js:", err);
        return { total_pages: 1, previews: [] };
      }
    },

    /**
     * Renders a single specific page on demand in ~50ms.
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
        const dataUrl = canvas.toDataURL('image/jpeg', 0.85);
        return {
          page: pageNumber,
          url: dataUrl
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
      // 1. If running inside Android APK with JavascriptInterface
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

      // 2. Standard Browser Blob Download
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
