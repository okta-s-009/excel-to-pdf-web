/**
 * Excel2PDF Studio - Offline Client-Side Conversion Engine
 * Runs 100% locally in the browser or inside an Android WebView APK.
 * Uses: SheetJS (xlsx), jsPDF, jsPDF-AutoTable, and PDF.js
 */

(function (window) {
  'use strict';

  // Configure PDF.js worker if available
  if (window.pdfjsLib) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = '/static/libs/pdf.worker.min.js';
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
     */
    convertToPdf: async function (fileOrWorkbook, options) {
      if (!this.isReady()) {
        throw new Error("Pustaka konversi offline belum siap (XLSX / jsPDF).");
      }

      const startTime = performance.now();

      // Resolve workbook
      let workbook = null;
      let filename = "dokumen.xlsx";

      if (fileOrWorkbook instanceof File || fileOrWorkbook instanceof Blob) {
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

      const orientation = (options.orientation || 'portrait').toLowerCase();
      const paperSizeName = options.paper_size || 'A4';
      const paperDim = PAPER_DIMENSIONS_MM[paperSizeName] || PAPER_DIMENSIONS_MM.A4;

      const pageWidth = orientation === 'landscape' ? paperDim.height : paperDim.width;
      const pageHeight = orientation === 'landscape' ? paperDim.width : paperDim.height;

      const fitMode = options.fit_mode || 'fit_width';
      const showGridlines = Boolean(options.show_gridlines);
      const showHeadings = Boolean(options.show_headings);
      const selectedSheet = options.sheet_name || 'all';

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
      const printableHeight = Math.max(pageHeight - topMm - bottomMm, 20);

      // Determine sheets to process
      let sheetsToRender = [];
      if (selectedSheet !== 'all' && workbook.Sheets[selectedSheet]) {
        sheetsToRender.push(selectedSheet);
      } else {
        sheetsToRender = workbook.SheetNames || [];
      }

      let isFirstTable = true;

      for (let sIdx = 0; sIdx < sheetsToRender.length; sIdx++) {
        const sName = sheetsToRender[sIdx];
        const ws = workbook.Sheets[sName];
        if (!ws || !ws['!ref']) continue;

        // Convert worksheet to 2D array
        const rawData = window.XLSX.utils.sheet_to_json(ws, {
          header: 1,
          defval: '',
          raw: false
        });

        if (!rawData || rawData.length === 0) continue;

        // Filter out completely empty trailing rows
        let cleanRows = rawData.filter(row => row.some(cell => String(cell).trim() !== ''));
        if (cleanRows.length === 0) continue;

        // Calculate maximum column count
        let maxCols = 0;
        cleanRows.forEach(r => {
          if (r.length > maxCols) maxCols = r.length;
        });

        // Ensure every row has maxCols entries
        cleanRows = cleanRows.map(row => {
          const filled = [...row];
          while (filled.length < maxCols) filled.push('');
          return filled;
        });

        // If show headings is true, prepend row indices and column letters
        let headerRow = null;
        let bodyRows = cleanRows;

        if (showHeadings) {
          const colLetters = ['#'];
          for (let c = 0; c < maxCols; c++) {
            colLetters.push(window.XLSX.utils.encode_col(c));
          }
          headerRow = colLetters;

          bodyRows = cleanRows.map((r, rIdx) => [String(rIdx + 1), ...r]);
          maxCols += 1;
        } else {
          // Use first row as header if more than 1 row exists
          if (cleanRows.length > 1) {
            headerRow = cleanRows[0];
            bodyRows = cleanRows.slice(1);
          } else {
            headerRow = null;
            bodyRows = cleanRows;
          }
        }

        // Add page if not the first table
        if (!isFirstTable) {
          doc.addPage([pageWidth, pageHeight], orientation);
        }
        isFirstTable = false;

        // Dynamic font size and cell padding based on column count and fitMode
        let fontSize = 9;
        let cellPadding = 1.8;

        if (fitMode === 'fit_width' || fitMode === 'fit_page') {
          if (maxCols > 20) {
            fontSize = 5.5;
            cellPadding = 0.8;
          } else if (maxCols > 14) {
            fontSize = 6.5;
            cellPadding = 1.0;
          } else if (maxCols > 8) {
            fontSize = 7.5;
            cellPadding = 1.3;
          }
        }

        // AutoTable options
        const autoTableConfig = {
          head: headerRow ? [headerRow] : [],
          body: bodyRows,
          startY: topMm,
          margin: {
            top: topMm,
            bottom: bottomMm,
            left: leftMm,
            right: rightMm
          },
          tableWidth: printableWidth,
          theme: showGridlines ? 'grid' : 'plain',
          styles: {
            fontSize: fontSize,
            cellPadding: cellPadding,
            overflow: 'linebreak',
            font: 'helvetica',
            lineWidth: showGridlines ? 0.15 : 0,
            lineColor: showGridlines ? [210, 215, 220] : [255, 255, 255],
            textColor: [30, 41, 59]
          },
          headStyles: {
            fillColor: [30, 41, 59],
            textColor: [255, 255, 255],
            fontStyle: 'bold',
            halign: 'left',
            lineWidth: showGridlines ? 0.2 : 0,
            lineColor: [15, 23, 42]
          },
          alternateRowStyles: {
            fillColor: [248, 250, 252]
          },
          didDrawPage: function (data) {
            // Draw header / footer pagination
            const pageNum = doc.internal.getNumberOfPages();
            doc.setFontSize(8);
            doc.setTextColor(140, 150, 160);

            // Sheet title in header margin if multi-sheet
            if (sheetsToRender.length > 1) {
              doc.text(`Sheet: ${sName}`, leftMm, Math.max(topMm - 3, 5));
            }

            // Page number in footer margin
            const footerText = `Hal ${pageNum}`;
            const textWidth = doc.getTextWidth(footerText);
            doc.text(footerText, pageWidth - rightMm - textWidth, pageHeight - Math.max(bottomMm / 2, 5));
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
        engine: "Offline Client-Side Engine (SheetJS + jsPDF)"
      };
    },

    /**
     * Generates visual page preview thumbnails from PDF buffer using PDF.js
     */
    generatePreviews: async function (pdfArrayBuffer, maxPages = 5) {
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
          const viewport = page.getViewport({ scale: 1.5 });

          const canvas = document.createElement('canvas');
          const context = canvas.getContext('2d');
          canvas.height = viewport.height;
          canvas.width = viewport.width;

          const renderContext = {
            canvasContext: context,
            viewport: viewport
          };

          await page.render(renderContext).promise;
          const dataUrl = canvas.toDataURL('image/png');

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
        console.warn("Gagal membuat thumbnail preview dengan PDF.js:", err);
        return { total_pages: 1, previews: [] };
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
