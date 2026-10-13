/**
 * Excel2PDF Studio - Interactive Client Logic
 * Supports both FastAPI Desktop Server and Standalone Offline / Android APK modes.
 */

document.addEventListener('DOMContentLoaded', () => {
  // --- App State ---
  const state = {
    currentFile: null,
    uploadId: null,
    cachedWorkbook: null,
    pdfArrayBuffer: null,
    unit: 'cm', // 'cm' or 'inch'
    preset: 'normal',
    orientation: 'portrait',
    paperSize: 'A4',
    fitMode: 'fit_width',
    convertedTask: null,
    currentPage: 1,
    zoomLevel: 100, // percentage
    isOfflineMode: false,
    lastPdfBlob: null,
    lastPdfUrl: null
  };

  // Presets in CM and INCH
  const PRESETS = {
    cm: {
      normal: { top: 1.91, bottom: 1.91, left: 1.78, right: 1.78, header: 0.76, footer: 0.76, desc: 'A: 1.91 | B: 1.78' },
      narrow: { top: 1.27, bottom: 1.27, left: 0.64, right: 0.64, header: 0.76, footer: 0.76, desc: 'A: 1.27 | B: 0.64' },
      wide:   { top: 2.54, bottom: 2.54, left: 2.54, right: 2.54, header: 1.27, footer: 1.27, desc: 'A: 2.54 | B: 2.54' },
      custom: { desc: 'Atur Bebas' }
    },
    inch: {
      normal: { top: 0.75, bottom: 0.75, left: 0.70, right: 0.70, header: 0.30, footer: 0.30, desc: 'A: 0.75 | B: 0.70' },
      narrow: { top: 0.50, bottom: 0.50, left: 0.25, right: 0.25, header: 0.30, footer: 0.30, desc: 'A: 0.50 | B: 0.25' },
      wide:   { top: 1.00, bottom: 1.00, left: 1.00, right: 1.00, header: 0.50, footer: 0.50, desc: 'A: 1.00 | B: 1.00' },
      custom: { desc: 'Atur Bebas' }
    }
  };

  // --- DOM Elements ---
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');
  const fileInfoCard = document.getElementById('fileInfoCard');
  const fileNameDisplay = document.getElementById('fileNameDisplay');
  const fileSizeDisplay = document.getElementById('fileSizeDisplay');
  const fileSheetsCountDisplay = document.getElementById('fileSheetsCountDisplay');
  const removeFileBtn = document.getElementById('removeFileBtn');
  const sheetSelectGroup = document.getElementById('sheetSelectGroup');
  const sheetSelect = document.getElementById('sheetSelect');

  const unitCmBtn = document.getElementById('unitCmBtn');
  const unitInchBtn = document.getElementById('unitInchBtn');
  const presetCards = document.querySelectorAll('.preset-card');

  const marginTop = document.getElementById('marginTop');
  const marginBottom = document.getElementById('marginBottom');
  const marginLeft = document.getElementById('marginLeft');
  const marginRight = document.getElementById('marginRight');
  const marginHeader = document.getElementById('marginHeader');
  const marginFooter = document.getElementById('marginFooter');
  const unitLabels = document.querySelectorAll('.current-unit');

  const orientationCards = document.querySelectorAll('.orientation-card');
  const paperSizeSelect = document.getElementById('paperSize');
  const fitModeSelect = document.getElementById('fitMode');
  const centerHorizontally = document.getElementById('centerHorizontally');
  const centerVertically = document.getElementById('centerVertically');
  const showGridlines = document.getElementById('showGridlines');
  const showHeadings = document.getElementById('showHeadings');

  const convertBtn = document.getElementById('convertBtn');
  const convertBtnText = document.getElementById('convertBtnText');
  const convertIcon = document.getElementById('convertIcon');
  const convertSpinner = document.getElementById('convertSpinner');
  const convertHintText = document.getElementById('convertHintText');

  const tabVisualizerBtn = document.getElementById('tabVisualizerBtn');
  const tabPdfResultBtn = document.getElementById('tabPdfResultBtn');
  const tabVisualizerContent = document.getElementById('tabVisualizerContent');
  const tabPdfResultContent = document.getElementById('tabPdfResultContent');
  const resultBadgeCounter = document.getElementById('resultBadgeCounter');

  // Visualizer elements
  const paperMockup = document.getElementById('paperMockup');
  const paperPrintableArea = document.getElementById('paperPrintableArea');
  const visValTop = document.getElementById('visValTop');
  const visValBottom = document.getElementById('visValBottom');
  const visValLeft = document.getElementById('visValLeft');
  const visValRight = document.getElementById('visValRight');

  // Result elements
  const emptyResultPlaceholder = document.getElementById('emptyResultPlaceholder');
  const conversionSuccessArea = document.getElementById('conversionSuccessArea');
  const resPdfName = document.getElementById('resPdfName');
  const resPdfSize = document.getElementById('resPdfSize');
  const resPdfPages = document.getElementById('resPdfPages');
  const resPdfDuration = document.getElementById('resPdfDuration');
  const downloadPdfBtn = document.getElementById('downloadPdfBtn');
  const sharePdfBtn = document.getElementById('sharePdfBtn');
  const openTabPdfBtn = document.getElementById('openTabPdfBtn');
  const resetAllBtn = document.getElementById('resetAllBtn');

  // Viewer controls
  const prevPageBtn = document.getElementById('prevPageBtn');
  const nextPageBtn = document.getElementById('nextPageBtn');
  const pageIndicatorText = document.getElementById('pageIndicatorText');
  const pdfPageImage = document.getElementById('pdfPageImage');
  const previewImageWrapper = document.getElementById('previewImageWrapper');
  const zoomInBtn = document.getElementById('zoomInBtn');
  const zoomOutBtn = document.getElementById('zoomOutBtn');
  const zoomResetBtn = document.getElementById('zoomResetBtn');
  const zoomLevelText = document.getElementById('zoomLevelText');

  // Modals & Header
  const loadingOverlay = document.getElementById('loadingOverlay');
  const themeToggleBtn = document.getElementById('themeToggleBtn');
  const engineStatusBadge = document.getElementById('engineStatusBadge');
  const engineStatusText = document.getElementById('engineStatusText');
  const toastContainer = document.getElementById('toastContainer');

  // Mobile Connect & PWA Elements
  const mobileConnectBtn = document.getElementById('mobileConnectBtn');
  const installPwaBtn = document.getElementById('installPwaBtn');
  const mobileModalOverlay = document.getElementById('mobileModalOverlay');
  const mobileModalBackdrop = document.getElementById('mobileModalBackdrop');
  const closeMobileModalBtn = document.getElementById('closeMobileModalBtn');
  const mobileUrlDisplay = document.getElementById('mobileUrlDisplay');
  const copyMobileUrlBtn = document.getElementById('copyMobileUrlBtn');
  const qrcodeBox = document.getElementById('qrcodeBox');

  // --- Initial Setup ---
  initTheme();
  checkEngineStatus();
  updateVisualMockup();
  initPwaAndMobileConnect();
  initAboutModal();

  // --- Engine Status Check ---
  async function checkEngineStatus() {
    // If AndroidBridge or running from local file/asset, prioritize offline native engine immediately
    if (window.AndroidBridge || window.location.protocol === 'file:') {
      enableOfflineMode('Android Native Engine (Offline)');
      return;
    }

    // If running on Vercel Cloud, activate super-fast client-side engine immediately (sub-second conversion)
    if (window.location.hostname.includes('vercel.app')) {
      enableOfflineMode('Vercel Cloud (Super Cepat)');
      return;
    }

    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 2000);
      const res = await fetch('/api/status', { signal: controller.signal });
      clearTimeout(timeoutId);

      if (res.ok) {
        const contentType = res.headers.get('content-type') || '';
        if (contentType.includes('application/json')) {
          const data = await res.json();
          if (data.excel_engine_ready) {
            state.isOfflineMode = false;
            engineStatusBadge.className = 'status-pill status-ready';
            engineStatusText.textContent = 'Microsoft Excel COM (PC)';
            engineStatusBadge.title = 'Menggunakan Microsoft Excel COM Native untuk hasil presisi 100%';
            return;
          }
        }
      }
      enableOfflineMode('Mode Mandiri (Super Cepat)');
    } catch (e) {
      enableOfflineMode('Mode Mandiri (Super Cepat)');
    }
  }

  function enableOfflineMode(label) {
    state.isOfflineMode = true;
    engineStatusBadge.className = 'status-pill status-ready';
    engineStatusText.textContent = label || 'Mode Offline (Super Cepat)';
    engineStatusBadge.title = 'Konversi instan langsung di peramban tanpa delay antrean server.';
  }

  // --- Theme Toggle ---
  function initTheme() {
    const savedTheme = localStorage.getItem('e2p_theme') || 'theme-dark';
    document.body.className = savedTheme;
    updateThemeIcon(savedTheme);

    themeToggleBtn.addEventListener('click', () => {
      const isDark = document.body.classList.contains('theme-dark');
      const newTheme = isDark ? 'theme-light' : 'theme-dark';
      document.body.className = newTheme;
      localStorage.setItem('e2p_theme', newTheme);
      updateThemeIcon(newTheme);
    });
  }

  function updateThemeIcon(theme) {
    themeToggleBtn.innerHTML = theme === 'theme-dark' 
      ? '<i class="fa-solid fa-sun"></i>' 
      : '<i class="fa-solid fa-moon"></i>';
  }

  // --- Mobile Connect & PWA Setup ---
  function initPwaAndMobileConnect() {
    // 1. Register Service Worker for PWA (Installable on HP/Tablet)
    if ('serviceWorker' in navigator) {
      window.addEventListener('load', () => {
        navigator.serviceWorker.register('/sw.js')
          .then(reg => console.log('PWA Service Worker registered:', reg.scope))
          .catch(err => console.warn('PWA SW error (ignored):', err));
      });
    }

    // 2. Handle PWA Install Prompt (Chrome/Android)
    let deferredPrompt = null;
    window.addEventListener('beforeinstallprompt', (e) => {
      e.preventDefault();
      deferredPrompt = e;
      if (installPwaBtn) {
        installPwaBtn.classList.remove('hidden');
      }
    });

    if (installPwaBtn) {
      installPwaBtn.addEventListener('click', async () => {
        if (deferredPrompt) {
          deferredPrompt.prompt();
          const { outcome } = await deferredPrompt.userChoice;
          if (outcome === 'accepted') {
            showToast('Aplikasi berhasil dipasang di layar utama!', 'success');
            installPwaBtn.classList.add('hidden');
          }
          deferredPrompt = null;
        } else {
          showToast('Di HP: buka menu browser (titik 3) lalu pilih "Tambahkan ke Layar Utama"', 'info');
        }
      });
    }

    // 3. Fetch Network IP & Setup QR Code
    let currentNetworkUrl = window.location.origin;
    let qrGenerated = false;

    async function fetchNetworkInfo() {
      try {
        const res = await fetch('/api/network-info');
        if (res.ok) {
          const info = await res.json();
          currentNetworkUrl = info.network_url;
          if (mobileUrlDisplay) {
            mobileUrlDisplay.value = currentNetworkUrl;
          }
        }
      } catch (err) {
        if (mobileUrlDisplay) {
          mobileUrlDisplay.value = window.location.href;
        }
      }
    }
    fetchNetworkInfo();

    function renderQrCode() {
      if (!qrcodeBox || typeof QRCode === 'undefined') return;
      const targetUrl = (mobileUrlDisplay && mobileUrlDisplay.value) || currentNetworkUrl || window.location.href;
      qrcodeBox.innerHTML = '';
      try {
        new QRCode(qrcodeBox, {
          text: targetUrl,
          width: 176,
          height: 176,
          colorDark: '#0f172a',
          colorLight: '#ffffff',
          correctLevel: QRCode.CorrectLevel.M
        });
        qrGenerated = true;
      } catch (e) {
        console.warn('QR Code generation error:', e);
      }
    }

    function openMobileModal() {
      if (!mobileModalOverlay) return;
      mobileModalOverlay.classList.remove('hidden');
      if (!qrGenerated) {
        renderQrCode();
      }
    }

    function closeMobileModal() {
      if (!mobileModalOverlay) return;
      mobileModalOverlay.classList.add('hidden');
    }

    if (mobileConnectBtn) {
      mobileConnectBtn.addEventListener('click', openMobileModal);
    }
    if (closeMobileModalBtn) {
      closeMobileModalBtn.addEventListener('click', closeMobileModal);
    }
    if (mobileModalBackdrop) {
      mobileModalBackdrop.addEventListener('click', closeMobileModal);
    }

    // Close on Escape key
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && mobileModalOverlay && !mobileModalOverlay.classList.contains('hidden')) {
        closeMobileModal();
      }
    });

    // Copy URL Button
    if (copyMobileUrlBtn && mobileUrlDisplay) {
      copyMobileUrlBtn.addEventListener('click', async () => {
        try {
          if (navigator.clipboard && navigator.clipboard.writeText) {
            await navigator.clipboard.writeText(mobileUrlDisplay.value);
          } else {
            mobileUrlDisplay.select();
            document.execCommand('copy');
          }
          showToast('Alamat URL berhasil disalin!', 'success');
        } catch (e) {
          showToast('Alamat URL: ' + mobileUrlDisplay.value, 'info');
        }
      });
    }
  }

  // --- Toast Notifications ---
  function showToast(message, type = 'info') {
    if (window.AndroidBridge && window.AndroidBridge.showToast && (type === 'error' || type === 'success')) {
      try { window.AndroidBridge.showToast(message); } catch (e) {}
    }

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    const iconClass = type === 'success' 
      ? 'fa-circle-check' 
      : type === 'error' 
      ? 'fa-circle-exclamation' 
      : 'fa-circle-info';

    toast.innerHTML = `
      <i class="fa-solid ${iconClass} toast-icon"></i>
      <span class="toast-text">${escapeHtml(message)}</span>
    `;

    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.classList.add('toast-show');
    }, 10);

    setTimeout(() => {
      toast.classList.remove('toast-show');
      setTimeout(() => toast.remove(), 300);
    }, 4000);
  }

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  // --- Drag and Drop Handling ---
  ['dragenter', 'dragover'].forEach(evt => {
    dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add('drag-active');
    });
  });

  ['dragleave', 'drop'].forEach(evt => {
    dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('drag-active');
    });
  });

  dropZone.addEventListener('click', () => {
    fileInput.click();
  });

  dropZone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      fileInput.click();
    }
  });

  dropZone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      handleSelectedFile(files[0]);
    }
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) {
      handleSelectedFile(fileInput.files[0]);
    }
  });

  removeFileBtn.addEventListener('click', () => {
    resetUploadedFile();
  });

  function resetUploadedFile() {
    state.currentFile = null;
    state.uploadId = null;
    state.lastPdfBlob = null;
    state.lastPdfUrl = null;
    fileInput.value = '';
    dropZone.classList.remove('hidden');
    fileInfoCard.classList.add('hidden');
    sheetSelectGroup.classList.add('hidden');
    sheetSelect.innerHTML = '<option value="all">Semua Lembar Kerja (Gabungkan ke 1 PDF)</option>';
    convertBtn.disabled = true;
    convertHintText.textContent = 'Pilih file Excel terlebih dahulu untuk memulai konversi';
  }

  async function handleSelectedFile(file) {
    const validExtensions = ['.xlsx', '.xls', '.xlsm', '.csv'];
    const fileName = file.name.toLowerCase();
    const isValid = validExtensions.some(ext => fileName.endsWith(ext));

    if (!isValid) {
      showToast('Format file tidak didukung. Harap pilih file .xlsx, .xls, .xlsm, atau .csv', 'error');
      return;
    }

    state.currentFile = file;

    // Show temporary file info while inspecting
    fileNameDisplay.textContent = file.name;
    fileSizeDisplay.textContent = formatBytes(file.size);
    fileSheetsCountDisplay.textContent = 'Memeriksa lembar...';
    dropZone.classList.add('hidden');
    fileInfoCard.classList.remove('hidden');

    // 1. If in offline mode, inspect directly client-side
    if (state.isOfflineMode || !navigator.onLine) {
      inspectFileLocally(file);
      return;
    }

    // 2. Otherwise try server first, fall back to offline
    const formData = new FormData();
    formData.append('file', file);

    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 3000);
      const res = await fetch('/api/inspect', {
        method: 'POST',
        body: formData,
        signal: controller.signal
      });
      clearTimeout(timeoutId);
      const data = await res.json();

      if (data.success) {
        state.uploadId = data.upload_id;
        fileSizeDisplay.textContent = data.filesize_formatted || formatBytes(file.size);
        fileSheetsCountDisplay.textContent = `${data.sheets.length} Sheet${data.sheets.length > 1 ? 's' : ''}`;

        sheetSelect.innerHTML = '<option value="all">Semua Lembar Kerja (Gabungkan ke 1 PDF)</option>';
        data.sheets.forEach(shName => {
          const opt = document.createElement('option');
          opt.value = shName;
          opt.textContent = `Sheet: ${shName}`;
          sheetSelect.appendChild(opt);
        });

        sheetSelectGroup.classList.remove('hidden');
        convertBtn.disabled = false;
        convertHintText.textContent = 'Siap dikonversi dengan pengaturan margin Anda';

        // AUTO-LANDSCAPE: detect wide tables and auto-switch orientation
        if (data.details && data.details.length > 0) {
          const maxCol = Math.max(...data.details.map(d => d.max_column || 0));
          if (maxCol > 7 && state.orientation !== 'landscape') {
            setOrientation('landscape');
            showToast(`Tabel memiliki ${maxCol} kolom. Orientasi otomatis diatur ke Landscape untuk hasil cetak optimal.`, 'info');
          }
        }

        showToast(`File "${file.name}" berhasil dimuat!`, 'success');
      } else {
        inspectFileLocally(file);
      }
    } catch (err) {
      console.warn("Gagal menghubungi server, beralih ke inspeksi lokal:", err);
      inspectFileLocally(file);
    }
  }

  async function inspectFileLocally(file) {
    if (window.OfflineConverter) {
      try {
        const data = await window.OfflineConverter.inspectFile(file);
        if (data.workbook) {
          state.cachedWorkbook = data.workbook;
        }
        fileSizeDisplay.textContent = data.filesize_formatted;
        fileSheetsCountDisplay.textContent = `${data.sheets.length} Sheet${data.sheets.length > 1 ? 's' : ''}`;

        sheetSelect.innerHTML = '<option value="all">Semua Lembar Kerja (Gabungkan ke 1 PDF)</option>';
        data.sheets.forEach(shName => {
          const opt = document.createElement('option');
          opt.value = shName;
          opt.textContent = `Sheet: ${shName}`;
          sheetSelect.appendChild(opt);
        });

        sheetSelectGroup.classList.remove('hidden');
        convertBtn.disabled = false;
        convertHintText.textContent = 'Siap dikonversi secara instan!';

        // AUTO-LANDSCAPE: detect wide tables and auto-switch orientation
        if (data.details && data.details.length > 0) {
          const maxCol = Math.max(...data.details.map(d => d.max_column || 0));
          if (maxCol > 7 && state.orientation !== 'landscape') {
            setOrientation('landscape');
            showToast(`Tabel memiliki ${maxCol} kolom. Orientasi otomatis diatur ke Landscape untuk hasil cetak optimal.`, 'info');
          }
        }

        showToast(`File "${file.name}" siap dikonversi!`, 'success');
        return;
      } catch (err) {
        console.error("Gagal membaca file lokal:", err);
      }
    }

    fileSheetsCountDisplay.textContent = '1 Sheet';
    convertBtn.disabled = false;
    convertHintText.textContent = 'Siap dikonversi dengan pengaturan margin Anda';
  }

  function formatBytes(bytes) {
    if (!bytes || bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  // --- Margin Unit Switching (cm <-> inch) ---
  unitCmBtn.addEventListener('click', () => switchUnit('cm'));
  unitInchBtn.addEventListener('click', () => switchUnit('inch'));

  function switchUnit(targetUnit) {
    if (state.unit === targetUnit) return;

    state.unit = targetUnit;

    unitCmBtn.classList.toggle('active', targetUnit === 'cm');
    unitInchBtn.classList.toggle('active', targetUnit === 'inch');

    // Update unit labels on inputs
    unitLabels.forEach(lbl => lbl.textContent = targetUnit);

    // Convert values
    const factor = targetUnit === 'inch' ? (1 / 2.54) : 2.54;
    [marginTop, marginBottom, marginLeft, marginRight, marginHeader, marginFooter].forEach(input => {
      const val = parseFloat(input.value) || 0;
      input.value = (val * factor).toFixed(2);
    });

    updatePresetLabels();
    updateVisualMockup();
    showToast(`Satuan diubah ke ${targetUnit.toUpperCase()}`, 'info');
  }

  function updatePresetLabels() {
    const pSet = PRESETS[state.unit];
    presetCards.forEach(card => {
      const pKey = card.dataset.preset;
      const detailEl = card.querySelector('.preset-detail');
      if (pSet[pKey] && detailEl) {
        detailEl.textContent = pSet[pKey].desc;
      }
    });
  }

  // --- Preset Selection ---
  presetCards.forEach(card => {
    card.addEventListener('click', () => {
      const pKey = card.dataset.preset;
      applyPreset(pKey);
    });
  });

  function applyPreset(pKey) {
    state.preset = pKey;
    presetCards.forEach(c => c.classList.toggle('active', c.dataset.preset === pKey));

    if (pKey === 'custom') return;

    const currentPresetValues = PRESETS[state.unit][pKey];
    if (currentPresetValues) {
      marginTop.value = currentPresetValues.top.toFixed(2);
      marginBottom.value = currentPresetValues.bottom.toFixed(2);
      marginLeft.value = currentPresetValues.left.toFixed(2);
      marginRight.value = currentPresetValues.right.toFixed(2);
      marginHeader.value = currentPresetValues.header.toFixed(2);
      marginFooter.value = currentPresetValues.footer.toFixed(2);
      updateVisualMockup();
    }
  }

  // Detect custom typing in margin inputs
  [marginTop, marginBottom, marginLeft, marginRight, marginHeader, marginFooter].forEach(input => {
    input.addEventListener('input', () => {
      setCustomPresetActive();
      updateVisualMockup();
    });
  });

  function setCustomPresetActive() {
    state.preset = 'custom';
    presetCards.forEach(c => c.classList.toggle('active', c.dataset.preset === 'custom'));
  }

  // --- Orientation Toggle ---
  const orientationRadios = document.querySelectorAll('input[name="orientation"]');

  function setOrientation(ori) {
    if (!ori) ori = 'portrait';
    ori = String(ori).trim().toLowerCase();
    state.orientation = ori;

    orientationCards.forEach(c => {
      const radio = c.querySelector('input[type="radio"]');
      const val = (c.dataset.orientation || (radio ? radio.value : '')).toLowerCase();
      const isMatch = (val === ori);
      c.classList.toggle('active', isMatch);
      if (radio) radio.checked = isMatch;
    });

    updateVisualMockup();
  }

  orientationCards.forEach(card => {
    card.addEventListener('click', (e) => {
      const radio = card.querySelector('input[type="radio"]');
      const ori = card.dataset.orientation || (radio ? radio.value : null);
      if (ori) {
        setOrientation(ori);
      }
    });
  });

  orientationRadios.forEach(radio => {
    radio.addEventListener('change', () => {
      if (radio.checked) {
        setOrientation(radio.value);
      }
    });
  });

  // Paper & Fit dropdown change listeners
  paperSizeSelect.addEventListener('change', () => {
    state.paperSize = paperSizeSelect.value;
    updateVisualMockup();
  });

  fitModeSelect.addEventListener('change', () => {
    state.fitMode = fitModeSelect.value;
  });

  // --- Visual Margin Mockup Updater ---
  function updateVisualMockup() {
    const isLandscape = state.orientation === 'landscape';
    paperMockup.classList.toggle('portrait-mode', !isLandscape);
    paperMockup.classList.toggle('landscape-mode', isLandscape);
    paperMockup.classList.toggle('orientation-landscape', isLandscape);

    const topVal = parseFloat(marginTop.value) || 0;
    const botVal = parseFloat(marginBottom.value) || 0;
    const leftVal = parseFloat(marginLeft.value) || 0;
    const rightVal = parseFloat(marginRight.value) || 0;

    const cmMultiplier = isLandscape ? 6.5 : 9;
    const unitScale = state.unit === 'inch' ? 2.54 : 1.0;

    const padTop = Math.min(Math.max(topVal * unitScale * cmMultiplier, 10), 80);
    const padBot = Math.min(Math.max(botVal * unitScale * cmMultiplier, 10), 80);
    const padLeft = Math.min(Math.max(leftVal * unitScale * cmMultiplier, 10), 80);
    const padRight = Math.min(Math.max(rightVal * unitScale * cmMultiplier, 10), 80);

    paperPrintableArea.style.margin = `${padTop}px ${padRight}px ${padBot}px ${padLeft}px`;

    visValTop.textContent = `${topVal} ${state.unit}`;
    visValBottom.textContent = `${botVal} ${state.unit}`;
    visValLeft.textContent = `${leftVal} ${state.unit}`;
    visValRight.textContent = `${rightVal} ${state.unit}`;
  }

  // --- Tabs Navigation ---
  tabVisualizerBtn.addEventListener('click', () => switchTab('visualizer'));
  tabPdfResultBtn.addEventListener('click', () => switchTab('pdfResult'));

  function switchTab(tabKey) {
    const isVis = tabKey === 'visualizer';
    tabVisualizerBtn.classList.toggle('active', isVis);
    tabVisualizerBtn.setAttribute('aria-selected', isVis);
    tabPdfResultBtn.classList.toggle('active', !isVis);
    tabPdfResultBtn.setAttribute('aria-selected', !isVis);

    tabVisualizerContent.classList.toggle('hidden', !isVis);
    tabPdfResultContent.classList.toggle('hidden', isVis);
  }

  // --- Convert Execution ---
  convertBtn.addEventListener('click', async () => {
    if (!state.currentFile && !state.uploadId) {
      showToast('Pilih file Excel terlebih dahulu!', 'error');
      return;
    }

    const convOptions = {
      filename: state.currentFile ? state.currentFile.name : 'dokumen.xlsx',
      unit: state.unit,
      top_margin: parseFloat(marginTop.value) || 1.91,
      bottom_margin: parseFloat(marginBottom.value) || 1.91,
      left_margin: parseFloat(marginLeft.value) || 1.78,
      right_margin: parseFloat(marginRight.value) || 1.78,
      header_margin: parseFloat(marginHeader.value) || 0.76,
      footer_margin: parseFloat(marginFooter.value) || 0.76,
      orientation: state.orientation,
      paper_size: paperSizeSelect.value,
      fit_mode: fitModeSelect.value,
      center_horizontally: centerHorizontally.checked,
      center_vertically: centerVertically.checked,
      sheet_name: sheetSelect.value,
      show_gridlines: showGridlines.checked,
      show_headings: showHeadings.checked
    };

    // Show loading state
    loadingOverlay.classList.remove('hidden');
    convertBtn.disabled = true;
    convertSpinner.classList.remove('hidden');
    convertIcon.classList.add('hidden');
    convertBtnText.textContent = 'Mengonversi...';

    const loadingStatusText = document.getElementById('loadingStatusText');
    const progressBarFill = document.querySelector('.progress-bar-fill');

    function updateProgress(msg, percent) {
      if (loadingStatusText) loadingStatusText.textContent = msg;
      if (progressBarFill) progressBarFill.style.width = percent;
    }

    function cleanupLoading() {
      loadingOverlay.classList.add('hidden');
      convertBtn.disabled = false;
      convertSpinner.classList.add('hidden');
      convertIcon.classList.remove('hidden');
      convertBtnText.textContent = 'Konversi ke PDF';
      if (progressBarFill) progressBarFill.style.width = '45%';
    }

    updateProgress('Menyiapkan parameter konversi...', '15%');

    // 1. If in Cloud (Vercel), Offline Mode, or file not uploaded to server: perform INSTANT client-side conversion
    const isCloudOrOffline = state.isOfflineMode || window.location.hostname.includes('vercel.app') || !state.uploadId;

    if (isCloudOrOffline) {
      try {
        await executeOfflineConversion(convOptions, updateProgress);
        return;
      } catch (err) {
        console.warn("Client-side conversion error, attempting server fallback:", err);
      } finally {
        cleanupLoading();
      }
    }

    // 2. Otherwise try Native Windows Server COM (with strict 8s timeout to avoid hanging)
    updateProgress('Mengirim file ke Microsoft Excel Native COM...', '35%');
    const formData = new FormData();
    if (state.uploadId) {
      formData.append('upload_id', state.uploadId);
    } else {
      formData.append('file', state.currentFile);
    }

    Object.keys(convOptions).forEach(k => {
      formData.append(k, convOptions[k]);
    });

    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 8000); // 8s max

      const res = await fetch('/api/convert', {
        method: 'POST',
        body: formData,
        signal: controller.signal
      });
      clearTimeout(timeoutId);

      const contentType = res.headers.get('content-type') || '';
      if (!res.ok || !contentType.includes('application/json')) {
        throw new Error("Server tidak merespons dengan JSON valid");
      }

      updateProgress('Menerima hasil cetak PDF...', '85%');
      const result = await res.json();
      state.convertedTask = result;
      state.currentPage = 1;
      state.zoomLevel = 100;

      updateProgress('Selesai!', '100%');
      displayConversionResult(result);
      showToast('Berhasil mengonversi Excel ke PDF via Microsoft Excel COM!', 'success');

      resultBadgeCounter.classList.remove('hidden');
      switchTab('pdfResult');
    } catch (err) {
      console.warn("Server timeout atau gagal, beralih instan ke engine client-side...", err);
      try {
        await executeOfflineConversion(convOptions, updateProgress);
      } catch (offErr) {
        showToast(`Error: ${offErr.message}`, 'error');
      }
    } finally {
      cleanupLoading();
    }
  });

  async function executeOfflineConversion(options, updateProgress) {
    if (!window.OfflineConverter) {
      throw new Error("Engine konversi offline belum siap.");
    }

    if (updateProgress) updateProgress('Membaca data lembar kerja Excel...', '35%');

    // Use cached workbook for instant processing
    if (state.cachedWorkbook) {
      options.cachedWorkbook = state.cachedWorkbook;
    }

    if (updateProgress) updateProgress('Menata margin & merender PDF...', '65%');
    const result = await window.OfflineConverter.convertToPdf(state.currentFile, options);
    
    if (updateProgress) updateProgress('Menyiapkan pratinjau halaman 1...', '90%');
    // Generate page 1 preview immediately (fast ~50ms)
    let previewData = { total_pages: result.total_pages, previews: [] };
    try {
      previewData = await window.OfflineConverter.generatePreviews(result.arrayBuffer, 1);
    } catch (e) {
      console.warn("Preview warning:", e);
    }

    if (updateProgress) updateProgress('Selesai!', '100%');

    state.lastPdfBlob = result.blob;
    state.lastPdfUrl = URL.createObjectURL(result.blob);
    state.pdfArrayBuffer = result.arrayBuffer;

    const taskResult = {
      task_id: 'fast-' + Date.now(),
      filename: result.filename,
      pdf_size_formatted: result.pdf_size_formatted,
      duration_seconds: result.duration_seconds,
      engine: result.engine,
      total_pages: result.total_pages,
      previews: previewData.previews,
      download_url: state.lastPdfUrl,
      view_url: state.lastPdfUrl,
      is_local_blob: true
    };

    state.convertedTask = taskResult;
    state.currentPage = 1;
    state.zoomLevel = 100;

    displayConversionResult(taskResult);
    if (result.auto_landscape) {
      showToast(`Konversi selesai (${result.duration_seconds}s)! Orientasi otomatis disesuaikan ke Landscape agar tabel rapi.`, 'success');
    } else {
      showToast('Konversi selesai dalam ' + result.duration_seconds + ' detik!', 'success');
    }

    resultBadgeCounter.classList.remove('hidden');
    switchTab('pdfResult');

    // Background pre-cache pages 2..5 if multi-page
    if (result.total_pages > 1) {
      setTimeout(async () => {
        try {
          const bgPreviews = await window.OfflineConverter.generatePreviews(result.arrayBuffer, Math.min(result.total_pages, 5));
          if (bgPreviews.previews && bgPreviews.previews.length > 0) {
            state.convertedTask.previews = bgPreviews.previews;
          }
        } catch (e) {
          console.warn(e);
        }
      }, 150);
    }
  }

  // --- Display Conversion Result ---
  function displayConversionResult(data) {
    emptyResultPlaceholder.classList.add('hidden');
    conversionSuccessArea.classList.remove('hidden');

    resPdfName.textContent = data.filename;
    resPdfSize.textContent = data.pdf_size_formatted;
    resPdfPages.textContent = `${data.total_pages} Halaman`;
    resPdfDuration.textContent = `${data.duration_seconds}s`;

    downloadPdfBtn.href = data.download_url;
    downloadPdfBtn.setAttribute('download', data.filename);
    openTabPdfBtn.href = data.view_url;

    // Direct download click handling
    downloadPdfBtn.onclick = (e) => {
      if (state.lastPdfBlob && window.OfflineConverter) {
        e.preventDefault();
        window.OfflineConverter.deliverPdf(state.lastPdfBlob, data.filename);
      }
    };

    updateViewerPage();
  }

  // --- Share PDF Button ---
  if (sharePdfBtn) {
    sharePdfBtn.addEventListener('click', async () => {
      if (!state.convertedTask) {
        showToast('Belum ada dokumen PDF untuk dibagikan.', 'error');
        return;
      }

      try {
        let blob = state.lastPdfBlob;
        if (!blob && state.convertedTask.download_url) {
          showToast('Menyiapkan file untuk dibagikan...', 'info');
          const res = await fetch(state.convertedTask.download_url);
          blob = await res.blob();
          state.lastPdfBlob = blob;
        }

        if (blob && window.OfflineConverter) {
          window.OfflineConverter.sharePdf(blob, state.convertedTask.filename);
        } else {
          showToast('Gunakan tombol Unduh untuk menyimpan dokumen.', 'info');
        }
      } catch (err) {
        console.warn('Share error:', err);
        showToast('Gagal membagikan: ' + err.message, 'error');
      }
    });
  }

  function updateViewerPage() {
    if (!state.convertedTask) return;

    const totalPages = state.convertedTask.total_pages;
    pageIndicatorText.textContent = `Halaman ${state.currentPage} dari ${totalPages}`;
    prevPageBtn.disabled = state.currentPage <= 1;
    nextPageBtn.disabled = state.currentPage >= totalPages;

    // Load preview image: Check local previews array first, fallback to on-demand render or API
    if (state.convertedTask.previews && state.convertedTask.previews[state.currentPage - 1]) {
      const p = state.convertedTask.previews[state.currentPage - 1];
      pdfPageImage.src = p.url;
    } else if (state.convertedTask.is_local_blob && state.pdfArrayBuffer && window.OfflineConverter) {
      window.OfflineConverter.renderSinglePage(state.pdfArrayBuffer, state.currentPage).then(p => {
        if (p) {
          state.convertedTask.previews = state.convertedTask.previews || [];
          state.convertedTask.previews[state.currentPage - 1] = p;
          pdfPageImage.src = p.url;
        }
      });
    } else if (!state.convertedTask.is_local_blob) {
      pdfPageImage.src = `/api/preview/${state.convertedTask.task_id}/${state.currentPage}`;
    }

    applyZoom();
  }

  // --- Page Navigation ---
  prevPageBtn.addEventListener('click', () => {
    if (state.currentPage > 1) {
      state.currentPage--;
      updateViewerPage();
    }
  });

  nextPageBtn.addEventListener('click', () => {
    if (state.convertedTask && state.currentPage < state.convertedTask.total_pages) {
      state.currentPage++;
      updateViewerPage();
    }
  });

  // --- Zoom Controls ---
  zoomInBtn.addEventListener('click', () => {
    if (state.zoomLevel < 250) {
      state.zoomLevel += 25;
      applyZoom();
    }
  });

  zoomOutBtn.addEventListener('click', () => {
    if (state.zoomLevel > 50) {
      state.zoomLevel -= 25;
      applyZoom();
    }
  });

  zoomResetBtn.addEventListener('click', () => {
    state.zoomLevel = 100;
    applyZoom();
  });

  function applyZoom() {
    zoomLevelText.textContent = `${state.zoomLevel}%`;
    previewImageWrapper.style.transform = `scale(${state.zoomLevel / 100})`;
  }

  // --- About / Tentang Sistem Modal ---
  function initAboutModal() {
    const aboutBtn = document.getElementById('aboutBtn');
    const headerDevBadge = document.getElementById('headerDevBadge');
    const aboutModalOverlay = document.getElementById('aboutModalOverlay');
    const aboutModalBackdrop = document.getElementById('aboutModalBackdrop');
    const closeAboutModalBtn = document.getElementById('closeAboutModalBtn');
    const closeAboutModalBtn2 = document.getElementById('closeAboutModalBtn2');

    if (!aboutModalOverlay) return;

    function openAboutModal() {
      aboutModalOverlay.classList.remove('hidden');
      document.body.style.overflow = 'hidden';
    }

    function closeAboutModal() {
      aboutModalOverlay.classList.add('hidden');
      document.body.style.overflow = '';
    }

    if (aboutBtn) {
      aboutBtn.addEventListener('click', openAboutModal);
    }

    if (headerDevBadge) {
      headerDevBadge.addEventListener('click', openAboutModal);
    }

    if (closeAboutModalBtn) {
      closeAboutModalBtn.addEventListener('click', closeAboutModal);
    }

    if (closeAboutModalBtn2) {
      closeAboutModalBtn2.addEventListener('click', closeAboutModal);
    }

    if (aboutModalBackdrop) {
      aboutModalBackdrop.addEventListener('click', closeAboutModal);
    }

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && aboutModalOverlay && !aboutModalOverlay.classList.contains('hidden')) {
        closeAboutModal();
      }
    });
  }

  // --- Reset All Button ---
  resetAllBtn.addEventListener('click', () => {
    resetUploadedFile();
    switchTab('visualizer');
    showToast('Silakan pilih file Excel baru.', 'info');
  });
});
