/* ═══════════════════════════════════════════
   DOM ELEMENTS
   ═══════════════════════════════════════════ */
const openGuide = document.getElementById('openGuide');
const closeGuide = document.getElementById('closeGuide');
const modalOverlay = document.getElementById('modalOverlay');

// Modal Logic
if (openGuide) {
    openGuide.addEventListener('click', () => modalOverlay.classList.add('active'));
}
if (closeGuide) {
    closeGuide.addEventListener('click', () => modalOverlay.classList.remove('active'));
}
if (modalOverlay) {
    modalOverlay.addEventListener('click', (e) => {
        if (e.target === modalOverlay) modalOverlay.classList.remove('active');
    });
}

/* ═══════════════════════════════════════════
   STATE
   ═══════════════════════════════════════════ */
let selectedFile = null;

/* ═══════════════════════════════════════════
   DOM ELEMENTS (UPLOAD)
   ═══════════════════════════════════════════ */
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const imagePreview = document.getElementById('imagePreview');
const previewImg = document.getElementById('previewImg');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const removeBtn = document.getElementById('removeBtn');
const convertBtn = document.getElementById('convertBtn');
const errorMsg = document.getElementById('errorMsg');
const errorText = document.getElementById('errorText');
const resultSection = document.getElementById('resultSection');
const previewTable = document.getElementById('previewTable');
const downloadBtn = document.getElementById('downloadBtn');
const statsContainer = document.getElementById('statsContainer');

/* ═══════════════════════════════════════════
   FILE UPLOAD HANDLERS
   ═══════════════════════════════════════════ */

// Click to upload
if (dropZone) {
    dropZone.addEventListener('click', () => fileInput.click());
}

// File selected
if (fileInput) {
    fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
            handleFile(e.target.files[0]);
        }
    });
}

// Drag & Drop
if (dropZone) {
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            handleFile(e.dataTransfer.files[0]);
        }
    });
}

// Remove file
if (removeBtn) {
    removeBtn.addEventListener('click', () => {
        resetUpload();
    });
}

function handleFile(file) {
    // Validate
    const validTypes = ['image/jpeg', 'image/png', 'image/bmp'];
    if (!validTypes.includes(file.type)) {
        showError('Format file tidak didukung. Gunakan JPG atau PNG.');
        return;
    }

    if (file.size > 16 * 1024 * 1024) {
        showError('Ukuran file terlalu besar (max 16MB).');
        return;
    }

    selectedFile = file;
    hideError();

    // Show preview
    const reader = new FileReader();
    reader.onload = (e) => {
        previewImg.src = e.target.result;
        fileName.textContent = file.name;
        fileSize.textContent = formatSize(file.size);

        dropZone.style.display = 'none';
        imagePreview.classList.add('active');
        convertBtn.disabled = false;
    };
    reader.readAsDataURL(file);

    // Hide previous results
    resultSection.classList.remove('active');
}

function resetUpload() {
    selectedFile = null;
    fileInput.value = '';
    previewImg.src = '';

    dropZone.style.display = '';
    imagePreview.classList.remove('active');
    convertBtn.disabled = true;
    resultSection.classList.remove('active');
    hideError();
}

function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

/* ═══════════════════════════════════════════
   CONVERSION
   ═══════════════════════════════════════════ */
if (convertBtn) {
    convertBtn.addEventListener('click', async () => {
        if (!selectedFile) return;

        // Set loading state
        convertBtn.classList.add('loading');
        convertBtn.disabled = true;
        hideError();
        resultSection.classList.remove('active');

        try {
            const formData = new FormData();
            formData.append('file', selectedFile);

            const response = await fetch('/upload', {
                method: 'POST',
                body: formData,
            });

            const data = await response.json();

            if (!data.success) {
                showError(data.error || 'Terjadi kesalahan saat memproses gambar.');
                return;
            }

            // Render results
            renderPreview(data.preview);
            renderStats(data.stats);
            downloadBtn.href = data.downloadUrl;

            resultSection.classList.add('active');

            // Scroll to results
            setTimeout(() => {
                resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }, 100);

            // Show success toast
            showToast('✅ Konversi berhasil!');

        } catch (err) {
            console.error(err);
            showError('Gagal terhubung ke server. Pastikan server berjalan.');
        } finally {
            convertBtn.classList.remove('loading');
            convertBtn.disabled = false;
        }
    });
}

/* ═══════════════════════════════════════════
   RENDER PREVIEW TABLE
   ═══════════════════════════════════════════ */
function renderPreview(preview) {
    let html = '';

    for (const row of preview.rows) {
        const rowClass = row.isHeader ? ' class="header-row"' : '';
        html += `<tr${rowClass}>`;

        for (const cell of row.cells) {
            const bg = cell.color || '#FFFFFF';
            const text = cell.text || '';

            // Text color based on background brightness
            const textColor = getTextColor(bg);

            html += `<td style="background-color: ${bg}; color: ${textColor};">${escapeHtml(text)}</td>`;
        }

        html += '</tr>';
    }

    previewTable.innerHTML = html;
}

function renderStats(stats) {
    statsContainer.innerHTML = `
        <span class="stat-badge">📐 ${stats.rows} × ${stats.cols}</span>
        <span class="stat-badge">⚡ ${stats.time}s</span>
    `;
}

function getTextColor(hex) {
    // Parse hex color
    hex = hex.replace('#', '');
    if (hex.length === 3) {
        hex = hex[0]+hex[0]+hex[1]+hex[1]+hex[2]+hex[2];
    }
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);
    const brightness = (r * 299 + g * 587 + b * 114) / 1000;
    return brightness > 128 ? '#000000' : '#FFFFFF';
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

/* ═══════════════════════════════════════════
   ERROR / TOAST
   ═══════════════════════════════════════════ */
function showError(msg) {
    if (errorText && errorMsg) {
        errorText.textContent = msg;
        errorMsg.classList.add('active');
    }
}

function hideError() {
    if (errorMsg) {
        errorMsg.classList.remove('active');
    }
}

function showToast(msg) {
    const toast = document.createElement('div');
    toast.className = 'success-toast';
    toast.textContent = msg;
    document.body.appendChild(toast);

    setTimeout(() => {
        toast.remove();
    }, 3200);
}
