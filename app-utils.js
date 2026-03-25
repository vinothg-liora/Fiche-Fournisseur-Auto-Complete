/* ===== Utils & Toast ===== */
function showToast(message, type = 'success') {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 4000);
}

function showLoading(text = 'Analyse du document en cours...') {
  document.getElementById('loadingText').textContent = text;
  document.getElementById('loadingOverlay').classList.add('active');
}

function hideLoading() {
  document.getElementById('loadingOverlay').classList.remove('active');
}

function toggleVisibility(inputId) {
  const input = document.getElementById(inputId);
  input.type = input.type === 'password' ? 'text' : 'password';
}

function formatDate(dateStr) {
  if (!dateStr) return '—';
  const d = new Date(dateStr);
  return d.toLocaleDateString('fr-FR', { day: '2-digit', month: 'long', year: 'numeric' });
}

function daysBetween(d1, d2) {
  return Math.ceil((d2 - d1) / (1000 * 60 * 60 * 24));
}

/* ===== State ===== */
const APP = {
  companyData: null,
  uploadedFile: null,
  uploadedFileName: '',
  fileType: '', // 'xlsx', 'pdf-form', 'pdf-scan'
  fileContent: null, // raw ArrayBuffer
  workbook: null, // for Excel
  analysisResult: null,
  fieldValues: {},
  networkPath: localStorage.getItem('networkPath') || 'company_data.json',
};
