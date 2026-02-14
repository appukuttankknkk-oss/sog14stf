
import { DossierState } from '../types';

declare const html2pdf: any;

/**
 * Export specific HTML element to a high-resolution PDF
 */
export const exportToPdf = async (elementId: string, filename: string) => {
  const element = document.getElementById(elementId);
  if (!element) return;

  const opt = {
    margin: [10, 0, 10, 0], // Top, Left, Bottom, Right in mm
    filename: filename,
    image: { type: 'jpeg', quality: 1.0 },
    html2canvas: { 
      scale: 4, // High resolution
      useCORS: true, 
      letterRendering: true,
      dpi: 300,
      logging: false
    },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait', compress: true },
    pagebreak: { 
      mode: ['avoid-all', 'css', 'legacy'],
      before: '.page-break-before',
      avoid: ['tr', 'td', 'img', 'video', 'audio']
    }
  };

  try {
    await html2pdf().from(element).set(opt).save();
  } catch (err) {
    console.error('PDF Generation failed:', err);
  }
};

/**
 * Export specific HTML content to a Word-compatible .doc file
 */
export const exportToWord = (elementId: string, filename: string) => {
  const element = document.getElementById(elementId);
  if (!element) return;

  // Simple Word XML wrapper for HTML content
  const preHtml = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title><style>table { border-collapse: collapse; width: 100%; } td { border: 1px solid black; padding: 4px; font-family: Arial; font-size: 10pt; } tr { page-break-inside: avoid; break-inside: avoid; }</style></head><body>";
  const postHtml = "</body></html>";
  const html = preHtml + element.innerHTML + postHtml;

  const blob = new Blob(['\ufeff', html], {
    type: 'application/msword'
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

/**
 * Save current state as a JSON file
 */
export const saveProject = (state: DossierState) => {
  const dataStr = JSON.stringify(state, null, 2);
  const blob = new Blob([dataStr], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  
  const now = new Date();
  const d = now.getDate().toString().padStart(2, '0');
  const m = (now.getMonth() + 1).toString().padStart(2, '0');
  const y = now.getFullYear();
  const dateStr = `${d}${m}${y}`;

  const link = document.createElement('a');
  link.href = url;
  link.download = `dossier_backup_${dateStr}.json`;
  link.click();
  URL.revokeObjectURL(url);
};

/**
 * Load state from a JSON file
 */
export const loadProject = (file: File, callback: (state: DossierState) => void) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const json = JSON.parse(e.target?.result as string);
      callback(json);
    } catch (err) {
      alert('Invalid project file');
    }
  };
  reader.readAsText(file);
};
