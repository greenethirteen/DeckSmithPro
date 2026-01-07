import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import mammoth from 'mammoth';
import WordExtractor from 'word-extractor';
import { extractPdfText } from './pdf.js';

/**
 * Extract brief text from supported file types.
 * Supported: .pdf, .docx, .doc, .pptx
 */
export async function extractBriefText(filePath, originalName = '') {
  const ext = path.extname(originalName || filePath).toLowerCase();
  if (ext === '.pdf') return extractPdfText(filePath);
  if (ext === '.docx') return extractDocxText(filePath);
  if (ext === '.doc') return extractDocText(filePath);
  if (ext === '.pptx') return extractPptxText(filePath);
  throw new Error(`Unsupported brief file type: ${ext || '(unknown)'}. Please upload PDF, DOC/DOCX, or PPTX.`);
}

export async function extractDocxText(docxPath) {
  const buf = fs.readFileSync(docxPath);
  const { value } = await mammoth.extractRawText({ buffer: buf });
  return normalizeText(value || '');
}

export async function extractDocText(docPath) {
  const extractor = new WordExtractor();
  const doc = await extractor.extract(docPath);
  const text = (doc && typeof doc.getBody === 'function') ? doc.getBody() : '';
  return normalizeText(text || '');
}

/**
 * Extract text from PPTX by reading slide XML and pulling <a:t> runs.
 * Keeps slide boundaries to preserve some structure.
 */
export async function extractPptxText(pptxPath) {
  const buf = fs.readFileSync(pptxPath);
  const zip = await JSZip.loadAsync(buf);

  const slideFiles = Object.keys(zip.files)
    .filter((p) => /^ppt\/slides\/slide\d+\.xml$/i.test(p))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)\.xml/i)?.[1] || '0', 10);
      const nb = parseInt(b.match(/slide(\d+)\.xml/i)?.[1] || '0', 10);
      return na - nb;
    });

  const notesFiles = Object.keys(zip.files)
    .filter((p) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(p))
    .sort((a, b) => {
      const na = parseInt(a.match(/notesSlide(\d+)\.xml/i)?.[1] || '0', 10);
      const nb = parseInt(b.match(/notesSlide(\d+)\.xml/i)?.[1] || '0', 10);
      return na - nb;
    });

  const extractRuns = (xml) => {
    const runs = [];
    const re = /<a:t[^>]*>([\s\S]*?)<\/a:t>/gi;
    let m;
    while ((m = re.exec(xml)) !== null) {
      const raw = m[1]
        .replace(/<[^>]+>/g, '')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .trim();
      if (raw) runs.push(raw);
    }
    return runs;
  };

  const chunks = [];
  for (let i = 0; i < slideFiles.length; i++) {
    const slideXml = await zip.file(slideFiles[i]).async('string');
    const slideText = extractRuns(slideXml).join('\n');

    let notesText = '';
    if (i < notesFiles.length && zip.file(notesFiles[i])) {
      const notesXml = await zip.file(notesFiles[i]).async('string');
      notesText = extractRuns(notesXml).join('\n');
    }

    const block = [
      `Slide ${i + 1}`,
      slideText,
      notesText ? `Notes:\n${notesText}` : ''
    ].filter(Boolean).join('\n');

    chunks.push(block.trim());
  }

  return normalizeText(chunks.join('\n\n---\n\n'));
}

function normalizeText(text) {
  return (text || '')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}
