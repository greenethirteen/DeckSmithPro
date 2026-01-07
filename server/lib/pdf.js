import fs from 'fs';
import pdfParse from 'pdf-parse';

/**
 * Extract text from a PDF file path.
 * Strips repetitive whitespace and preserves paragraphs.
 */
export async function extractPdfText(pdfPath) {
  const dataBuffer = fs.readFileSync(pdfPath);
  const data = await pdfParse(dataBuffer);
  const text = (data.text || '')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  return text;
}
