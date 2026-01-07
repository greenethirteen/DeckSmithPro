import sharp from 'sharp';
import fs from 'fs-extra';

/**
 * Resize/crop image to cover target size.
 */
export async function coverTo(fileIn, fileOut, width, height) {
  await sharp(fileIn)
    .resize(width, height, { fit: 'cover', position: 'entropy' })
    .png({ quality: 90 })
    .toFile(fileOut);
  return fileOut;
}

export function clamp(n, a, b) { return Math.max(a, Math.min(b, n)); }

/**
 * Returns luminance (0..1) for an image file by averaging a tiny resize.
 */
export async function getLuminance(filePath) {
  const buf = await sharp(filePath).resize(1,1, { fit:'cover' }).raw().toBuffer();
  const r = buf[0] ?? 0, g = buf[1] ?? 0, b = buf[2] ?? 0;
  // relative luminance (sRGB)
  const L = (0.2126*r + 0.7152*g + 0.0722*b) / 255;
  return L;
}

/**
 * Ensures directory exists, empties older temp files (best-effort).
 */
export async function ensureCleanTmp(tmpDir, maxAgeMinutes = 60) {
  await fs.ensureDir(tmpDir);
  const files = await fs.readdir(tmpDir);
  const now = Date.now();
  await Promise.all(files.map(async f => {
    const p = `${tmpDir}/${f}`;
    try {
      const st = await fs.stat(p);
      if (st.isFile() && (now - st.mtimeMs) > maxAgeMinutes*60*1000) {
        await fs.remove(p);
      }
    } catch {}
  }));
}
