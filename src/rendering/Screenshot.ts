/**
 * Screenshot — renders HTML to PNG using Puppeteer (optional peer dependency).
 */

import { Worksheet } from '../core/Worksheet';
import { worksheetToHtml, type PngRenderOptions } from './SheetRenderer';

export async function worksheetToPng(ws: Worksheet, options: PngRenderOptions = {}): Promise<Buffer> {
  const html = worksheetToHtml(ws, { ...options, fullPage: true });
  const viewportWidth = options.viewportWidth || 1200;

  let puppeteer: any;
  try {
    puppeteer = await import('puppeteer');
  } catch {
    throw new Error(
      'puppeteer is required for worksheetToPng(). Install it: npm install puppeteer'
    );
  }

  const browser = await puppeteer.default.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  try {
    const page = await browser.newPage();
    await page.setViewport({ width: viewportWidth, height: 800 });
    await page.setContent(html, { waitUntil: 'networkidle0' });

    // Auto-size to content
    const bodyHandle = await page.$('body');
    const boundingBox = await bodyHandle?.boundingBox();
    if (boundingBox) {
      await page.setViewport({
        width: Math.max(viewportWidth, Math.ceil(boundingBox.width) + 20),
        height: Math.ceil(boundingBox.height) + 20,
      });
    }

    const buffer = await page.screenshot({
      type: 'png',
      fullPage: true,
    });

    return Buffer.from(buffer);
  } finally {
    await browser.close();
  }
}
