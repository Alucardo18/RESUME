import puppeteer from 'puppeteer';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const htmlPath = path.join(__dirname, 'Emmanuel_Gonzalez_Antigravity_CV.html');
const pdfPath = path.join(__dirname, 'Emmanuel_Gonzalez_Antigravity_CV.pdf');

(async () => {
  console.log('🚀 Launching browser...');
  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const page = await browser.newPage();

  console.log('📄 Loading CV HTML...');
  await page.goto(`file://${htmlPath}`, { waitUntil: 'networkidle0', timeout: 30000 });

  // Wait for fonts to load
  await page.evaluateHandle('document.fonts.ready');
  console.log('✅ Fonts loaded');

  // Wait an extra moment for any animations to settle
  await new Promise(r => setTimeout(r, 1500));

  console.log('🖨️  Generating PDF...');
  await page.pdf({
    path: pdfPath,
    format: 'A4',
    printBackground: true,
    preferCSSPageSize: false,
    margin: { top: '0', right: '0', bottom: '0', left: '0' },
    scale: 0.73,
    pageRanges: '1'
  });

  console.log(`✅ PDF saved to: ${pdfPath}`);
  await browser.close();
})();
