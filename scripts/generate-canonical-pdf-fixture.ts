#!/usr/bin/env bun
import fontkit from "@pdf-lib/fontkit";
import { PDFDocument, rgb } from "pdf-lib";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = dirname(__dirname);
const FIXTURE_DIR = join(PROJECT_ROOT, "tests", "fixtures");
const FIXTURE_PATH = join(FIXTURE_DIR, "zh-TW-editable.pdf");

// 689 x 960 pt is the canonical Desktop export page size.
const PAGE_WIDTH_PT = 689;
const PAGE_HEIGHT_PT = 960;
const PAGE_COUNT = 13;

// Arial Unicode is a system CJK font available on macOS; we embed a tiny subset.
const FONT_PATH = "/System/Library/Fonts/Supplemental/Arial Unicode.ttf";

async function main(): Promise<void> {
  await mkdir(FIXTURE_DIR, { recursive: true });

  const pdfDoc = await PDFDocument.create();
  pdfDoc.registerFontkit(fontkit);
  const fontBytes = await readFile(FONT_PATH);
  const font = await pdfDoc.embedFont(fontBytes, { subset: true });

  // Page 1: title + body text matching extractPdfTextRuns expectations.
  const page1 = pdfDoc.addPage([PAGE_WIDTH_PT, PAGE_HEIGHT_PT]);
  const titleSize = 32;
  const bodySize = 16;
  const titleText = "真的傾聽";
  const bodyText = "不是：只聽";

  const titleWidth = font.widthOfTextAtSize(titleText, titleSize);
  const bodyWidth = font.widthOfTextAtSize(bodyText, bodySize);

  page1.drawText(titleText, {
    x: (PAGE_WIDTH_PT - titleWidth) / 2,
    y: PAGE_HEIGHT_PT - 120,
    size: titleSize,
    font,
    color: rgb(0x19 / 255, 0x19 / 255, 0x19 / 255),
  });

  page1.drawText(bodyText, {
    x: (PAGE_WIDTH_PT - bodyWidth) / 2,
    y: PAGE_HEIGHT_PT - 180,
    size: bodySize,
    font,
    color: rgb(0x68 / 255, 0x68 / 255, 0x68 / 255),
  });

  // Pages 2-13: blank placeholder pages.
  for (let i = 2; i <= PAGE_COUNT; i++) {
    pdfDoc.addPage([PAGE_WIDTH_PT, PAGE_HEIGHT_PT]);
  }

  const pdfBytes = await pdfDoc.save();
  await writeFile(FIXTURE_PATH, pdfBytes);
  console.log(`Wrote ${FIXTURE_PATH} (${pdfBytes.length} bytes)`);
  console.log(`Pages: ${pdfDoc.getPageCount()}, size: ${PAGE_WIDTH_PT}x${PAGE_HEIGHT_PT} pt`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
