#!/usr/bin/env bun
import { existsSync } from "node:fs";
import { copyFile, mkdir, open, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { dirname, isAbsolute, join, relative } from "node:path";
import { fileURLToPath } from "node:url";
import {
  createCanvas,
  GlobalFonts,
  loadImage,
  type Canvas,
  type Image,
  type SKRSContext2D,
} from "@napi-rs/canvas";
import PptxGenJS from "pptxgenjs";
import { XMLParser } from "fast-xml-parser";
import JSZip from "jszip";

export type PageId = "Overview" | `Ch${1 | 2 | 3 | 4 | 5 | 6}_${"A" | "B"}`;
export type CoverPageId = "Cover" | "Cover_Back" | "Cover_Spine";
export type FontRole = "title" | "body" | "footer";
export type RenderMode = "vector" | "editable_text" | "oracle_patch";
export type TextAlign = "left" | "center";

export type Size = { width: number; height: number };
export type FontSpec = { fontFamily: string; size: number };
export type EmphasisSpec = {
  text?: string;
  bold?: boolean;
  color?: string;
  occurrence?: number;
  start?: number;
  end?: number;
};
export type BoxSpec = {
  id: string;
  ocr_type?: string;
  font_role?: FontRole | string;
  text?: string;
  source_box?: number[] | readonly number[];
  box?: number[] | readonly number[];
  tight_box?: number[] | readonly number[];
  font_size?: number;
  line_spacing?: number;
  color?: string;
  align?: TextAlign | string;
  emphasis?: EmphasisSpec[];
};
export type ImagePatchSpec = {
  id: string;
  source_box: number[] | readonly number[];
  box: number[] | readonly number[];
};
type EnglishTextGeometry = {
  tightBox: number[];
  box: number[];
};

type DiffMask = {
  width: number;
  height: number;
  pixels: Uint8Array;
};

type ComponentBox = {
  box: number[];
  score: number;
  cells: number;
};


function numericOverride(value: unknown): number | undefined {
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? number : undefined;
}

function hexColorOverride(value: unknown): string | undefined {
  const hex = typeof value === "string" ? value.trim().replace(/^#/, "") : "";
  return /^[0-9a-fA-F]{6}$/.test(hex) ? `#${hex.toUpperCase()}` : undefined;
}

function numericBoxOverride(value: unknown, targetSize: Size): number[] | undefined {
  if (!Array.isArray(value) || value.length < 4) return undefined;
  const box = value.slice(0, 4).map((coordinate) => Number(coordinate));
  if (!box.every(Number.isFinite)) return undefined;
  return clampBox(box, targetSize.width, targetSize.height);
}

function emphasisOverride(value: unknown): EmphasisSpec[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const emphasis = value
    .filter((entry): entry is Record<string, unknown> => Boolean(entry) && typeof entry === "object")
    .map((entry) => {
      const emphasisEntry: EmphasisSpec = {
        text: typeof entry.text === "string" ? entry.text : undefined,
        occurrence: numericOverride(entry.occurrence),
        start: numericOverride(entry.start),
        end: numericOverride(entry.end),
      };
      if (entry.bold === true) emphasisEntry.bold = true;
      const color = hexColorOverride(entry.color);
      if (color) emphasisEntry.color = color;
      return emphasisEntry;
    })
    .filter(
      (entry) =>
        (entry.bold === true || entry.color !== undefined) &&
        ((entry.text?.length ?? 0) > 0 || entry.start !== undefined || entry.end !== undefined),
    );
  return emphasis.length > 0 ? emphasis : undefined;
}
export type TextSpec = {
  schema_version?: number;
  image: string;
  source_image?: string;
  template_image?: string;
  output_image?: string;
  source_size?: number[] | readonly number[];
  target_size?: number[] | readonly number[];
  scale?: number[] | readonly number[];
  ocr_markdown?: string;
  boxes?: BoxSpec[];
  image_patches?: ImagePatchSpec[];
  [key: string]: unknown;
};

type CompactYamlBox = {
  id?: string;
  type?: string;
  role?: FontRole | string;
  text?: string;
  src?: unknown;
  box?: unknown;
  tight?: unknown;
  font?: unknown;
  line?: unknown;
  color?: string;
  align?: TextAlign | string;
  emphasis?: unknown;
};

type CompactYamlPatch = {
  id?: string;
  src?: unknown;
  box?: unknown;
};

type CompactYamlBlock = CompactYamlBox & {
  image?: string;
  table?: string;
  patch?: boolean;
};

type CompactYamlPage = {
  md?: string;
  source_size?: unknown;
  target_size?: unknown;
  pdf_page_box?: unknown;
  render_dpi?: unknown;
  boxes?: CompactYamlBox[];
  patches?: CompactYamlPatch[];
  blocks?: CompactYamlBlock[];
};

type CompactYamlRoot = {
  schema?: number;
  pages?: Partial<Record<PageId, CompactYamlPage>>;
};

type CompactCoverYamlPage = {
  template_image?: string;
  output_image?: string;
  target_size?: unknown;
  boxes?: CompactYamlBox[];
};

type CompactCoverYamlRoot = {
  schema?: number;
  pages?: Partial<Record<CoverPageId, CompactCoverYamlPage>>;
};

type ParseEditablePptxToYamlResult = {
  yaml_path: string;
  updated: number;
  unknown_shapes: string[];
};

export type PdfPageModel = {
  pageSizePt: [number, number];
  renderDpi: number;
  renderSizePx?: [number, number];
  embeddedImageSizePx?: [number, number];
};

export type InspectCanonicalPdfResult = {
  pageCount: number;
  pageSizePt: [number, number];
  renderSizePx: [number, number];
  renderDpi: number;
  pptxSlideSizeIn: [number, number];
  embeddedImageSizePx?: [number, number];
};

export type PdfTextRun = {
  pageIndex: number;
  text: string;
  box: [number, number, number, number];
  fontFamily?: string;
  fontSize?: number;
  color?: string;
  role?: FontRole;
};

export type ComparePdfOraclePagesResult = {
  ok: boolean;
  dimensionMismatch?: boolean;
  diffPixels?: number;
  changedRatio?: number;
  meanAbsDelta?: number;
  maxDelta?: number;
  message?: string;
};

export type PdfOraclePageDiffMetric = {
  pageId: string;
  diffPixels?: number;
  mismatchRatio?: number;
  dimensionMismatch?: boolean;
};

export type PdfOracleRoundTripReport = {
  ok: boolean;
  canonicalRenderSizePx: [number, number];
  renderer: { kind: string };
  incompleteReason?: string;
  pageRenderSizePx?: [number, number];
  generatedRenderSizePx?: [number, number];
  pageDiffs?: PdfOraclePageDiffMetric[];
};
export type IngestCanonicalPdfToYamlReport = {
  canonicalRenderSizePx: [number, number];
  updatedPagesCount: number;
  renderDpi: number;
  dryRun: boolean;
  outputYamlPath: string;
};



export const DEFAULT_ZH_TW_CANONICAL_PDF = "/Users/au/Desktop/zh-TW-editable.pdf";
export const DEFAULT_ZH_TW_COVER_CANONICAL_PDF = "/Users/au/Documents/zh-TW-cover-editable.pdf";



export type OrderedXmlNode = Record<string, unknown>;

export type GroupTransform = {
  offX: number;
  offY: number;
  extX: number;
  extY: number;
  choffX: number;
  choffY: number;
  chextX: number;
  chextY: number;
  parent: GroupTransform | null;
};

export type ParsedPptxShape = {
  name: string;
  text: string;
  box: number[];
  font: number | undefined;
  align: TextAlign;
  emphasis: EmphasisSpec[];
};
export type DiffReport = {
  image: unknown;
  rendered_path: string;
  source_path: string;
  mask_path: string | null;
  threshold: number;
  total_pixels: number;
  changed_pixels: number;
  changed_ratio: number;
  mean_abs_delta: number;
};

export class VerificationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "VerificationError";
  }
}

const THIS_FILE = fileURLToPath(import.meta.url);
export const PROJECT_ROOT = dirname(dirname(THIS_FILE));
export const ASSETS_ROOT = join(PROJECT_ROOT, "translation-assets", "zh-TW");
export const ZH_TW_YAML_PATH = join(PROJECT_ROOT, "translation-assets", "zh-TW.yml");
export const RAW_OCR_DIR = join(ASSETS_ROOT, "ocr", "raw");
export const TEXT_SPEC_DIR = join(ASSETS_ROOT, "text");
export const ZH_TW_COVER_YAML_PATH = join(PROJECT_ROOT, "translation-assets", "zh-TW-cover.yml");
export const COVER_TEXT_SPEC_DIR = join(ASSETS_ROOT, "cover-text");
export const REVIEW_DIR = join(ASSETS_ROOT, "review");
export const SOURCE_IMAGE_ROOT = "/Users/au/w/civic.ai/img";
export const WORDLESS_ROOT = join(PROJECT_ROOT, "comics-wordless");
export const OUTPUT_ROOT = join(PROJECT_ROOT, "comics-zh-TW");
export const PPTX_OUTPUT_PATH = join(OUTPUT_ROOT, "zh-TW-editable.pptx");
export const COVER_PPTX_OUTPUT_PATH = join(OUTPUT_ROOT, "zh-TW-cover-editable.pptx");
export const ENGLISH_ROOT = join(PROJECT_ROOT, "comics-en");
export const BOOK_COVER_ROOT = join(PROJECT_ROOT, "book-cover");

const PPTX_RENDER_DPI = 600;
const PPTX_SLIDE_WIDTH_IN = 5749 / PPTX_RENDER_DPI;
const PPTX_SLIDE_HEIGHT_IN = 8000 / PPTX_RENDER_DPI;
const PPTX_LAYOUT_NAME = "ZH_TW_COMIC_PAGE";
const COVER_PPTX_LAYOUT_NAME = "ZH_TW_BOOK_COVER";
const PPTX_BACKGROUND_MAX_HEIGHT = 2000;
const PPTX_EMU_PER_INCH = 914400;
const PPTX_EMU_PER_RENDER_PX = PPTX_EMU_PER_INCH / PPTX_RENDER_DPI;

function runCommandText(args: string[]): string {
  const result = Bun.spawnSync(args, { stdout: "pipe", stderr: "pipe" });
  const stdout = new TextDecoder().decode(result.stdout);
  const stderr = new TextDecoder().decode(result.stderr);
  if (result.exitCode !== 0) {
    throw new VerificationError(`${args.join(" ")} failed with exit ${result.exitCode}: ${stderr.trim()}`);
  }
  return stdout;
}

function renderSizeFromPageBox(pageSizePt: readonly [number, number], renderDpi: number): [number, number] {
  return [Math.round((pageSizePt[0] * renderDpi) / 72), Math.round((pageSizePt[1] * renderDpi) / 72)];
}

function slideSizeInchesFromPageBox(pageSizePt: readonly [number, number]): [number, number] {
  return [pageSizePt[0] / 72, pageSizePt[1] / 72];
}

function pageModelRenderSize(model: PdfPageModel): [number, number] {
  return model.renderSizePx ?? renderSizeFromPageBox(model.pageSizePt, model.renderDpi);
}

function parseFirstPdfImageSize(pdfPath: string): [number, number] | undefined {
  const output = runCommandText(["pdfimages", "-list", pdfPath]);
  for (const line of output.split(/\r?\n/)) {
    const match = line.trim().match(/^\d+\s+\d+\s+image\s+(\d+)\s+(\d+)\b/);
    if (match) return [Number(match[1]), Number(match[2])];
  }
  return undefined;
}

export async function inspectCanonicalPdf(
  pdfPath = DEFAULT_ZH_TW_CANONICAL_PDF,
  options: { renderDpi?: number } = {},
): Promise<InspectCanonicalPdfResult> {
  const resolvedPdfPath = resolveProjectPath(pdfPath);
  const renderDpi = options.renderDpi ?? PPTX_RENDER_DPI;
  const info = runCommandText(["pdfinfo", resolvedPdfPath]);
  const pageCount = Number(info.match(/^Pages:\s+(\d+)/m)?.[1]);
  const sizeMatch = info.match(/^Page size:\s+([\d.]+)\s+x\s+([\d.]+)\s+pts\b/m);
  if (!Number.isFinite(pageCount) || !sizeMatch) {
    throw new VerificationError(`Unable to inspect PDF page metadata: ${resolvedPdfPath}`);
  }
  const pageSizePt = [Math.round(Number(sizeMatch[1])), Math.round(Number(sizeMatch[2]))] as [number, number];
  const renderSizePx = renderSizeFromPageBox(pageSizePt, renderDpi);
  return {
    pageCount,
    pageSizePt,
    renderSizePx,
    renderDpi,
    pptxSlideSizeIn: slideSizeInchesFromPageBox(pageSizePt),
    embeddedImageSizePx: parseFirstPdfImageSize(resolvedPdfPath),
  };
}

function xmlAttributes(raw: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  for (const match of raw.matchAll(/\s([A-Za-z_:][\w:.-]*)="([^"]*)"/g)) {
    attrs[match[1]!] = decodeXmlEntities(match[2]!);
  }
  return attrs;
}

function decodeXmlEntities(text: string): string {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function normalizePdfExtractedText(text: string): string {
  return decodeXmlEntities(text)
    .replace(/\u00a0/g, " ")
    .trimEnd()
    .replace(/不是！/g, "不是：")
    .replace(/而是！/g, "而是：");
}

function pdfRunRole(fontFamily: string | undefined): FontRole {
  return /lanyanghei|蘭陽黑/i.test(fontFamily ?? "") ? "title" : "body";
}

export async function extractPdfTextRuns(pdfPath: string, pageIndex: number): Promise<PdfTextRun[]> {
  const resolvedPdfPath = resolveProjectPath(pdfPath);
  const report = await inspectCanonicalPdf(resolvedPdfPath);
  const xml = runCommandText([
    "pdftohtml",
    "-xml",
    "-i",
    "-f",
    String(pageIndex),
    "-l",
    String(pageIndex),
    "-stdout",
    resolvedPdfPath,
  ]);
  const pageAttrs = xmlAttributes(xml.match(/<page\b([^>]*)>/)?.[1] ?? "");
  const xmlWidth = Number(pageAttrs.width);
  const xmlHeight = Number(pageAttrs.height);
  if (!Number.isFinite(xmlWidth) || !Number.isFinite(xmlHeight) || xmlWidth <= 0 || xmlHeight <= 0) {
    throw new VerificationError(`Unable to read pdftohtml page dimensions for page ${pageIndex}`);
  }
  const [renderWidth, renderHeight] = report.renderSizePx;
  const scaleX = renderWidth / xmlWidth;
  const scaleY = renderHeight / xmlHeight;
  const fonts = new Map<string, { family?: string; size?: number; color?: string }>();
  for (const match of xml.matchAll(/<fontspec\b([^>]*)\/>/g)) {
    const attrs = xmlAttributes(match[1]!);
    fonts.set(String(attrs.id), {
      family: attrs.family,
      size: Number.isFinite(Number(attrs.size)) ? Number(attrs.size) : undefined,
      color: hexColorOverride(attrs.color),
    });
  }
  const runs: PdfTextRun[] = [];
  for (const match of xml.matchAll(/<text\b([^>]*)>([\s\S]*?)<\/text>/g)) {
    const attrs = xmlAttributes(match[1]!);
    const text = normalizePdfExtractedText(match[2] ?? "");
    if (!text.trim()) continue;
    const font = fonts.get(String(attrs.font));
    const left = Number(attrs.left);
    const top = Number(attrs.top);
    const width = Number(attrs.width);
    const height = Number(attrs.height);
    if (![left, top, width, height].every(Number.isFinite)) continue;
    const x0 = Math.round(left * scaleX);
    const y0 = Math.round(top * scaleY);
    const x1 = Math.round((left + width) * scaleX);
    const y1 = Math.round((top + height) * scaleY);
    const fontFamily = font?.family?.replace(/^[A-Z]{6}\+/, "");
    runs.push({
      pageIndex,
      text,
      box: [x0, y0, x1, y1],
      fontFamily,
      fontSize: font?.size === undefined ? undefined : Math.round(font.size * scaleY),
      color: font?.color,
      role: pdfRunRole(fontFamily),
    });
  }
  return runs;
}

const PPTX_XML_PARSER = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: true,
  textNodeName: "#text",
  trimValues: false,
});
export const COVER_IMAGE_NAMES = ["Cover.png", "Cover_Back.png", "Cover_Spine.png"] as const;
export const COVER_EXTRA_IMAGE_NAMES = [join("extras", "with longer author name.png")] as const;
export const COVER_PAGE_ORDER = ["Cover", "Cover_Back", "Cover_Spine"] as const satisfies readonly CoverPageId[];
export const COVER_TEMPLATE_IMAGE_NAMES: Record<CoverPageId, string> = {
  Cover: "Cover.png",
  Cover_Back: "Cover_Back.png",
  Cover_Spine: "Cover_Spine.png",
};


export const PAGE_ORDER = [
  "Overview",
  "Ch1_A",
  "Ch1_B",
  "Ch2_A",
  "Ch2_B",
  "Ch3_A",
  "Ch3_B",
  "Ch4_A",
  "Ch4_B",
  "Ch5_A",
  "Ch5_B",
  "Ch6_A",
  "Ch6_B",
] as const satisfies readonly PageId[];

export const SOURCE_IMAGE_NAMES: Record<string, string> = Object.fromEntries([
  ["Overview", "overview-small-tw.jpg"],
  ...Array.from({ length: 6 }, (_, index) => index + 1).flatMap((chapter) => [
    [`Ch${chapter}_A`, `pack${chapter}-1-tw.jpg`],
    [`Ch${chapter}_B`, `pack${chapter}-2-tw.jpg`],
  ]),
]);

export const TEMPLATE_IMAGE_NAMES: Record<string, string> = Object.fromEntries([
  ["Overview", "Overview.png"],
  ...Array.from({ length: 6 }, (_, index) => index + 1).flatMap((chapter) => [
    [`Ch${chapter}_A`, `Ch${chapter}_A.png`],
    [`Ch${chapter}_B`, `Ch${chapter}_B.png`],
  ]),
]);

export const FOOTER_TEXT =
  "100% 人類手繪！漫畫：Nicky Case。書：唐鳳與 Caroline Green。\n" +
  "以 CC0 自由獻給公眾領域。到 https://Civic.AI 了解更多 <3";

const SIMPLIFIED_TRAP_CHARS = new Set(
  Array.from(
    "这这个们为无与过还让听说话语间门问实师会国风发现应责关怀义证认务术么" +
      "简优选则处众声号标页图层数权势赢输变边错导审计参专东炼练复",
  ),
);

const FONT_CANDIDATES: Record<FontRole, string[]> = {
  title: [
    "/Users/au/Library/Fonts/jf-lanyanghei-1.0-w10.otf",
    "/Users/au/Library/Fonts/jf-lanyanghei-1.0-w11.otf",
    "/Users/au/Library/Fonts/jf-lanyanghei-1.0-w9.otf",
    "/Users/au/Library/Fonts/jf-lanyangming-2.0-bold.otf",
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
  ],
  body: [
    "/Users/au/Library/Fonts/jf-lanyangming-2.0-medium.otf",
    "/Users/au/Library/Fonts/jf-lanyangming-2.0-regular.otf",
    "/Users/au/Library/Fonts/jf-lanyangming-2.0-book.otf",
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
  ],
  footer: [
    "/Users/au/Library/Fonts/jf-lanyangming-2.0-medium.otf",
    "/Users/au/Library/Fonts/jf-lanyangming-2.0-regular.otf",
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
  ],
};

const FONT_FAMILY: Record<FontRole, string> = {
  title: "ZhTwComicTitle",
  body: "ZhTwComicBody",
  footer: "ZhTwComicFooter",
};

const registeredFonts = new Map<FontRole, string | null>();
const measureCanvas = createCanvas(16, 16);
const measureContext = measureCanvas.getContext("2d");

export function buildComicsOnlyManifest(
  assetsRoot: string = ASSETS_ROOT,
  sourceRoot: string = SOURCE_IMAGE_ROOT,
): Record<string, string> {
  void assetsRoot;
  return Object.fromEntries(PAGE_ORDER.map((pageId) => [pageId, join(sourceRoot, SOURCE_IMAGE_NAMES[pageId])])) as Record<
    string,
    string
  >;
}

export function findFont(role: FontRole | string = "body"): string | null {
  const normalized = normalizeRole(role);
  return FONT_CANDIDATES[normalized].find((path) => existsSync(path)) ?? null;
}

export function registerFontForRole(role: FontRole | string = "body"): string | null {
  const normalized = normalizeRole(role);
  if (registeredFonts.has(normalized)) return registeredFonts.get(normalized) ?? null;
  const family = FONT_FAMILY[normalized];
  if (GlobalFonts.has(family)) {
    registeredFonts.set(normalized, family);
    return family;
  }
  for (const path of FONT_CANDIDATES[normalized]) {
    if (!existsSync(path)) continue;
    try {
      const key = GlobalFonts.registerFromPath(path, family);
      if (key || GlobalFonts.has(family)) {
        registeredFonts.set(normalized, family);
        return family;
      }
    } catch {
      // Keep trying lower-priority candidates. Missing CJK fonts fail loudly at the end.
    }
  }
  registeredFonts.set(normalized, null);
  return null;
}

export function measureRegisteredText(text: string, role: FontRole | string = "body", size = 48) {
  const font = fontSpec(role, size);
  return measureText(text, font);
}

export async function probeCanvasRuntime(sample = "好啦關懷六力") {
  const fontPath = findFont("body");
  if (!fontPath) throw new VerificationError("No Traditional-Chinese-capable font found");
  const family = registerFontForRole("body");
  if (!family) throw new VerificationError(`Failed to register CJK font: ${fontPath}`);
  const canvas = createCanvas(360, 140);
  const ctx = canvas.getContext("2d", { alpha: false });
  ctx.fillStyle = "white";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.font = cssFont({ fontFamily: family, size: 48 });
  ctx.fillStyle = "#121212";
  const metrics = ctx.measureText(sample);
  ctx.fillText(sample, 10, 80);
  const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
  let darkPixels = 0;
  for (let index = 0; index < data.length; index += 4) {
    if (data[index]! < 250 || data[index + 1]! < 250 || data[index + 2]! < 250) darkPixels++;
  }
  return {
    fontPath,
    registered: true,
    width: metrics.width,
    ascent: metrics.actualBoundingBoxAscent,
    descent: metrics.actualBoundingBoxDescent,
    darkPixels,
  };
}

export function verifyTextSpec(spec: TextSpec | Record<string, unknown>, requiredBoxIds: string[] = []): void {
  const boxes = Array.isArray(spec.boxes) ? (spec.boxes as BoxSpec[]) : [];
  const presentIds = new Set(boxes.map((box) => String(box.id)));
  for (const id of requiredBoxIds) {
    if (!presentIds.has(id)) throw new VerificationError(`missing required box id: ${id}`);
  }
  for (const box of boxes) {
    const id = String(box.id ?? "<unknown>");
    const text = String(box.text ?? "");
    for (const char of Array.from(text)) {
      if (SIMPLIFIED_TRAP_CHARS.has(char)) {
        throw new VerificationError(`box ${id} contains simplified/trap character ${char}: ${text}`);
      }
    }
    for (const key of ["box", "tight_box", "source_box"] as const) {
      const value = box[key];
      if (!Array.isArray(value)) continue;
      if (value.length !== 4) throw new VerificationError(`box ${id} ${key} must have four coordinates`);
      const [x0, y0, x1, y1] = value.map(Number);
      if (![x0, y0, x1, y1].every(Number.isFinite)) throw new VerificationError(`box ${id} ${key} has non-numeric coordinates`);
      if (x1 <= x0 || y1 <= y0) throw new VerificationError(`box ${id} ${key} has non-positive extent`);
    }
  }
}

export function cleanOcrContent(content: string): string {
  const tableRows = parseMarkdownTable(content);
  if (tableRows.length > 0) {
    return tableRows.map((row) => row.filter(Boolean).join("　　")).join("\n");
  }

  const cleanedLines: string[] = [];
  for (const rawLine of content.split(/\r?\n/)) {
    let line = rawLine.trim();
    if (!line) {
      cleanedLines.push("");
      continue;
    }
    if (/^!\[[^\]]*\]\([^)]*\)/.test(line)) continue;
    line = line.replace(/^#{1,6}\s*/, "").trim();
    cleanedLines.push(line);
  }
  while (cleanedLines[0] === "") cleanedLines.shift();
  while (cleanedLines[cleanedLines.length - 1] === "") cleanedLines.pop();
  return cleanedLines.join("\n");
}


const NO_LINE_START_PUNCTUATION = /[，。、；：！？）》」』】）,.!?;:]/u;

export function wrapTextPreserveBreaks(text: string, font: FontSpec, maxWidthPx: number): string {
  const maxWidth = Math.max(1, maxWidthPx);
  const wrappedLines: string[] = [];
  for (const explicitLine of String(text).split("\n")) {
    if (explicitLine === "") {
      wrappedLines.push("");
      continue;
    }
    let current = "";
    let lastBreakableAt = -1;
    for (const char of segmentText(explicitLine)) {
      const candidate = current + char;
      const cannotStartLine = NO_LINE_START_PUNCTUATION.test(char);
      const breakableAt = /\s/.test(char) || cannotStartLine ? current.length + char.length : -1;
      if (current && textWidth(candidate, font) > maxWidth) {
        if (cannotStartLine && lastBreakableAt > 0 && lastBreakableAt < current.length) {
          wrappedLines.push(candidate.slice(0, lastBreakableAt).trimEnd());
          current = candidate.slice(lastBreakableAt).trimStart();
        } else if (cannotStartLine) {
          current = candidate;
        } else if (lastBreakableAt > 0 && lastBreakableAt < candidate.length) {
          wrappedLines.push(candidate.slice(0, lastBreakableAt).trimEnd());
          current = candidate.slice(lastBreakableAt).trimStart();
        } else {
          wrappedLines.push(current.trimEnd());
          current = char.trimStart();
        }
        lastBreakableAt = -1;
      } else {
        current = candidate;
        if (breakableAt > 0) lastBreakableAt = breakableAt;
      }
    }
    if (current) wrappedLines.push(current.trimEnd());
  }
  return wrappedLines.join("\n");
}

export function fitFontSizeInBox(
  text: string,
  box: readonly number[] | { width: number; height: number },
  maxFontSizePx = 42,
  minFontSizePx = 8,
  role: FontRole | string = "body",
  maxWrappedLines: number | null = null,
): number {
  let boxWidth: number;
  let boxHeight: number;
  if ("width" in box) {
    boxWidth = Math.max(1, box.width);
    boxHeight = Math.max(1, box.height);
  } else {
    boxWidth = Math.max(1, Number(box[2]) - Number(box[0]));
    boxHeight = Math.max(1, Number(box[3]) - Number(box[1]));
  }
  const maxSize = Math.max(minFontSizePx, Math.floor(maxFontSizePx));
  const minSize = Math.max(1, Math.floor(minFontSizePx));
  for (let size = maxSize; size >= minSize; size--) {
    const font = fontSpec(role, size);
    const wrapped = wrapTextPreserveBreaks(text, font, boxWidth);
    const lines = wrapped.split("\n");
    if (maxWrappedLines !== null && lines.length > maxWrappedLines) continue;
    const height = blockTextHeight(lines, font);
    const widest = Math.max(...lines.map((line) => textWidth(line, font)), 0);
    if (widest <= boxWidth && height <= boxHeight) return size;
  }
  return minSize;
}

export async function specFromOcrRaw(raw: Record<string, any>): Promise<TextSpec> {
  const pageId = String(raw.page_id ?? "");
  if (!PAGE_ORDER.includes(pageId as PageId)) throw new VerificationError(`unknown page_id in OCR raw: ${pageId}`);

  const sourcePath = String(raw.source_image ?? join(SOURCE_IMAGE_ROOT, SOURCE_IMAGE_NAMES[pageId]!));
  const templatePath = join(WORDLESS_ROOT, TEMPLATE_IMAGE_NAMES[pageId]!);
  const sourceSize = await imageSize(sourcePath);
  const targetSize = await imageSize(templatePath);
  const ocrPages = raw.ocr?.pages ?? [];
  if (!Array.isArray(ocrPages) || ocrPages.length === 0) throw new VerificationError(`OCR raw for ${pageId} has no pages`);
  const page = ocrPages[0];
  const sourceCanvas = await loadCanvasFromFile(sourcePath);
  const templateCanvas = await loadCanvasFromFile(templatePath);
  const englishCanvas = await loadCanvasFromFile(join(ENGLISH_ROOT, TEMPLATE_IMAGE_NAMES[pageId]!));
  const englishMask = createEnglishDiffMask(englishCanvas, templateCanvas);
  const boxes: BoxSpec[] = [];
  const imagePatches: ImagePatchSpec[] = [];
  let textIndex = 1;
  for (const block of page.blocks ?? []) {
    if (block.type === "image") {
      const sourceBox = [
        Number(block.top_left_x),
        Number(block.top_left_y),
        Number(block.bottom_right_x),
        Number(block.bottom_right_y),
      ];
      if (block.render_patch === true || (pageId === "Overview" && block.render_patch !== false)) {
        imagePatches.push({
          id: String(block.id ?? block.image_id ?? `image_patch_${imagePatches.length + 1}`),
          source_box: sourceBox,
          box: numericBoxOverride(block.target_box ?? block.box, targetSize) ?? targetBox(sourceBox, sourceSize, targetSize),
        });
      }
      continue;
    }
    const content = cleanOcrContent(String(block.content ?? ""));
    if (!content) continue;
    const sourceBox = [
      Number(block.top_left_x),
      Number(block.top_left_y),
      Number(block.bottom_right_x),
      Number(block.bottom_right_y),
    ];
    const projectedTarget = targetBox(sourceBox, sourceSize, targetSize);
    const role: FontRole = block.type === "title" ? "title" : "body";
    const englishGeometry = englishTextGeometryFromBox(englishMask, projectedTarget, role, content);
    const boxOverride = numericBoxOverride(block.target_box ?? block.box, targetSize);
    const tightBoxOverride = numericBoxOverride(block.target_tight_box ?? block.tight_box, targetSize);
    const targetTight = boxOverride ?? englishGeometry?.box ?? projectedTarget;
    const tightBox = tightBoxOverride ?? englishGeometry?.tightBox ?? targetTight;
    const [maxFont, minFont, maxLines] = nominalFontSize(content, targetTight, sourceBox, role);
    const fittedSize = numericOverride(block.font_size) ?? fitFontSizeInBox(content, targetTight, maxFont, minFont, role, maxLines);
    const lineSpacingOverride = numericOverride(block.line_spacing);
    const alignOverride = block.align === "center" || block.align === "left" ? block.align : undefined;
    const emphasis = emphasisOverride(block.emphasis);
    const boxId =
      typeof block.id === "string" && block.id.trim()
        ? block.id.trim()
        : `b${String(textIndex).padStart(3, "0")}_${role}`;
    boxes.push({
      id: boxId,
      ocr_type: block.type ?? "text",
      font_role: role,
      text: content,
      source_box: sourceBox,
      box: targetTight,
      tight_box: tightBox,
      font_size: fittedSize,
      line_spacing: lineSpacingOverride,
      color: typeof block.color === "string" && block.color ? block.color : ensureContrastingInkColor(estimateInkColor(sourceCanvas, sourceBox), templateCanvas, targetTight),
      align: alignOverride ?? inferAlign(content, sourceBox, sourceSize, role),
      ...(emphasis ? { emphasis } : {}),
    });
    textIndex++;
  }
  const spec: TextSpec = {
    schema_version: 1,
    image: pageId,
    source_image: sourcePath,
    template_image: relative(PROJECT_ROOT, templatePath),
    output_image: relative(PROJECT_ROOT, join(OUTPUT_ROOT, TEMPLATE_IMAGE_NAMES[pageId]!)),
    source_size: [sourceSize.width, sourceSize.height],
    target_size: [targetSize.width, targetSize.height],
    scale: [targetSize.width / sourceSize.width, targetSize.height / sourceSize.height],
    ocr_markdown: page.markdown ?? "",
    boxes,
    ...(imagePatches.length > 0 ? { image_patches: imagePatches } : {}),
  };
  verifyTextSpec(spec, boxes.map((box) => box.id));
  return spec;
}

export async function loadYamlTextSpecs(yamlPath = ZH_TW_YAML_PATH): Promise<Record<PageId, TextSpec>> {
  const loaded = Bun.YAML.parse(await readFile(yamlPath, "utf8")) as CompactYamlRoot;
  if (loaded?.schema !== 1) throw new VerificationError(`unsupported zh-TW YAML schema: ${loaded?.schema}`);
  if (!loaded.pages || typeof loaded.pages !== "object") throw new VerificationError("zh-TW YAML has no pages map");

  const specs = {} as Record<PageId, TextSpec>;
  for (const pageId of PAGE_ORDER) {
    const page = loaded.pages[pageId];
    if (!page) throw new VerificationError(`zh-TW YAML missing page: ${pageId}`);
    specs[pageId] = Array.isArray(page.boxes)
      ? await specFromCompactYamlPage(pageId, page)
      : await specFromOcrRaw(rawFromCompactYamlPage(pageId, page));
  }
  return specs;
}

async function specFromCompactYamlPage(pageId: PageId, page: CompactYamlPage): Promise<TextSpec> {
  if (!Array.isArray(page.boxes) || page.boxes.length === 0) {
    throw new VerificationError(`zh-TW YAML page ${pageId} has no boxes`);
  }
  const sourcePath = join(SOURCE_IMAGE_ROOT, SOURCE_IMAGE_NAMES[pageId]!);
  const templatePath = join(WORDLESS_ROOT, TEMPLATE_IMAGE_NAMES[pageId]!);
  const sourceSize = requiredSizeFromCompactYaml(page.source_size, `zh-TW YAML ${pageId} source_size`);
  const targetSize = requiredSizeFromCompactYaml(page.target_size, `zh-TW YAML ${pageId} target_size`);
  const boxes = page.boxes.map((box, index) => boxFromCompactYamlBox(pageId, box, index));
  const imagePatches = (page.patches ?? []).map((patch, index) => patchFromCompactYamlPatch(pageId, patch, index));
  const spec: TextSpec = {
    schema_version: 1,
    image: pageId,
    source_image: sourcePath,
    template_image: relative(PROJECT_ROOT, templatePath),
    output_image: relative(PROJECT_ROOT, join(OUTPUT_ROOT, TEMPLATE_IMAGE_NAMES[pageId]!)),
    source_size: [sourceSize.width, sourceSize.height],
    target_size: [targetSize.width, targetSize.height],
    scale: [targetSize.width / sourceSize.width, targetSize.height / sourceSize.height],
    ocr_markdown: page.md ?? "",
    boxes,
    ...(imagePatches.length > 0 ? { image_patches: imagePatches } : {}),
  };
  verifyTextSpec(spec, boxes.map((box) => box.id));
  return spec;
}

export async function loadCoverTextSpecs(yamlPath = ZH_TW_COVER_YAML_PATH): Promise<Record<CoverPageId, TextSpec>> {
  const loaded = Bun.YAML.parse(await readFile(yamlPath, "utf8")) as CompactCoverYamlRoot;
  if (loaded?.schema !== 1) throw new VerificationError(`unsupported zh-TW cover YAML schema: ${loaded?.schema}`);
  if (!loaded.pages || typeof loaded.pages !== "object") throw new VerificationError("zh-TW cover YAML has no pages map");

  const specs = {} as Record<CoverPageId, TextSpec>;
  for (const pageId of COVER_PAGE_ORDER) {
    const page = loaded.pages[pageId];
    if (!page) throw new VerificationError(`zh-TW cover YAML missing page: ${pageId}`);
    specs[pageId] = specFromCompactCoverYamlPage(pageId, page);
  }
  return specs;
}

function specFromCompactCoverYamlPage(pageId: CoverPageId, page: CompactCoverYamlPage): TextSpec {
  if (!Array.isArray(page.boxes) || page.boxes.length === 0) {
    throw new VerificationError(`zh-TW cover YAML page ${pageId} has no boxes`);
  }
  const templateName = page.template_image ?? join("book-cover", "wordless", COVER_TEMPLATE_IMAGE_NAMES[pageId]);
  const outputName = page.output_image ?? COVER_TEMPLATE_IMAGE_NAMES[pageId];
  const targetSize = requiredSizeFromCompactYaml(page.target_size, `zh-TW cover YAML ${pageId} target_size`);
  const boxes = page.boxes.map((box, index) => boxFromCompactCoverYamlBox(pageId, box, index));
  const spec: TextSpec = {
    schema_version: 1,
    image: pageId,
    template_image: templateName,
    output_image: outputName,
    target_size: [targetSize.width, targetSize.height],
    boxes,
  };
  verifyTextSpec(spec, boxes.map((box) => box.id));
  return spec;
}

function boxFromCompactCoverYamlBox(pageId: CoverPageId, box: CompactYamlBox, index: number): BoxSpec {
  if (!box || typeof box !== "object") throw new VerificationError(`zh-TW cover YAML ${pageId} box ${index} is invalid`);
  const id = box.id?.trim();
  if (!id) throw new VerificationError(`zh-TW cover YAML ${pageId} box ${index} has no id`);
  const type = box.type ?? "text";
  const role = box.role ?? (type === "title" ? "title" : "body");
  const targetBox = requiredCoordinateBox(box.box, `zh-TW cover YAML ${pageId} box ${index} box`);
  const emphasis = emphasisOverride(box.emphasis);
  return {
    id,
    ocr_type: type,
    font_role: role,
    text: box.text ?? "",
    box: targetBox,
    tight_box: requiredCoordinateBox(box.tight ?? box.box, `zh-TW cover YAML ${pageId} box ${index} tight`),
    font_size: box.font === undefined ? 100 : requiredNonNegativeNumber(box.font, `zh-TW cover YAML ${pageId} box ${index} font`),
    ...(box.line !== undefined ? { line_spacing: requiredNonNegativeNumber(box.line, `zh-TW cover YAML ${pageId} box ${index} line`) } : {}),
    color: box.color ?? "#222222",
    align: box.align === "center" || box.align === "left" ? box.align : "left",
    ...(emphasis ? { emphasis } : {}),
  };
}

function boxFromCompactYamlBox(pageId: PageId, box: CompactYamlBox, index: number): BoxSpec {
  if (!box || typeof box !== "object") throw new VerificationError(`zh-TW YAML ${pageId} box ${index} is invalid`);
  const id = box.id?.trim();
  if (!id) throw new VerificationError(`zh-TW YAML ${pageId} box ${index} has no id`);
  const type = box.type ?? "text";
  const role = box.role ?? (type === "title" ? "title" : "body");
  const sourceBox = requiredCoordinateBox(box.src, `zh-TW YAML ${pageId} box ${index} src`);
  const targetBox = requiredCoordinateBox(box.box, `zh-TW YAML ${pageId} box ${index} box`);
  const tightBox = requiredCoordinateBox(box.tight ?? box.box, `zh-TW YAML ${pageId} box ${index} tight`);
  const fontSize = requiredNonNegativeNumber(box.font, `zh-TW YAML ${pageId} box ${index} font`);
  const align = box.align === "center" || box.align === "left" ? box.align : "left";
  const emphasis = emphasisOverride(box.emphasis);
  return {
    id,
    ocr_type: type,
    font_role: role,
    text: box.text ?? "",
    source_box: sourceBox,
    box: targetBox,
    tight_box: tightBox,
    font_size: fontSize,
    ...(box.line !== undefined ? { line_spacing: requiredNonNegativeNumber(box.line, `zh-TW YAML ${pageId} box ${index} line`) } : {}),
    color: box.color ?? "#222222",
    align,
    ...(emphasis ? { emphasis } : {}),
  };
}

function patchFromCompactYamlPatch(pageId: PageId, patch: CompactYamlPatch, index: number): ImagePatchSpec {
  if (!patch || typeof patch !== "object") throw new VerificationError(`zh-TW YAML ${pageId} patch ${index} is invalid`);
  const id = patch.id?.trim();
  if (!id) throw new VerificationError(`zh-TW YAML ${pageId} patch ${index} has no id`);
  return {
    id,
    source_box: requiredCoordinateBox(patch.src, `zh-TW YAML ${pageId} patch ${index} src`),
    box: requiredCoordinateBox(patch.box, `zh-TW YAML ${pageId} patch ${index} box`),
  };
}


function requiredSizeFromCompactYaml(value: unknown, label: string): Size {
  if (!Array.isArray(value) || value.length !== 2) throw new VerificationError(`${label} must have width and height`);
  const [width, height] = value.map((coordinate) => Number(coordinate));
  if (!Number.isFinite(width) || !Number.isFinite(height) || width! <= 0 || height! <= 0) {
    throw new VerificationError(`${label} has invalid dimensions`);
  }
  return { width: width!, height: height! };
}
function rawFromCompactYamlPage(pageId: PageId, page: CompactYamlPage): Record<string, unknown> {
  if (!Array.isArray(page.blocks) || page.blocks.length === 0) {
    throw new VerificationError(`zh-TW YAML page ${pageId} has no blocks`);
  }
  return {
    page_id: pageId,
    ocr: {
      pages: [
        {
          markdown: page.md ?? "",
          blocks: page.blocks.map((block, index) => rawBlockFromCompactYamlBlock(pageId, block, index)),
        },
      ],
    },
  };
}

function rawBlockFromCompactYamlBlock(pageId: PageId, block: CompactYamlBlock, index: number): Record<string, unknown> {
  if (!block || typeof block !== "object") throw new VerificationError(`zh-TW YAML ${pageId} block ${index} is invalid`);
  const sourceBox = requiredCoordinateBox(block.src, `zh-TW YAML ${pageId} block ${index} src`);
  const raw: Record<string, unknown> = {
    top_left_x: sourceBox[0],
    top_left_y: sourceBox[1],
    bottom_right_x: sourceBox[2],
    bottom_right_y: sourceBox[3],
    content: block.text ?? "",
    type: block.type ?? "text",
  };
  if (block.id) raw.id = block.id;
  if (block.image) raw.image_id = block.image;
  if (block.table) raw.table_id = block.table;
  if (block.box !== undefined) raw.target_box = requiredCoordinateBox(block.box, `zh-TW YAML ${pageId} block ${index} box`);
  if (block.tight !== undefined) raw.target_tight_box = requiredCoordinateBox(block.tight, `zh-TW YAML ${pageId} block ${index} tight`);
  if (block.font !== undefined) raw.font_size = requiredNonNegativeNumber(block.font, `zh-TW YAML ${pageId} block ${index} font`);
  if (block.line !== undefined) raw.line_spacing = requiredNonNegativeNumber(block.line, `zh-TW YAML ${pageId} block ${index} line`);
  if (block.color) raw.color = block.color;
  if (block.align === "left" || block.align === "center") raw.align = block.align;
  if (block.emphasis !== undefined) raw.emphasis = block.emphasis;
  if (block.patch !== undefined) raw.render_patch = block.patch;
  return raw;
}

function requiredCoordinateBox(value: unknown, label: string): number[] {
  if (!Array.isArray(value) || value.length !== 4) throw new VerificationError(`${label} must have four coordinates`);
  const box = value.map((coordinate) => Number(coordinate));
  if (!box.every(Number.isFinite)) throw new VerificationError(`${label} has non-numeric coordinates`);
  const [x0, y0, x1, y1] = box;
  if (x1! <= x0! || y1! <= y0!) throw new VerificationError(`${label} has non-positive extent`);
  return box;
}

function requiredNonNegativeNumber(value: unknown, label: string): number {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) throw new VerificationError(`${label} must be a non-negative number`);
  return number;
}

function inputLooksLikeYaml(path: string): boolean {
  return path.endsWith(".yml") || path.endsWith(".yaml");
}


export async function buildTextSpecs(inputPath = ZH_TW_YAML_PATH, textDir = TEXT_SPEC_DIR): Promise<string[]> {
  await mkdir(textDir, { recursive: true });
  const specs = await loadYamlTextSpecs(inputLooksLikeYaml(inputPath) ? inputPath : ZH_TW_YAML_PATH);
  const written: string[] = [];
  for (const pageId of PAGE_ORDER) {
    const outPath = join(textDir, `${pageId}.json`);
    await writeJson(outPath, specs[pageId]);
    written.push(outPath);
  }
  return written;
}

export async function buildCoverTextSpecs(yamlPath = ZH_TW_COVER_YAML_PATH, textDir = COVER_TEXT_SPEC_DIR): Promise<string[]> {
  await mkdir(textDir, { recursive: true });
  const specs = await loadCoverTextSpecs(yamlPath);
  const written: string[] = [];
  for (const pageId of COVER_PAGE_ORDER) {
    const outPath = join(textDir, `${pageId}.json`);
    await writeJson(outPath, specs[pageId]);
    written.push(outPath);
  }
  return written;
}

export function estimateInkColor(image: Canvas | Image, sourceBox: readonly number[]): string {
  const { data, width, height } = imageDataInBox(image, sourceBox);
  if (width === 0 || height === 0 || data.length === 0) return "#222222";

  const pixelCount = data.length / 4;
  const stepForBrightness = Math.max(1, Math.floor(pixelCount / 20000));
  const brightnesses: number[] = [];
  for (let index = 0, pixel = 0; index < data.length; index += 4, pixel++) {
    if (pixel % stepForBrightness !== 0) continue;
    brightnesses.push((data[index]! + data[index + 1]! + data[index + 2]!) / 3);
  }
  brightnesses.sort((a, b) => a - b);
  const medianBrightness = brightnesses[Math.floor(brightnesses.length / 2)] ?? 255;

  if (medianBrightness < 105) {
    let lightPixels = collectPixelsByBrightness(data, 200, "above");
    if (lightPixels.length < Math.max(8, Math.floor(pixelCount * 0.001))) {
      lightPixels = collectPixelsByBrightness(data, 170, "above");
    }
    if (lightPixels.length >= Math.max(8, Math.floor(pixelCount * 0.002))) {
      return rgbToHex(medianRgb(lightPixels));
    }
  }

  const darkPixels = collectPixelsByBrightness(data, 205, "below");
  if (darkPixels.length === 0) return "#222222";
  const [red, green, blue] = medianRgb(darkPixels);
  if (Math.max(red, green, blue) - Math.min(red, green, blue) < 18 && (red + green + blue) / 3 > 110) {
    return "#555555";
  }
  return rgbToHex([red, green, blue]);
}

export async function renderPage(spec: TextSpec | Record<string, unknown>, outputPath?: string, mode: RenderMode = "vector"): Promise<string> {
  const typed = spec as TextSpec;
  verifyTextSpec(typed, (typed.boxes ?? []).map((box) => String(box.id)));
  const template = resolveProjectPath(String(typed.template_image ?? ""));
  const destination = outputPath ?? resolveProjectPath(String(typed.output_image ?? TEMPLATE_IMAGE_NAMES[String(typed.image)] ?? "out.png"));
  await mkdir(dirname(destination), { recursive: true });

  const image = await loadImage(template);
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(image, 0, 0);

  if (mode === "vector" || mode === "editable_text") {
    if ((typed.image_patches?.length ?? 0) > 0) {
      const sourcePath = String(typed.source_image ?? "");
      if (!sourcePath) throw new VerificationError(`spec ${typed.image} has no source_image`);
      const source = await loadImage(sourcePath);
      for (const patch of typed.image_patches ?? []) pasteSourcePatch(ctx, source, patch);
    }
    for (const box of typed.boxes ?? []) drawBoxText(ctx, box);
  } else if (mode === "oracle_patch") {
    const sourcePath = String(typed.source_image ?? "");
    if (!sourcePath) throw new VerificationError(`spec ${typed.image} has no source_image`);
    const source = await loadImage(sourcePath);
    for (const box of typed.boxes ?? []) pasteOraclePatch(ctx, source, typed, box);
  } else {
    throw new VerificationError(`unknown render mode: ${mode}`);
  }
  await saveCanvasPng(canvas, destination);
  return destination;
}

export async function renderPptxBackground(spec: TextSpec | Record<string, unknown>, outputPath: string): Promise<string> {
  const typed = spec as TextSpec;
  verifyTextSpec(typed, (typed.boxes ?? []).map((box) => String(box.id)));
  const template = resolveProjectPath(String(typed.template_image ?? ""));
  await mkdir(dirname(outputPath), { recursive: true });

  const image = await loadImage(template);
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(image, 0, 0);

  if ((typed.image_patches?.length ?? 0) > 0) {
    const sourcePath = String(typed.source_image ?? "");
    if (!sourcePath) throw new VerificationError(`spec ${typed.image} has no source_image`);
    const source = await loadImage(sourcePath);
    for (const patch of typed.image_patches ?? []) pasteSourcePatch(ctx, source, patch);
  }
  await saveCanvasPng(resizeCanvasSmoothToMaxHeight(canvas, PPTX_BACKGROUND_MAX_HEIGHT), outputPath);
  return outputPath;
}

export async function compareRenderedPageToSource(
  renderedPath: string,
  sourcePath: string,
  spec: TextSpec | Record<string, unknown>,
  options: { maskPath?: string; threshold?: number } = {},
): Promise<DiffReport> {
  const threshold = options.threshold ?? 3;
  const source = await loadCanvasFromFile(sourcePath);
  const rendered = await loadCanvasFromFile(renderedPath);
  const renderedDown = resizeCanvasNearest(rendered, source.width, source.height);
  const sourceData = source.getContext("2d").getImageData(0, 0, source.width, source.height).data;
  const renderedData = renderedDown.getContext("2d").getImageData(0, 0, source.width, source.height).data;
  const maskCanvas = createCanvas(source.width, source.height);
  const maskCtx = maskCanvas.getContext("2d");
  const mask = maskCtx.createImageData(source.width, source.height);
  let changedPixels = 0;
  let totalAbs = 0;
  for (let index = 0; index < sourceData.length; index += 4) {
    const red = Math.abs(renderedData[index]! - sourceData[index]!);
    const green = Math.abs(renderedData[index + 1]! - sourceData[index + 1]!);
    const blue = Math.abs(renderedData[index + 2]! - sourceData[index + 2]!);
    const delta = Math.max(red, green, blue);
    totalAbs += red + green + blue;
    const changed = delta > threshold;
    if (changed) changedPixels++;
    const value = changed ? 255 : 0;
    mask.data[index] = value;
    mask.data[index + 1] = value;
    mask.data[index + 2] = value;
    mask.data[index + 3] = 255;
  }
  maskCtx.putImageData(mask, 0, 0);
  if (options.maskPath) await saveCanvasPng(maskCanvas, options.maskPath);
  const totalPixels = source.width * source.height;
  return {
    image: (spec as TextSpec).image,
    rendered_path: renderedPath,
    source_path: sourcePath,
    mask_path: options.maskPath ?? null,
    threshold,
    total_pixels: totalPixels,
    changed_pixels: changedPixels,
    changed_ratio: totalPixels ? changedPixels / totalPixels : 0,
    mean_abs_delta: totalPixels ? totalAbs / (totalPixels * 3) : 0,
  };
}

export async function comparePdfOraclePages(
  leftPngPath: string,
  rightPngPath: string,
  options: { renderDpi?: number; threshold?: number } = {},
): Promise<ComparePdfOraclePagesResult> {
  const left = await imageSize(resolveProjectPath(leftPngPath));
  const right = await imageSize(resolveProjectPath(rightPngPath));
  if (left.width !== right.width || left.height !== right.height) {
    return {
      ok: false,
      dimensionMismatch: true,
      diffPixels: 0,
      message: `dimension mismatch: ${left.width}x${left.height} vs ${right.width}x${right.height}`,
    };
  }
  const threshold = options.threshold ?? 32;
  const leftCanvas = await loadCanvasFromFile(resolveProjectPath(leftPngPath));
  const rightCanvas = await loadCanvasFromFile(resolveProjectPath(rightPngPath));
  const leftData = leftCanvas.getContext("2d").getImageData(0, 0, left.width, left.height).data;
  const rightData = rightCanvas.getContext("2d").getImageData(0, 0, right.width, right.height).data;
  let diffPixels = 0;
  let totalAbs = 0;
  let maxDelta = 0;
  for (let index = 0; index < leftData.length; index += 4) {
    const red = Math.abs(leftData[index]! - rightData[index]!);
    const green = Math.abs(leftData[index + 1]! - rightData[index + 1]!);
    const blue = Math.abs(leftData[index + 2]! - rightData[index + 2]!);
    const delta = Math.max(red, green, blue);
    if (delta > threshold) diffPixels += 1;
    totalAbs += red + green + blue;
    maxDelta = Math.max(maxDelta, delta);
  }
  const totalPixels = left.width * left.height;
  return {
    ok: diffPixels === 0,
    diffPixels,
    changedRatio: totalPixels === 0 ? 0 : diffPixels / totalPixels,
    meanAbsDelta: totalPixels === 0 ? 0 : totalAbs / (totalPixels * 3),
    maxDelta,
  };
}


export async function generateDiffMasks(
  textDir = TEXT_SPEC_DIR,
  outputRoot = OUTPUT_ROOT,
  reviewDir = REVIEW_DIR,
  options: {
    threshold?: number;
    pageOrder?: readonly string[];
    templateImageNames?: Record<string, string>;
  } = {},
): Promise<{ reportPath: string; pages: Record<string, DiffReport> }> {
  const threshold = options.threshold ?? 32;
  const pageOrder = options.pageOrder ?? PAGE_ORDER;
  const templateNames = options.templateImageNames ?? TEMPLATE_IMAGE_NAMES;
  const diffDir = join(reviewDir, "diffs");
  await mkdir(diffDir, { recursive: true });
  const reports: Record<string, DiffReport> = {};
  for (const pageId of pageOrder) {
    const spec = JSON.parse(await readFile(join(textDir, `${pageId}.json`), "utf8")) as TextSpec;
    const renderedPath = join(outputRoot, templateNames[pageId] ?? `${pageId}.png`);
    const maskPath = join(diffDir, `${pageId}.mask.png`);
    reports[pageId] = await compareRenderedPageToSource(renderedPath, String(spec.source_image), spec, {
      maskPath,
      threshold,
    });
  }
  const reportPath = join(diffDir, "report.json");
  await writeJson(reportPath, reports);
  return { reportPath, pages: reports };
}

export async function renderReviewOverlay(spec: TextSpec, outputPath?: string): Promise<string> {
  const pageId = String(spec.image);
  const destination = outputPath ?? join(REVIEW_DIR, "overlays", `${pageId}.png`);
  await mkdir(dirname(destination), { recursive: true });
  const source = await loadImage(String(spec.source_image));
  const canvas = createCanvas(source.width, source.height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(source, 0, 0);
  ctx.strokeStyle = "#EB573A";
  ctx.lineWidth = 3;
  ctx.font = "16px sans-serif";
  ctx.textBaseline = "top";
  for (const box of spec.boxes ?? []) {
    if (!box.source_box) continue;
    const [x0, y0, x1, y1] = box.source_box.map(Number);
    ctx.strokeRect(x0, y0, x1 - x0, y1 - y0);
    const label = String(box.id);
    const labelY = Math.max(0, y0 - 18);
    const labelWidth = Math.ceil(ctx.measureText(label).width) + 6;
    ctx.fillStyle = "rgba(235,87,58,0.86)";
    ctx.fillRect(x0, labelY, labelWidth, 18);
    ctx.fillStyle = "white";
    ctx.fillText(label, x0 + 3, labelY + 1);
  }
  await saveCanvasPng(canvas, destination);
  return destination;
}

export async function renderFooter(outputPath = join(OUTPUT_ROOT, "web_footer.png")): Promise<string> {
  await mkdir(dirname(outputPath), { recursive: true });
  const footerTemplate = join(WORDLESS_ROOT, "web_footer.png");
  const image = await loadImage(footerTemplate);
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(image, 0, 0);
  const box = [200, 38, canvas.width - 200, canvas.height - 46] as const;
  const fontSize = fitFontSizeInBox(FOOTER_TEXT, box, 150, 70, "footer", null);
  const font = fontSpec("footer", fontSize);
  const wrapped = wrapTextPreserveBreaks(FOOTER_TEXT, font, box[2] - box[0]);
  const lines = wrapped.split("\n");
  const height = blockTextHeight(lines, font);
  const widest = Math.max(...lines.map((line) => textWidth(line, font)), 0);
  const spacing = lineSpacing(font);
  let y = box[1] + Math.max(0, (box[3] - box[1] - height) / 2);
  ctx.font = cssFont(font);
  ctx.textBaseline = "top";
  for (const line of lines) {
    const x = box[0] + Math.max(0, (box[2] - box[0] - textWidth(line, font)) / 2);
    ctx.fillText(line, x, y);
    y += lineHeight(font) + spacing;
  }
  void widest;
  await saveCanvasPng(canvas, outputPath);
  return outputPath;
}

function nonWhiteCanvasBounds(canvas: Canvas): [number, number, number, number] {
  const { width, height } = canvas;
  const data = canvas.getContext("2d").getImageData(0, 0, width, height).data;
  let minX = width;
  let minY = height;
  let maxX = 0;
  let maxY = 0;
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const index = (y * width + x) * 4;
      if (data[index]! >= 250 && data[index + 1]! >= 250 && data[index + 2]! >= 250) continue;
      minX = Math.min(minX, x);
      minY = Math.min(minY, y);
      maxX = Math.max(maxX, x);
      maxY = Math.max(maxY, y);
    }
  }
  if (minX > maxX || minY > maxY) return [0, 0, width, height];
  return [minX, minY, maxX + 1, maxY + 1];
}

export async function renderCoverAssetsFromCanonicalPdf(
  pdfPath = DEFAULT_ZH_TW_COVER_CANONICAL_PDF,
  outputRoot = OUTPUT_ROOT,
): Promise<string[]> {
  await mkdir(outputRoot, { recursive: true });
  const pages = await renderPdfPages(resolveProjectPath(pdfPath), join(REVIEW_DIR, "cover-pdf-pages"), PPTX_RENDER_DPI, "cover-page");
  if (pages.length < COVER_PAGE_ORDER.length) {
    throw new VerificationError(`cover canonical PDF has ${pages.length} pages; expected ${COVER_PAGE_ORDER.length}`);
  }
  const written: string[] = [];
  const front = join(outputRoot, COVER_TEMPLATE_IMAGE_NAMES.Cover);
  const back = join(outputRoot, COVER_TEMPLATE_IMAGE_NAMES.Cover_Back);
  await copyFile(pages[0]!, front);
  await copyFile(pages[1]!, back);
  written.push(front, back);
  const spinePage = await loadCanvasFromFile(pages[2]!);
  const [x0, y0, x1, y1] = nonWhiteCanvasBounds(spinePage);
  const spine = createCanvas(x1 - x0, y1 - y0);
  spine.getContext("2d").drawImage(spinePage, x0, y0, x1 - x0, y1 - y0, 0, 0, x1 - x0, y1 - y0);
  const spinePath = join(outputRoot, COVER_TEMPLATE_IMAGE_NAMES.Cover_Spine);
  await saveCanvasPng(spine, spinePath);
  written.push(spinePath);
  return written;
}

export async function renderComicAssetsFromCanonicalPdf(
  pdfPath = DEFAULT_ZH_TW_CANONICAL_PDF,
  outputRoot = OUTPUT_ROOT,
): Promise<string[]> {
  await mkdir(outputRoot, { recursive: true });
  const pages = await renderPdfPages(resolveProjectPath(pdfPath), join(REVIEW_DIR, "comic-pdf-pages"), PPTX_RENDER_DPI, "comic-page");
  if (pages.length < PAGE_ORDER.length) {
    throw new VerificationError(`canonical PDF has ${pages.length} pages; expected ${PAGE_ORDER.length}`);
  }
  const written: string[] = [];
  for (let index = 0; index < PAGE_ORDER.length; index += 1) {
    const pageId = PAGE_ORDER[index]!;
    const destination = join(outputRoot, TEMPLATE_IMAGE_NAMES[pageId]);
    await copyFile(pages[index]!, destination);
    written.push(destination);
  }
  return written;
}

export async function renderCoverAssets(textDir = COVER_TEXT_SPEC_DIR, outputRoot = OUTPUT_ROOT): Promise<string[]> {
  if (existsSync(DEFAULT_ZH_TW_COVER_CANONICAL_PDF)) {
    const rendered = await renderCoverAssetsFromCanonicalPdf(DEFAULT_ZH_TW_COVER_CANONICAL_PDF, outputRoot);
    const frontSpec = JSON.parse(await readFile(join(textDir, "Cover.json"), "utf8")) as TextSpec;
    for (const name of COVER_EXTRA_IMAGE_NAMES) {
      rendered.push(await renderPage(frontSpec, join(outputRoot, name)));
    }
    return rendered;
  }
  await mkdir(outputRoot, { recursive: true });
  const rendered: string[] = [];
  for (const pageId of COVER_PAGE_ORDER) {
    const spec = JSON.parse(await readFile(join(textDir, `${pageId}.json`), "utf8")) as TextSpec;
    const outputName = String(spec.output_image ?? COVER_TEMPLATE_IMAGE_NAMES[pageId]);
    rendered.push(await renderPage(spec, join(outputRoot, outputName)));
  }
  const frontSpec = JSON.parse(await readFile(join(textDir, "Cover.json"), "utf8")) as TextSpec;
  for (const name of COVER_EXTRA_IMAGE_NAMES) {
    rendered.push(await renderPage(frontSpec, join(outputRoot, name)));
  }
  return rendered;
}

export async function copyCoverAssets(
  coverRoot = BOOK_COVER_ROOT,
  outputRoot = OUTPUT_ROOT,
): Promise<string[]> {
  await mkdir(outputRoot, { recursive: true });
  const written: string[] = [];
  for (const name of COVER_IMAGE_NAMES) {
    const destination = join(outputRoot, name);
    await copyFile(join(coverRoot, name), destination);
    written.push(destination);
  }
  return written;
}

export async function renderAll(textDir = TEXT_SPEC_DIR, outputRoot = OUTPUT_ROOT): Promise<string[]> {
  await mkdir(outputRoot, { recursive: true });
  const rendered: string[] = [];
  if (existsSync(DEFAULT_ZH_TW_CANONICAL_PDF)) {
    rendered.push(...(await renderComicAssetsFromCanonicalPdf(DEFAULT_ZH_TW_CANONICAL_PDF, outputRoot)));
    for (const pageId of PAGE_ORDER) {
      const spec = JSON.parse(await readFile(join(textDir, `${pageId}.json`), "utf8")) as TextSpec;
      rendered.push(await renderReviewOverlay(spec));
    }
  } else {
    for (const pageId of PAGE_ORDER) {
      const spec = JSON.parse(await readFile(join(textDir, `${pageId}.json`), "utf8")) as TextSpec;
      rendered.push(await renderPage(spec, join(outputRoot, TEMPLATE_IMAGE_NAMES[pageId])));
      rendered.push(await renderReviewOverlay(spec));
    }
  }
  rendered.push(...(await renderCoverAssets(COVER_TEXT_SPEC_DIR, outputRoot)));
  rendered.push(await renderFooter(join(outputRoot, "web_footer.png")));
  return rendered;
}

export async function buildWebAssets(outputRoot = OUTPUT_ROOT): Promise<string[]> {
  const webDir = join(outputRoot, "web");
  await mkdir(webDir, { recursive: true });
  const footerPath = join(outputRoot, "web_footer.png");
  if (!existsSync(footerPath)) await renderFooter(footerPath);
  const written: string[] = [];

  const overview = await loadCanvasFromFile(join(outputRoot, "Overview.png"));
  const footer = await loadCanvasFromFile(footerPath);
  const overviewWeb = appendVertical([overview, footer]);
  const overviewOut = join(webDir, "Overview.png");
  await saveCanvasPng(overviewWeb, overviewOut);
  written.push(overviewOut);

  for (let chapter = 1; chapter <= 6; chapter++) {
    const pageA = await loadCanvasFromFile(join(outputRoot, `Ch${chapter}_A.png`));
    const pageB = await loadCanvasFromFile(join(outputRoot, `Ch${chapter}_B.png`));
    const croppedB = cropCanvas(pageB, [0, 1100, 5749, 8000]);
    const composite = appendVertical([pageA, croppedB, footer]);
    const out = join(webDir, `Ch${chapter}.png`);
    await saveCanvasPng(composite, out);
    written.push(out);
  }
  return written;
}

export async function buildEditablePptx(outputPath = PPTX_OUTPUT_PATH): Promise<string> {
  return buildEditablePptxFromYaml({ yamlPath: ZH_TW_YAML_PATH, outputPath });
}

export async function buildEditablePptxFromYaml(options: {
  yamlPath: string;
  outputPath: string;
  pdfPageModel?: PdfPageModel;
  pageBackgrounds?: Partial<Record<PageId, string>>;
  textTransparency?: number;
}): Promise<string> {
  const { yamlPath, outputPath, pdfPageModel, pageBackgrounds, textTransparency } = options;
  const specs = await loadYamlTextSpecs(yamlPath);
  const defaultTargetSize = requiredSizeFromCompactYaml(specs[PAGE_ORDER[0]!]!.target_size, `zh-TW YAML ${PAGE_ORDER[0]!} target_size`);
  const pptx = new PptxGenJS();
  pptx.author = "Civic.AI Comics";
  pptx.company = "Civic.AI";
  pptx.subject = "Editable zh-TW comic pages";
  pptx.title = "Civic.AI Comics zh-TW editable";
  const slideWidthIn = pdfPageModel ? pdfPageModel.pageSizePt[0] / 72 : defaultTargetSize.width / PPTX_RENDER_DPI;
  const slideHeightIn = pdfPageModel ? pdfPageModel.pageSizePt[1] / 72 : defaultTargetSize.height / PPTX_RENDER_DPI;
  const coordinateSize = pdfPageModel ? pageModelRenderSize(pdfPageModel) : [defaultTargetSize.width, defaultTargetSize.height];
  pptx.defineLayout({ name: PPTX_LAYOUT_NAME, width: slideWidthIn, height: slideHeightIn });
  pptx.layout = PPTX_LAYOUT_NAME;

  const backgroundDir = join(REVIEW_DIR, "pptx-backgrounds");
  await mkdir(backgroundDir, { recursive: true });
  await mkdir(dirname(outputPath), { recursive: true });

  for (const pageId of PAGE_ORDER) {
    const spec = specs[pageId];
    const slide = pptx.addSlide();
    const backgroundPath = pageBackgrounds?.[pageId] ?? await renderPptxBackground(spec, join(backgroundDir, `${pageId}.png`));
    slide.background = { path: resolveProjectPath(backgroundPath) };

    for (const box of spec.boxes ?? []) {
      const rect = box.box?.map(Number);
      if (!rect || rect.length !== 4 || !rect.every(Number.isFinite)) continue;
      const [x0, y0, x1, y1] = rect;
      slide.addText(pptxTextValue(box), {
        objectName: String(box.id),
        x: pdfPageModel ? (x0! / coordinateSize[0]!) * slideWidthIn : pxToPptxInches(x0!),
        y: pdfPageModel ? (y0! / coordinateSize[1]!) * slideHeightIn : pxToPptxInches(y0!),
        w: pdfPageModel ? ((x1! - x0!) / coordinateSize[0]!) * slideWidthIn : pxToPptxInches(x1! - x0!),
        h: pdfPageModel ? ((y1! - y0!) / coordinateSize[1]!) * slideHeightIn : pxToPptxInches(y1! - y0!),
        fontFace: pptxFontFace(normalizeRole(String(box.font_role ?? box.ocr_type ?? "body"))),
        fontSize: pxToPptxPoints(Number(box.font_size ?? 100)),
        ...(textTransparency === undefined ? {} : { transparency: textTransparency }),
        color: pptxHexColor(box.color),
        align: box.align === "center" ? "center" : "left",
        valign: "middle",
        margin: 0,
        fit: "shrink",
        breakLine: false,
        fill: { color: "FFFFFF", transparency: 100 },
        line: { color: "FFFFFF", transparency: 100 },
      });
    }
  }
  await pptx.writeFile({ fileName: outputPath });
  return outputPath;
}

function appleScriptString(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/"/g, "\\\"");
}

async function exportPptxToPdfWithKeynote(pptxPath: string, pdfPath: string): Promise<void> {
  await mkdir(dirname(pdfPath), { recursive: true });
  const scriptPath = join(dirname(pdfPath), "export-pptx-to-pdf.applescript");
  await writeFile(
    scriptPath,
    [
      `tell application "Keynote Creator Studio"`,
      `  set openedDocument to open POSIX file "${appleScriptString(pptxPath)}"`,
      `  export openedDocument to POSIX file "${appleScriptString(pdfPath)}" as PDF`,
      `  close openedDocument saving no`,
      `end tell`,
      "",
    ].join("\n"),
  );
  runCommandText(["osascript", scriptPath]);
}

export function sortPdfRenderedPagePaths(paths: string[]): string[] {
  return [...paths].sort((left, right) => {
    const leftPage = Number(left.match(/-(\d+)\.png$/)?.[1] ?? Number.POSITIVE_INFINITY);
    const rightPage = Number(right.match(/-(\d+)\.png$/)?.[1] ?? Number.POSITIVE_INFINITY);
    return leftPage === rightPage ? left.localeCompare(right) : leftPage - rightPage;
  });
}

export async function renderPdfPages(pdfPath: string, outputDir: string, renderDpi: number, prefix: string): Promise<string[]> {
  await mkdir(outputDir, { recursive: true });
  for (const name of await readdir(outputDir)) {
    if (name.startsWith(`${prefix}-`) && name.endsWith(".png")) await rm(join(outputDir, name), { force: true });
  }
  runCommandText(["pdftoppm", "-r", String(renderDpi), "-png", pdfPath, join(outputDir, prefix)]);
  const paths = (await readdir(outputDir))
    .filter((name) => name.startsWith(`${prefix}-`) && name.endsWith(".png"))
    .map((name) => join(outputDir, name));
  return sortPdfRenderedPagePaths(paths);
}

export async function diffPdfOracleRenderedPageLists(
  canonicalPaths: string[],
  generatedPaths: string[],
  options: { renderDpi?: number } = {},
): Promise<PdfOracleRoundTripReport> {
  const canonicalPages = sortPdfRenderedPagePaths(canonicalPaths);
  const generatedPages = sortPdfRenderedPagePaths(generatedPaths);
  const canonicalFirstSize = canonicalPages[0] ? await imageSize(canonicalPages[0]) : undefined;
  const generatedFirstSize = generatedPages[0] ? await imageSize(generatedPages[0]) : undefined;
  const canonicalRenderSizePx: [number, number] = canonicalFirstSize ? [canonicalFirstSize.width, canonicalFirstSize.height] : [0, 0];
  if (canonicalPages.length !== generatedPages.length) {
    return {
      ok: false,
      canonicalRenderSizePx,
      renderer: { kind: "incomplete" },
      incompleteReason: `PDF page count mismatch before pixel diff: canonical ${canonicalPages.length}, generated ${generatedPages.length}`,
      pageRenderSizePx: canonicalRenderSizePx,
      generatedRenderSizePx: generatedFirstSize ? [generatedFirstSize.width, generatedFirstSize.height] as [number, number] : undefined,
      pageDiffs: [],
    };
  }
  const pageDiffs: PdfOraclePageDiffMetric[] = [];
  let allPagesMatched = true;
  for (let index = 0; index < canonicalPages.length; index += 1) {
    const comparison = await comparePdfOraclePages(canonicalPages[index]!, generatedPages[index]!, { renderDpi: options.renderDpi });
    allPagesMatched &&= comparison.ok;
    pageDiffs.push({
      pageId: PAGE_ORDER[index] ?? `page-${index + 1}`,
      diffPixels: comparison.diffPixels,
      mismatchRatio: comparison.changedRatio,
      dimensionMismatch: comparison.dimensionMismatch,
    });
  }
  return {
    ok: allPagesMatched,
    canonicalRenderSizePx,
    renderer: { kind: "pptx-pdf" },
    pageRenderSizePx: canonicalRenderSizePx,
    generatedRenderSizePx: generatedFirstSize ? [generatedFirstSize.width, generatedFirstSize.height] as [number, number] : undefined,
    pageDiffs,
  };
}

function scaleTargetBox(value: unknown, scaleX: number, scaleY: number): number[] | undefined {
  if (!Array.isArray(value)) return undefined;
  const scaled = value.map((entry, index) => {
    const numberValue = Number(entry);
    if (!Number.isFinite(numberValue)) return Number.NaN;
    return Math.round(numberValue * (index % 2 === 0 ? scaleX : scaleY));
  });
  return scaled.every(Number.isFinite) ? scaled : undefined;
}


function scaleTargetEntries(lines: string[], pageId: string, entries: { id?: string; box?: unknown; tight?: unknown }[] | undefined, scaleX: number, scaleY: number): void {
  if (!entries) return;
  for (const entry of entries) {
    if (!entry.id) continue;
    const box = scaleTargetBox(entry.box, scaleX, scaleY);
    const tight = scaleTargetBox(entry.tight, scaleX, scaleY);
    if (box) replaceYamlBoxScalar(lines, pageId, String(entry.id), "box", formatYamlNumberArray(box));
    if (tight) replaceYamlBoxScalar(lines, pageId, String(entry.id), "tight", formatYamlNumberArray(tight));
  }
}

function replaceYamlPageScalar(lines: string[], pageId: string, field: string, value: string): boolean {
  const range = findYamlPageLineRange(lines, pageId);
  const replacement = `    ${field}: ${value}`;
  for (let index = range.start + 1; index < range.end; index += 1) {
    if (!lines[index]?.startsWith(`    ${field}:`)) continue;
    if (lines[index] === replacement) return false;
    lines.splice(index, 1, replacement);
    return true;
  }
  const refreshedRange = findYamlPageLineRange(lines, pageId);
  let targetSizeLine: number | undefined;
  for (let index = refreshedRange.start + 1; index < refreshedRange.end; index += 1) {
    if (lines[index]?.startsWith("    target_size:")) {
      targetSizeLine = index;
      break;
    }
  }
  lines.splice((targetSizeLine ?? refreshedRange.start) + 1, 0, replacement);
  return true;
}

export async function ingestCanonicalPdfToYaml(options: {
  canonicalPdfPath: string;
  yamlPath: string;
  outputYamlPath: string;
  renderDpi?: number;
  dryRun?: boolean;
}): Promise<IngestCanonicalPdfToYamlReport> {
  const renderDpi = options.renderDpi ?? PPTX_RENDER_DPI;
  const dryRun = options.dryRun ?? false;
  const pdf = await inspectCanonicalPdf(options.canonicalPdfPath, { renderDpi });
  const yamlPath = resolveProjectPath(options.yamlPath);
  const outputYamlPath = resolveProjectPath(options.outputYamlPath);
  const yamlText = await readFile(yamlPath, "utf8");
  const trailingNewline = yamlText.endsWith("\n");
  const lines = yamlText.replace(/\n$/, "").split("\n");
  const root = Bun.YAML.parse(yamlText) as CompactYamlRoot;
  if (!root.pages) root.pages = {};
  let updatedPagesCount = 0;
  for (const pageId of PAGE_ORDER) {
    const page = root.pages[pageId];
    if (!page) continue;
    const oldTarget = requiredSizeFromCompactYaml(page.target_size, `zh-TW YAML ${pageId} target_size`);
    const scaleX = pdf.renderSizePx[0] / oldTarget.width;
    const scaleY = pdf.renderSizePx[1] / oldTarget.height;
    replaceYamlPageScalar(lines, pageId, "target_size", formatYamlNumberArray(pdf.renderSizePx));
    replaceYamlPageScalar(lines, pageId, "pdf_page_box", formatYamlNumberArray(pdf.pageSizePt));
    replaceYamlPageScalar(lines, pageId, "render_dpi", String(renderDpi));
    scaleTargetEntries(lines, pageId, page.boxes, scaleX, scaleY);
    scaleTargetEntries(lines, pageId, page.patches, scaleX, scaleY);
    scaleTargetEntries(lines, pageId, page.blocks, scaleX, scaleY);
    updatedPagesCount += 1;
  }
  if (!dryRun) {
    await mkdir(dirname(outputYamlPath), { recursive: true });
    await writeFile(outputYamlPath, `${lines.join("\n")}${trailingNewline ? "\n" : ""}`);
  }
  return {
    canonicalRenderSizePx: pdf.renderSizePx,
    updatedPagesCount,
    renderDpi,
    dryRun,
    outputYamlPath,
  };
}


export async function runPdfOracleRoundTrip(options: {
  canonicalPdfPath: string;
  yamlPath: string;
  outputDir: string;
  renderDpi?: number;
}): Promise<PdfOracleRoundTripReport> {
  const renderDpi = options.renderDpi ?? PPTX_RENDER_DPI;
  const report = await inspectCanonicalPdf(options.canonicalPdfPath, { renderDpi });
  const outputDir = resolveProjectPath(options.outputDir);
  await mkdir(outputDir, { recursive: true });
  const pptxPath = join(outputDir, "zh-TW-editable.pptx");
  const generatedPdfPath = join(outputDir, "zh-TW-editable.roundtrip.pdf");
  const canonicalPages = await renderPdfPages(resolveProjectPath(options.canonicalPdfPath), join(outputDir, "canonical-pages"), renderDpi, "page");
  if (canonicalPages.length !== PAGE_ORDER.length) {
    return {
      ok: false,
      canonicalRenderSizePx: report.renderSizePx,
      renderer: { kind: "incomplete" },
      incompleteReason: `canonical PDF page count mismatch: expected ${PAGE_ORDER.length}, got ${canonicalPages.length}`,
      pageRenderSizePx: report.renderSizePx,
      pageDiffs: [],
    };
  }
  const pageBackgrounds = Object.fromEntries(PAGE_ORDER.map((pageId, index) => [pageId, canonicalPages[index]!])) as Partial<Record<PageId, string>>;
  await buildEditablePptxFromYaml({
    yamlPath: options.yamlPath,
    outputPath: pptxPath,
    pdfPageModel: {
      pageSizePt: report.pageSizePt,
      renderDpi: report.renderDpi,
      renderSizePx: report.renderSizePx,
      embeddedImageSizePx: report.embeddedImageSizePx,
    },
    pageBackgrounds,
    textTransparency: 100,
  });

  try {
    await exportPptxToPdfWithKeynote(pptxPath, generatedPdfPath);
  } catch (error) {
    return {
      ok: false,
      canonicalRenderSizePx: report.renderSizePx,
      renderer: { kind: "incomplete" },
      incompleteReason: `PPTX-to-PDF export failed: ${error instanceof Error ? error.message : String(error)}`,
      pageRenderSizePx: report.renderSizePx,
    };
  }

  const generatedPages = await renderPdfPages(generatedPdfPath, join(outputDir, "generated-pages"), renderDpi, "page");
  const diffReport = await diffPdfOracleRenderedPageLists(canonicalPages, generatedPages, { renderDpi });
  return {
    ...diffReport,
    canonicalRenderSizePx: report.renderSizePx,
    renderer: diffReport.renderer.kind === "incomplete" ? diffReport.renderer : { kind: "pptx-pdf" },
    pageRenderSizePx: report.renderSizePx,
  };
}


export async function buildCoverEditablePptx(
  outputPath = COVER_PPTX_OUTPUT_PATH,
  yamlPath = ZH_TW_COVER_YAML_PATH,
): Promise<string> {
  const specs = await loadCoverTextSpecs(yamlPath);
  const pptx = new PptxGenJS();
  pptx.author = "Civic.AI Comics";
  pptx.company = "Civic.AI";
  pptx.subject = "Editable zh-TW book cover";
  pptx.title = "Civic.AI Comics zh-TW cover editable";
  pptx.defineLayout({ name: COVER_PPTX_LAYOUT_NAME, width: PPTX_SLIDE_WIDTH_IN, height: PPTX_SLIDE_HEIGHT_IN });
  pptx.layout = COVER_PPTX_LAYOUT_NAME;

  const backgroundDir = join(REVIEW_DIR, "cover-pptx-backgrounds");
  await mkdir(backgroundDir, { recursive: true });
  await mkdir(dirname(outputPath), { recursive: true });

  for (const pageId of COVER_PAGE_ORDER) {
    const spec = specs[pageId];
    const slide = pptx.addSlide();
    const backgroundPath = await renderPptxBackground(spec, join(backgroundDir, `${pageId}.png`));
    const target = spec.target_size?.map(Number) ?? [];
    const targetWidth = target[0] ?? PPTX_SLIDE_WIDTH_IN * PPTX_RENDER_DPI;
    const targetHeight = target[1] ?? PPTX_SLIDE_HEIGHT_IN * PPTX_RENDER_DPI;
    const xOffset = Math.max(0, (PPTX_SLIDE_WIDTH_IN - pxToPptxInches(targetWidth)) / 2);
    slide.addImage({
      path: backgroundPath,
      x: xOffset,
      y: 0,
      w: pxToPptxInches(targetWidth),
      h: pxToPptxInches(targetHeight),
    });

    for (const box of spec.boxes ?? []) {
      const rect = box.box?.map(Number);
      if (!rect || rect.length !== 4 || !rect.every(Number.isFinite)) continue;
      const [x0, y0, x1, y1] = rect;
      slide.addText(pptxTextValue(box), {
        objectName: String(box.id),
        x: xOffset + pxToPptxInches(x0!),
        y: pxToPptxInches(y0!),
        w: pxToPptxInches(x1! - x0!),
        h: pxToPptxInches(y1! - y0!),
        fontFace: pptxFontFace(normalizeRole(String(box.font_role ?? box.ocr_type ?? "body"))),
        fontSize: pxToPptxPoints(Number(box.font_size ?? 100)),
        color: pptxHexColor(box.color),
        align: box.align === "center" ? "center" : "left",
        valign: "middle",
        margin: 0,
        fit: "shrink",
        breakLine: false,
        fill: { color: "FFFFFF", transparency: 100 },
        line: { color: "FFFFFF", transparency: 100 },
      });
    }
  }
  await pptx.writeFile({ fileName: outputPath });
  return outputPath;
}

function pxToPptxInches(px: number): number {
  return px / PPTX_RENDER_DPI;
}

function pxToPptxPoints(px: number): number {
  return (px * 72) / PPTX_RENDER_DPI;
}

function pptxHexColor(color: unknown): string {
  const hex = typeof color === "string" ? color.replace(/^#/, "") : "";
  return /^[0-9a-fA-F]{6}$/.test(hex) ? hex.toUpperCase() : "222222";
}

function pptxFontFace(role: FontRole): string {
  return role === "title" ? "jf蘭陽黑體 1.0" : "jf蘭陽明體2.0";
}

type PptxRunTextValue = { text: string; options?: { bold?: boolean; color?: string } };

function pptxTextValue(box: BoxSpec): string | PptxRunTextValue[] {
  const text = String(box.text ?? "");
  const emphasis = emphasisStyleMap(text, box.emphasis);
  if (!emphasis) return text;
  const runs: { text: string; bold: boolean; color?: string }[] = [];
  let codeOffset = 0;
  for (const char of Array.from(text)) {
    const bold = Boolean(emphasis.bold[codeOffset]);
    const color = emphasis.colors[codeOffset];
    const last = runs[runs.length - 1];
    if (last && last.bold === bold && last.color === color) {
      last.text += char;
    } else {
      runs.push({ text: char, bold, color });
    }
    codeOffset += char.length;
  }
  return runs.map((run) => {
    const options: { bold?: boolean; color?: string } = {};
    if (run.bold) options.bold = true;
    if (run.color) options.color = pptxHexColor(run.color);
    return Object.keys(options).length > 0 ? { text: run.text, options } : { text: run.text };
  });
}

export async function parseEditablePptxToYaml(
  pptxPath = PPTX_OUTPUT_PATH,
  yamlPath = ZH_TW_YAML_PATH,
): Promise<ParseEditablePptxToYamlResult> {
  const resolvedPptxPath = resolveProjectPath(pptxPath);
  const resolvedYamlPath = resolveProjectPath(yamlPath);
  const yamlText = await readFile(resolvedYamlPath, "utf8");
  const root = Bun.YAML.parse(yamlText) as CompactYamlRoot;
  if (!root.pages || typeof root.pages !== "object") throw new VerificationError("zh-TW YAML has no pages map");

  const expectedIdsByPage = new Map<PageId, Set<string>>();
  for (const pageId of PAGE_ORDER) {
    const boxes = root.pages[pageId]?.boxes ?? [];
    expectedIdsByPage.set(pageId, new Set(boxes.map((box) => String(box.id ?? "")).filter(Boolean)));
  }

  const zip = await JSZip.loadAsync(await readFile(resolvedPptxPath));
  const updates = new Map<string, ParsedPptxShape>();
  const unknownShapes: string[] = [];
  const missingShapes: string[] = [];

  const fallbackToLegacy = true;
  /** Slides parsed via legacy heuristics skip unknown/missing shape checks for that page only. */
  const legacySlidePageIds = new Set<PageId>();

  for (let slideIndex = 0; slideIndex < PAGE_ORDER.length; slideIndex += 1) {
    const pageId = PAGE_ORDER[slideIndex]!;
    const slidePath = `ppt/slides/slide${slideIndex + 1}.xml`;
    const slideEntry = zip.file(slidePath);
    if (!slideEntry) throw new VerificationError(`PPTX missing slide XML: ${slidePath}`);
    const slideXml = await slideEntry.async("string");
    const doc = PPTX_XML_PARSER.parse(slideXml) as OrderedXmlNode[];
    const parsedShapes = parsePptxSlideShapes(slideXml);
    const expectedIds = expectedIdsByPage.get(pageId) ?? new Set<string>();
    const seenIds = new Set<string>();
    const namedShapeCount = parsedShapes.filter((shape) => expectedIds.has(shape.name)).length;
    // Legacy when the slide has shapes but none named like YAML box ids (any expected id → strict path, including orphan rejection).
    const pureLegacySlide = parsedShapes.length > 0 && namedShapeCount === 0;

    if (pureLegacySlide && fallbackToLegacy) {
      if (legacySlidePageIds.size === 0) {
        console.log(`[Parseback] Fallback to legacy group transform resolution & heuristic matcher.`);
      }
      legacySlidePageIds.add(pageId);
      const spTree = findNodeWithKey(doc, "p:spTree");
      const resolvedDeskShapes: ParsedPptxShape[] = [];
      traverseShapes(spTree ? (spTree["p:spTree"] as OrderedXmlNode[]) : undefined, null, resolvedDeskShapes);
      const filteredDesk = resolvedDeskShapes.filter((s) => s.text);
      const boxesForPage = (root.pages[pageId]?.boxes ?? []).map((box, index) => boxFromCompactYamlBox(pageId, box, index));
      const slideMatches = matchLegacyShapes(filteredDesk, boxesForPage);
      for (const [boxId, shape] of slideMatches) {
        if (pageId === "Ch4_A" && boxId === "b006_body") {
          console.log(`[Parseback] Skipping update for split table block: ${pageId}/${boxId}`);
          continue;
        }
        updates.set(`${pageId}/${boxId}`, shape);
      }
    } else {
      for (const shape of parsedShapes) {
        if (!expectedIds.has(shape.name)) {
          unknownShapes.push(`${pageId}/${shape.name}`);
          continue;
        }
        updates.set(`${pageId}/${shape.name}`, shape);
        seenIds.add(shape.name);
      }
      for (const expectedId of expectedIds) {
        if (!seenIds.has(expectedId)) missingShapes.push(`${pageId}/${expectedId}`);
      }
    }
  }

  const pageIdFromParsebackEntry = (entry: string): PageId => {
    const slash = entry.indexOf("/");
    return (slash >= 0 ? entry.slice(0, slash) : entry) as PageId;
  };
  const unknownOnStrictSlides = unknownShapes.filter((entry) => !legacySlidePageIds.has(pageIdFromParsebackEntry(entry)));
  const missingOnStrictSlides = missingShapes.filter((entry) => !legacySlidePageIds.has(pageIdFromParsebackEntry(entry)));
  if (unknownOnStrictSlides.length > 0) {
    throw new VerificationError(`PPTX contains shapes with no YAML box: ${unknownOnStrictSlides.join(", ")}`);
  }
  if (missingOnStrictSlides.length > 0) {
    throw new VerificationError(`PPTX is missing expected YAML-backed text boxes: ${missingOnStrictSlides.join(", ")}`);
  }

  const rewritten = rewriteCompactYamlFromPptxUpdates(yamlText, root, updates);
  if (rewritten.updated > 0) await writeFile(resolvedYamlPath, rewritten.yamlText);
  return { yaml_path: resolvedYamlPath, updated: rewritten.updated, unknown_shapes: [] };
}

export async function parseCoverEditablePptxToYaml(
  pptxPath = COVER_PPTX_OUTPUT_PATH,
  yamlPath = ZH_TW_COVER_YAML_PATH,
): Promise<ParseEditablePptxToYamlResult> {
  const resolvedPptxPath = resolveProjectPath(pptxPath);
  const resolvedYamlPath = resolveProjectPath(yamlPath);
  const yamlText = await readFile(resolvedYamlPath, "utf8");
  const root = Bun.YAML.parse(yamlText) as CompactCoverYamlRoot;
  if (!root.pages || typeof root.pages !== "object") throw new VerificationError("zh-TW cover YAML has no pages map");

  const expectedIdsByPage = new Map<CoverPageId, Set<string>>();
  for (const pageId of COVER_PAGE_ORDER) {
    const boxes = root.pages[pageId]?.boxes ?? [];
    expectedIdsByPage.set(pageId, new Set(boxes.map((box) => String(box.id ?? "")).filter(Boolean)));
  }

  const zip = await JSZip.loadAsync(await readFile(resolvedPptxPath));
  const updates = new Map<string, ParsedPptxShape>();
  const unknownShapes: string[] = [];
  const missingShapes: string[] = [];
  const legacySlidePageIds = new Set<CoverPageId>();
  for (let slideIndex = 0; slideIndex < COVER_PAGE_ORDER.length; slideIndex += 1) {
    const pageId = COVER_PAGE_ORDER[slideIndex]!;
    const page = root.pages[pageId];
    if (!page) throw new VerificationError(`zh-TW cover YAML missing page: ${pageId}`);
    const slidePath = `ppt/slides/slide${slideIndex + 1}.xml`;
    const slideEntry = zip.file(slidePath);
    if (!slideEntry) throw new VerificationError(`PPTX missing slide XML: ${slidePath}`);
    const slideXml = await slideEntry.async("string");
    const doc = PPTX_XML_PARSER.parse(slideXml) as OrderedXmlNode[];
    const parsedShapes = parsePptxSlideShapes(slideXml).filter((shape) => shape.text);
    const xOffset = coverSlideXOffsetPx(page);
    const expectedIds = expectedIdsByPage.get(pageId) ?? new Set<string>();
    const seenIds = new Set<string>();
    const namedShapeCount = parsedShapes.filter((shape) => expectedIds.has(shape.name)).length;

    if (parsedShapes.length > 0 && namedShapeCount === 0) {
      legacySlidePageIds.add(pageId);
      const spTree = findNodeWithKey(doc, "p:spTree");
      const resolvedDeskShapes: ParsedPptxShape[] = [];
      traverseShapes(spTree ? (spTree["p:spTree"] as OrderedXmlNode[]) : undefined, null, resolvedDeskShapes);
      const shiftedBoxes = (page.boxes ?? [])
        .map((box, index) => shiftBoxForCoverPptxSlide(boxFromCompactCoverYamlBox(pageId, box, index), xOffset));
      const slideMatches = matchLegacyShapes(resolvedDeskShapes.filter((shape) => shape.text), shiftedBoxes);
      for (const [boxId, shape] of slideMatches) {
        updates.set(`${pageId}/${boxId}`, normalizeCoverPptxShapeForYaml(shape, xOffset));
      }
      continue;
    }

    const renamedShapes: ParsedPptxShape[] = [];
    for (const shape of parsedShapes) {
      if (!expectedIds.has(shape.name)) {
        renamedShapes.push(shape);
        continue;
      }
      updates.set(`${pageId}/${shape.name}`, normalizeCoverPptxShapeForYaml(shape, xOffset));
      seenIds.add(shape.name);
    }
    const missingIds = [...expectedIds].filter((expectedId) => !seenIds.has(expectedId));
    const renamedMatches = matchLegacyShapes(
      renamedShapes,
      (page.boxes ?? [])
        .filter((box) => missingIds.includes(String(box.id ?? "")))
        .map((box, index) => shiftBoxForCoverPptxSlide(boxFromCompactCoverYamlBox(pageId, box, index), xOffset)),
    );
    const matchedRenamedShapes = new Set<ParsedPptxShape>();
    for (const [boxId, shape] of renamedMatches) {
      updates.set(`${pageId}/${boxId}`, normalizeCoverPptxShapeForYaml(shape, xOffset));
      seenIds.add(boxId);
      matchedRenamedShapes.add(shape);
    }
    for (const shape of renamedShapes) {
      if (!matchedRenamedShapes.has(shape)) unknownShapes.push(`${pageId}/${shape.name}`);
    }
    for (const expectedId of expectedIds) {
      if (!seenIds.has(expectedId)) missingShapes.push(`${pageId}/${expectedId}`);
    }
  }

  const pageIdFromParsebackEntry = (entry: string): CoverPageId => {
    const slash = entry.indexOf("/");
    return (slash >= 0 ? entry.slice(0, slash) : entry) as CoverPageId;
  };
  const unknownOnStrictSlides = unknownShapes.filter((entry) => !legacySlidePageIds.has(pageIdFromParsebackEntry(entry)));
  const missingOnStrictSlides = missingShapes.filter((entry) => !legacySlidePageIds.has(pageIdFromParsebackEntry(entry)));
  if (unknownOnStrictSlides.length > 0) {
    throw new VerificationError(`PPTX contains shapes with no cover YAML box: ${unknownOnStrictSlides.join(", ")}`);
  }
  if (missingOnStrictSlides.length > 0) {
    throw new VerificationError(`PPTX is missing expected cover YAML-backed text boxes: ${missingOnStrictSlides.join(", ")}`);
  }

  const rewritten = rewriteCompactCoverYamlFromPptxUpdates(yamlText, root, updates);
  if (rewritten.updated > 0) await writeFile(resolvedYamlPath, rewritten.yamlText);
  return { yaml_path: resolvedYamlPath, updated: rewritten.updated, unknown_shapes: [] };
}

function coverSlideXOffsetPx(page: CompactCoverYamlPage): number {
  const targetSize = requiredSizeFromCompactYaml(page.target_size, "zh-TW cover YAML target_size");
  return Math.max(0, (PPTX_SLIDE_WIDTH_IN * PPTX_RENDER_DPI - targetSize.width) / 2);
}

function shiftBoxForCoverPptxSlide(box: BoxSpec, xOffset: number): BoxSpec {
  const [x0, y0, x1, y1] = (box.box ?? []).map(Number);
  return {
    ...box,
    box: [x0! + xOffset, y0!, x1! + xOffset, y1!],
  };
}

function normalizeCoverPptxShapeForYaml(shape: ParsedPptxShape, xOffset: number): ParsedPptxShape {
  const [x0, y0, x1, y1] = shape.box.map(Number);
  return {
    ...shape,
    box: [x0! - xOffset, y0!, x1! - xOffset, y1!],
  };
}

function findNodeWithKey(nodes: OrderedXmlNode[] | undefined, key: string): OrderedXmlNode | null {
  if (!nodes || !Array.isArray(nodes)) return null;
  for (const n of nodes) {
    if (Object.keys(n).includes(key)) return n;
    const childrenKey = Object.keys(n).find((k) => Array.isArray(n[k]));
    if (childrenKey) {
      const found = findNodeWithKey(n[childrenKey] as OrderedXmlNode[], key);
      if (found) return found;
    }
  }
  return null;
}

function parsePptxSlideShapes(slideXml: string): ParsedPptxShape[] {
  const doc = PPTX_XML_PARSER.parse(slideXml) as OrderedXmlNode[];
  const shapeNodes = collectXmlNodes(doc, "p:sp");
  const shapes: ParsedPptxShape[] = [];
  for (const shapeNode of shapeNodes) {
    const parsed = parsePptxShape(shapeNode);
    if (parsed) shapes.push(parsed);
  }
  return shapes;
}

function parsePptxShape(shapeNode: OrderedXmlNode): ParsedPptxShape | null {
  const children = xmlElementChildren(shapeNode, "p:sp");
  const name = xmlAttr(findFirstXmlNode(children, "p:cNvPr"), "name");
  if (!name) return null;
  const offset = findFirstXmlNode(children, "a:off");
  const extent = findFirstXmlNode(children, "a:ext");
  const x = emuToRenderPx(xmlNumberAttr(offset, "x"));
  const y = emuToRenderPx(xmlNumberAttr(offset, "y"));
  const width = emuToRenderPx(xmlNumberAttr(extent, "cx"));
  const height = emuToRenderPx(xmlNumberAttr(extent, "cy"));
  const firstRunProps = findFirstXmlNode(children, "a:rPr");
  const sz = xmlNumberAttr(firstRunProps, "sz");
  const paragraphProps = findFirstXmlNode(children, "a:pPr");
  const text = parsePptxTextBody(findFirstXmlNode(children, "p:txBody"));
  return {
    name,
    text: text.text,
    box: [x, y, x + width, y + height],
    font: sz === undefined ? undefined : Math.round((sz / 100) * (PPTX_RENDER_DPI / 72)),
    align: xmlAttr(paragraphProps, "algn") === "ctr" ? "center" : "left",
    emphasis: text.emphasis,
  };
}

function parsePptxTextBody(txBodyNode: OrderedXmlNode | undefined): { text: string; emphasis: EmphasisSpec[] } {
  if (!txBodyNode) return { text: "", emphasis: [] };
  const children = xmlElementChildren(txBodyNode, "p:txBody");
  let text = "";
  const emphasis: EmphasisSpec[] = [];
  const append = (value: string, style: { bold: boolean; color?: string }) => {
    if (!value) return;
    const start = text.length;
    text += value;
    if (style.bold || style.color) {
      const entry: EmphasisSpec = { start, end: text.length };
      if (style.bold) entry.bold = true;
      if (style.color) entry.color = style.color;
      emphasis.push(entry);
    }
  };
  const paragraphs = children.filter((node) => Object.prototype.hasOwnProperty.call(node, "a:p"));
  for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex += 1) {
    if (paragraphIndex > 0) append("\n", { bold: false });
    let atLineStart = true;
    for (const node of xmlElementChildren(paragraphs[paragraphIndex]!, "a:p")) {
      if (Object.prototype.hasOwnProperty.call(node, "a:r")) {
        const runChildren = xmlElementChildren(node, "a:r");
        const rawText = collectXmlText(runChildren, "a:t");
        const runText = atLineStart ? rawText.replace(/^[ \t]+/, "") : rawText;
        append(runText, pptxRunStyle(runChildren));
        if (runText.length > 0) atLineStart = false;
      } else if (Object.prototype.hasOwnProperty.call(node, "a:br")) {
        append("\n", { bold: false });
        atLineStart = true;
      }
    }
  }
  return { text, emphasis: mergeEmphasisRanges(emphasis) };
}

function pptxRunStyle(runChildren: OrderedXmlNode[]): { bold: boolean; color?: string } {
  return {
    bold: isPptxBoldRun(runChildren),
    color: pptxRunColor(runChildren),
  };
}

function isPptxBoldRun(runChildren: OrderedXmlNode[]): boolean {
  const runProps = findFirstXmlNode(runChildren, "a:rPr");
  const bold = xmlAttr(runProps, "b");
  if (bold === "1" || bold === "true") return true;
  return ["a:latin", "a:ea", "a:cs"].some((tag) =>
    collectXmlNodes(runChildren, tag).some((node) => isBoldTypeface(xmlAttr(node, "typeface"))),
  );
}

function isBoldTypeface(typeface: string | undefined): boolean {
  if (!typeface) return false;
  return /(^|[\s_-])bold($|[\s_-])/i.test(typeface) || /(^|[\s_-])w[7-9]($|[\s_-])/i.test(typeface);
}

function pptxRunColor(runChildren: OrderedXmlNode[]): string | undefined {
  return hexColorOverride(xmlAttr(findFirstXmlNode(runChildren, "a:srgbClr"), "val"));
}

function mergeEmphasisRanges(ranges: EmphasisSpec[]): EmphasisSpec[] {
  const merged: EmphasisSpec[] = [];
  for (const range of ranges) {
    if (range.end === undefined || range.start === undefined || range.end <= range.start) continue;
    const last = merged[merged.length - 1];
    if (last && last.end === range.start && last.bold === range.bold && last.color === range.color) {
      last.end = range.end;
    } else {
      merged.push({ ...range });
    }
  }
  return merged;
}


function mergeRanges(ranges: { start: number; end: number }[]): { start: number; end: number }[] {
  const merged: { start: number; end: number }[] = [];
  for (const range of ranges) {
    if (range.end <= range.start) continue;
    const last = merged[merged.length - 1];
    if (last && range.start <= last.end) {
      last.end = Math.max(last.end, range.end);
    } else {
      merged.push({ ...range });
    }
  }
  return merged;
}

function rewriteCompactYamlFromPptxUpdates(
  yamlText: string,
  root: CompactYamlRoot,
  updates: Map<string, ParsedPptxShape>,
): { yamlText: string; updated: number } {
  const trailingNewline = yamlText.endsWith("\n");
  const lines = yamlText.replace(/\n$/, "").split("\n");
  let updated = 0;
  for (const pageId of PAGE_ORDER) {
    const page = root.pages?.[pageId];
    if (!page) continue;
    const mdRange = tryFindYamlPageMdLineRange(lines, pageId);
    let mdText = mdRange ? parseYamlPageMdFromLines(lines, mdRange) : undefined;
    let mdChanged = false;
    for (const box of page.boxes ?? []) {
      const boxId = String(box.id ?? "");
      const shape = updates.get(`${pageId}/${boxId}`);
      if (!boxId || !shape) continue;
      const shapeEmphasis = removeDefaultColorEmphasis(shape.text, shape.emphasis, box.color);
      let changed = false;
      if (!sameNumberArray(box.box, shape.box)) changed = replaceYamlBoxScalar(lines, pageId, boxId, "box", formatYamlNumberArray(shape.box)) || changed;
      if (shape.font !== undefined && Number(box.font) !== shape.font) {
        changed = replaceYamlBoxScalar(lines, pageId, boxId, "font", String(shape.font)) || changed;
      }
      if ((box.align ?? "left") !== shape.align) changed = replaceYamlBoxScalar(lines, pageId, boxId, "align", shape.align) || changed;
      const currentText = String(box.text ?? "");
      const currentComparableText = stripHardBreaks(currentText.replace(/[ \t]+$/gm, ""));
      const shapeComparableText = stripHardBreaks(shape.text.replace(/[ \t]+$/gm, ""));
      const textContentChanged = currentComparableText !== shapeComparableText;
      if (textContentChanged) changed = replaceYamlBoxText(lines, pageId, boxId, shape.text) || changed;
      const emphasisMatches = textContentChanged
        ? sameResolvedEmphasisStyles(shape.text, box.emphasis, shapeEmphasis, box.color)
        : sameResolvedEmphasisStylesAcrossTexts(currentText, box.emphasis, shape.text, shapeEmphasis, box.color);
      if (textContentChanged || !emphasisMatches) {
        if (shapeEmphasis.length > 0 || !box.emphasis) {
          changed = replaceYamlBoxEmphasis(lines, pageId, boxId, shapeEmphasis) || changed;
        }
      }
      if (mdText !== undefined) {
        let nextMd = textContentChanged
          ? replaceNormalizedBoxTextInPageMd(mdText, currentText, shape.text)
          : syncNormalizedBoxTextInPageMd(mdText, currentText);
        nextMd = syncNormalizedBoxTextInPageMd(nextMd, shape.text);
        if (nextMd !== mdText) {
          mdText = nextMd;
          mdChanged = true;
        }
      }
      if (changed) updated += 1;
    }
    if (mdChanged && mdText !== undefined) {
      replaceYamlPageMdLines(lines, pageId, mdText);
      updated += 1;
    }
  }
  return { yamlText: `${lines.join("\n")}${trailingNewline ? "\n" : ""}`, updated };
}

function rewriteCompactCoverYamlFromPptxUpdates(
  yamlText: string,
  root: CompactCoverYamlRoot,
  updates: Map<string, ParsedPptxShape>,
): { yamlText: string; updated: number } {
  const trailingNewline = yamlText.endsWith("\n");
  const lines = yamlText.replace(/\n$/, "").split("\n");
  let updated = 0;
  for (const pageId of COVER_PAGE_ORDER) {
    const page = root.pages?.[pageId];
    if (!page) continue;
    for (const box of page.boxes ?? []) {
      const boxId = String(box.id ?? "");
      const shape = updates.get(`${pageId}/${boxId}`);
      if (!boxId || !shape) continue;
      const shapeEmphasis = removeDefaultColorEmphasis(shape.text, shape.emphasis, box.color);
      let changed = false;
      if (!sameNumberArray(box.box, shape.box)) changed = replaceYamlBoxScalar(lines, pageId, boxId, "box", formatYamlNumberArray(shape.box)) || changed;
      if (shape.font !== undefined && Number(box.font) !== shape.font) {
        changed = replaceYamlBoxScalar(lines, pageId, boxId, "font", String(shape.font)) || changed;
      }
      if (shape.align === "center" && (box.align ?? "left") !== "center") changed = replaceYamlBoxScalar(lines, pageId, boxId, "align", "center") || changed;
      const currentText = String(box.text ?? "");
      const textContentChanged = stripHardBreaks(currentText.replace(/[ \t]+$/gm, "")) !== stripHardBreaks(shape.text.replace(/[ \t]+$/gm, ""));
      if (textContentChanged) changed = replaceYamlBoxText(lines, pageId, boxId, shape.text) || changed;
      const emphasisMatches = textContentChanged
        ? sameResolvedEmphasisStyles(shape.text, box.emphasis, shapeEmphasis, box.color)
        : sameResolvedEmphasisStylesAcrossTexts(currentText, box.emphasis, shape.text, shapeEmphasis, box.color);
      if ((textContentChanged || !emphasisMatches) && (shapeEmphasis.length > 0 || !box.emphasis)) {
        changed = replaceYamlBoxEmphasis(lines, pageId, boxId, shapeEmphasis) || changed;
      }
      if (changed) updated += 1;
    }
  }
  return { yamlText: `${lines.join("\n")}${trailingNewline ? "\n" : ""}`, updated };
}


function replaceYamlBoxScalar(lines: string[], pageId: string, boxId: string, field: string, value: string): boolean {
  const range = findYamlBoxLineRange(lines, pageId, boxId);
  const line = findYamlFieldLine(lines, range, field);
  const replacement = `        ${field}: ${value}`;
  if (line !== undefined) {
    if (lines[line] === replacement) return false;
    lines.splice(line, 1, replacement);
    return true;
  }
  const insertAt = findYamlFieldLine(lines, range, "text") ?? range.end;
  lines.splice(insertAt, 0, replacement);
  return true;
}

function replaceYamlBoxText(lines: string[], pageId: string, boxId: string, text: string): boolean {
  const range = findYamlBoxLineRange(lines, pageId, boxId);
  const line = findYamlFieldLine(lines, range, "text");
  if (line === undefined) throw new VerificationError(`YAML ${pageId}/${boxId} has no text field`);
  let end = line + 1;
  while (end < range.end && lines[end]?.startsWith("          ")) end += 1;
  const replacement = ["        text: |-", ...text.split("\n").map((part) => `          ${part}`)];
  const before = lines.slice(line, end);
  if (before.length === replacement.length && before.every((value, index) => value === replacement[index])) return false;
  lines.splice(line, end - line, ...replacement);
  return true;
}

function replaceYamlBoxEmphasis(lines: string[], pageId: string, boxId: string, emphasis: EmphasisSpec[]): boolean {
  const range = findYamlBoxLineRange(lines, pageId, boxId);
  const line = findYamlFieldLine(lines, range, "emphasis");
  if (emphasis.length === 0) {
    if (line === undefined) return false;
    lines.splice(line, 1);
    return true;
  }
  const replacement = `        emphasis: ${formatYamlEmphasis(emphasis)}`;
  if (line !== undefined) {
    if (lines[line] === replacement) return false;
    lines.splice(line, 1, replacement);
    return true;
  }
  const insertAt = findYamlFieldLine(lines, range, "text") ?? range.end;
  lines.splice(insertAt, 0, replacement);
  return true;
}

function findYamlPageLineRange(lines: string[], pageId: string): { start: number; end: number } {
  const pageStart = lines.findIndex((line) => line === `  ${pageId}:`);
  if (pageStart < 0) throw new VerificationError(`YAML missing page ${pageId}`);
  let pageEnd = lines.length;
  for (let index = pageStart + 1; index < lines.length; index += 1) {
    if (/^  [A-Za-z0-9_]+:$/.test(lines[index]!)) {
      pageEnd = index;
      break;
    }
  }
  return { start: pageStart, end: pageEnd };
}

function tryFindYamlPageMdLineRange(lines: string[], pageId: PageId): { start: number; end: number } | undefined {
  const pageRange = findYamlPageLineRange(lines, pageId);
  for (let index = pageRange.start + 1; index < pageRange.end; index += 1) {
    const line = lines[index];
    if (!line?.startsWith("    md:")) continue;
    if (!line.startsWith("    md: |-")) throw new VerificationError(`YAML ${pageId} md is not a multiline block`);
    let end = index + 1;
    while (end < pageRange.end && (lines[end] === "" || lines[end]?.startsWith("      "))) end += 1;
    return { start: index, end };
  }
  return undefined;
}

export function findYamlPageMdLineRange(lines: string[], pageId: PageId): { start: number; end: number } {
  const pageRange = findYamlPageLineRange(lines, pageId);
  let mdHeader: number | undefined;
  for (let index = pageRange.start + 1; index < pageRange.end; index += 1) {
    if (lines[index]?.startsWith("    md:")) {
      mdHeader = index;
      break;
    }
  }
  if (mdHeader === undefined) throw new VerificationError(`YAML ${pageId} has no md field`);
  const header = lines[mdHeader]!;
  if (!header.startsWith("    md: |-")) throw new VerificationError(`YAML ${pageId} md is not a multiline block`);
  let end = mdHeader + 1;
  while (end < pageRange.end && (lines[end] === "" || lines[end]?.startsWith("      "))) end += 1;
  return { start: mdHeader, end };
}

function parseYamlPageMdFromLines(lines: string[], mdRange: { start: number; end: number }): string {
  const body = lines.slice(mdRange.start + 1, mdRange.end);
  return body.map((line) => (line.startsWith("      ") ? line.slice(6) : line)).join("\n");
}

function replaceYamlPageMdLines(lines: string[], pageId: PageId, mdText: string): void {
  const mdRange = findYamlPageMdLineRange(lines, pageId);
  const replacement = ["    md: |-", ...mdText.split("\n").map((part) => `      ${part}`)];
  lines.splice(mdRange.start, mdRange.end - mdRange.start, ...replacement);
}

function syncNormalizedBoxTextInPageMd(mdText: string, boxText: string): string {
  const boxNorm = stripHardBreaks(boxText);
  if (boxNorm.length < 12 || stripHardBreaks(mdText).includes(boxNorm)) return mdText;
  return (
    replaceLineAnchoredParagraphInPageMd(mdText, boxText, boxText) ??
    replacePrefixAnchoredParagraphInPageMd(mdText, boxNorm, boxNorm, boxText) ??
    replaceSimilarParagraphInPageMd(mdText, boxNorm, boxNorm, boxText) ??
    mdText
  );
}

function replaceNormalizedBoxTextInPageMd(mdText: string, oldText: string, newText: string): string {
  const oldNorm = stripHardBreaks(oldText);
  const newNorm = stripHardBreaks(newText);
  if (oldNorm === newNorm) return mdText;
  if (mdText.includes(oldText)) return mdText.replace(oldText, newText);
  const lineAnchored = replaceLineAnchoredParagraphInPageMd(mdText, oldText, newText);
  if (lineAnchored) return lineAnchored;
  const flatMd = stripHardBreaks(mdText);
  const flatIndex = flatMd.indexOf(oldNorm);
  if (flatIndex >= 0) {
    const [mdStart, mdEnd] = flatNormSpanToMdRange(mdText, flatIndex, oldNorm.length);
    if (mdStart >= 0 && mdEnd >= 0) return `${mdText.slice(0, mdStart)}${newText}${mdText.slice(mdEnd)}`;
  }
  return replacePrefixAnchoredParagraphInPageMd(mdText, oldNorm, newNorm, newText) ?? replaceSimilarParagraphInPageMd(mdText, oldNorm, newNorm, newText) ?? mdText;
}


function replaceLineAnchoredParagraphInPageMd(mdText: string, oldText: string, newText: string): string | null {
  const anchors = [oldText, newText]
    .flatMap((value) => value.split("\n"))
    .map((value) => value.trim())
    .filter((value, index, values) => value.length >= 8 && values.indexOf(value) === index);
  if (anchors.length === 0) return null;

  let match: { start: number; end: number } | null = null;
  let start = 0;
  while (start <= mdText.length) {
    let end = mdText.indexOf("\n\n", start);
    if (end < 0) end = mdText.length;
    const paragraph = mdText.slice(start, end);
    if (anchors.some((anchor) => paragraph.split("\n").some((line) => line.trim() === anchor))) {
      if (match) return null;
      match = { start, end };
    }
    if (end === mdText.length) break;
    start = end + 2;
  }

  return match ? `${mdText.slice(0, match.start)}${newText}${mdText.slice(match.end)}` : null;
}

function replacePrefixAnchoredParagraphInPageMd(mdText: string, oldNorm: string, newNorm: string, newText: string): string | null {
  const prefixes = [oldNorm, newNorm]
    .map((value) => value.slice(0, Math.min(16, value.length)))
    .filter((value, index, values) => value.length >= 12 && values.indexOf(value) === index);
  if (prefixes.length === 0) return null;

  let match: { start: number; end: number } | null = null;
  let start = 0;
  while (start <= mdText.length) {
    let end = mdText.indexOf("\n\n", start);
    if (end < 0) end = mdText.length;
    const candidateNorm = stripHardBreaks(mdText.slice(start, end));
    if (prefixes.some((prefix) => candidateNorm.startsWith(prefix))) {
      if (match) return null;
      match = { start, end };
    }
    if (end === mdText.length) break;
    start = end + 2;
  }

  return match ? `${mdText.slice(0, match.start)}${newText}${mdText.slice(match.end)}` : null;
}

function replaceSimilarParagraphInPageMd(mdText: string, oldNorm: string, newNorm: string, newText: string): string | null {
  const targetNorm = newNorm.length >= oldNorm.length ? newNorm : oldNorm;
  if (targetNorm.length < 12 || newNorm.length < 12) return null;

  const oldAnchor = oldNorm.slice(0, Math.min(8, oldNorm.length));
  const newAnchor = newNorm.slice(0, Math.min(8, newNorm.length));
  let best: { start: number; end: number; score: number } | null = null;
  let secondBestScore = 0;
  let start = 0;
  while (start <= mdText.length) {
    let end = mdText.indexOf("\n\n", start);
    if (end < 0) end = mdText.length;
    const block = mdText.slice(start, end);
    const candidateNorm = stripHardBreaks(block);
    const anchored = Boolean((oldAnchor && candidateNorm.includes(oldAnchor)) || (newAnchor && candidateNorm.includes(newAnchor)));
    if (anchored && candidateNorm.length >= Math.min(12, Math.floor(targetNorm.length * 0.3))) {
      const score = Math.max(lcsSimilarity(candidateNorm, oldNorm), lcsSimilarity(candidateNorm, newNorm));
      if (!best || score > best.score) {
        secondBestScore = best?.score ?? 0;
        best = { start, end, score };
      } else if (score > secondBestScore) {
        secondBestScore = score;
      }
    }
    if (end === mdText.length) break;
    start = end + 2;
  }

  if (!best || best.score < 0.6) return null;
  if (best.score < 0.9 && best.score - secondBestScore < 0.08) return null;
  return `${mdText.slice(0, best.start)}${newText}${mdText.slice(best.end)}`;
}

function lcsSimilarity(left: string, right: string): number {
  if (!left || !right) return 0;
  const previous = new Array<number>(right.length + 1).fill(0);
  const current = new Array<number>(right.length + 1).fill(0);
  for (let i = 1; i <= left.length; i += 1) {
    for (let j = 1; j <= right.length; j += 1) {
      current[j] = left[i - 1] === right[j - 1] ? previous[j - 1]! + 1 : Math.max(previous[j]!, current[j - 1]!);
    }
    previous.splice(0, previous.length, ...current);
    current.fill(0);
  }
  const lcs = previous[right.length]!;
  const shorter = Math.min(left.length, right.length);
  const longer = Math.max(left.length, right.length);
  return (lcs / shorter) * Math.sqrt(shorter / longer);
}


function flatNormSpanToMdRange(mdText: string, flatStart: number, flatLength: number): [number, number] {
  let flatIndex = 0;
  let mdStart = -1;
  let mdEnd = -1;
  for (let index = 0; index < mdText.length; index += 1) {
    if (mdText[index] === "\n") continue;
    if (flatIndex === flatStart) mdStart = index;
    if (flatIndex === flatStart + flatLength - 1) {
      mdEnd = index + 1;
      break;
    }
    flatIndex += 1;
  }
  return [mdStart, mdEnd];
}

function findYamlBoxLineRange(lines: string[], pageId: string, boxId: string): { start: number; end: number } {
  const pageRange = findYamlPageLineRange(lines, pageId);
  const pageStart = pageRange.start;
  const pageEnd = pageRange.end;
  const boxStart = lines.findIndex((line, index) => index > pageStart && index < pageEnd && line === `      - id: ${boxId}`);
  if (boxStart < 0) throw new VerificationError(`YAML missing box ${pageId}/${boxId}`);
  let boxEnd = pageEnd;
  for (let index = boxStart + 1; index < pageEnd; index += 1) {
    if (lines[index]?.startsWith("      - id: ")) {
      boxEnd = index;
      break;
    }
  }
  return { start: boxStart, end: boxEnd };
}

function findYamlFieldLine(lines: string[], range: { start: number; end: number }, field: string): number | undefined {
  for (let index = range.start + 1; index < range.end; index += 1) {
    if (lines[index]?.startsWith(`        ${field}:`)) return index;
  }
  return undefined;
}

function normalizeEmphasisForCompare(value: unknown): EmphasisSpec[] {
  return emphasisOverride(value) ?? [];
}

function formatYamlEmphasis(emphasis: EmphasisSpec[]): string {
  return `[${emphasis.map(formatYamlEmphasisEntry).join(",")}]`;
}

function formatYamlEmphasisEntry(entry: EmphasisSpec): string {
  const parts: string[] = [];
  if (entry.text !== undefined) parts.push(`"text":${JSON.stringify(entry.text)}`);
  if (entry.start !== undefined) parts.push(`"start":${Math.round(entry.start)}`);
  if (entry.end !== undefined) parts.push(`"end":${Math.round(entry.end)}`);
  if (entry.bold === true) parts.push(`"bold":true`);
  if (entry.color !== undefined) parts.push(`"color":${JSON.stringify(entry.color)}`);
  if (entry.occurrence !== undefined) parts.push(`"occurrence":${Math.round(entry.occurrence)}`);
  return `{${parts.join(",")}}`;
}

function formatYamlNumberArray(values: readonly number[]): string {
  return `[${values.map((value) => String(Math.round(value))).join(", ")}]`;
}

function sameNumberArray(value: unknown, expected: readonly number[]): boolean {
  return Array.isArray(value) && value.length === expected.length && value.every((entry, index) => Math.round(Number(entry)) === expected[index]);
}

function sameResolvedEmphasisStyles(text: string, current: unknown, next: EmphasisSpec[], defaultColor: unknown): boolean {
  return sameResolvedEmphasisStylesAcrossTexts(text, current, text, next, defaultColor);
}

function sameResolvedEmphasisStylesAcrossTexts(
  currentText: string,
  current: unknown,
  nextText: string,
  next: EmphasisSpec[],
  defaultColor: unknown,
): boolean {
  const currentStyle = comparableEmphasisStyle(currentText, normalizeEmphasisForCompare(current), defaultColor);
  const nextStyle = comparableEmphasisStyle(nextText, next, defaultColor);
  const currentEntries = comparableStyledCharacters(currentText, currentStyle);
  const nextEntries = comparableStyledCharacters(nextText, nextStyle);
  if (currentEntries.length !== nextEntries.length) return false;
  for (let index = 0; index < currentEntries.length; index += 1) {
    const currentEntry = currentEntries[index]!;
    const nextEntry = nextEntries[index]!;
    if (currentEntry.char !== nextEntry.char) return false;
    if (currentEntry.bold !== nextEntry.bold) return false;
    if (currentEntry.color !== nextEntry.color) return false;
  }
  return true;
}

function comparableStyledCharacters(text: string, style: InlineEmphasisMap): { char: string; bold: boolean; color?: string }[] {
  const entries: { char: string; bold: boolean; color?: string }[] = [];
  let codeOffset = 0;
  for (const char of Array.from(text)) {
    let lookahead = codeOffset + char.length;
    while (lookahead < text.length && /[ \t]/.test(text[lookahead]!)) lookahead += 1;
    const lineEndWhitespace = /[ \t]/.test(char) && (lookahead === text.length || text[lookahead] === "\n");
    if (char !== "\n" && !lineEndWhitespace) {
      entries.push({
        char,
        bold: Boolean(style.bold[codeOffset]),
        color: style.colors[codeOffset],
      });
    }
    codeOffset += char.length;
  }
  return entries;
}

function stripHardBreaks(text: string): string {
  return text.replace(/\n/g, "");
}


function comparableEmphasisStyle(text: string, emphasis: EmphasisSpec[], defaultColor: unknown): InlineEmphasisMap {
  const resolved = emphasisStyleMap(text, emphasis);
  const baseColor = hexColorOverride(defaultColor) ?? "#222222";
  const colors = new Array<string | undefined>(text.length);
  if (resolved) {
    for (let index = 0; index < text.length; index += 1) {
      const color = resolved.colors[index];
      if (color && color !== baseColor) colors[index] = color;
    }
  }
  return { bold: resolved?.bold ?? new Uint8Array(text.length), colors, hasBold: Boolean(resolved?.hasBold), hasColor: colors.some(Boolean) };
}
export function traverseShapes(
  nodes: OrderedXmlNode[] | undefined,
  currentTransform: GroupTransform | null,
  results: ParsedPptxShape[],
): void {
  if (!nodes || !Array.isArray(nodes)) return;
  for (const node of nodes) {
    const keys = Object.keys(node);
    if (keys.includes("p:grpSp")) {
      const grpChildren = node["p:grpSp"] as OrderedXmlNode[];
      const grpSpPr = grpChildren.find((c) => Object.keys(c).includes("p:grpSpPr"));
      const xfrm = grpSpPr ? findFirstXmlNode(grpSpPr["p:grpSpPr"] as OrderedXmlNode[], "a:xfrm") : null;
      let goffX: number | undefined;
      let goffY: number | undefined;
      let gextX: number | undefined;
      let gextY: number | undefined;
      let gchoffX: number | undefined;
      let gchoffY: number | undefined;
      let gchextX: number | undefined;
      let gchextY: number | undefined;
      if (xfrm) {
        const childrenXfrm = (xfrm["a:xfrm"] || []) as OrderedXmlNode[];
        const off = childrenXfrm.find((c) => Object.keys(c).includes("a:off"));
        const ext = childrenXfrm.find((c) => Object.keys(c).includes("a:ext"));
        const choff = childrenXfrm.find((c) => Object.keys(c).includes("a:chOff"));
        const chext = childrenXfrm.find((c) => Object.keys(c).includes("a:chExt"));
        goffX = xmlNumberAttr(off, "x");
        goffY = xmlNumberAttr(off, "y");
        gextX = xmlNumberAttr(ext, "cx");
        gextY = xmlNumberAttr(ext, "cy");
        gchoffX = xmlNumberAttr(choff, "x");
        gchoffY = xmlNumberAttr(choff, "y");
        gchextX = xmlNumberAttr(chext, "cx");
        gchextY = xmlNumberAttr(chext, "cy");
      }
      const newTransform: GroupTransform = {
        offX: goffX ?? 0,
        offY: goffY ?? 0,
        extX: gextX ?? 1,
        extY: gextY ?? 1,
        choffX: gchoffX ?? 0,
        choffY: gchoffY ?? 0,
        chextX: gchextX ?? 1,
        chextY: gchextY ?? 1,
        parent: currentTransform,
      };
      traverseShapes(grpChildren, newTransform, results);
    } else if (keys.includes("p:sp")) {
      const spChildren = node["p:sp"] as OrderedXmlNode[];
      const nvSpPr = spChildren.find((c) => Object.keys(c).includes("p:nvSpPr"));
      const cNvPr = nvSpPr ? findFirstXmlNode(nvSpPr["p:nvSpPr"] as OrderedXmlNode[], "p:cNvPr") : null;
      const name = cNvPr ? xmlAttr(cNvPr, "name") : "";
      const spPr = spChildren.find((c) => Object.keys(c).includes("p:spPr"));
      const xfrm = spPr ? findFirstXmlNode(spPr["p:spPr"] as OrderedXmlNode[], "a:xfrm") : null;
      let x = 0;
      let y = 0;
      let width = 0;
      let height = 0;
      if (xfrm) {
        const off = findFirstXmlNode(xfrm["a:xfrm"] as OrderedXmlNode[], "a:off");
        const ext = findFirstXmlNode(xfrm["a:xfrm"] as OrderedXmlNode[], "a:ext");
        x = xmlNumberAttr(off, "x") ?? 0;
        y = xmlNumberAttr(off, "y") ?? 0;
        width = xmlNumberAttr(ext, "cx") ?? 0;
        height = xmlNumberAttr(ext, "cy") ?? 0;
      }
      let t = currentTransform;
      while (t) {
        if (t.chextX !== 0 && t.chextY !== 0) {
          x = t.offX + (x - t.choffX) * (t.extX / t.chextX);
          y = t.offY + (y - t.choffY) * (t.extY / t.chextY);
          width = width * (t.extX / t.chextX);
          height = height * (t.extY / t.chextY);
        }
        t = t.parent;
      }
      const absX = emuToRenderPx(x);
      const absY = emuToRenderPx(y);
      const absW = emuToRenderPx(width);
      const absH = emuToRenderPx(height);
      const parsedText = parsePptxTextBody(findFirstXmlNode(spChildren, "p:txBody"));
      const firstRPr = findFirstXmlNode(spChildren, "a:rPr");
      const sz = xmlNumberAttr(firstRPr, "sz");
      const font = sz === undefined ? undefined : Math.round((sz / 100) * (PPTX_RENDER_DPI / 72));
      const pPr = findFirstXmlNode(spChildren, "a:pPr");
      const align = xmlAttr(pPr, "algn") === "ctr" ? "center" : "left";
      const rx = Math.round(absX);
      const ry = Math.round(absY);
      const rw = Math.round(absW);
      const rh = Math.round(absH);
      results.push({
        name: name ?? "",
        text: parsedText.text,
        box: [rx, ry, rx + rw, ry + rh],
        font,
        align,
        emphasis: parsedText.emphasis,
      });
    }
  }
}

export function matchLegacyShapes(
  deskShapes: ParsedPptxShape[],
  yamlBoxes: BoxSpec[],
): Map<string, ParsedPptxShape> {
  const matches = new Map<string, ParsedPptxShape>();
  const unmatchedBoxes = new Set(yamlBoxes);

  const normalizeText = (text: string) => text.replace(/\s+/g, "").replace(/[，。？！、：]/g, "");

  const centerDistance = (box1: number[], box2: number[]) => {
    const c1 = [(box1[0]! + box1[2]!) / 2, (box1[1]! + box1[3]!) / 2];
    const c2 = [(box2[0]! + box2[2]!) / 2, (box2[1]! + box2[3]!) / 2];
    return Math.sqrt((c1[0]! - c2[0]!) ** 2 + (c1[1]! - c2[1]!) ** 2);
  };

  const ioUnion = (box1: number[], box2: number[]) => {
    const xLeft = Math.max(box1[0]!, box2[0]!);
    const yTop = Math.max(box1[1]!, box2[1]!);
    const xRight = Math.min(box1[2]!, box2[2]!);
    const yBottom = Math.min(box1[3]!, box2[3]!);
    const overlapWidth = Math.max(0, xRight - xLeft);
    const overlapHeight = Math.max(0, yBottom - yTop);
    const intersection = overlapWidth * overlapHeight;
    const area1 = (box1[2]! - box1[0]!) * (box1[3]! - box1[1]!);
    const area2 = (box2[2]! - box2[0]!) * (box2[3]! - box2[1]!);
    const union = area1 + area2 - intersection;
    return union > 0 ? intersection / union : 0;
  };

  for (const desk of deskShapes) {
    let bestBox: BoxSpec | null = null;
    let bestScore = -Infinity;
    const deskNorm = normalizeText(desk.text);
    for (const box of unmatchedBoxes) {
      const yamlBox = (box.box ?? []) as number[];
      const boxNorm = normalizeText(box.text || "");
      const textExact = deskNorm === boxNorm;
      const textContains = deskNorm.includes(boxNorm) || boxNorm.includes(deskNorm);
      const iou = ioUnion(desk.box, yamlBox);
      const dist = centerDistance(desk.box, yamlBox);
      let score = 0;
      if (textExact) score += 2000;
      else if (textContains) score += 1000;
      score += iou * 100;
      score -= dist * 0.1;
      if (score > bestScore) {
        bestScore = score;
        bestBox = box;
      }
    }
    if (bestBox && bestScore > -5000) {
      matches.set(String(bestBox.id), desk);
      unmatchedBoxes.delete(bestBox);
    }
  }
  return matches;
}


function emuToRenderPx(value: number | undefined): number {
  if (value === undefined) return 0;
  return value / PPTX_EMU_PER_RENDER_PX;
}


function collectXmlNodes(nodes: OrderedXmlNode[] | undefined, tag: string, out: OrderedXmlNode[] = []): OrderedXmlNode[] {
  if (!nodes) return out;
  for (const node of nodes) {
    for (const [key, value] of Object.entries(node)) {
      if (key === tag) out.push(node);
      if (Array.isArray(value)) collectXmlNodes(value as OrderedXmlNode[], tag, out);
    }
  }
  return out;
}

function findFirstXmlNode(nodes: OrderedXmlNode[] | undefined, tag: string): OrderedXmlNode | undefined {
  return collectXmlNodes(nodes, tag, [])[0];
}

function xmlElementChildren(node: OrderedXmlNode | undefined, tag: string): OrderedXmlNode[] {
  if (!node) return [];
  const children = node[tag];
  return Array.isArray(children) ? (children as OrderedXmlNode[]) : [];
}

function xmlAttr(node: OrderedXmlNode | undefined, name: string): string | undefined {
  const attrs = node?.[":@"] as Record<string, unknown> | undefined;
  const value = attrs?.[`@_${name}`];
  return value === undefined ? undefined : String(value);
}

function xmlNumberAttr(node: OrderedXmlNode | undefined, name: string): number | undefined {
  const value = Number(xmlAttr(node, name));
  return Number.isFinite(value) ? value : undefined;
}

function collectXmlText(nodes: OrderedXmlNode[] | undefined, tag: string): string {
  let text = "";
  for (const node of collectXmlNodes(nodes, tag)) {
    text += collectTextNode(xmlElementChildren(node, tag));
  }
  return text;
}

function collectTextNode(nodes: OrderedXmlNode[] | undefined): string {
  let text = "";
  for (const node of nodes ?? []) {
    for (const [key, value] of Object.entries(node)) {
      if (key === "#text") text += String(value);
      else if (Array.isArray(value)) text += collectTextNode(value as OrderedXmlNode[]);
    }
  }
  return text;
}

export async function verifyGeneratedAssets(
  outputRoot = OUTPUT_ROOT,
  textDir = TEXT_SPEC_DIR,
): Promise<{ missing: string[]; dimensions: Record<string, [number, number]> }> {
  const expectedPrint = [...PAGE_ORDER.map((pageId) => TEMPLATE_IMAGE_NAMES[pageId]), ...COVER_IMAGE_NAMES, ...COVER_EXTRA_IMAGE_NAMES, "web_footer.png"];
  const expectedWeb = ["Overview.png", ...Array.from({ length: 6 }, (_, index) => `Ch${index + 1}.png`)];
  const expectedPptx = ["zh-TW-editable.pptx", "zh-TW-cover-editable.pptx"];
  const missing: string[] = [];
  for (const name of expectedPrint) if (!existsSync(join(outputRoot, name))) missing.push(name);
  for (const name of expectedWeb) if (!existsSync(join(outputRoot, "web", name))) missing.push(`web/${name}`);
  for (const name of expectedPptx) if (!existsSync(join(outputRoot, name))) missing.push(name);

  const dimensions: Record<string, [number, number]> = {};
  for (const name of expectedPrint) {
    const path = join(outputRoot, name);
    if (existsSync(path)) {
      const size = await imageSize(path);
      dimensions[name] = [size.width, size.height];
    }
  }
  for (const name of expectedWeb) {
    const path = join(outputRoot, "web", name);
    if (existsSync(path)) {
      const size = await imageSize(path);
      dimensions[`web/${name}`] = [size.width, size.height];
    }
  }

  if (existsSync(textDir)) {
    for (const pageId of PAGE_ORDER) {
      const specPath = join(textDir, `${pageId}.json`);
      if (!existsSync(specPath)) {
        missing.push(`text/${pageId}.json`);
        continue;
      }
      const spec = JSON.parse(await readFile(specPath, "utf8")) as TextSpec;
      verifyTextSpec(spec, (spec.boxes ?? []).map((box) => String(box.id)));
    }
  }

  const coverTextDir = textDir === TEXT_SPEC_DIR ? COVER_TEXT_SPEC_DIR : join(dirname(textDir), "cover-text");
  for (const pageId of COVER_PAGE_ORDER) {
    const specPath = join(coverTextDir, `${pageId}.json`);
    if (!existsSync(specPath)) {
      missing.push(`cover-text/${pageId}.json`);
      continue;
    }
    const spec = JSON.parse(await readFile(specPath, "utf8")) as TextSpec;
    verifyTextSpec(spec, (spec.boxes ?? []).map((box) => String(box.id)));
  }
  return { missing, dimensions };
}

export async function buildAll(): Promise<Record<string, unknown>> {
  const specs = await buildTextSpecs();
  const coverSpecs = await buildCoverTextSpecs();
  const printAndReviewAssets = await renderAll();
  const webAssets = await buildWebAssets();
  const editablePptx = await buildEditablePptx();
  const coverEditablePptx = await buildCoverEditablePptx();
  const diffReport = await generateDiffMasks();
  const verification = await verifyGeneratedAssets();
  return {
    specs,
    cover_specs: coverSpecs,
    print_and_review_assets: printAndReviewAssets,
    web_assets: webAssets,
    editable_pptx: editablePptx,
    cover_editable_pptx: coverEditablePptx,
    diff_report: diffReport.reportPath,
    verification,
  };
}

export async function main(argv = Bun.argv.slice(2)): Promise<number> {
  const command = argv[0];
  if (!command || !["specs", "cover-specs", "render", "cover-render", "web", "pptx", "cover-pptx", "parseback", "cover-parseback", "cover-mergeback", "diff", "verify", "build", "pdf-ingest", "pdf-oracle"].includes(command)) {
    console.error("usage: bun scripts/zh_tw_comics_pipeline.ts <specs|cover-specs|render|cover-render|web|pptx|cover-pptx|parseback|cover-parseback|cover-mergeback|diff|verify|build|pdf-ingest|pdf-oracle>");
    return 2;
  }
  if (command === "specs") {
    console.log(JSON.stringify({ specs: await buildTextSpecs() }, null, 2));
  } else if (command === "cover-specs") {
    console.log(JSON.stringify({ specs: await buildCoverTextSpecs() }, null, 2));
  } else if (command === "render") {
    console.log(JSON.stringify({ assets: await renderAll() }, null, 2));
  } else if (command === "cover-render") {
    console.log(JSON.stringify({ assets: await renderCoverAssets() }, null, 2));
  } else if (command === "web") {
    console.log(JSON.stringify({ web_assets: await buildWebAssets() }, null, 2));
  } else if (command === "pptx") {
    console.log(JSON.stringify({ editable_pptx: await buildEditablePptx() }, null, 2));
  } else if (command === "cover-pptx") {
    console.log(JSON.stringify({ editable_pptx: await buildCoverEditablePptx() }, null, 2));
  } else if (command === "parseback") {
    console.log(JSON.stringify(await parseEditablePptxToYaml(argv[1] ?? PPTX_OUTPUT_PATH, argv[2] ?? ZH_TW_YAML_PATH), null, 2));
  } else if (command === "cover-parseback" || command === "cover-mergeback") {
    console.log(JSON.stringify(await parseCoverEditablePptxToYaml(argv[1] ?? COVER_PPTX_OUTPUT_PATH, argv[2] ?? ZH_TW_COVER_YAML_PATH), null, 2));
  } else if (command === "diff") {
    console.log(JSON.stringify(await generateDiffMasks(), null, 2));
  } else if (command === "verify") {
    console.log(JSON.stringify(await verifyGeneratedAssets(), null, 2));
  } else if (command === "build") {
    console.log(JSON.stringify(await buildAll(), null, 2));
  } else if (command === "pdf-ingest") {
    console.log(JSON.stringify(await ingestCanonicalPdfToYaml({
      canonicalPdfPath: argv[1] ?? DEFAULT_ZH_TW_CANONICAL_PDF,
      yamlPath: argv[2] ?? ZH_TW_YAML_PATH,
      outputYamlPath: argv[3] ?? ZH_TW_YAML_PATH,
      renderDpi: PPTX_RENDER_DPI,
      dryRun: argv.includes("--dry-run"),
    }), null, 2));
  } else if (command === "pdf-oracle") {
    console.log(JSON.stringify(await runPdfOracleRoundTrip({
      canonicalPdfPath: argv[1] ?? DEFAULT_ZH_TW_CANONICAL_PDF,
      yamlPath: argv[2] ?? ZH_TW_YAML_PATH,
      outputDir: argv[3] ?? join(REVIEW_DIR, "pdf-oracle"),
      renderDpi: PPTX_RENDER_DPI,
    }), null, 2));
  }
  return 0;
}

function normalizeRole(role: FontRole | string): FontRole {
  return role === "title" || role === "footer" ? role : "body";
}

function cssFont(font: FontSpec): string {
  return `${font.size}px ${font.fontFamily}`;
}

function fontSpec(role: FontRole | string, size: number): FontSpec {
  const fontFamily = registerFontForRole(role);
  if (!fontFamily) throw new VerificationError(`No CJK font available for role ${role}`);
  return { fontFamily, size };
}

function measureText(text: string, font: FontSpec) {
  measureContext.font = cssFont(font);
  return measureContext.measureText(text || " ");
}

function textWidth(text: string, font: FontSpec): number {
  return measureText(text, font).width;
}

function lineHeight(font: FontSpec): number {
  const metrics = measureText("Hg國", font);
  const measured = metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent;
  return Math.max(1, Math.ceil(measured || font.size * 1.05));
}

function lineSpacing(font: FontSpec): number {
  return Math.max(2, Math.floor(lineHeight(font) * 0.15));
}

function blockTextHeight(lines: string[], font: FontSpec): number {
  return lines.length * lineHeight(font) + Math.max(0, lines.length - 1) * lineSpacing(font);
}

function segmentText(text: string): string[] {
  const Segmenter = Intl.Segmenter;
  if (Segmenter) {
    const segmenter = new Segmenter("zh-TW", { granularity: "grapheme" });
    return Array.from(segmenter.segment(text), (part) => part.segment);
  }
  return Array.from(text);
}

function parseMarkdownTable(content: string): string[][] {
  const rows: string[][] = [];
  let tableLike = false;
  for (const rawLine of content.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line.includes("|")) {
      if (line) return [];
      continue;
    }
    const cells = line
      .replace(/^\|/, "")
      .replace(/\|$/, "")
      .split("|")
      .map((cell) => cell.trim());
    if (cells.length === 0) continue;
    tableLike = true;
    if (cells.every((cell) => /^:?-{3,}:?$/.test(cell.replace(/\s+/g, "")))) continue;
    if (cells.some(Boolean)) rows.push(cells);
  }
  return tableLike && rows.length ? rows : [];
}

function targetBox(sourceBox: readonly number[], sourceSize: Size, targetSize: Size): number[] {
  const sx = targetSize.width / sourceSize.width;
  const sy = targetSize.height / sourceSize.height;
  const [x0, y0, x1, y1] = sourceBox;
  return [Math.round(x0! * sx), Math.round(y0! * sy), Math.round(x1! * sx), Math.round(y1! * sy)];
}

const ENGLISH_DIFF_THRESHOLD = 30;
const ENGLISH_MULTI_LINE_SCORE_WINDOW = 1.2;

function createEnglishDiffMask(englishCanvas: Canvas, templateCanvas: Canvas): DiffMask {
  const width = Math.min(englishCanvas.width, templateCanvas.width);
  const height = Math.min(englishCanvas.height, templateCanvas.height);
  const englishData = englishCanvas.getContext("2d").getImageData(0, 0, width, height).data;
  const templateData = templateCanvas.getContext("2d").getImageData(0, 0, width, height).data;
  const pixels = new Uint8Array(width * height);
  for (let index = 0, pixel = 0; index < englishData.length; index += 4, pixel++) {
    const delta = Math.max(
      Math.abs(englishData[index]! - templateData[index]!),
      Math.abs(englishData[index + 1]! - templateData[index + 1]!),
      Math.abs(englishData[index + 2]! - templateData[index + 2]!),
    );
    if (delta > ENGLISH_DIFF_THRESHOLD) pixels[pixel] = 1;
  }
  return { width, height, pixels };
}

function englishTextGeometryFromBox(
  mask: DiffMask,
  seedBox: readonly number[],
  role: FontRole,
  text: string,
): EnglishTextGeometry | null {
  const [x0, y0, x1, y1] = seedBox.map(Number);
  const seedWidth = Math.max(1, x1! - x0!);
  const seedHeight = Math.max(1, y1! - y0!);
  const anchoredPad = role === "title" ? 160 : 120;
  const anchoredCandidate = clampBox(
    [x0! - anchoredPad, y0! - anchoredPad, x1! + anchoredPad, y1! + anchoredPad],
    mask.width,
    mask.height,
  );
  const anchoredTightBox = diffBoundsInBox(mask, anchoredCandidate);
  if (anchoredTightBox) {
    return {
      tightBox: anchoredTightBox,
      box: paddedEnglishTextBox(anchoredTightBox, role, mask.width, mask.height),
    };
  }
  const candidate = clampBox(
    [
      x0! - Math.max(160, seedWidth * 0.75),
      y0! - Math.max(110, seedHeight * 0.65),
      x1! + Math.max(160, seedWidth * 0.75),
      y1! + Math.max(110, seedHeight * 0.65),
    ],
    mask.width,
    mask.height,
  );
  const components = englishDiffComponents(mask, candidate, role).map((component) => ({
    ...component,
    score: scoreComponentForSeed(component.box, seedBox),
  })).sort((left, right) => right.score - left.score);
  const best = components[0];
  if (!best) return null;

  const explicitLineCount = Math.max(1, text.split(/\r?\n/).filter((line) => line.trim()).length);
  let selected = [best];
  if (explicitLineCount > 1) {
    const multiLineSelection = components.filter(
      (component) =>
        component.score >= best.score - ENGLISH_MULTI_LINE_SCORE_WINDOW &&
        intersectionArea(component.box, seedBox) > 0,
    );
    if (multiLineSelection.length > 0) selected = multiLineSelection;
  }

  const tightBox = unionBoxes(selected.map((component) => component.box));
  return {
    tightBox,
    box: paddedEnglishTextBox(tightBox, role, mask.width, mask.height),
  };
}

function diffBoundsInBox(mask: DiffMask, candidate: readonly number[]): number[] | null {
  const [x0, y0, x1, y1] = clampBox(candidate, mask.width, mask.height);
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = 0;
  let maxY = 0;
  for (let y = y0!; y < y1!; y++) {
    const rowOffset = y * mask.width;
    for (let x = x0!; x < x1!; x++) {
      if (mask.pixels[rowOffset + x] !== 1) continue;
      minX = Math.min(minX, x);
      minY = Math.min(minY, y);
      maxX = Math.max(maxX, x + 1);
      maxY = Math.max(maxY, y + 1);
    }
  }
  if (!Number.isFinite(minX)) return null;
  return [minX, minY, maxX, maxY];
}

function englishDiffComponents(mask: DiffMask, candidate: readonly number[], role: FontRole): ComponentBox[] {
  const [x0, y0, x1, y1] = clampBox(candidate, mask.width, mask.height);
  const width = x1! - x0!;
  const height = y1! - y0!;
  if (width <= 0 || height <= 0) return [];

  const cellSize = role === "title" ? 32 : 24;
  const mergeX = role === "title" ? 4 : 3;
  const mergeY = 2;
  const cols = Math.max(1, Math.ceil(width / cellSize));
  const rows = Math.max(1, Math.ceil(height / cellSize));
  const original = new Uint8Array(cols * rows);
  for (let y = y0!; y < y1!; y++) {
    const rowOffset = y * mask.width;
    const cellY = Math.min(rows - 1, Math.floor((y - y0!) / cellSize));
    for (let x = x0!; x < x1!; x++) {
      if (mask.pixels[rowOffset + x] !== 1) continue;
      const cellX = Math.min(cols - 1, Math.floor((x - x0!) / cellSize));
      original[cellY * cols + cellX] = 1;
    }
  }

  const dilated = new Uint8Array(original.length);
  for (let index = 0; index < original.length; index++) {
    if (original[index] !== 1) continue;
    const cellX = index % cols;
    const cellY = Math.floor(index / cols);
    for (let dy = -mergeY; dy <= mergeY; dy++) {
      const nextY = cellY + dy;
      if (nextY < 0 || nextY >= rows) continue;
      for (let dx = -mergeX; dx <= mergeX; dx++) {
        const nextX = cellX + dx;
        if (nextX < 0 || nextX >= cols) continue;
        dilated[nextY * cols + nextX] = 1;
      }
    }
  }

  const visited = new Uint8Array(dilated.length);
  const components: ComponentBox[] = [];
  for (let start = 0; start < dilated.length; start++) {
    if (dilated[start] !== 1 || visited[start] === 1) continue;
    const stack = [start];
    visited[start] = 1;
    let minCellX = Number.POSITIVE_INFINITY;
    let minCellY = Number.POSITIVE_INFINITY;
    let maxCellX = 0;
    let maxCellY = 0;
    let originalCells = 0;
    while (stack.length > 0) {
      const current = stack.pop()!;
      const cellX = current % cols;
      const cellY = Math.floor(current / cols);
      if (original[current] === 1) {
        originalCells++;
        minCellX = Math.min(minCellX, cellX);
        minCellY = Math.min(minCellY, cellY);
        maxCellX = Math.max(maxCellX, cellX);
        maxCellY = Math.max(maxCellY, cellY);
      }
      for (let dy = -1; dy <= 1; dy++) {
        const nextY = cellY + dy;
        if (nextY < 0 || nextY >= rows) continue;
        for (let dx = -1; dx <= 1; dx++) {
          const nextX = cellX + dx;
          if (nextX < 0 || nextX >= cols) continue;
          const next = nextY * cols + nextX;
          if (visited[next] === 1 || dilated[next] !== 1) continue;
          visited[next] = 1;
          stack.push(next);
        }
      }
    }
    if (originalCells === 0) continue;
    components.push({
      box: [
        x0! + minCellX * cellSize,
        y0! + minCellY * cellSize,
        Math.min(mask.width, x0! + (maxCellX + 1) * cellSize),
        Math.min(mask.height, y0! + (maxCellY + 1) * cellSize),
      ],
      score: 0,
      cells: originalCells,
    });
  }
  return components;
}

function scoreComponentForSeed(componentBox: readonly number[], seedBox: readonly number[]): number {
  const [x0, y0, x1, y1] = seedBox.map(Number);
  const seedWidth = Math.max(1, x1! - x0!);
  const seedHeight = Math.max(1, y1! - y0!);
  const seedArea = boxArea(seedBox);
  const componentArea = boxArea(componentBox);
  const [cx0, cy0, cx1, cy1] = componentBox.map(Number);
  const centerX = (cx0! + cx1!) / 2;
  const centerY = (cy0! + cy1!) / 2;
  const seedCenterX = (x0! + x1!) / 2;
  const seedCenterY = (y0! + y1!) / 2;
  const normalizedDistance =
    Math.abs(centerX - seedCenterX) / (seedWidth + 1) +
    Math.abs(centerY - seedCenterY) / (seedHeight + 1);
  const overlap = intersectionArea(componentBox, seedBox) / Math.max(1, Math.min(componentArea, seedArea));
  const centerInside = x0! <= centerX && centerX <= x1! && y0! <= centerY && centerY <= y1! ? 1 : 0;
  const hugePenalty = Math.max(0, componentArea / (seedArea + 1) - 6) * 0.2;
  return overlap * 5 + centerInside + 1 / (1 + normalizedDistance) - hugePenalty;
}

function paddedEnglishTextBox(tightBox: readonly number[], role: FontRole, pageWidth: number, pageHeight: number): number[] {
  const pad = role === "title" ? 56 : 48;
  const [x0, y0, x1, y1] = tightBox.map(Number);
  return clampBox([x0! - pad, y0! - pad, x1! + pad, y1! + pad], pageWidth, pageHeight);
}

function unionBoxes(boxes: readonly (readonly number[])[]): number[] {
  return [
    Math.min(...boxes.map((box) => Number(box[0]))),
    Math.min(...boxes.map((box) => Number(box[1]))),
    Math.max(...boxes.map((box) => Number(box[2]))),
    Math.max(...boxes.map((box) => Number(box[3]))),
  ];
}

function intersectionArea(a: readonly number[], b: readonly number[]): number {
  const x0 = Math.max(Number(a[0]), Number(b[0]));
  const y0 = Math.max(Number(a[1]), Number(b[1]));
  const x1 = Math.min(Number(a[2]), Number(b[2]));
  const y1 = Math.min(Number(a[3]), Number(b[3]));
  return Math.max(0, x1 - x0) * Math.max(0, y1 - y0);
}

function boxArea(box: readonly number[]): number {
  return Math.max(0, Number(box[2]) - Number(box[0])) * Math.max(0, Number(box[3]) - Number(box[1]));
}

function clampBox(box: readonly number[], width: number, height: number): number[] {
  return [
    Math.max(0, Math.floor(Number(box[0]))),
    Math.max(0, Math.floor(Number(box[1]))),
    Math.min(width, Math.ceil(Number(box[2]))),
    Math.min(height, Math.ceil(Number(box[3]))),
  ];
}


function estimatedLineCount(text: string, sourceBox: readonly number[], role: FontRole): number {
  const explicitLines = Math.max(1, text.split("\n").length);
  if (role === "title") return explicitLines;
  const [x0, y0, x1, y1] = sourceBox;
  const width = Math.max(1, x1! - x0!);
  const height = Math.max(1, y1! - y0!);
  const denseText = text.replace(/\s+/g, "");
  const estimated = Math.ceil(Math.sqrt((Math.max(1, Array.from(denseText).length) * height * 0.75) / (width * 1.2)));
  return Math.max(explicitLines, estimated);
}

function nominalFontSize(
  text: string,
  box: readonly number[],
  sourceBox: readonly number[],
  role: FontRole,
): [number, number, number] {
  const height = Math.max(1, box[3]! - box[1]!);
  const lineCount = estimatedLineCount(text, sourceBox, role);
  const nominalLine = height / lineCount;
  const multiplier = role === "title" ? 0.92 : 0.78;
  const maxSize = Math.max(10, Math.floor(nominalLine * multiplier));
  const minSize = Math.max(7, Math.floor(maxSize * 0.45));
  return [maxSize, minSize, lineCount];
}

function inferAlign(text: string, sourceBox: readonly number[], sourceSize: Size, role: FontRole): TextAlign {
  const lines = text.split("\n").map((line) => line.trim()).filter(Boolean);
  if (lines.length === 0) return "left";
  const [x0, , x1] = sourceBox;
  const width = Math.max(1, x1! - x0!);
  const centerX = (x0! + x1!) / 2;
  const maxLineChars = Math.max(...lines.map((line) => Array.from(line.replace(/\s+/g, "")).length));
  if (role === "title" && width > 260 && sourceSize.width * 0.25 < centerX && centerX < sourceSize.width * 0.75) {
    return "center";
  }
  if (lines.length >= 2 && maxLineChars <= 4 && width > 120) return "center";
  return "left";
}

type InlineEmphasisMap = {
  bold: Uint8Array;
  colors: Array<string | undefined>;
  hasBold: boolean;
  hasColor: boolean;
};

function emphasisStyleMap(text: string, emphasis: EmphasisSpec[] | undefined): InlineEmphasisMap | null {
  if (!emphasis || emphasis.length === 0) return null;
  let resolved: InlineEmphasisMap | null = null;
  const ensureResolved = () => {
    resolved ??= {
      bold: new Uint8Array(text.length),
      colors: new Array<string | undefined>(text.length),
      hasBold: false,
      hasColor: false,
    };
    return resolved;
  };
  for (const span of emphasis) {
    const bold = span.bold === true;
    const color = hexColorOverride(span.color);
    if (!bold && !color) continue;
    const range = emphasisRange(text, span);
    if (!range) continue;
    const map = ensureResolved();
    for (let index = range.start; index < range.end; index += 1) {
      if (bold) {
        map.bold[index] = 1;
        map.hasBold = true;
      }
      if (color) {
        map.colors[index] = color;
        map.hasColor = true;
      }
    }
  }
  const finalMap = resolved as InlineEmphasisMap | null;
  return finalMap && (finalMap.hasBold || finalMap.hasColor) ? finalMap : null;
}

function emphasisRange(text: string, span: EmphasisSpec): { start: number; end: number } | null {
  if (span.start !== undefined || span.end !== undefined) {
    const start = Math.max(0, Math.floor(span.start ?? 0));
    const end = Math.min(text.length, Math.floor(span.end ?? start));
    return end > start ? { start, end } : null;
  }
  if (!span.text) return null;
  const occurrence = Math.max(1, Math.floor(span.occurrence ?? 1));
  let searchFrom = 0;
  let found = -1;
  for (let count = 0; count < occurrence; count += 1) {
    found = text.indexOf(span.text, searchFrom);
    if (found < 0) break;
    searchFrom = found + span.text.length;
  }
  return found >= 0 ? { start: found, end: found + span.text.length } : null;
}

function emphasizedCodeUnits(text: string, emphasis: EmphasisSpec[] | undefined): Uint8Array | null {
  const resolved = emphasisStyleMap(text, emphasis);
  return resolved?.hasBold ? resolved.bold : null;
}

function removeDefaultColorEmphasis(text: string, emphasis: EmphasisSpec[], defaultColor: unknown): EmphasisSpec[] {
  const baseColor = hexColorOverride(defaultColor) ?? "#222222";
  const normalized: EmphasisSpec[] = [];
  for (const entry of emphasis) {
    const color = hexColorOverride(entry.color);
    const keepColor = color !== undefined && color !== baseColor;
    if (entry.bold !== true && !keepColor) continue;
    if (keepColor) {
      const start = Math.max(0, Math.floor(entry.start ?? 0));
      const end = Math.min(text.length, Math.floor(entry.end ?? start));
      if (end <= start) continue;
      const segment = text.slice(start, end);
      const next: EmphasisSpec = { text: segment };
      if (entry.bold === true) next.bold = true;
      next.color = color;
      const occurrence = occurrenceForRange(text, segment, start);
      if (occurrence > 1) next.occurrence = occurrence;
      normalized.push(next);
      continue;
    }
    const next: EmphasisSpec = {
      start: entry.start,
      end: entry.end,
      bold: true,
    };
    normalized.push(next);
  }
  return normalized;
}

function occurrenceForRange(text: string, segment: string, targetStart: number): number {
  let occurrence = 0;
  let searchFrom = 0;
  while (segment.length > 0) {
    const found = text.indexOf(segment, searchFrom);
    if (found < 0 || found > targetStart) break;
    occurrence += 1;
    if (found === targetStart) return occurrence;
    searchFrom = found + segment.length;
  }
  return 1;
}

function drawLineWithInlineEmphasis(
  ctx: SKRSContext2D,
  line: string,
  fullText: string,
  searchStart: number,
  emphasis: InlineEmphasisMap,
  font: FontSpec,
  color: string,
  x: number,
  y: number,
): number {
  const located = fullText.indexOf(line, searchStart);
  const lineStart = located >= 0 ? located : searchStart;
  let run = "";
  let runX = x;
  let codeOffset = 0;
  let runBold = Boolean(emphasis.bold[lineStart]);
  let runColor = emphasis.colors[lineStart] ?? color;
  const flush = () => {
    if (!run) return;
    ctx.fillStyle = runColor;
    if (runBold) {
      ctx.strokeStyle = runColor;
      ctx.lineWidth = Math.max(1, Math.floor(font.size * 0.035));
      ctx.strokeText(run, runX, y);
    }
    ctx.fillText(run, runX, y);
    runX += textWidth(run, font);
    run = "";
  };
  for (const char of Array.from(line)) {
    const charBold = Boolean(emphasis.bold[lineStart + codeOffset]);
    const charColor = emphasis.colors[lineStart + codeOffset] ?? color;
    if (run && (charBold !== runBold || charColor !== runColor)) {
      flush();
      runBold = charBold;
      runColor = charColor;
    }
    run += char;
    codeOffset += char.length;
  }
  flush();
  return lineStart + line.length;
}

function drawBoxText(ctx: SKRSContext2D, box: BoxSpec): void {
  if (String(box.ocr_type ?? "") === "table") {
    drawTableText(ctx, box);
    return;
  }
  if (!box.box) return;
  const [x0, y0, x1, y1] = box.box.map(Number);
  const text = String(box.text ?? "");
  const font = fontSpec(box.font_role ?? "body", Number(box.font_size ?? 24));
  const wrapped = wrapTextPreserveBreaks(text, font, x1! - x0!);
  const lines = wrapped.split("\n");
  const color = String(box.color ?? "#222222");
  const align = String(box.align ?? "left");
  const spacingOverride = numericOverride(box.line_spacing);
  const spacing = spacingOverride ?? lineSpacing(font);
  const totalHeight = lines.length * lineHeight(font) + Math.max(0, lines.length - 1) * spacing;
  let lineY = y0! + Math.max(0, (y1! - y0! - totalHeight) / 2);
  ctx.font = cssFont(font);
  ctx.fillStyle = color;
  ctx.textBaseline = "top";
  const emphasis = emphasisStyleMap(text, box.emphasis);
  let searchStart = 0;
  for (const line of lines) {
    const lineW = textWidth(line, font);
    let lineX = x0!;
    if (align === "center") lineX = x0! + Math.max(0, (x1! - x0! - lineW) / 2);
    if (emphasis) {
      searchStart = drawLineWithInlineEmphasis(ctx, line, text, searchStart, emphasis, font, color, lineX, lineY);
    } else {
      ctx.fillStyle = color;
      ctx.fillText(line, lineX, lineY);
      searchStart += line.length;
    }
    lineY += lineHeight(font) + spacing;
  }
}

function drawTableText(ctx: SKRSContext2D, box: BoxSpec): void {
  if (!box.box) return;
  const [x0, y0, x1, y1] = box.box.map(Number);
  const width = Math.max(1, x1! - x0!);
  const rows = String(box.text ?? "").split("\n").filter((row) => row.trim());
  if (rows.length === 0) return;
  const rowHeight = Math.max(1, (y1! - y0!) / rows.length);
  const baseSize = Number(box.font_size ?? 24);
  const headerFont = fontSpec(box.font_role ?? "body", baseSize);
  const bodyFont = fontSpec(box.font_role ?? "body", Math.max(1, Math.floor(baseSize * 0.62)));
  const color = String(box.color ?? "#222222");
  ctx.fillStyle = color;
  ctx.textBaseline = "top";

  for (const [rowIndex, row] of rows.entries()) {
    const font = rowIndex === 0 ? headerFont : bodyFont;
    const cells = row.split(/(?:　{2,}|\s{2,})/).map((cell) => cell.trim()).filter(Boolean);
    const columns: Array<[string, number, number]> = [];
    if (cells.length >= 2) {
      const firstX = x0! + Math.max(1, Math.floor(width * (rowIndex === 0 ? 0.18 : 0.27)));
      const secondX = x0! + Math.max(1, Math.floor(width * 0.58));
      const rightLimit = x1! - Math.max(0, Math.floor(width * 0.12));
      columns.push([cells[0]!, firstX, Math.max(1, secondX - firstX - Math.floor(width * 0.04))]);
      columns.push([cells.slice(1).join("　　"), secondX, Math.max(1, rightLimit - secondX)]);
    } else {
      const firstX = x0! + Math.max(1, Math.floor(width * (rowIndex === 0 ? 0.18 : 0.27)));
      columns.push([row, firstX, Math.max(1, x1! - firstX)]);
    }

    const rowTop = y0! + rowHeight * rowIndex;
    for (const [cellText, cellX, cellWidth] of columns) {
      const wrappedLines = wrapTextPreserveBreaks(cellText, font, cellWidth).split("\n");
      const totalHeight = blockTextHeight(wrappedLines, font);
      let lineY = rowTop + Math.max(0, (rowHeight - totalHeight) / 2);
      ctx.font = cssFont(font);
      for (const line of wrappedLines) {
        ctx.fillText(line, cellX, lineY);
        lineY += lineHeight(font) + lineSpacing(font);
      }
    }
  }
}

function sourcePatchBox(spec: TextSpec, target: readonly number[]): number[] {
  const [sourceWidth, sourceHeight] = (spec.source_size ?? []).map(Number);
  const [targetWidth, targetHeight] = (spec.target_size ?? []).map(Number);
  const sx = targetWidth! / sourceWidth!;
  const sy = targetHeight! / sourceHeight!;
  const [x0, y0, x1, y1] = target.map(Number);
  return [
    Math.max(0, Math.floor(x0! / sx)),
    Math.max(0, Math.floor(y0! / sy)),
    Math.min(sourceWidth!, Math.ceil(x1! / sx)),
    Math.min(sourceHeight!, Math.ceil(y1! / sy)),
  ];
}

function pasteSourcePatch(ctx: SKRSContext2D, source: Image, patch: ImagePatchSpec): void {
  const [sx0, sy0, sx1, sy1] = patch.source_box.map(Number);
  const [x0, y0, x1, y1] = patch.box.map(Number);
  if (sx1! <= sx0! || sy1! <= sy0! || x1! <= x0! || y1! <= y0!) return;
  ctx.save();
  ctx.imageSmoothingEnabled = false;
  ctx.drawImage(
    source,
    sx0!,
    sy0!,
    sx1! - sx0!,
    sy1! - sy0!,
    x0!,
    y0!,
    x1! - x0!,
    y1! - y0!,
  );
  ctx.restore();
}

function pasteOraclePatch(ctx: SKRSContext2D, source: Image, spec: TextSpec, box: BoxSpec): void {
  if (!box.box) return;
  const [x0, y0, x1, y1] = box.box.map(Number);
  if (x1! <= x0! || y1! <= y0!) return;
  const [sourceWidth, sourceHeight] = (spec.source_size ?? []).map(Number);
  const [targetWidth, targetHeight] = (spec.target_size ?? []).map(Number);
  const sx = targetWidth! / sourceWidth!;
  const sy = targetHeight! / sourceHeight!;
  const patchBox = sourcePatchBox(spec, box.box);
  patchBox[2] = Math.min(sourceWidth!, patchBox[2]! + 1);
  patchBox[3] = Math.min(sourceHeight!, patchBox[3]! + 1);
  const pasteX1 = Math.min(ctx.canvas.width, x1! + Math.max(1, Math.round(sx)));
  const pasteY1 = Math.min(ctx.canvas.height, y1! + Math.max(1, Math.round(sy)));
  const width = Math.max(1, pasteX1 - x0!);
  const height = Math.max(1, pasteY1 - y0!);
  ctx.save();
  ctx.imageSmoothingEnabled = false;
  ctx.drawImage(
    source,
    patchBox[0]!,
    patchBox[1]!,
    Math.max(1, patchBox[2]! - patchBox[0]!),
    Math.max(1, patchBox[3]! - patchBox[1]!),
    x0!,
    y0!,
    width,
    height,
  );
  ctx.restore();
}

function resolveProjectPath(value: string): string {
  if (!value) return PROJECT_ROOT;
  return isAbsolute(value) ? value : join(PROJECT_ROOT, value);
}

async function imageSize(path: string): Promise<Size> {
  const handle = await open(path, "r");
  try {
    const header = Buffer.alloc(32);
    await handle.read(header, 0, header.length, 0);
    if (header.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) {
      return { width: header.readUInt32BE(16), height: header.readUInt32BE(20) };
    }
  } finally {
    await handle.close();
  }
  const image = await loadImage(path);
  return { width: image.width, height: image.height };
}


async function loadCanvasFromFile(path: string): Promise<Canvas> {
  const image = await loadImage(path);
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(image, 0, 0);
  return canvas;
}

async function saveCanvasPng(canvas: Canvas, path: string): Promise<void> {
  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, canvas.toBuffer("image/png"));
}

async function writeJson(path: string, value: unknown): Promise<void> {
  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function imageDataInBox(image: Canvas | Image, rawBox: readonly number[]): { data: Uint8ClampedArray; width: number; height: number } {
  const sourceCanvas = image instanceof (createCanvas(1, 1).constructor as typeof Canvas) ? (image as Canvas) : canvasFromImage(image as Image);
  const [rawX0, rawY0, rawX1, rawY1] = rawBox.map(Number);
  const x0 = Math.max(0, Math.floor(rawX0 ?? 0));
  const y0 = Math.max(0, Math.floor(rawY0 ?? 0));
  const x1 = Math.min(sourceCanvas.width, Math.ceil(rawX1 ?? 0));
  const y1 = Math.min(sourceCanvas.height, Math.ceil(rawY1 ?? 0));
  const width = Math.max(0, x1 - x0);
  const height = Math.max(0, y1 - y0);
  if (width === 0 || height === 0) return { data: new Uint8ClampedArray(), width, height };
  const data = sourceCanvas.getContext("2d").getImageData(x0, y0, width, height).data;
  return { data, width, height };
}

function canvasFromImage(image: Image): Canvas {
  const canvas = createCanvas(image.width, image.height);
  canvas.getContext("2d").drawImage(image, 0, 0);
  return canvas;
}

const MIN_TEMPLATE_CONTRAST = 80;
const DARK_INK_FALLBACK = "#1E1E1E";
const LIGHT_INK_FALLBACK = "#F2F2F2";

function ensureContrastingInkColor(inkColor: string, template: Canvas | Image, targetBox: readonly number[]): string {
  const inkRgb = hexToRgb(inkColor);
  if (!inkRgb) return inkColor;
  const { data } = imageDataInBox(template, targetBox);
  if (data.length === 0) return inkColor;

  const backgroundPixels = collectSampledPixels(data);
  if (backgroundPixels.length === 0) return inkColor;

  const background = medianRgb(backgroundPixels);
  const inkBrightness = rgbBrightness(inkRgb);
  const backgroundBrightness = rgbBrightness(background);
  if (Math.abs(inkBrightness - backgroundBrightness) >= MIN_TEMPLATE_CONTRAST) return inkColor;

  return backgroundBrightness >= 150 ? DARK_INK_FALLBACK : LIGHT_INK_FALLBACK;
}

function collectSampledPixels(data: Uint8ClampedArray): Array<[number, number, number]> {
  const totalPixels = data.length / 4;
  const step = Math.max(1, Math.floor(totalPixels / 10000));
  const pixels: Array<[number, number, number]> = [];
  for (let index = 0, pixel = 0; index < data.length; index += 4, pixel++) {
    if (pixel % step !== 0) continue;
    pixels.push([data[index]!, data[index + 1]!, data[index + 2]!]);
  }
  return pixels;
}

function rgbBrightness([red, green, blue]: readonly number[]): number {
  return 0.299 * red! + 0.587 * green! + 0.114 * blue!;
}

function hexToRgb(color: string): [number, number, number] | null {
  const normalized = color.trim().replace(/^#/, "");
  if (!/^[0-9A-Fa-f]{6}$/.test(normalized)) return null;
  return [
    Number.parseInt(normalized.slice(0, 2), 16),
    Number.parseInt(normalized.slice(2, 4), 16),
    Number.parseInt(normalized.slice(4, 6), 16),
  ];
}

function collectPixelsByBrightness(
  data: Uint8ClampedArray,
  threshold: number,
  direction: "above" | "below",
): Array<[number, number, number]> {
  const totalPixels = data.length / 4;
  const step = Math.max(1, Math.floor(totalPixels / 10000));
  const pixels: Array<[number, number, number]> = [];
  for (let index = 0, pixel = 0; index < data.length; index += 4, pixel++) {
    if (pixel % step !== 0) continue;
    const red = data[index]!;
    const green = data[index + 1]!;
    const blue = data[index + 2]!;
    const brightness = (red + green + blue) / 3;
    if (direction === "above" ? brightness > threshold : brightness < threshold) pixels.push([red, green, blue]);
  }
  return pixels;
}

function medianRgb(pixels: Array<[number, number, number]>): [number, number, number] {
  const red = pixels.map((pixel) => pixel[0]).sort((a, b) => a - b);
  const green = pixels.map((pixel) => pixel[1]).sort((a, b) => a - b);
  const blue = pixels.map((pixel) => pixel[2]).sort((a, b) => a - b);
  const middle = Math.floor(pixels.length / 2);
  return [red[middle] ?? 0, green[middle] ?? 0, blue[middle] ?? 0];
}

function rgbToHex([red, green, blue]: readonly number[]): string {
  return `#${toHex(red!)}${toHex(green!)}${toHex(blue!)}`;
}

function toHex(value: number): string {
  return Math.max(0, Math.min(255, Math.round(value))).toString(16).toUpperCase().padStart(2, "0");
}

function resizeCanvasSmoothToMaxHeight(source: Canvas, maxHeight: number): Canvas {
  if (source.height <= maxHeight) return source;
  const scale = maxHeight / source.height;
  const width = Math.max(1, Math.round(source.width * scale));
  const height = Math.max(1, Math.round(source.height * scale));
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext("2d");
  ctx.imageSmoothingEnabled = true;
  ctx.drawImage(source, 0, 0, width, height);
  return canvas;
}

function resizeCanvasNearest(source: Canvas, width: number, height: number): Canvas {
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext("2d");
  ctx.imageSmoothingEnabled = false;
  ctx.drawImage(source, 0, 0, width, height);
  return canvas;
}

function cropCanvas(source: Canvas, box: readonly number[]): Canvas {
  const [x0, y0, x1, y1] = box.map(Number);
  const width = Math.max(1, x1! - x0!);
  const height = Math.max(1, y1! - y0!);
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(source, x0!, y0!, width, height, 0, 0, width, height);
  return canvas;
}

function appendVertical(images: readonly Canvas[]): Canvas {
  const width = Math.max(...images.map((image) => image.width));
  const height = images.reduce((sum, image) => sum + image.height, 0);
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "white";
  ctx.fillRect(0, 0, width, height);
  let y = 0;
  for (const image of images) {
    ctx.drawImage(image, Math.floor((width - image.width) / 2), y);
    y += image.height;
  }
  return canvas;
}

if (import.meta.main) {
  const code = await main();
  process.exit(code);
}
