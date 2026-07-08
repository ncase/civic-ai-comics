/**
 * Failing-first contracts for the zh-TW comics Bun pipeline.
 * Targets ../scripts/zh_tw_comics_pipeline (TypeScript port).
 */
import { beforeAll, describe, expect, test } from "bun:test";
import { existsSync } from "node:fs";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { tmpdir } from "node:os";
import { createCanvas, loadImage, type Canvas } from "@napi-rs/canvas";
import PptxGenJS from "pptxgenjs";
import JSZip from "jszip";
import * as zhTwPipeline from "../scripts/zh_tw_comics_pipeline";
import {
  type BoxSpec,
  type EmphasisSpec,
  type OrderedXmlNode,
  type ParsedPptxShape,
  type TextSpec,
  VerificationError,
  PAGE_ORDER,
  SOURCE_IMAGE_NAMES,
  TEMPLATE_IMAGE_NAMES,
  PROJECT_ROOT,
  OUTPUT_ROOT,
  PPTX_OUTPUT_PATH,
  WORDLESS_ROOT,
  RAW_OCR_DIR,
  TEXT_SPEC_DIR,
  ZH_TW_YAML_PATH,
  REVIEW_DIR,
  renderPptxBackground,
  parseEditablePptxToYaml,
  type PageId,
  buildComicsOnlyManifest,
  buildEditablePptx,
  buildTextSpecs,
  cleanOcrContent,
  compareRenderedPageToSource,
  estimateInkColor,
  findFont,
  fitFontSizeInBox,
  generateDiffMasks,
  loadYamlTextSpecs,
  measureRegisteredText,
  probeCanvasRuntime,
  registerFontForRole,
  renderPage,
  renderPdfPages,
  verifyGeneratedAssets,
  specFromOcrRaw,
  verifyTextSpec,
  wrapTextPreserveBreaks,
} from "../scripts/zh_tw_comics_pipeline";
const hasOcrRaw = existsSync(join(RAW_OCR_DIR, "Overview.json"));
const ocrTest = hasOcrRaw ? test : test.skip;

const EMPHASIS_COLOR_COMPASS = "#4F8F00";
const EMPHASIS_COLOR_COMPASS_SPAN = "「羅盤」";
const EMPHASIS_COLOR_CH1_BALLOT = "#DF5D46";
const EMPHASIS_COLOR_CH1_BALLOT_SPAN = "壞人";
const CH4_A_EMPHASIS_COLOR_BOX = "b008_body";
const PPTX_TEST_RENDER_DPI = 600;

const TABLE_OCR =
  "|  指標 | 失靈  |\n| --- | --- |\n|  點擊 | 標題黨  |\n";

const TABLE_TEXT = "指標　　失靈\n點擊　　標題黨";
const TABLE_BOX = [40, 60, 520, 320] as const;
const GUTTER_WIDTH = 72;
const DARK_LUMA_MAX = 200;
const DIFF_THRESHOLD = 32;
const SYNTHETIC_TEXT = "測試";
const LIGHT_INK_LUMA_MIN = 170;
const ORACLE_PATCH_COLOR = { r: 18, g: 64, b: 200 };

const ZH_TW_COVER_OUTPUT_NAMES = ["Cover.png", "Cover_Back.png", "Cover_Spine.png"] as const;
const ZH_TW_COVER_VERIFY_TEXT_SPECS = [
  "cover-text/Cover.json",
  "cover-text/Cover_Back.json",
  "cover-text/Cover_Spine.json",
] as const;
const ZH_TW_COVER_EDITABLE_PPTX = "zh-TW-cover-editable.pptx";

const CH5_B_PPTX_CAPTION_SNIPPET = "舉例來說，也許我們的兩位主廚會對彼此說";
const CANONICAL_ZH_TW_EDITABLE_PDF = join(PROJECT_ROOT, "tests", "fixtures", "zh-TW-editable.pdf");
/** Rendered from {@link CANONICAL_ZH_TW_EDITABLE_PDF} on-the-fly for aspect-ratio tests. */
let DESKTOP_ZH_TW_ORACLE_ROOT = "";
let ORACLE_PAGE_PATHS: string[] = [];
/** Max |width/height(gen) − width/height(oracle)|; print vs slide export scales differ ~0.00036. */
const ORACLE_ASPECT_RATIO_TOLERANCE = 0.001;
const PDF_ORACLE_PAGE_COUNT = 13;
const PDF_ORACLE_PAGE_SIZE_PT: readonly [number, number] = [689, 960];
const PDF_ORACLE_RENDER_DPI = 600;
const PDF_ORACLE_RENDER_SIZE_PX: readonly [number, number] = [5742, 8000];
const PDF_ORACLE_PPTX_SLIDE_WIDTH_IN = PDF_ORACLE_PAGE_SIZE_PT[0] / 72;
const PDF_ORACLE_PPTX_SLIDE_HEIGHT_IN = PDF_ORACLE_PAGE_SIZE_PT[1] / 72;
const LEGACY_PPTX_SLIDE_WIDTH_IN = 5749 / PPTX_TEST_RENDER_DPI;
const PPTX_EMU_PER_INCH_TEST = 914400;
const LEGACY_ZH_TW_TARGET_WIDTH_PX = 5749;
const INGESTED_ZH_TW_TARGET_WIDTH_PX = PDF_ORACLE_RENDER_SIZE_PX[0];
const INGESTED_ZH_TW_TARGET_HEIGHT_PX = PDF_ORACLE_RENDER_SIZE_PX[1];
const INGESTED_ZH_TW_PPTX_SLIDE_WIDTH_IN = INGESTED_ZH_TW_TARGET_WIDTH_PX / PPTX_TEST_RENDER_DPI;
const INGESTED_ZH_TW_PPTX_SLIDE_HEIGHT_IN = INGESTED_ZH_TW_TARGET_HEIGHT_PX / PPTX_TEST_RENDER_DPI;
/** Overview/b001_title legacy target-space box[0] and tight[0] in translation-assets/zh-TW.yml. */
const OVERVIEW_B001_TITLE_BOX_X_LEGACY = 965;
const OVERVIEW_B001_TIGHT_X_LEGACY = 1001;
const OVERVIEW_B001_SRC: readonly [number, number, number, number] = [242, 89, 489, 184];


type InspectCanonicalPdfResult = {
  pageCount: number;
  pageSizePt: [number, number];
  renderSizePx: [number, number];
  renderDpi: number;
  pptxSlideSizeIn: [number, number];
};

type PdfTextRun = {
  text: string;
  fontFamily?: string;
  color?: string;
  role?: "title" | "body" | "footer";
};

type ComparePdfOraclePagesResult = {
  ok: boolean;
  dimensionMismatch?: boolean;
  diffPixels?: number;
  message?: string;
};

type PdfOraclePageDiffMetric = {
  pageId: string;
  diffPixels?: number;
  mismatchRatio?: number;
};

type PdfOracleRoundTripReport = {
  ok: boolean;
  canonicalRenderSizePx: [number, number];
  renderer: { kind: string };
  incompleteReason?: string;
  pageRenderSizePx?: [number, number];
  generatedRenderSizePx?: [number, number];
  pageDiffs?: PdfOraclePageDiffMetric[];
};

type IngestCanonicalPdfToYamlReport = {
  canonicalRenderSizePx: [number, number];
  updatedPagesCount: number;
  renderDpi: number;
  dryRun: boolean;
  outputYamlPath: string;
};

function expectPdfOracleRoundTripRendererContract(report: PdfOracleRoundTripReport): void {
  if (report.renderer?.kind === "pptx-pdf") {
    return;
  }
  expect(report.ok).toBe(false);
  expect(String(report.incompleteReason ?? "")).toMatch(/pptx/i);
  expect(String(report.incompleteReason ?? "")).toMatch(/pdf/i);
}

function requirePdfOracleExport<T>(exportName: string): T {
  return requireCoverPipelineExport<T>(exportName);
}

function pptxSlideSizeEmuFromInches(widthIn: number, heightIn: number): { cx: number; cy: number } {
  return {
    cx: Math.round(widthIn * PPTX_EMU_PER_INCH_TEST),
    cy: Math.round(heightIn * PPTX_EMU_PER_INCH_TEST),
  };
}

async function readPptxPresentationSlideSizeEmu(pptxPath: string): Promise<{ cx: number; cy: number }> {
  const presXml = await readPptxZipEntryText(pptxPath, "ppt/presentation.xml");
  const cxMatch = presXml.match(/<p:sldSz[^>]*\bcx="(\d+)"/);
  const cyMatch = presXml.match(/<p:sldSz[^>]*\bcy="(\d+)"/);
  if (!cxMatch || !cyMatch) {
    throw new Error(`ppt/presentation.xml missing p:sldSz cx/cy in ${pptxPath}`);
  }
  return { cx: Number(cxMatch[1]), cy: Number(cyMatch[1]) };
}

async function listPptxZipEntryNames(pptxPath: string): Promise<string[]> {
  const zip = await new JSZip().loadAsync(await Bun.file(pptxPath).arrayBuffer());
  return Object.keys(zip.files).filter((name) => !zip.files[name]!.dir);
}

async function readPptxZipEntryText(pptxPath: string, entryPath: string): Promise<string> {
  const zip = await new JSZip().loadAsync(await Bun.file(pptxPath).arrayBuffer());
  const entry = zip.file(entryPath);
  if (!entry) throw new Error(`pptx missing entry: ${entryPath}`);
  return entry.async("string");
}

async function readPptxZipEntryBytes(pptxPath: string, entryPath: string): Promise<Uint8Array> {
  const zip = await new JSZip().loadAsync(await Bun.file(pptxPath).arrayBuffer());
  const entry = zip.file(entryPath);
  if (!entry) throw new Error(`pptx missing entry: ${entryPath}`);
  return entry.async("uint8array");
}

const PARSEBACK_PPTX_DPI = 600;
const PARSEBACK_EMU_PER_INCH = 914400;
const CH2_A_PARSEBACK_NOBODY = "當沒有人負責，就沒有人負責。";
const CH5_B_PARSEBACK_SOLIDARITY_OLD = `雙方「為何」的、新的「做什麼」？`;
const CH5_B_PARSEBACK_SOLIDARITY_NEW_MID = `雙方「為何」的、新的「共同目標」？`;
const CH5_B_PARSEBACK_EARLIER_MD_DECOY = "在你想要「做什麼」背後";
const CH5_B_PARSEBACK_STALE_B004_OPENING = "          退一步想。在你想要「做什麼」背後，\n          問問自己";
const CH5_B_PARSEBACK_STALE_B004_OPENING_TRIM = "          問問自己";
const CH5_B_PARSEBACK_EARLIER_MD_DECOY_PARAGRAPH =
  "先前討論時，我們也會先問：在你想要「做什麼」背後，真正驅動你的是什麼？";
const CH5_B_PARSEBACK_STALE_B004_BODY_FOR_MD_SYNC =
  "問問自己：你「為何」想要它？接著：你能不能創造一個同時滿足";

type ParseEditablePptxToYamlResult = {
  yaml_path: string;
  updated: number;
  unknown_shapes: string[];
};

type ParseEditablePptxToYamlFn = (
  pptxPath?: string,
  yamlPath?: string,
) => Promise<ParseEditablePptxToYamlResult>;

function requireParseEditablePptxToYaml(): ParseEditablePptxToYamlFn {
  if (!("parseEditablePptxToYaml" in zhTwPipeline)) {
    throw new Error("parseEditablePptxToYaml is not exported from ../scripts/zh_tw_comics_pipeline");
  }
  const fn = zhTwPipeline.parseEditablePptxToYaml;
  if (typeof fn !== "function") {
    throw new Error("parseEditablePptxToYaml is not exported from ../scripts/zh_tw_comics_pipeline");
  }
  return fn as ParseEditablePptxToYamlFn;
}

function requireParseCoverEditablePptxToYaml(): ParseEditablePptxToYamlFn {
  if (!("parseCoverEditablePptxToYaml" in zhTwPipeline)) {
    throw new Error("parseCoverEditablePptxToYaml is not exported from ../scripts/zh_tw_comics_pipeline");
  }
  const fn = (zhTwPipeline as Record<string, unknown>).parseCoverEditablePptxToYaml;
  if (typeof fn !== "function") {
    throw new Error("parseCoverEditablePptxToYaml is not exported from ../scripts/zh_tw_comics_pipeline");
  }
  return fn as ParseEditablePptxToYamlFn;
}

type CoverPageId = "Cover" | "Cover_Back" | "Cover_Spine";

type LoadCoverTextSpecsFn = (yamlPath?: string) => Promise<Record<CoverPageId, TextSpec>>;
type BuildCoverTextSpecsFn = (yamlPath?: string, textDir?: string) => Promise<string[]>;
type RenderCoverAssetsFn = (textDir?: string, outputRoot?: string) => Promise<string[]>;
type BuildCoverEditablePptxFn = (outputPath?: string, yamlPath?: string) => Promise<string>;

const COVER_PAGE_ORDER = ["Cover", "Cover_Back", "Cover_Spine"] as const satisfies readonly CoverPageId[];

const COVER_PAGE_OUTPUT_PNG: Record<CoverPageId, string> = {
  Cover: "Cover.png",
  Cover_Back: "Cover_Back.png",
  Cover_Spine: "Cover_Spine.png",
};

/** Small sentinel subset of pipeline SIMPLIFIED_TRAP_CHARS for cover copy checks. */
const COVER_SIMPLIFIED_TRAP_SENTINEL = "这为国" as const;

function requireCoverPipelineExport<T>(exportName: string): T {
  if (!(exportName in zhTwPipeline)) {
    throw new Error(`${exportName} is not exported from ../scripts/zh_tw_comics_pipeline`);
  }
  const value = (zhTwPipeline as Record<string, unknown>)[exportName];
  if (value === undefined) {
    throw new Error(`${exportName} is not exported from ../scripts/zh_tw_comics_pipeline`);
  }
  return value as T;
}

function requireLoadCoverTextSpecs(): LoadCoverTextSpecsFn {
  const fn = requireCoverPipelineExport<LoadCoverTextSpecsFn>("loadCoverTextSpecs");
  if (typeof fn !== "function") {
    throw new Error("loadCoverTextSpecs is not exported from ../scripts/zh_tw_comics_pipeline");
  }
  return fn;
}

function requireBuildCoverTextSpecs(): BuildCoverTextSpecsFn {
  const fn = requireCoverPipelineExport<BuildCoverTextSpecsFn>("buildCoverTextSpecs");
  if (typeof fn !== "function") {
    throw new Error("buildCoverTextSpecs is not exported from ../scripts/zh_tw_comics_pipeline");
  }
  return fn;
}

function requireRenderCoverAssets(): RenderCoverAssetsFn {
  const fn = requireCoverPipelineExport<RenderCoverAssetsFn>("renderCoverAssets");
  if (typeof fn !== "function") {
    throw new Error("renderCoverAssets is not exported from ../scripts/zh_tw_comics_pipeline");
  }
  return fn;
}

function requireBuildCoverEditablePptx(): BuildCoverEditablePptxFn {
  const fn = requireCoverPipelineExport<BuildCoverEditablePptxFn>("buildCoverEditablePptx");
  if (typeof fn !== "function") {
    throw new Error("buildCoverEditablePptx is not exported from ../scripts/zh_tw_comics_pipeline");
  }
  return fn;
}

function collectCoverBoxText(specs: Record<CoverPageId, TextSpec>): string {
  return COVER_PAGE_ORDER.flatMap((pageId) => (specs[pageId]?.boxes ?? []).map((box) => String(box.text ?? ""))).join(
    "\n",
  );
}

async function readPngDimensions(pngPath: string): Promise<{ width: number; height: number }> {
  const image = await loadImage(Buffer.from(await readFile(pngPath)));
  return { width: image.width, height: image.height };
}

function expectCoverFrontDimensions(dim: { width: number; height: number }): void {
  const canonicalCoverPdf = requireCoverPipelineExport<string>("DEFAULT_ZH_TW_COVER_CANONICAL_PDF");
  if (existsSync(canonicalCoverPdf)) {
    expect(dim).toEqual({ width: 5742, height: 8000 });
  } else {
    expect(dim.height).toBe(8000);
    expect([5748, 5749]).toContain(dim.width);
  }
}


async function writePptxZipEntry(
  pptxPath: string,
  entryPath: string,
  content: string | Uint8Array,
): Promise<void> {
  const zip = await new JSZip().loadAsync(await readFile(pptxPath));
  zip.file(entryPath, content);
  const out = await zip.generateAsync({ type: "nodebuffer" });
  await writeFile(pptxPath, out);
}

function extractPptxShapeXmlByName(slideXml: string, shapeName: string): string | undefined {
  for (const chunk of slideXml.split("<p:sp>")) {
    if (!chunk.includes(`name="${shapeName}"`)) continue;
    return `<p:sp>${chunk.split("</p:sp>")[0]}</p:sp>`;
  }
  return undefined;
}

function replacePptxShapeXml(slideXml: string, shapeName: string, replacement: string): string {
  const existing = extractPptxShapeXmlByName(slideXml, shapeName);
  if (!existing) throw new Error(`slide missing shape: ${shapeName}`);
  return slideXml.replace(existing, replacement);
}

function extractYamlPageBoxBlock(yamlText: string, pageId: string, boxId: string): string {
  const pageKey = `${pageId}:`;
  const pageStart = yamlText.indexOf(`  ${pageKey}`);
  if (pageStart < 0) throw new Error(`yaml missing page ${pageId}`);
  const nextPage = yamlText.indexOf("\n  Ch", pageStart + pageKey.length);
  const pageSlice = nextPage < 0 ? yamlText.slice(pageStart) : yamlText.slice(pageStart, nextPage);
  const boxMarker = `      - id: ${boxId}`;
  const boxStart = pageSlice.indexOf(boxMarker);
  if (boxStart < 0) throw new Error(`yaml missing ${pageId} box ${boxId}`);
  const boxTail = pageSlice.slice(boxStart);
  const nextBox = boxTail.indexOf("\n      - id:", boxMarker.length);
  return nextBox < 0 ? boxTail.trimEnd() : boxTail.slice(0, nextBox).trimEnd();
}

function replaceYamlPageBoxBlock(yamlText: string, pageId: string, boxId: string, newBlock: string): string {
  const block = extractYamlPageBoxBlock(yamlText, pageId, boxId);
  return yamlText.replace(block, newBlock.trimEnd());
}

async function writeYamlPageBoxBlockReplacement(
  yamlPath: string,
  pageId: string,
  boxId: string,
  newBlock: string,
): Promise<void> {
  const yamlText = await readFile(yamlPath, "utf8");
  await writeFile(yamlPath, replaceYamlPageBoxBlock(yamlText, pageId, boxId, newBlock));
}

/** Replaces a `text: |-` block body (indented content lines) inside one YAML box block. */
function patchYamlBoxLiteralTextBlock(block: string, indentedBody: string): string {
  const body = indentedBody.replace(/\n/g, "\n          ");
  if (block.includes("text: |-")) {
    return block.replace(/        text: \|-\n(?:          .+\n)+/, `        text: |-\n          ${body}\n`);
  }
  return block.replace(/        text: .+/, `        text: |-\n          ${body}`);
}

/** Sets align on a box block whether or not align was already present (idempotent). */
function ensureYamlBoxAlign(block: string, align: "left" | "center" | "right"): string {
  if (/\n        align: /.test(block)) {
    return block.replace(/\n        align: (?:left|center|right)/, `\n        align: ${align}`);
  }
  return `${block}\n        align: ${align}`;
}


/** Patches Ch5_B b004_body text in-place so line-based parseback still resolves page md. */
async function writeCh5BStaleB004BodyOpeningTrim(yamlPath: string): Promise<void> {
  let yamlText = await readFile(yamlPath, "utf8");
  const block = extractYamlPageBoxBlock(yamlText, "Ch5_B", "b004_body");
  const staleBlock = block.replace(
    CH5_B_PARSEBACK_STALE_B004_OPENING,
    CH5_B_PARSEBACK_STALE_B004_OPENING_TRIM,
  );
  if (staleBlock === block) throw new Error("Ch5_B b004_body stale opening trim did not apply");
  await writeFile(yamlPath, yamlText.replace(block, staleBlock));
}


/** Drops the solidarity line from Ch5_B page md so stale b004_body text is not a flat-md substring. */

/** Sets Ch5_B b004_body text to match trimmed page md (no YAML text-indent padding) for paragraph sync. */
async function writeCh5BStaleB004BodyAlignWithPageMd(yamlPath: string): Promise<void> {
  let yamlText = await readFile(yamlPath, "utf8");
  const block = extractYamlPageBoxBlock(yamlText, "Ch5_B", "b004_body");
  const indentedBody = CH5_B_PARSEBACK_STALE_B004_BODY_FOR_MD_SYNC.split("\n")
    .map((line) => `          ${line}`)
    .join("\n");
  const newTextBlock = `        text: |-\n${indentedBody}`;
  const staleBlock = block.replace(/        text: \|-[\s\S]*?(?=\n        [a-z_]+:|\n      - id:|\s*$)/, newTextBlock);
  if (staleBlock === block) throw new Error("Ch5_B b004_body text align with page md did not apply");
  await writeFile(yamlPath, yamlText.replace(block, staleBlock));
}

/** Drops the solidarity line from Ch5_B b004_body so stale box text matches trimmed page md. */
async function writeCh5BStaleB004BodyDropSolidarityLine(yamlPath: string): Promise<void> {
  let yamlText = await readFile(yamlPath, "utf8");
  const block = extractYamlPageBoxBlock(yamlText, "Ch5_B", "b004_body");
  const solidarityInBlock = `\n          ${CH5_B_PARSEBACK_SOLIDARITY_OLD}`;
  if (!block.includes(solidarityInBlock) && !block.includes(`          ${CH5_B_PARSEBACK_SOLIDARITY_OLD}`)) {
    throw new Error("Ch5_B b004_body solidarity line missing for stale parseback fixture");
  }
  const staleBlock = block
    .replace(solidarityInBlock, "")
    .replace(`          ${CH5_B_PARSEBACK_SOLIDARITY_OLD}`, "");
  await writeFile(yamlPath, yamlText.replace(block, staleBlock));
}
async function writeCh5BStalePageMdDropSolidarityLine(yamlPath: string): Promise<void> {
  let yamlText = await readFile(yamlPath, "utf8");
  const pageStart = yamlText.indexOf("  Ch5_B:");
  if (pageStart < 0) throw new Error("missing Ch5_B page");
  const mdMarker = yamlText.indexOf("    md: |-", pageStart);
  if (mdMarker < 0) throw new Error("missing Ch5_B md");
  const mdBodyStart = yamlText.indexOf("\n", mdMarker) + 1;
  const boxesMarker = yamlText.indexOf("    boxes:", pageStart);
  if (boxesMarker < 0) throw new Error("missing Ch5_B boxes");
  const mdSection = yamlText.slice(mdBodyStart, boxesMarker);
  const solidarityMdLine = `      ${CH5_B_PARSEBACK_SOLIDARITY_OLD}\n`;
  if (!mdSection.includes(solidarityMdLine)) {
    throw new Error("Ch5_B md solidarity line missing for stale parseback fixture");
  }
  const trimmedMd = mdSection.replace(solidarityMdLine, "");
  await writeFile(
    yamlPath,
    yamlText.slice(0, mdBodyStart) + trimmedMd + yamlText.slice(boxesMarker),
  );
}

/** Inserts an unrelated earlier md paragraph with the 做什麼 decoy before the b004 markdown block. */
async function writeCh5BStalePageMdInsertEarlierDecoyParagraph(yamlPath: string): Promise<void> {
  let yamlText = await readFile(yamlPath, "utf8");
  const pageStart = yamlText.indexOf("  Ch5_B:");
  if (pageStart < 0) throw new Error("missing Ch5_B page");
  const mdMarker = yamlText.indexOf("    md: |-", pageStart);
  if (mdMarker < 0) throw new Error("missing Ch5_B md");
  const mdBodyStart = yamlText.indexOf("\n", mdMarker) + 1;
  const boxesMarker = yamlText.indexOf("    boxes:", pageStart);
  if (boxesMarker < 0) throw new Error("missing Ch5_B boxes");
  const mdSection = yamlText.slice(mdBodyStart, boxesMarker);
  const b004MdLine = "      退一步想。在你想要「做什麼」背後，";
  const b004LineIndex = mdSection.indexOf(b004MdLine);
  if (b004LineIndex < 0) throw new Error("Ch5_B md missing b004 paragraph for decoy insert");
  const decoyBlock = `      ${CH5_B_PARSEBACK_EARLIER_MD_DECOY_PARAGRAPH}\n\n`;
  const nextMdSection = mdSection.slice(0, b004LineIndex) + decoyBlock + mdSection.slice(b004LineIndex);
  await writeFile(
    yamlPath,
    yamlText.slice(0, mdBodyStart) + nextMdSection + yamlText.slice(boxesMarker),
  );
}

/** Rewrites Ch5_B md b004 lines to exactly match stale box text for flat-md parseback replacement. */
async function writeCh5BStalePageMdAlignB004Paragraph(yamlPath: string): Promise<void> {
  let yamlText = await readFile(yamlPath, "utf8");
  const pageStart = yamlText.indexOf("  Ch5_B:");
  if (pageStart < 0) throw new Error("missing Ch5_B page");
  const mdMarker = yamlText.indexOf("    md: |-", pageStart);
  if (mdMarker < 0) throw new Error("missing Ch5_B md");
  const mdBodyStart = yamlText.indexOf("\n", mdMarker) + 1;
  const boxesMarker = yamlText.indexOf("    boxes:", pageStart);
  if (boxesMarker < 0) throw new Error("missing Ch5_B boxes");
  const mdSection = yamlText.slice(mdBodyStart, boxesMarker);
  const alignedBlock = CH5_B_PARSEBACK_STALE_B004_BODY_FOR_MD_SYNC.split("\n")
    .map((line) => `      ${line}`)
    .join("\n");
  const startMarker = "      問問自己：你「為何」想要它？";
  const endMarker = "      接著：你能不能創造一個同時滿足";
  const startIndex = mdSection.indexOf(startMarker);
  const endIndex = mdSection.indexOf(endMarker);
  if (startIndex < 0 || endIndex < 0) throw new Error("Ch5_B md missing b004 paragraph markers for align");
  const endLineEnd = mdSection.indexOf("\n", endIndex + endMarker.length);
  const sliceEnd = endLineEnd < 0 ? mdSection.length : endLineEnd + 1;
  const nextMdSection = mdSection.slice(0, startIndex) + `${alignedBlock}\n` + mdSection.slice(sliceEnd);
  await writeFile(
    yamlPath,
    yamlText.slice(0, mdBodyStart) + nextMdSection + yamlText.slice(boxesMarker),
  );
}

/** Drops the distant pol.is paragraph that also anchors on 雙方「為何」 so b004 md sync stays unambiguous. */
async function writeCh5BStalePageMdDropPolisParagraph(yamlPath: string): Promise<void> {
  let yamlText = await readFile(yamlPath, "utf8");
  const pageStart = yamlText.indexOf("  Ch5_B:");
  if (pageStart < 0) throw new Error("missing Ch5_B page");
  const mdMarker = yamlText.indexOf("    md: |-", pageStart);
  if (mdMarker < 0) throw new Error("missing Ch5_B md");
  const mdBodyStart = yamlText.indexOf("\n", mdMarker) + 1;
  const boxesMarker = yamlText.indexOf("    boxes:", pageStart);
  if (boxesMarker < 0) throw new Error("missing Ch5_B boxes");
  const mdSection = yamlText.slice(mdBodyStart, boxesMarker);
  const polisLine =
    "      但在 pol.is 與參與官的協助下，大家找到了能同時滿足雙方「為何」的新「做什麼」：就業金卡，把科學家與藝術家帶進台灣。\n";
  if (!mdSection.includes(polisLine)) {
    throw new Error("Ch5_B md pol.is paragraph missing for stale parseback fixture");
  }
  const trimmedMd = mdSection.replace(polisLine, "");
  await writeFile(
    yamlPath,
    yamlText.slice(0, mdBodyStart) + trimmedMd + yamlText.slice(boxesMarker),
  );
}
type ParsebackEmphasisSpan = { start: number; end: number; bold?: boolean; color?: string };
type ParsebackEmphasisEntry = ParsebackEmphasisSpan | { text?: string; color?: string; bold?: boolean };

type CompactYamlBox = {
  id: string;
  box?: number[];
  src?: number[];
  tight?: number[];
  font?: number;
  align?: string;
  emphasis?: ParsebackEmphasisEntry[];
  text?: string;
};

function readCompactYamlBox(yamlText: string, pageId: string, boxId: string): CompactYamlBox {
  const doc = Bun.YAML.parse(yamlText) as {
    pages: Record<string, { boxes?: CompactYamlBox[] }>;
  };
  const box = doc.pages[pageId]?.boxes?.find((entry) => entry.id === boxId);
  if (!box) throw new Error(`parsed yaml missing ${pageId}/${boxId}`);
  return box;
}
type PptxTextRun = { text: string; bold: boolean; hex?: string };

function extractPptxTextRunsFromShape(shapeXml: string): PptxTextRun[] {
  const runs: PptxTextRun[] = [];
  const re = /<a:r>[\s\S]*?<a:t>([^<]*)<\/a:t>[\s\S]*?<\/a:r>/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(shapeXml)) !== null) {
    const hexMatch = match[0].match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
    runs.push({
      text: match[1]!,
      bold: /\bb="1"/.test(match[0]),
      hex: hexMatch?.[1]?.toUpperCase(),
    });
  }
  return runs;
}

function rewritePptxShapeTextRuns(shapeXml: string, runs: PptxTextRun[]): string {
  let index = 0;
  return shapeXml.replace(/<a:r>[\s\S]*?<\/a:r>/g, (runXml) => {
    if (!runXml.includes("<a:t>")) return runXml;
    const spec = runs[index++];
    if (!spec) throw new Error("rewritePptxShapeTextRuns: not enough run specs");
    let next = runXml.replace(/<a:t>[^<]*<\/a:t>/, `<a:t>${spec.text}</a:t>`);
    if (spec.bold) {
      if (!/\bb="1"/.test(next)) {
        next = next.replace(/<a:rPr([^>]*)>/, (_m, attrs: string) => {
          const trimmed = attrs.trim();
          return trimmed ? `<a:rPr ${trimmed} b="1">` : `<a:rPr b="1">`;
        });
      }
    } else {
      next = next.replace(/\s*b="1"/, "");
    }
    return next;
  });
}

/** Simulates Keynote/PowerPoint resave: same text, no bold or hex color on runs. */
function rewritePptxShapeToPlainUnstyledRuns(shapeXml: string): string {
  const runs = extractPptxTextRunsFromShape(shapeXml);
  const plain = runs.map((run) => ({ text: run.text, bold: false }));
  let stripped = rewritePptxShapeTextRuns(shapeXml, plain);
  stripped = stripped.replace(/<a:srgbClr\s+val="[0-9A-Fa-f]{6}"\s*\/>/g, "");
  return stripped;
}

const COVER_SPINE_AUTHORS_DESKTOP_PARSED_TEXT = "唐\n鳳\n、\n格\n林";
const COVER_SPINE_AUTHORS_PRE_MERGE_LITERAL = "唐鳳\n格林";


/** Mirrors parsePptxTextBody: paragraph boundaries and <a:br/> become newlines. */
function extractPptxShapeTextWithBreaks(shapeXml: string): string {
  const txBody = shapeXml.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/)?.[1];
  if (!txBody) return "";
  const paragraphs = [...txBody.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)].map((match) => match[1]!);
  const chunks: string[] = [];
  for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex += 1) {
    if (paragraphIndex > 0) chunks.push("\n");
    const paragraph = paragraphs[paragraphIndex]!;
    const tokens = paragraph.match(/<a:r>[\s\S]*?<\/a:r>|<a:br\s*\/>/g) ?? [];
    for (const token of tokens) {
      if (token.includes("<a:br")) {
        chunks.push("\n");
        continue;
      }
      const text = token.match(/<a:t>([^<]*)<\/a:t>/)?.[1] ?? "";
      chunks.push(text);
    }
  }
  return chunks.join("");
}

/** Simulates PowerPoint resave: shape id renamed to visible text; spine authors split per character/line. */
function simulateDesktopResavedCoverSpineAuthorsShape(shapeXml: string): string {
  const lines = ["唐", "鳳", "、", "格", "林"];
  const txBodyMatch = shapeXml.match(/<p:txBody>[\s\S]*?<\/p:txBody>/);
  if (!txBodyMatch) throw new Error("simulateDesktopResavedCoverSpineAuthorsShape: missing p:txBody");
  const txBody = txBodyMatch[0];
  const bodyPr = txBody.match(/<a:bodyPr[\s\S]*?<\/a:bodyPr>/)?.[0];
  const lstStyle = txBody.match(/<a:lstStyle\s*\/>/)?.[0] ?? "<a:lstStyle/>";
  const firstParagraph = txBody.match(/<a:p>[\s\S]*?<\/a:p>/)?.[0];
  if (!bodyPr || !firstParagraph) {
    throw new Error("simulateDesktopResavedCoverSpineAuthorsShape: missing bodyPr or paragraph template");
  }
  const pPr = firstParagraph.match(/<a:pPr[\s\S]*?<\/a:pPr>/)?.[0] ?? "<a:pPr/>";
  const firstRun = firstParagraph.match(/<a:r>[\s\S]*?<\/a:r>/)?.[0];
  if (!firstRun) throw new Error("simulateDesktopResavedCoverSpineAuthorsShape: missing a:r template");
  const runTemplate = firstRun.replace(/<a:t>[^<]*<\/a:t>/, "<a:t>__TEXT__</a:t>");
  const runXml: string[] = [];
  for (let index = 0; index < lines.length; index += 1) {
    if (index > 0) runXml.push("<a:br/>");
    runXml.push(runTemplate.replace("__TEXT__", lines[index]!));
  }
  const endPara =
    firstParagraph.match(/<a:endParaRPr[\s\S]*?\/>/)?.[0] ?? '<a:endParaRPr lang="en-US" sz="1800" dirty="0"/>';
  const newTxBody = `<p:txBody>${bodyPr}${lstStyle}<a:p>${pPr}${runXml.join("")}${endPara}</a:p></p:txBody>`;
  let next = shapeXml.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, newTxBody);
  next = next.replace(/name="spine_authors"/, 'name="唐"');
  return next;
}

function pptxRunsToUtf16Ranges(runs: PptxTextRun[]): ParsebackEmphasisSpan[] {
  const spans: ParsebackEmphasisSpan[] = [];
  let offset = 0;
  const fullText = runs.map(r => r.text).join("");
  
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

  for (const run of runs) {
    const start = offset;
    offset += run.text.length;
    if (run.bold || run.hex) {
      if (run.hex) {
        const span: any = { text: run.text, color: `#${run.hex.toUpperCase()}` };
        if (run.bold) span.bold = true;
        const occurrence = occurrenceForRange(fullText, run.text, start);
        if (occurrence > 1) span.occurrence = occurrence;
        spans.push(span);
      } else {
        spans.push({ start, end: offset, bold: true });
      }
    }
  }
  return spans;
}




async function writeMinimalYamlPageBoxPatch(
  yamlPath: string,
  pageId: string,
  boxId: string,
  patch: Record<string, unknown>,
): Promise<void> {
  const root = Bun.YAML.parse(await readFile(yamlPath, "utf8")) as {
    schema: number;
    pages: Record<string, { boxes?: Array<Record<string, unknown>> }>;
  };
  const page = structuredClone(root.pages[pageId]!);
  const boxes = page.boxes ?? [];
  const index = boxes.findIndex((entry) => String(entry.id) === boxId);
  if (index < 0) throw new Error(`missing YAML box ${pageId}/${boxId}`);
  boxes[index] = { ...boxes[index]!, ...patch };
  page.boxes = boxes;
  await writeFile(
    yamlPath,
    Bun.YAML.stringify({ schema: root.schema, pages: { ...root.pages, [pageId]: page } }),
  );
}

function pptxTestPxToInches(px: number): number {
  return px / PPTX_TEST_RENDER_DPI;
}

function pptxTestPxToPoints(px: number): number {
  return (px * 72) / PPTX_TEST_RENDER_DPI;
}

function pptxTestHexColor(color: unknown): string {
  const hex = typeof color === "string" ? color.replace(/^#/, "") : "";
  return /^[0-9a-fA-F]{6}$/.test(hex) ? hex.toUpperCase() : "222222";
}

function pptxTestFontFace(role: string): string {
  return role === "title" ? "jf蘭陽黑體 1.0" : "jf蘭陽明體2.0";
}

type PptxTestRunTextValue = { text: string; options?: { bold?: boolean; color?: string } };
type PptxTestEmphasisMap = {
  bold: Uint8Array;
  colors: Array<string | undefined>;
  hasBold: boolean;
  hasColor: boolean;
};



function pptxTestEmphasisRange(text: string, span: EmphasisSpec): { start: number; end: number } | null {
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

function pptxTestEmphasisStyleMap(
  text: string,
  emphasis: EmphasisSpec[] | undefined,
): PptxTestEmphasisMap | null {
  if (!emphasis?.length) return null;
  let resolved: PptxTestEmphasisMap | null = null;
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
    const rawColor = typeof span.color === "string" ? span.color : undefined;
    const color =
      rawColor && /^#?[0-9a-fA-F]{6}$/.test(rawColor.replace(/^#/, ""))
        ? rawColor.startsWith("#")
          ? rawColor
          : `#${rawColor}`
        : undefined;
    if (!bold && !color) continue;
    const range = pptxTestEmphasisRange(text, span);
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
  const finalMap = resolved as PptxTestEmphasisMap | null;
  return finalMap && (finalMap.hasBold || finalMap.hasColor) ? finalMap : null;
}

function pptxTestTextValue(box: BoxSpec): string | PptxTestRunTextValue[] {
  const text = String(box.text ?? "");
  const emphasis = pptxTestEmphasisStyleMap(text, box.emphasis);
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
    if (run.color) options.color = pptxTestHexColor(run.color);
    return Object.keys(options).length > 0 ? { text: run.text, options } : { text: run.text };
  });
}

async function buildEditablePptxFromYaml(yamlPath: string, outputPath: string): Promise<string> {
  const specs = await loadYamlTextSpecs(yamlPath);
  const pptx = new PptxGenJS();
  const slideWidthIn = 5749 / PPTX_TEST_RENDER_DPI;
  const slideHeightIn = 8000 / PPTX_TEST_RENDER_DPI;
  pptx.defineLayout({ name: "ZH_TW_COMIC_PAGE_TEST", width: slideWidthIn, height: slideHeightIn });
  pptx.layout = "ZH_TW_COMIC_PAGE_TEST";
  const backgroundDir = join(dirname(outputPath), "pptx-backgrounds");
  await mkdir(backgroundDir, { recursive: true });
  await mkdir(dirname(outputPath), { recursive: true });
  for (const pageId of PAGE_ORDER) {
    const spec = specs[pageId as PageId];
    const slide = pptx.addSlide();
    const backgroundPath = await renderPptxBackground(spec, join(backgroundDir, `${pageId}.png`));
    slide.background = { path: backgroundPath };
    for (const box of spec.boxes ?? []) {
      const rect = box.box?.map(Number);
      if (!rect || rect.length !== 4 || !rect.every(Number.isFinite)) continue;
      const [x0, y0, x1, y1] = rect;
      const role = String(box.font_role ?? box.ocr_type ?? "body");
      slide.addText(pptxTestTextValue(box), {
        objectName: String(box.id),
        x: pptxTestPxToInches(x0!),
        y: pptxTestPxToInches(y0!),
        w: pptxTestPxToInches(x1! - x0!),
        h: pptxTestPxToInches(y1! - y0!),
        fontFace: pptxTestFontFace(role),
        fontSize: pptxTestPxToPoints(Number(box.font_size ?? 100)),
        color: pptxTestHexColor(box.color),
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

function extractPptxSrgbClrRuns(shapeXml: string): { text: string; hex: string | undefined }[] {
  const runs: { text: string; hex: string | undefined }[] = [];
  const re = /<a:r>[\s\S]*?<\/a:r>/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(shapeXml)) !== null) {
    const runXml = match[0];
    const textMatch = runXml.match(/<a:t>([^<]*)<\/a:t>/);
    if (!textMatch) continue;
    const hexMatch = runXml.match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
    runs.push({ text: textMatch[1]!, hex: hexMatch?.[1]?.toUpperCase() });
  }
  return runs;
}

function setPptxRunText(runXml: string, text: string): string {
  return runXml.replace(/<a:t>[^<]*<\/a:t>/, `<a:t>${text}</a:t>`);
}

function applyPptxRunHexColor(runXml: string, hex: string): string {
  const hexVal = hex.replace(/^#/, "").toUpperCase();
  if (runXml.includes("<a:srgbClr")) {
    return runXml.replace(/<a:srgbClr\s+val="[0-9A-Fa-f]{6}"/, `<a:srgbClr val="${hexVal}"`);
  }
  return runXml.replace(/<a:rPr([^>]*)>/, (_m, attrs: string) => {
    const trimmed = String(attrs).trim();
    const open = trimmed ? `<a:rPr ${trimmed}>` : `<a:rPr>`;
    return `${open}<a:solidFill><a:srgbClr val="${hexVal}"/></a:solidFill>`;
  });
}

function applyPptxRunBold(runXml: string, bold: boolean): string {
  if (bold) {
    if (!/\bb="1"/.test(runXml)) {
      return runXml.replace(/<a:rPr([^>]*)>/, (_m, attrs: string) => {
        const trimmed = String(attrs).trim();
        return trimmed ? `<a:rPr ${trimmed} b="1">` : `<a:rPr b="1">`;
      });
    }
    return runXml;
  }
  return runXml.replace(/\s*b="1"/, "");
}

function rewritePptxShapeRunSubstringStyle(
  shapeXml: string,
  substring: string,
  style: { hex?: string; bold?: boolean },
): string {
  let replaced = false;
  return shapeXml.replace(/<a:r>[\s\S]*?<\/a:r>/g, (runXml) => {
    if (replaced || !runXml.includes("<a:t>")) return runXml;
    const textMatch = runXml.match(/<a:t>([^<]*)<\/a:t>/);
    if (!textMatch) return runXml;
    const fullText = textMatch[1]!;
    const idx = fullText.indexOf(substring);
    if (idx < 0) return runXml;
    replaced = true;

    const segments: { text: string; target: boolean }[] = [];
    const before = fullText.slice(0, idx);
    if (before) segments.push({ text: before, target: false });
    segments.push({ text: substring, target: true });
    const after = fullText.slice(idx + substring.length);
    if (after) segments.push({ text: after, target: false });

    return segments
      .map(({ text, target }) => {
        let next = setPptxRunText(runXml, text);
        if (target) {
          if (style.bold !== undefined) next = applyPptxRunBold(next, style.bold);
          if (style.hex) next = applyPptxRunHexColor(next, style.hex);
        }
        return next;
      })
      .join("");
  });
}

function rewritePptxShapeColorRun(shapeXml: string, runText: string, hex: string): string {
  return rewritePptxShapeRunSubstringStyle(shapeXml, runText, { hex });
}

function rewritePptxShapeBoldColorRun(
  shapeXml: string,
  runText: string,
  hex: string,
  bold = true,
): string {
  return rewritePptxShapeRunSubstringStyle(shapeXml, runText, { hex, bold });
}


async function writeMinimalYamlPageSizes(
  yamlPath: string,
  sizes: { source_size?: unknown; target_size?: unknown },
  pageId = "Overview",
): Promise<void> {
  const root = Bun.YAML.parse(await readFile(ZH_TW_YAML_PATH, "utf8")) as {
    schema: number;
    pages: Record<string, Record<string, unknown>>;
  };
  const page = { ...root.pages[pageId]! };
  if ("source_size" in sizes) page.source_size = sizes.source_size;
  if ("target_size" in sizes) page.target_size = sizes.target_size;
  await writeFile(
    yamlPath,
    Bun.YAML.stringify({ schema: root.schema, pages: { ...root.pages, [pageId]: page } }),
  );
}

function yamlBoxComparableFields(box: BoxSpec) {
  return {
    id: box.id,
    ocr_type: box.ocr_type,
    font_role: box.font_role,
    text: box.text,
    source_box: box.source_box,
    box: box.box,
    tight_box: box.tight_box,
    font_size: box.font_size,
    line_spacing: box.line_spacing,
    color: box.color,
    align: box.align,
    emphasis: box.emphasis,
  };
}

function yamlPatchComparableFields(patch: { id: string; source_box?: number[] | readonly number[]; box?: number[] | readonly number[] }) {
  return {
    id: patch.id,
    source_box: patch.source_box,
    box: patch.box,
  };
}


function hexBrightness(hex: string): number {
  const normalized = hex.replace(/^#/, "");
  if (normalized.length !== 6) return 0;
  const r = Number.parseInt(normalized.slice(0, 2), 16);
  const g = Number.parseInt(normalized.slice(2, 4), 16);
  const b = Number.parseInt(normalized.slice(4, 6), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

function luma(r: number, g: number, b: number): number {
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

const EMPHASIS_COLOR_COMPASS_RGB = { r: 0x4f, g: 0x8f, b: 0x00 } as const;
const EMPHASIS_COLOR_RGB_TOLERANCE = 48;
const MIN_EMPHASIS_GREEN_PIXEL_CLUSTER = 40;

function countPixelsCloseToRgb(
  canvas: Canvas,
  box: readonly [number, number, number, number],
  target: { r: number; g: number; b: number },
  tolerance = EMPHASIS_COLOR_RGB_TOLERANCE,
): number {
  const [x0, y0, x1, y1] = box;
  const w = Math.max(0, x1 - x0);
  const h = Math.max(0, y1 - y0);
  if (w === 0 || h === 0) return 0;
  const ctx = canvas.getContext("2d");
  const data = ctx.getImageData(x0, y0, w, h).data;
  let hits = 0;
  for (let i = 0; i < data.length; i += 4) {
    const r = data[i]!;
    const g = data[i + 1]!;
    const b = data[i + 2]!;
    const a = data[i + 3]!;
    if (a < 32) continue;
    if (luma(r, g, b) > 244) continue;
    if (Math.abs(r - target.r) > tolerance) continue;
    if (Math.abs(g - target.g) > tolerance) continue;
    if (Math.abs(b - target.b) > tolerance) continue;
    if (g <= r + 12 || g <= b + 12) continue;
    hits += 1;
  }
  return hits;
}

async function renderSyntheticCh4AEmphasisColorVector(
  root: string,
): Promise<{ canvas: Canvas; boxRect: [number, number, number, number] }> {
  const templatePath = join(root, "template.png");
  const outputPath = join(root, "vector.png");
  const width = 800;
  const height = 600;
  const boxRect: [number, number, number, number] = [40, 380, 760, 560];
  await writeWhitePng(templatePath, width, height);
  const text =
    "那我們該如何衡量我們所重視的——我們的「羅盤」——而不被誤導？";
  const spec: TextSpec = {
    schema_version: 1,
    image: "Ch4_A_emphasis_render",
    template_image: templatePath,
    output_image: outputPath,
    source_size: [width, height],
    target_size: [width, height],
    scale: [1, 1],
    boxes: [
      {
        id: CH4_A_EMPHASIS_COLOR_BOX,
        ocr_type: "text",
        font_role: "body",
        text,
        source_box: boxRect,
        box: boxRect,
        tight_box: boxRect,
        font_size: 28,
        color: "#202020",
        align: "left",
        emphasis: [{ text: EMPHASIS_COLOR_COMPASS_SPAN, color: EMPHASIS_COLOR_COMPASS }],
      },
    ],
  };
  verifyTextSpec(spec, [CH4_A_EMPHASIS_COLOR_BOX]);
  await renderPage(spec, outputPath, "vector");
  const canvas = await loadCanvasFromFile(outputPath);
  return { canvas, boxRect };
}

function findBoxByText(boxes: BoxSpec[] | undefined, text: string): BoxSpec | undefined {
  return boxes?.find((box) => box.text === text);
}

function findBoxById(boxes: BoxSpec[] | undefined, id: string): BoxSpec | undefined {
  return boxes?.find((box) => box.id === id);
}

function boxRect(box: BoxSpec | undefined): [number, number, number, number] | undefined {
  if (!box?.box || box.box.length < 4) return undefined;
  const [x0, y0, x1, y1] = box.box.map(Number);
  return [x0!, y0!, x1!, y1!];
}

function intersectionArea(
  a: readonly [number, number, number, number],
  b: readonly [number, number, number, number],
): number {
  const x0 = Math.max(a[0], b[0]);
  const y0 = Math.max(a[1], b[1]);
  const x1 = Math.min(a[2], b[2]);
  const y1 = Math.min(a[3], b[3]);
  return Math.max(0, x1 - x0) * Math.max(0, y1 - y0);
}
function verticalOverlapArea(
  a: readonly [number, number, number, number],
  b: readonly [number, number, number, number],
): number {
  const yOverlap = Math.max(0, Math.min(a[3], b[3]) - Math.max(a[1], b[1]));
  if (yOverlap <= 0) return 0;
  const xOverlap = Math.max(0, Math.min(a[2], b[2]) - Math.max(a[0], b[0]));
  return xOverlap * yOverlap;
}

function isDummyUnitSourceBox(box: BoxSpec | undefined): boolean {
  const sb = box?.source_box?.map(Number) ?? [];
  return sb.length === 4 && sb[0] === 0 && sb[1] === 0 && sb[2]! <= 1 && sb[3]! <= 1;
}

function imagePatchRegionCount(spec: TextSpec): number {
  const patches = spec.image_patches;
  if (Array.isArray(patches)) return patches.length;
  const regions = spec.image_patch_regions;
  if (Array.isArray(regions)) return regions.length;
  return 0;
}

function countMarkdownIllustrationRefs(md: string | undefined): number {
  if (!md) return 0;
  return [...md.matchAll(/!\[[^\]]*\]\(img-\d/g)].length;
}

async function loadOcrRawImageCount(pageId: string): Promise<number> {
  const raw = JSON.parse(await readFile(join(RAW_OCR_DIR, `${pageId}.json`), "utf8"));
  return raw.ocr?.pages?.[0]?.images?.length ?? 0;
}

async function loadCommittedTextSpec(pageId: string): Promise<TextSpec> {
  const raw = await readFile(join(TEXT_SPEC_DIR, `${pageId}.json`), "utf8");
  return JSON.parse(raw) as TextSpec;
}

function countDarkPixels(
  canvas: Canvas,
  box: readonly [number, number, number, number],
): number {
  const [x0, y0, x1, y1] = box;
  const w = Math.max(0, x1 - x0);
  const h = Math.max(0, y1 - y0);
  if (w === 0 || h === 0) return 0;
  const ctx = canvas.getContext("2d");
  const data = ctx.getImageData(x0, y0, w, h).data;
  let dark = 0;
  for (let i = 0; i < data.length; i += 4) {
    if (luma(data[i]!, data[i + 1]!, data[i + 2]!) < DARK_LUMA_MAX) dark++;
  }
  return dark;
}

async function loadCanvasFromFile(path: string): Promise<Canvas> {
  const image = await loadImage(path);
  const canvas = createCanvas(image.width, image.height);
  const ctx = canvas.getContext("2d");
  ctx.drawImage(image, 0, 0);
  return canvas;
}

async function writeWhitePng(path: string, width: number, height: number): Promise<void> {
  const canvas = createCanvas(width, height);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "white";
  ctx.fillRect(0, 0, width, height);
  await writeFile(path, canvas.toBuffer("image/png"));
}

type SyntheticFixture = {
  spec: Record<string, unknown>;
  templatePath: string;
  sourcePath: string;
  root: string;
};

async function writeSyntheticDiffFixture(
  options: {
    pageId?: string;
    inkPatch?: boolean;
    scaleX?: number;
    scaleY?: number;
  } = {},
): Promise<{
  textDir: string;
  outputRoot: string;
  reviewDir: string;
  pageId: string;
  outName: string;
  fixture: SyntheticFixture;
}> {
  const pageId = options.pageId ?? "Synthetic";
  const inkPatch = options.inkPatch ?? false;
  const root = await mkdtemp(join(tmpdir(), "zh-tw-diff-"));
  const textDir = join(root, "text");
  const outputRoot = join(root, "out");
  const reviewDir = join(root, "review");
  await mkdir(textDir, { recursive: true });
  await mkdir(outputRoot, { recursive: true });
  await mkdir(join(reviewDir, "diffs"), { recursive: true });

  const sourceSize: [number, number] = [200, 150];
  const targetSize: [number, number] = [400, 300];
  const sourceBox = [30, 40, 90, 80];
  const scaleX = options.scaleX ?? targetSize[0] / sourceSize[0];
  const scaleY = options.scaleY ?? targetSize[1] / sourceSize[1];
  const targetBox = [
    Math.round(sourceBox[0]! * scaleX),
    Math.round(sourceBox[1]! * scaleY),
    Math.round(sourceBox[2]! * scaleX),
    Math.round(sourceBox[3]! * scaleY),
  ];

  const templatePath = join(root, "template.png");
  const sourcePath = join(root, "source.jpg");
  await writeWhitePng(templatePath, targetSize[0], targetSize[1]);

  const sourceCanvas = createCanvas(sourceSize[0], sourceSize[1]);
  const sctx = sourceCanvas.getContext("2d");
  sctx.fillStyle = "white";
  sctx.fillRect(0, 0, sourceSize[0], sourceSize[1]);
  if (inkPatch) {
    sctx.fillStyle = `rgb(${ORACLE_PATCH_COLOR.r},${ORACLE_PATCH_COLOR.g},${ORACLE_PATCH_COLOR.b})`;
    sctx.fillRect(sourceBox[0]!, sourceBox[1]!, sourceBox[2]! - sourceBox[0]!, sourceBox[3]! - sourceBox[1]!);
  }
  await writeFile(sourcePath, sourceCanvas.toBuffer("image/jpeg", 95));

  const outName = pageId === "Synthetic" ? "Synthetic.png" : "Diag.png";
  const spec: Record<string, unknown> = {
    schema_version: 1,
    image: pageId,
    source_image: sourcePath,
    template_image: templatePath,
    output_image: join(outputRoot, outName),
    source_size: [...sourceSize],
    target_size: [...targetSize],
    scale: [scaleX, scaleY],
    boxes: [
      {
        id: "b001_body",
        ocr_type: "text",
        font_role: "body",
        text: SYNTHETIC_TEXT,
        source_box: sourceBox,
        box: targetBox,
        tight_box: targetBox,
        font_size: 24,
        color: "#141414",
        align: "left",
      },
    ],
  };

  verifyTextSpec(spec, ["b001_body"]);
  await writeFile(join(textDir, `${pageId}.json`), JSON.stringify(spec), "utf8");
  await renderPage(spec, join(outputRoot, outName), "vector");

  const fixture: SyntheticFixture = { spec, templatePath, sourcePath, root };
  return { textDir, outputRoot, reviewDir, pageId, outName, fixture };
}

async function writeSyntheticTableFixture(): Promise<{
  spec: Record<string, unknown>;
  outputRoot: string;
  templatePath: string;
}> {
  const root = await mkdtemp(join(tmpdir(), "zh-tw-table-"));
  const outputRoot = join(root, "out");
  await mkdir(outputRoot, { recursive: true });
  const templatePath = join(root, "template.png");
  await writeWhitePng(templatePath, 600, 400);

  const spec: Record<string, unknown> = {
    schema_version: 1,
    image: "TableGutter",
    source_image: join(root, "source.jpg"),
    template_image: templatePath,
    output_image: join(outputRoot, "TablePage.png"),
    source_size: [600, 400],
    target_size: [600, 400],
    scale: [1, 1],
    boxes: [
      {
        id: "table001",
        ocr_type: "table",
        font_role: "body",
        text: TABLE_TEXT,
        source_box: [...TABLE_BOX],
        box: [...TABLE_BOX],
        tight_box: [...TABLE_BOX],
        font_size: 28,
        color: "#121212",
        align: "left",
      },
    ],
  };
  verifyTextSpec(spec, ["table001"]);
  return { spec, outputRoot, templatePath };
}

describe("Bun canvas foundation", () => {
  test("canvas runtime loads, registers a real CJK font, measures text, and exposes pixels", async () => {
    const probe = await probeCanvasRuntime("好啦關懷六力");
    expect(probe.fontPath).toMatch(/\.(otf|ttf|ttc)$/i);
    expect(probe.registered).toBe(true);
    expect(probe.width).toBeGreaterThan(100);
    expect(probe.ascent).toBeGreaterThan(20);
    expect(probe.darkPixels).toBeGreaterThan(1000);
  });

  test("manifest constants name exactly the thirteen comic source pages", () => {
    expect([...PAGE_ORDER]).toEqual([
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
    ]);
    expect(Object.values(SOURCE_IMAGE_NAMES)).not.toContain("gpt-value-correlation-tw.jpg");
    expect(TEMPLATE_IMAGE_NAMES.Overview).toBe("Overview.png");
  });

  test("font discovery prefers user-local jf-lanyang CJK fonts before system fallbacks", () => {
    const font = findFont("body");
    expect(font).toBeTruthy();
    expect(font ?? "").toMatch(/(jf-lanyang|PingFang|STHeiti)/i);
  });

  test("measureRegisteredText returns positive width for CJK sample", () => {
    const metrics = measureRegisteredText("測試");
    expect(metrics.width).toBeGreaterThan(0);
  });
});

describe("comics-only manifest", () => {
  test("buildComicsOnlyManifest maps exactly thirteen comic source images", () => {
    const manifest = buildComicsOnlyManifest();
    expect(Object.keys(manifest).sort()).toEqual(
      [
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
        "Overview",
      ].sort(),
    );
    expect(Object.keys(manifest)).toHaveLength(13);
  });

  test("buildComicsOnlyManifest excludes gpt-value-correlation-tw.jpg", () => {
    const manifest = buildComicsOnlyManifest();
    const joined = Object.values(manifest).join(" ").toLowerCase();
    expect(joined).not.toContain("gpt-value-correlation-tw.jpg");
    for (const value of Object.values(manifest)) {
      const name = value.split("/").pop()?.toLowerCase() ?? "";
      expect(name).not.toBe("gpt-value-correlation-tw.jpg");
    }
  });
});

describe("text/spec contracts", () => {
  test("verifyTextSpec fails when a required box id is missing", () => {
    expect(() =>
      verifyTextSpec({ image: "Ch1_A", boxes: [{ id: "title", text: "標題" }] }, [
        "title",
        "caption",
      ]),
    ).toThrow(VerificationError);
    expect(() =>
      verifyTextSpec({ image: "Ch1_A", boxes: [{ id: "title", text: "標題" }] }, [
        "title",
        "caption",
      ]),
    ).toThrow(/caption/i);
  });

  test("verifyTextSpec fails on simplified trap characters", () => {
    expect(() =>
      verifyTextSpec({ image: "Ch2_A", boxes: [{ id: "line1", text: "这是什么" }] }, ["line1"]),
    ).toThrow(VerificationError);
    expect(() =>
      verifyTextSpec({ image: "Ch2_A", boxes: [{ id: "line1", text: "这是什么" }] }, ["line1"]),
    ).toThrow(/simplified|trap|这|什么|简体/i);
  });

  test("wrapTextPreserveBreaks preserves explicit line breaks", () => {
    const family = registerFontForRole("body");
    expect(family).toBeTruthy();
    const wrapped = wrapTextPreserveBreaks(
      "第一行\n第二行很長很長",
      { fontFamily: family!, size: 24 },
      80,
    );
    const lines = wrapped.split("\n");
    expect(lines[0]).toBe("第一行");
    expect(lines.some((line, index) => index > 0 && line.includes("第二"))).toBe(true);
  });

  test("wrapTextPreserveBreaks keeps fullwidth colon attached to short CJK title (解方：)", () => {
    const titleSize = 48;
    const family = registerFontForRole("title");
    expect(family).toBeTruthy();
    const font = { fontFamily: family!, size: titleSize };
    const maxWidth = measureRegisteredText("解方", "title", titleSize).width + 1;
    const wrapped = wrapTextPreserveBreaks("解方：", font, maxWidth);
    expect(wrapped).toBe("解方：");
    expect(wrapped).not.toContain("\n：");
  });

  test("wrapTextPreserveBreaks keeps ASCII exclamation attached to CJK (工資!)", () => {
    const bodySize = 24;
    const family = registerFontForRole("body");
    expect(family).toBeTruthy();
    const font = { fontFamily: family!, size: bodySize };
    const maxWidth = measureRegisteredText("工資", "body", bodySize).width + 1;
    const wrapped = wrapTextPreserveBreaks("工資!", font, maxWidth);
    expect(wrapped).toBe("工資!");
    expect(wrapped).not.toContain("\n!");
  });

  test("wrapTextPreserveBreaks wraps at earlier safe whitespace before ASCII punctuation (abc def!)", () => {
    const bodySize = 24;
    const family = registerFontForRole("body");
    expect(family).toBeTruthy();
    const font = { fontFamily: family!, size: bodySize };
    const text = "abc def!";
    const maxWidth = measureRegisteredText("abc def", "body", bodySize).width + 1;
    expect(measureRegisteredText("abc def", "body", bodySize).width).toBeLessThanOrEqual(maxWidth);
    expect(measureRegisteredText(text, "body", bodySize).width).toBeGreaterThan(maxWidth);
    const wrapped = wrapTextPreserveBreaks(text, font, maxWidth);
    expect(wrapped).toBe("abc\ndef!");
    expect(wrapped).not.toBe("abc def!");
    expect(wrapped).not.toBe("abc def\n!");
  });

  test("fitFontSizeInBox returns a size within max and min bounds", () => {
    const fitted = fitFontSizeInBox("這是一段需要縮放的文字", [0, 0, 200, 80], 42, 8, "body", 3);
    expect(fitted).toBeLessThanOrEqual(42);
    expect(fitted).toBeGreaterThanOrEqual(8);
  });

  test("cleanOcrContent strips markdown table syntax", () => {
    const cleaned = cleanOcrContent(TABLE_OCR);
    expect(cleaned).toContain("指標");
    expect(cleaned).toContain("點擊");
    expect(cleaned).toContain("標題黨");
    expect(cleaned).not.toContain("|");
    expect(cleaned).not.toContain("---");
  });

  ocrTest("specFromOcrRaw classifies Ch6_B opening contrast lines as title font_role", async () => {
    const rawPath = join(RAW_OCR_DIR, "Ch6_B.json");
    const raw = JSON.parse(await readFile(rawPath, "utf8"));
    const spec = await specFromOcrRaw(raw);
    const boxes = spec.boxes ?? [];
    const dontBox = boxes.find((box) => box.text === "不要：一體適用的霸主");
    const doBox = boxes.find((box) => box.text === "而要：在地解方給在地人");
    expect(dontBox).toBeDefined();
    expect(doBox).toBeDefined();
    expect(dontBox?.font_role).toBe("title");
    expect(doBox?.font_role).toBe("title");
    expect(dontBox?.id).toMatch(/_title$/);
    expect(doBox?.id).toMatch(/_title$/);
  });

  ocrTest("specFromOcrRaw preserves raw block font_size override over auto-fit", async () => {
    const rawPath = join(RAW_OCR_DIR, "Ch1_A.json");
    const baseline = JSON.parse(await readFile(rawPath, "utf8"));
    const baselineSpec = await specFromOcrRaw(baseline);
    const baselineBox = baselineSpec.boxes?.find((box) => box.text === "覺察力");
    expect(baselineBox?.font_size).toBeDefined();

    const raw = JSON.parse(await readFile(rawPath, "utf8"));
    const overrideSize = 37;
    const blocks = raw.ocr.pages[0].blocks as Array<{ content?: string; font_size?: number }>;
    const titleBlock = blocks.find((block) => String(block.content ?? "").includes("覺察力"));
    expect(titleBlock).toBeDefined();
    titleBlock!.font_size = overrideSize;

    const spec = await specFromOcrRaw(raw);
    const box = spec.boxes?.find((entry) => entry.text === "覺察力");
    expect(baselineBox!.font_size).not.toBe(overrideSize);
    expect(box?.font_size).toBe(overrideSize);
  });

  ocrTest("specFromOcrRaw preserves raw block line_spacing on generated BoxSpec", async () => {
    const raw = JSON.parse(await readFile(join(RAW_OCR_DIR, "Ch1_A.json"), "utf8"));
    const overrideSpacing = 11;
    const blocks = raw.ocr.pages[0].blocks as Array<{ content?: string; line_spacing?: number }>;
    const titleBlock = blocks.find((block) => String(block.content ?? "").includes("覺察力"));
    expect(titleBlock).toBeDefined();
    titleBlock!.line_spacing = overrideSpacing;

    const spec = await specFromOcrRaw(raw);
    const box = spec.boxes?.find((entry) => entry.text === "覺察力");
    expect(box).toBeDefined();
    let spacing: unknown;
    if (box && typeof box === "object" && "line_spacing" in box) {
      spacing = box.line_spacing;
    }
    expect(spacing).toBe(overrideSpacing);
  });
  ocrTest("specFromOcrRaw preserves raw block align override on generated BoxSpec", async () => {
    const bubbleText = "我只是聽命令／隨誘因／跟著大家走。";
    const rawPath = join(RAW_OCR_DIR, "Ch2_A.json");
    const baseline = JSON.parse(await readFile(rawPath, "utf8"));
    const baselineSpec = await specFromOcrRaw(baseline);
    const baselineBox = findBoxByText(baselineSpec.boxes, bubbleText);
    expect(baselineBox?.align).toBe("left");

    const raw = JSON.parse(await readFile(rawPath, "utf8"));
    const blocks = raw.ocr.pages[0].blocks as Array<{ content?: string; align?: string }>;
    const bubbleBlock = blocks.find((block) => String(block.content ?? "") === bubbleText);
    expect(bubbleBlock).toBeDefined();
    bubbleBlock!.align = "center";

    const spec = await specFromOcrRaw(raw);
    const box = findBoxByText(spec.boxes, bubbleText);
    expect(box).toBeDefined();
    expect(box!.align).toBe("center");
  });

  ocrTest("specFromOcrRaw infers legible light ink on Ch4_A and Ch4_B dark top-panel text", async () => {
    const ch4ARaw = JSON.parse(await readFile(join(RAW_OCR_DIR, "Ch4_A.json"), "utf8"));
    const ch4ASpec = await specFromOcrRaw(ch4ARaw);
    const ch4ABoxes = ch4ASpec.boxes ?? [];
    const ch4AProblem = findBoxByText(ch4ABoxes, "問題：");
    const ch4ANavigator = findBoxByText(
      ch4ABoxes,
      "就像航海者需要羅盤，社會也需要一種方法來衡量我們前進的方向。所以，我們使用指標！但可惜，指標往往會失靈。",
    );
    expect(ch4AProblem).toBeDefined();
    expect(ch4ANavigator).toBeDefined();
    for (const box of [ch4AProblem!, ch4ANavigator!]) {
      expect(hexBrightness(String(box.color))).toBeGreaterThanOrEqual(LIGHT_INK_LUMA_MIN);
    }

    const ch4BRaw = JSON.parse(await readFile(join(RAW_OCR_DIR, "Ch4_B.json"), "utf8"));
    const ch4BSpec = await specFromOcrRaw(ch4BRaw);
    const ch4BBoxes = ch4BSpec.boxes ?? [];
    const ch4BSolution = findBoxByText(ch4BBoxes, "解方：");
    const ch4BList = findBoxByText(
      ch4BBoxes,
      "1. 衡量我們真正在乎的事。\n2. 用多種方式衡量。\n3. 經常重新評估指標。",
    );
    expect(ch4BSolution).toBeDefined();
    expect(ch4BList).toBeDefined();
    for (const box of [ch4BSolution!, ch4BList!]) {
      expect(hexBrightness(String(box.color))).toBeGreaterThanOrEqual(LIGHT_INK_LUMA_MIN);
    }
  });

  ocrTest("specFromOcrRaw infers light ink on Ch3_A dark speech panel quote", async () => {
    const raw = JSON.parse(await readFile(join(RAW_OCR_DIR, "Ch3_A.json"), "utf8"));
    const spec = await specFromOcrRaw(raw);
    const quote = findBoxById(spec.boxes, "b010_body");
    expect(quote).toBeDefined();
    expect(hexBrightness(String(quote!.color))).toBeGreaterThanOrEqual(LIGHT_INK_LUMA_MIN);
  });

  const CH2_A_SPEECH_BUBBLES = [
    "我只是在市場\n上賣顧客要買\n的東西！",
    "我只是用薪水\n買我買得起的\n東西！",
    "我只是付給\n工人由市場\n決定的工資\n！",
    "我只是投給\n比較不爛的\n那個！",
    "我只是給選民\n他們想要的。",
  ] as const;

  ocrTest("specFromOcrRaw emits BoxSpec for each Ch2_A speech-bubble string", async () => {
    const raw = JSON.parse(await readFile(join(RAW_OCR_DIR, "Ch2_A.json"), "utf8"));
    const spec = await specFromOcrRaw(raw);
    const boxes = spec.boxes ?? [];
    for (const text of CH2_A_SPEECH_BUBBLES) {
      expect(findBoxByText(boxes, text)).toBeDefined();
    }
  });

  ocrTest("specFromOcrRaw splits Ch2_B lower 如何 into its own BoxSpec below poster block", async () => {
    const raw = JSON.parse(await readFile(join(RAW_OCR_DIR, "Ch2_B.json"), "utf8"));
    const spec = await specFromOcrRaw(raw);
    const boxes = spec.boxes ?? [];
    const posterStack = findBoxByText(boxes, "誰\n做什麼\n何時\n為何");
    const howLabel = findBoxByText(boxes, "如何");
    expect(posterStack).toBeDefined();
    expect(howLabel).toBeDefined();
    expect(howLabel!.text).toBe("如何");
    const posterTop = posterStack!.source_box?.[1];
    const howTop = howLabel!.source_box?.[1];
    expect(posterTop).toBeDefined();
    expect(howTop).toBeDefined();
    expect(howTop!).toBeGreaterThan(posterTop!);
    expect(howTop!).toBeGreaterThan(1000);
    expect(howLabel!.font_role).toBe("title");
  });
  describe("English comic layout oracle for specFromOcrRaw", () => {
    const CH1_A_AND_OR_TEXT = "且\n或";
    const CH2_A_SPEECH_FOR_LAYOUT = [
      "我只是在市場\n上賣顧客要買\n的東西！",
      "我只是用薪水\n買我買得起的\n東西！",
      "我只是付給\n工人由市場\n決定的工資\n！",
    ] as const;
    const ENGLISH_TEXT_INK_LUMA_MAX = 95;
    const ENGLISH_MINUS_WORDLESS_DELTA_MIN = 30;
    const ORACLE_EXTENT_PAD_PX = 48;
    const ORACLE_EDGE_TOLERANCE_PX = 56;

    function rawWithoutManualBoxOverridesForText(
      raw: { ocr: { pages: Array<{ blocks: Array<Record<string, unknown>> }> } },
      text: string,
    ): typeof raw {
      const cloned = JSON.parse(JSON.stringify(raw)) as typeof raw;
      const block = cloned.ocr.pages[0]!.blocks.find((b) => String(b.content ?? "") === text);
      if (block) {
        delete block.target_box;
        delete block.target_tight_box;
        delete block.box;
        delete block.tight_box;
      }
      return cloned;
    }


    ocrTest("specFromOcrRaw uses English page geometry for Ch1_A 且\\n或 instead of raw OCR projection", async () => {
      const rawFull = JSON.parse(await readFile(join(RAW_OCR_DIR, "Ch1_A.json"), "utf8"));
      const raw = rawWithoutManualBoxOverridesForText(rawFull, CH1_A_AND_OR_TEXT);
      const spec = await specFromOcrRaw(raw);
      const entry = findBoxByText(spec.boxes, CH1_A_AND_OR_TEXT);
      expect(entry?.box).toBeDefined();
      expect(entry?.source_box).toBeDefined();
      const scale = spec.scale as [number, number];
      const sourceBox = entry!.source_box!.map(Number);
      const ocrProjected = [
        Math.round(sourceBox[0]! * scale[0]),
        Math.round(sourceBox[1]! * scale[1]),
        Math.round(sourceBox[2]! * scale[0]),
        Math.round(sourceBox[3]! * scale[1]),
      ];
      const layoutBox = entry!.box!.map(Number);
      expect(layoutBox).not.toEqual(ocrProjected);

      const englishImage = await loadImage(join(PROJECT_ROOT, "comics-en", "Ch1_A.png"));
      const wordlessImage = await loadImage(join(WORDLESS_ROOT, "Ch1_A.png"));
      const englishCanvas = createCanvas(englishImage.width, englishImage.height);
      const wordlessCanvas = createCanvas(wordlessImage.width, wordlessImage.height);
      const englishCtx = englishCanvas.getContext("2d");
      const wordlessCtx = wordlessCanvas.getContext("2d");
      englishCtx.drawImage(englishImage, 0, 0);
      wordlessCtx.drawImage(wordlessImage, 0, 0);

      const searchPad = 220;
      const searchX0 = Math.max(0, ocrProjected[0]! - searchPad);
      const searchY0 = Math.max(0, ocrProjected[1]! - searchPad);
      const searchX1 = Math.min(englishImage.width, ocrProjected[2]! + searchPad);
      const searchY1 = Math.min(englishImage.height, ocrProjected[3]! + searchPad);
      let inkMinX = Number.POSITIVE_INFINITY;
      let inkMinY = Number.POSITIVE_INFINITY;
      let inkMaxX = 0;
      let inkMaxY = 0;
      let inkPixels = 0;
      for (let y = searchY0; y < searchY1; y += 2) {
        for (let x = searchX0; x < searchX1; x += 2) {
          const englishSample = englishCtx.getImageData(x, y, 1, 1).data;
          const wordlessSample = wordlessCtx.getImageData(x, y, 1, 1).data;
          const englishLuma =
            0.299 * englishSample[0]! + 0.587 * englishSample[1]! + 0.114 * englishSample[2]!;
          const wordlessLuma =
            0.299 * wordlessSample[0]! + 0.587 * wordlessSample[1]! + 0.114 * wordlessSample[2]!;
          if (englishLuma >= ENGLISH_TEXT_INK_LUMA_MAX) continue;
          if (wordlessLuma - englishLuma < ENGLISH_MINUS_WORDLESS_DELTA_MIN) continue;
          inkPixels += 1;
          inkMinX = Math.min(inkMinX, x);
          inkMinY = Math.min(inkMinY, y);
          inkMaxX = Math.max(inkMaxX, x);
          inkMaxY = Math.max(inkMaxY, y);
        }
      }
      expect(inkPixels).toBeGreaterThan(200);
      const oracleExtent = [
        inkMinX - ORACLE_EXTENT_PAD_PX,
        inkMinY - ORACLE_EXTENT_PAD_PX,
        inkMaxX + ORACLE_EXTENT_PAD_PX,
        inkMaxY + ORACLE_EXTENT_PAD_PX,
      ];

      const layoutEdgeSum =
        Math.abs(layoutBox[0]! - oracleExtent[0]!) +
        Math.abs(layoutBox[1]! - oracleExtent[1]!) +
        Math.abs(layoutBox[2]! - oracleExtent[2]!) +
        Math.abs(layoutBox[3]! - oracleExtent[3]!);
      const ocrEdgeSum =
        Math.abs(ocrProjected[0]! - oracleExtent[0]!) +
        Math.abs(ocrProjected[1]! - oracleExtent[1]!) +
        Math.abs(ocrProjected[2]! - oracleExtent[2]!) +
        Math.abs(ocrProjected[3]! - oracleExtent[3]!);
      expect(layoutEdgeSum).toBeLessThan(ocrEdgeSum - 80);
      expect(layoutEdgeSum).toBeLessThan(ORACLE_EDGE_TOLERANCE_PX * 4);
    });

    ocrTest("specFromOcrRaw fits Ch2_A speech bubbles to English comic geometry not OCR rectangles", async () => {
      const raw = JSON.parse(await readFile(join(RAW_OCR_DIR, "Ch2_A.json"), "utf8"));
      const spec = await specFromOcrRaw(raw);
      const scale = spec.scale as [number, number];
      const englishImage = await loadImage(join(PROJECT_ROOT, "comics-en", "Ch2_A.png"));
      const wordlessImage = await loadImage(join(WORDLESS_ROOT, "Ch2_A.png"));
      const englishCanvas = createCanvas(englishImage.width, englishImage.height);
      const wordlessCanvas = createCanvas(wordlessImage.width, wordlessImage.height);
      const englishCtx = englishCanvas.getContext("2d");
      const wordlessCtx = wordlessCanvas.getContext("2d");
      englishCtx.drawImage(englishImage, 0, 0);
      wordlessCtx.drawImage(wordlessImage, 0, 0);

      for (const text of CH2_A_SPEECH_FOR_LAYOUT) {
        const entry = findBoxByText(spec.boxes, text);
        expect(entry?.box).toBeDefined();
        expect(entry?.source_box).toBeDefined();
        const sourceBox = entry!.source_box!.map(Number);
        const ocrProjected = [
          Math.round(sourceBox[0]! * scale[0]),
          Math.round(sourceBox[1]! * scale[1]),
          Math.round(sourceBox[2]! * scale[0]),
          Math.round(sourceBox[3]! * scale[1]),
        ];
        const layoutBox = entry!.box!.map(Number);
        expect(layoutBox).not.toEqual(ocrProjected);

        const searchPad = 120;
        const searchX0 = Math.max(0, ocrProjected[0]! - searchPad);
        const searchY0 = Math.max(0, ocrProjected[1]! - searchPad);
        const searchX1 = Math.min(englishImage.width, ocrProjected[2]! + searchPad);
        const searchY1 = Math.min(englishImage.height, ocrProjected[3]! + searchPad);
        let inkMinX = Number.POSITIVE_INFINITY;
        let inkMinY = Number.POSITIVE_INFINITY;
        let inkMaxX = 0;
        let inkMaxY = 0;
        let inkPixels = 0;
        for (let y = searchY0; y < searchY1; y += 2) {
          for (let x = searchX0; x < searchX1; x += 2) {
            const englishSample = englishCtx.getImageData(x, y, 1, 1).data;
            const wordlessSample = wordlessCtx.getImageData(x, y, 1, 1).data;
            const englishLuma =
              0.299 * englishSample[0]! + 0.587 * englishSample[1]! + 0.114 * englishSample[2]!;
            const wordlessLuma =
              0.299 * wordlessSample[0]! + 0.587 * wordlessSample[1]! + 0.114 * wordlessSample[2]!;
            if (englishLuma >= ENGLISH_TEXT_INK_LUMA_MAX) continue;
            if (wordlessLuma - englishLuma < ENGLISH_MINUS_WORDLESS_DELTA_MIN) continue;
            inkPixels += 1;
            inkMinX = Math.min(inkMinX, x);
            inkMinY = Math.min(inkMinY, y);
            inkMaxX = Math.max(inkMaxX, x);
            inkMaxY = Math.max(inkMaxY, y);
          }
        }
        expect(inkPixels).toBeGreaterThan(200);
        const oracleExtent = [
          inkMinX - ORACLE_EXTENT_PAD_PX,
          inkMinY - ORACLE_EXTENT_PAD_PX,
          inkMaxX + ORACLE_EXTENT_PAD_PX,
          inkMaxY + ORACLE_EXTENT_PAD_PX,
        ];

        const layoutEdgeSum =
          Math.abs(layoutBox[0]! - oracleExtent[0]!) +
          Math.abs(layoutBox[1]! - oracleExtent[1]!) +
          Math.abs(layoutBox[2]! - oracleExtent[2]!) +
          Math.abs(layoutBox[3]! - oracleExtent[3]!);
        const ocrEdgeSum =
          Math.abs(ocrProjected[0]! - oracleExtent[0]!) +
          Math.abs(ocrProjected[1]! - oracleExtent[1]!) +
          Math.abs(ocrProjected[2]! - oracleExtent[2]!) +
          Math.abs(ocrProjected[3]! - oracleExtent[3]!);
        expect(layoutEdgeSum).toBeLessThan(ocrEdgeSum - 40);
        expect(layoutEdgeSum).toBeLessThan(ORACLE_EDGE_TOLERANCE_PX * 4);
      }
    });
  });
});

describe("zh-TW YAML text-spec source", () => {
  const CH5_B_B005_EXPECTED = {
    text: "舉例來說，也許我們的兩位主廚會對彼此說……",
    source_box: [128, 512, 650, 618],
    box: [345, 2276, 2666, 2530],
    font_size: 145,
    align: "left" as const,
  };
  const CH5_B_B006_EXPECTED = {
    text: "…而他們發現，雙方都能100% 滿足需求，毫無浪費！",
    source_box: [775, 514, 1290, 618],
    box: [2582, 2236, 5763, 2571],
    font_size: 145,
    align: "left" as const,
  };

  test("ZH_TW_YAML_PATH points at translation-assets/zh-TW.yml under the project root", () => {
    expect(ZH_TW_YAML_PATH).toBe(join(PROJECT_ROOT, "translation-assets", "zh-TW.yml"));
  });

  test("loadYamlTextSpecs returns PAGE_ORDER specs with Ch5_B chef caption boxes preserved", async () => {
    const specs = await loadYamlTextSpecs();
    expect(Object.keys(specs).sort()).toEqual([...PAGE_ORDER].sort());
    const ch5 = specs.Ch5_B!;
    const b005 = findBoxById(ch5.boxes, "b005_body");
    const b006 = findBoxById(ch5.boxes, "b006_body");
    expect(b005?.text).toBe(CH5_B_B005_EXPECTED.text);
    expect(b005?.source_box).toEqual(CH5_B_B005_EXPECTED.source_box);
    expect(b005?.box).toEqual(CH5_B_B005_EXPECTED.box);
    expect(b005?.font_size).toBe(CH5_B_B005_EXPECTED.font_size);
    expect(b005?.align).toBe(CH5_B_B005_EXPECTED.align);
    expect(b006?.text).toBe(CH5_B_B006_EXPECTED.text);
    expect(b006?.source_box).toEqual(CH5_B_B006_EXPECTED.source_box);
    expect(b006?.box).toEqual(CH5_B_B006_EXPECTED.box);
    expect(b006?.font_size).toBe(CH5_B_B006_EXPECTED.font_size);
    expect(b006?.align).toBe(CH5_B_B006_EXPECTED.align);
  });

  test("buildTextSpecs builds from YAML when raw OCR dir is missing", async () => {
    const root = await mkdtemp(join(tmpdir(), "zh-tw-yaml-build-"));
    const missingRawDir = join(root, "no-such-raw-ocr");
    const textDir = join(root, "derived-text");
    expect(existsSync(missingRawDir)).toBe(false);
    const written = await buildTextSpecs(missingRawDir, textDir);
    expect(written).toHaveLength(PAGE_ORDER.length);
    const ch5 = JSON.parse(await readFile(join(textDir, "Ch5_B.json"), "utf8")) as TextSpec;
    const b005 = findBoxById(ch5.boxes, "b005_body");
    expect(b005?.text).toBe(CH5_B_B005_EXPECTED.text);
    expect(b005?.source_box).toEqual(CH5_B_B005_EXPECTED.source_box);
    expect(b005?.box).toEqual(CH5_B_B005_EXPECTED.box);
    expect(b005?.font_size).toBe(CH5_B_B005_EXPECTED.font_size);
    expect(b005?.align).toBe(CH5_B_B005_EXPECTED.align);
  });

  test("loadYamlTextSpecs rejects missing target_size instead of probing template dimensions", async () => {
    const root = await mkdtemp(join(tmpdir(), "zh-tw-yaml-bad-target-"));
    const yamlPath = join(root, "bad-target.yml");
    await writeMinimalYamlPageSizes(yamlPath, { source_size: [1280, 1781], target_size: undefined });
    await expect(loadYamlTextSpecs(yamlPath)).rejects.toThrow(VerificationError);
  });

  test("loadYamlTextSpecs rejects malformed source_size instead of probing source dimensions", async () => {
    const root = await mkdtemp(join(tmpdir(), "zh-tw-yaml-bad-source-"));
    const yamlPath = join(root, "bad-source.yml");
    await writeMinimalYamlPageSizes(yamlPath, { source_size: [1280], target_size: [5749, 8000] });
    await expect(loadYamlTextSpecs(yamlPath)).rejects.toThrow(VerificationError);
  });

  test("loadYamlTextSpecs rejects malformed Ch5_B target_size instead of probing template dimensions", async () => {
    const root = await mkdtemp(join(tmpdir(), "zh-tw-yaml-bad-ch5b-target-"));
    const yamlPath = join(root, "bad-ch5b-target.yml");
    await writeMinimalYamlPageSizes(yamlPath, { target_size: [5749] }, "Ch5_B");
    await expect(loadYamlTextSpecs(yamlPath)).rejects.toThrow(VerificationError);
  });


  test("loadYamlTextSpecs matches committed text specs for box and patch fields on every page", async () => {
    const yamlSpecs = await loadYamlTextSpecs();
    for (const pageId of PAGE_ORDER) {
      const committed = await loadCommittedTextSpec(pageId);
      const yamlSpec = yamlSpecs[pageId as keyof typeof yamlSpecs]!;
      const committedBoxes = [...(committed.boxes ?? [])].sort((a, b) => a.id.localeCompare(b.id));
      const yamlBoxes = [...(yamlSpec.boxes ?? [])].sort((a, b) => a.id.localeCompare(b.id));
      expect(yamlBoxes.map((box) => box.id)).toEqual(committedBoxes.map((box) => box.id));
      for (let index = 0; index < committedBoxes.length; index += 1) {
        expect(yamlBoxComparableFields(yamlBoxes[index]!)).toEqual(
          yamlBoxComparableFields(committedBoxes[index]!),
        );
      }
      const committedPatches = [...(committed.image_patches ?? [])].sort((a, b) => a.id.localeCompare(b.id));
      const yamlPatches = [...(yamlSpec.image_patches ?? [])].sort((a, b) => a.id.localeCompare(b.id));
      expect(yamlPatches.map((patch) => patch.id)).toEqual(committedPatches.map((patch) => patch.id));
      for (let index = 0; index < committedPatches.length; index += 1) {
        expect(yamlPatchComparableFields(yamlPatches[index]!)).toEqual(
          yamlPatchComparableFields(committedPatches[index]!),
        );
      }
    }
  });
});

describe("zh-TW visual feedback (committed text specs)", () => {
  const CH1_A_AND_OR = "且\n/\n或";
  const CH2_A_NOBODY = "當沒有人負責，就沒有人負責。";

  const CH5_A_AXIS_CAPTIONS: { id: string; text: string }[] = [
    { id: "ch5_a_axis_i_win", text: "我贏" },
    { id: "ch5_a_axis_you_win", text: "你贏" },
    { id: "ch5_a_axis_you_lose", text: "你輸" },
    { id: "ch5_a_axis_i_lose", text: "我輸" },
  ];

  const CH6_A_BASKET_BASELINE_FONT = 236;
  const OVERVIEW_B012_MAX_Y = 7250;

  test("Ch1_A b007_body 且/或 uses layout box centered in divider gap, not scaled OCR source_box", async () => {
    const spec = await loadCommittedTextSpec("Ch1_A");
    const entry = findBoxById(spec.boxes, "b007_body");
    const rect = boxRect(entry);
    expect(rect).toBeDefined();
    expect(entry?.text).toBe(CH1_A_AND_OR);
    expect(entry?.align).toBe("center");

    const scale = spec.scale as [number, number];
    const source = entry?.source_box?.map((value, index) => Number(value) * scale[index % 2]!) ?? [];
    expect(source.length).toBe(4);
    expect(rect).not.toEqual([source[0], source[1], source[2], source[3]]);

    const b006 = boxRect(findBoxById(spec.boxes, "b006_body"));
    const b008 = boxRect(findBoxById(spec.boxes, "b008_body"));
    expect(b006).toBeDefined();
    expect(b008).toBeDefined();
    const gapLeft = b006![2];
    const gapRight = b008![0];
    const centerX = (rect![0] + rect![2]) / 2;
    const gapMid = (gapLeft + gapRight) / 2;
    expect(rect![0]).toBeGreaterThanOrEqual(gapLeft - 80);
    expect(rect![2]).toBeLessThanOrEqual(gapRight + 80);
    expect(Math.abs(centerX - gapMid)).toBeLessThan((gapRight - gapLeft) / 2);
  });

  test("Ch2_A b004_body carries bold emphasis metadata for first 人 and second 沒有人", async () => {
    const spec = await loadCommittedTextSpec("Ch2_A");
    const entry = findBoxById(spec.boxes, "b004_body");
    expect(entry?.text).toBe(CH2_A_NOBODY);
    const emphasis = (entry as BoxSpec & { emphasis?: { text: string; bold?: boolean; occurrence?: number }[] })
      .emphasis;
    expect(emphasis).toEqual([
      { text: "人", bold: true, color: "#F09444" },
      { text: "沒有人", bold: true, color: "#F09444", occurrence: 2 },
    ]);
  });

  test("Ch2_B b010_body 如何 paragraph keeps explicit newline before 應該", async () => {
    const spec = await loadCommittedTextSpec("Ch2_B");
    const entry = findBoxById(spec.boxes, "b010_body");
    expect(entry?.text).toContain("「如何」去做 ?\n應該");
  });

  test("Ch5_A b007_body keeps explicit newline after 也就是", async () => {
    const spec = await loadCommittedTextSpec("Ch5_A");
    const entry = findBoxById(spec.boxes, "b007_body");
    expect(entry?.text).toContain("也就是\n「正和」");
  });

  test("Ch5_A committed spec includes axis caption boxes for I win / you win graph labels", async () => {
    const spec = await loadCommittedTextSpec("Ch5_A");
    const boxes = spec.boxes ?? [];
    for (const { id, text } of CH5_A_AXIS_CAPTIONS) {
      const byId = findBoxById(boxes, id);
      expect(byId?.text).toBe(text);
      expect(byId?.box?.length).toBeGreaterThanOrEqual(4);
    }
  });

  test("Ch5_B b004_body drops 的 before 做什麼 for better wrap", async () => {
    const spec = await loadCommittedTextSpec("Ch5_B");
    const entry = findBoxById(spec.boxes, "b004_body");
    expect(entry?.text).toBeDefined();
    expect(entry!.text!).toContain("在你想要「做什麼」背後");
    expect(entry!.text!).not.toContain("在你想要的「做什麼」背後");
    expect(entry!.text!).not.toMatch(/b0/);
  });

  const CH5_B_SOLIDARITY_SOLUTION_PHRASE = `雙方「為何」的、新的「做什麼」？`;

  test("Ch5_B solidarity solution phrase appears in b004_body with md-free committed spec", async () => {
    const spec = await loadCommittedTextSpec("Ch5_B");
    const entry = findBoxById(spec.boxes, "b004_body");
    expect(entry?.text).toBeDefined();
    expect(entry!.text!).toContain(CH5_B_SOLIDARITY_SOLUTION_PHRASE);
    expect(spec.ocr_markdown ?? "").toBe("");
  });

  test("Ch5_B b005_body and b006_body preserve full page-markdown chef caption copy", async () => {
    const spec = await loadCommittedTextSpec("Ch5_B");
    expect(findBoxById(spec.boxes, "b005_body")?.text).toBe(
      "舉例來說，也許我們的兩位主廚會對彼此說……",
    );
    expect(findBoxById(spec.boxes, "b006_body")?.text).toBe("…而他們發現，雙方都能100% 滿足需求，毫無浪費！");
  });


  test("Ch6_A b003_body basket line is center-aligned, one-line, and larger than 236px baseline", async () => {
    const spec = await loadCommittedTextSpec("Ch6_A");
    const entry = findBoxById(spec.boxes, "b003_body");
    expect(entry?.align).toBe("left");
    const text = entry?.text ?? "";
    expect(text).not.toContain("\n");
    const fontSize = entry?.font_size ?? 0;
    expect(fontSize).toBeGreaterThan(CH6_A_BASKET_BASELINE_FONT);
    const rect = boxRect(entry);
    expect(rect).toBeDefined();
    const fittedOneLine = fitFontSizeInBox(text, rect!, fontSize + 8, 8, "body", 1);
    expect(fittedOneLine).toBeGreaterThanOrEqual(fontSize);
  });

  test("Overview contrast bodies keep explicit line breaks and b012 stays above bunny-ear band", async () => {
    const spec = await loadCommittedTextSpec("Overview");
    const b006 = findBoxById(spec.boxes, "b006_body");
    const b008 = findBoxById(spec.boxes, "b008_body");
    const b012 = findBoxById(spec.boxes, "b012_body");

    expect(b006?.text).toMatch(/不是：.*\n而是：/s);
    expect(b008?.text).toMatch(/不是：.*\n而是：/s);

    const b012Rect = boxRect(b012);
    expect(b012Rect).toBeDefined();
    expect(b012Rect![3]).toBeLessThanOrEqual(OVERVIEW_B012_MAX_Y);
  });
});
describe("zh-TW visual feedback regression (2026 user review)", () => {
  const CH5_B_WELCOME_SIGN_Y_MIN = 5400;
  const CH5_B_WELCOME_SIGN_Y_MAX = 6500;
  const CH6_B_CLOSING_MAX_BOTTOM_Y = 7150;
  const CH6_B_BOTTOM_ART_BAND: [number, number, number, number] = [0, 7000, 5749, 8000];
  const OVERVIEW_RAW_IMAGE_COUNT = 7;
  const OVERVIEW_EXPECTED_IMAGE_PATCHES = 6;

  const CH3_A_B008_SOURCE_LINE_FRAGMENTS = [
    "這是……淨讚數的",
    "對數，再依時間",
    "折扣……我想",
    "是吧。",
  ] as const;

  const CH5_B_CHEF_LEFT_SPEECH = "我需要做\n橙酒。";
  const CH5_B_CHEF_RIGHT_SPEECH = "噢！我需要\n柳橙汁。";
  const CH5_B_CHEF_PEEL_LABEL = "果皮與皮屑\n做利口酒";
  const CH5_B_CHEF_JUICE_LABEL = "果肉\n榨汁";
  const CH1_B_MID_PAGE_LABEL_Y_MIN = 2400;
  const CH1_B_MID_PAGE_LABEL_Y_MAX = 5900;
  const CH1_B_MID_PAGE_CIRCLE_LABELS = [
    "乘客",
    "計程車",
    "司機",
    "安全是\n第一優先",
    "Uber 必須出示車輛\n與駕駛證件",
  ] as const;

  test("Ch3_A b008_body uses source-like four-line wrap inside right speech bubble", async () => {
    const spec = await loadCommittedTextSpec("Ch3_A");
    const entry = findBoxById(spec.boxes, "b008_body");
    const rect = boxRect(entry);
    expect(entry?.text).toBeDefined();
    const lines = entry!.text!.split("\n");
    expect(lines.length).toBeGreaterThanOrEqual(4);
    for (const fragment of CH3_A_B008_SOURCE_LINE_FRAGMENTS) {
      expect(entry!.text!).toContain(fragment);
    }
    expect(rect).toBeDefined();
    expect(rect![3] - rect![1]).toBeLessThanOrEqual(1100);
  });

  test("Ch3_B b005_body keeps explicit newlines after Birdwatch and 認同之處", async () => {
    const spec = await loadCommittedTextSpec("Ch3_B");
    const entry = findBoxById(spec.boxes, "b005_body");
    expect(entry?.text).toBeDefined();
    expect(entry!.text!).toContain("Birdwatch）。\n");
    expect(entry!.text!).toContain("認同之處。\n");
  });

  test("Ch1_B committed spec includes editable mid-page circle labels for Uber taxi diagram", async () => {
    const spec = await loadCommittedTextSpec("Ch1_B");
    const boxes = spec.boxes ?? [];
    for (const text of CH1_B_MID_PAGE_CIRCLE_LABELS) {
      const entry = findBoxByText(boxes, text);
      expect(entry).toBeDefined();
      expect(entry!.text).toBe(text);
      const rect = boxRect(entry);
      expect(rect).toBeDefined();
      expect(rect![1]).toBeGreaterThanOrEqual(CH1_B_MID_PAGE_LABEL_Y_MIN);
      expect(rect![3]).toBeLessThanOrEqual(CH1_B_MID_PAGE_LABEL_Y_MAX);
    }
  });

  test("Ch5_B chef speech bubbles and peel/juice labels match source crops", async () => {
    const spec = await loadCommittedTextSpec("Ch5_B");
    const boxes = spec.boxes ?? [];
    const speechLeft =
      findBoxByText(boxes, CH5_B_CHEF_LEFT_SPEECH) ?? findBoxById(boxes, "ch5_b_chef_speech_left");
    const speechRight =
      findBoxByText(boxes, CH5_B_CHEF_RIGHT_SPEECH) ?? findBoxById(boxes, "ch5_b_chef_speech_right");
    const peelLabel =
      findBoxByText(boxes, CH5_B_CHEF_PEEL_LABEL) ??
      findBoxById(boxes, "ch5_b_chef_peel_label") ??
      findBoxById(boxes, "ch5_b_chef_peel");
    const juiceLabel =
      findBoxByText(boxes, CH5_B_CHEF_JUICE_LABEL) ??
      findBoxById(boxes, "ch5_b_chef_juice_label") ??
      findBoxById(boxes, "ch5_b_chef_juice");

    expect(speechLeft?.text?.trim().length).toBeGreaterThan(0);
    expect(speechRight?.text?.trim().length).toBeGreaterThan(0);
    expect(peelLabel?.text?.trim().length).toBeGreaterThan(0);
    expect(juiceLabel?.text?.trim().length).toBeGreaterThan(0);
    expect(speechLeft!.text!).toBe(CH5_B_CHEF_LEFT_SPEECH);
    expect(speechRight!.text!).toBe(CH5_B_CHEF_RIGHT_SPEECH);
    expect(peelLabel!.text!).toBe(CH5_B_CHEF_PEEL_LABEL);
    expect(juiceLabel!.text!).toBe(CH5_B_CHEF_JUICE_LABEL);
  });

  test("Ch5_B welcome sign box sits in approved lower sign band", async () => {
    const spec = await loadCommittedTextSpec("Ch5_B");
    const entry = findBoxById(spec.boxes, "ch5_b_welcome_creator");
    const rect = boxRect(entry);
    expect(rect).toBeDefined();
    expect(rect![1]).toBeGreaterThanOrEqual(CH5_B_WELCOME_SIGN_Y_MIN);
    expect(rect![3]).toBeLessThanOrEqual(CH5_B_WELCOME_SIGN_Y_MAX);
  });

  test("Ch6_B b007_body and b008_body layout boxes do not overlap vertically", async () => {
    const spec = await loadCommittedTextSpec("Ch6_B");
    const b007 = boxRect(findBoxById(spec.boxes, "b007_body"));
    const b008 = boxRect(findBoxById(spec.boxes, "b008_body"));
    expect(b007).toBeDefined();
    expect(b008).toBeDefined();
    expect(verticalOverlapArea(b007!, b008!)).toBe(0);
  });

  test("Ch6_B b007_body Kami paragraph uses explicit source-like line breaks", async () => {
    const spec = await loadCommittedTextSpec("Ch6_B");
    const entry = findBoxById(spec.boxes, "b007_body");
    expect(entry?.text).toBeDefined();
    expect(entry!.text!).toContain("在地精神」\n的 Kami");
    expect(entry!.text!).toContain("大公司廢料，\n而是一個");
    expect(entry!.text!).toContain("筆電上\n在地運行的");
  });

  test("Ch6_B b009_body closing line avoids wide cross-page band and bottom-art overlap", async () => {
    const spec = await loadCommittedTextSpec("Ch6_B");
    const entry = findBoxById(spec.boxes, "b009_body");
    const rect = boxRect(entry);
    expect(rect).toBeDefined();
    const width = rect![2] - rect![0];
    expect(width).toBeLessThan(3500);
    expect(intersectionArea(rect!, CH6_B_BOTTOM_ART_BAND)).toBeLessThanOrEqual(400000);
    expect(rect![3]).toBeLessThanOrEqual(CH6_B_CLOSING_MAX_BOTTOM_Y);
  });

  test("Ch6_B charter scroll art uses source image patch metadata or non-dummy scroll provenance", async () => {
    const spec = await loadCommittedTextSpec("Ch6_B");
    const patchList = [
      ...(Array.isArray(spec.image_patches) ? spec.image_patches : []),
      ...(Array.isArray(spec.image_patch_regions) ? spec.image_patch_regions : []),
    ] as Record<string, unknown>[];
    const charterPatch = patchList.find((row) => {
      const id = String(row.id ?? row.image_id ?? row.patch_id ?? "");
      return /charter|章程|scroll/i.test(id);
    });
    const patchHasRealSource =
      charterPatch != null &&
      (() => {
        const sb = charterPatch.source_box ?? charterPatch.box;
        if (!Array.isArray(sb) || sb.length !== 4) return false;
        const n = sb.map(Number);
        return n[2]! - n[0]! > 10 && n[3]! - n[1]! > 10;
      })();
    const scrollBox = findBoxById(spec.boxes, "ch6_b_charter_scroll");
    const scrollFromTextBox = scrollBox != null && !isDummyUnitSourceBox(scrollBox);
    expect(patchHasRealSource || scrollFromTextBox).toBe(true);
  });

  ocrTest("Overview committed spec exposes source illustration patch regions without page markdown", async () => {
    const spec = await loadCommittedTextSpec("Overview");
    const rawImageCount = await loadOcrRawImageCount("Overview");
    expect(rawImageCount).toBe(OVERVIEW_RAW_IMAGE_COUNT);
    expect(countMarkdownIllustrationRefs(spec.ocr_markdown)).toBe(0);
    expect(imagePatchRegionCount(spec)).toBe(OVERVIEW_EXPECTED_IMAGE_PATCHES);
  });
});

describe("round two visual feedback", () => {
  const CH1_A_AND_OR = "且\n/\n或";
  /** Divider-gap center for Ch1_A b006/b008 (committed geometry baseline). */
  const CH1_A_DIVIDER_GAP_CENTER_TOLERANCE_PX = 95;

  const CH2_A_NOBODY = "當沒有人負責，就沒有人負責。";

  const CH3_A_B008_EXPECTED_LINES = [
    "這是……淨讚數的",
    "對數，再依時間",
    "折扣……我想",
    "是吧。",
  ] as const;
  const CH3_A_B010_EXPECTED_LINES = [
    "覺悟並非自我覺察，",
    "而是自我不察。",
    "要成為一個如此浩瀚",
    "的心智，連它自己都",
    "再也說不出自己真正",
    "的名字。",
  ] as const;
  const CH3_A_B008_MIN_LEFT_X = 3980;

  const CH4_A_TABLE_LABEL_SNIPPET = "使用者好評";
  const CH4_A_TABLE_MAX_RIGHT_X = 2850;
  const CH4_A_TABLE_RIGHT_CELL_SNIPPET = "諂媚的聊天機器人";
  const CH4_A_TABLE_OVAL_TEXT = "之類";

  const CH5_B_CHEF_SPEECH_MIN_FONT = 120;
  const CH5_B_CHEF_LABEL_MIN_FONT = 110;

  const CH6_B_KAMI_MIN_FONT = 120;
  const CH6_B_CLOSING_MIN_FONT = 110;
  const CH6_B_SAFEGUARD_MIN_FONT = 150;

  const COMIC_PAGE_IDS = PAGE_ORDER.filter((id) => id !== "Overview");

  function emphasisSpans(box: BoxSpec | undefined): { text: string; bold?: boolean; color?: string; occurrence?: number }[] {
    return (box as BoxSpec & { emphasis?: { text: string; bold?: boolean; color?: string; occurrence?: number }[] }).emphasis ?? [];
  }

  test("Ch1_A b007_body keeps vertical 且/或 and centers on the divider gap", async () => {
    const spec = await loadCommittedTextSpec("Ch1_A");
    const entry = findBoxById(spec.boxes, "b007_body");
    const rect = boxRect(entry);
    expect(entry?.text).toBe(CH1_A_AND_OR);
    expect(entry?.text).toContain("\n");
    expect(rect).toBeDefined();

    const b006 = boxRect(findBoxById(spec.boxes, "b006_body"));
    const b008 = boxRect(findBoxById(spec.boxes, "b008_body"));
    expect(b006).toBeDefined();
    expect(b008).toBeDefined();
    const gapMid = (b006![2] + b008![0]) / 2;
    const centerX = (rect![0] + rect![2]) / 2;
    expect(Math.abs(centerX - gapMid)).toBeLessThanOrEqual(CH1_A_DIVIDER_GAP_CENTER_TOLERANCE_PX);
  });

  test("committed comic text specs contain no stray lowercase a artifact boxes", async () => {
    for (const pageId of COMIC_PAGE_IDS) {
      const spec = await loadCommittedTextSpec(pageId);
      for (const box of spec.boxes ?? []) {
        const text = box.text ?? "";
        expect(text).not.toMatch(/^\s*a\s*$/i);
      }
    }
  });

  test("Ch2_A b004_body bolds only first 人 and second 沒有人 via emphasis spans", async () => {
    const spec = await loadCommittedTextSpec("Ch2_A");
    const entry = findBoxById(spec.boxes, "b004_body");
    expect(entry?.text).toBe(CH2_A_NOBODY);
    const emphasis = emphasisSpans(entry);
    expect(emphasis).toEqual([
      { text: "人", bold: true, color: "#F09444" },
      { text: "沒有人", bold: true, color: "#F09444", occurrence: 2 },
    ]);
  });

  test("Ch3_A b008_body keeps source-like line breaks with inset from left bubble edge", async () => {
    const spec = await loadCommittedTextSpec("Ch3_A");
    const entry = findBoxById(spec.boxes, "b008_body");
    const rect = boxRect(entry);
    expect(rect).toBeDefined();
    expect(rect![0]).toBeGreaterThanOrEqual(CH3_A_B008_MIN_LEFT_X);
    const lines = (entry?.text ?? "").split("\n");
    expect(lines).toEqual([...CH3_A_B008_EXPECTED_LINES]);
  });

  test("Ch3_A b010_body gray quote box keeps centered source-like six-line wrap", async () => {
    const spec = await loadCommittedTextSpec("Ch3_A");
    const entry = findBoxById(spec.boxes, "b010_body");
    expect(entry?.align).toBe("center");
    expect((entry?.text ?? "").split("\n")).toEqual([...CH3_A_B010_EXPECTED_LINES]);
  });

  test("Ch4_A b006_body table copy includes 使用者好評 and stays inside table band", async () => {
    const spec = await loadCommittedTextSpec("Ch4_A");
    const entry = findBoxById(spec.boxes, "b006_body");
    const rect = boxRect(entry);
    expect(entry?.text).toContain(CH4_A_TABLE_LABEL_SNIPPET);
    expect(entry?.text).toContain(CH4_A_TABLE_RIGHT_CELL_SNIPPET);
    expect(rect).toBeDefined();
    expect(rect![2]).toBeLessThanOrEqual(CH4_A_TABLE_MAX_RIGHT_X);
    const oval = findBoxById(spec.boxes, "ch4_a_table_etc_bubble");
    expect(oval?.text).toBe(CH4_A_TABLE_OVAL_TEXT);
    const ovalRect = boxRect(oval);
    expect(ovalRect).toBeDefined();
    expect(ovalRect![0]).toBeGreaterThanOrEqual(2400);
    expect(ovalRect![2]).toBeLessThanOrEqual(3050);
    expect(ovalRect![1]).toBeGreaterThanOrEqual(7100);
    expect(ovalRect![3]).toBeLessThanOrEqual(7600);
  });

  test("Ch5_B chef speech and peel/juice labels use readable fonts above round-one minima", async () => {
    const spec = await loadCommittedTextSpec("Ch5_B");
    const boxes = spec.boxes ?? [];
    const speechLeft = findBoxById(boxes, "ch5_b_chef_speech_left");
    const speechRight = findBoxById(boxes, "ch5_b_chef_speech_right");
    const peelLabel = findBoxById(boxes, "ch5_b_chef_peel_label");
    const juiceLabel = findBoxById(boxes, "ch5_b_chef_juice_label");
    expect(speechLeft?.font_size ?? 0).toBeGreaterThanOrEqual(CH5_B_CHEF_SPEECH_MIN_FONT);
    expect(speechRight?.font_size ?? 0).toBeGreaterThanOrEqual(CH5_B_CHEF_SPEECH_MIN_FONT);
    expect(peelLabel?.font_size ?? 0).toBeGreaterThanOrEqual(CH5_B_CHEF_LABEL_MIN_FONT);
    expect(juiceLabel?.font_size ?? 0).toBeGreaterThanOrEqual(CH5_B_CHEF_LABEL_MIN_FONT);
  });

  test("Ch6_B lower Kami and closing blocks use larger readable fonts without vertical overlap", async () => {
    const spec = await loadCommittedTextSpec("Ch6_B");
    const b007 = findBoxById(spec.boxes, "b007_body");
    const b008 = findBoxById(spec.boxes, "b008_body");
    const b009 = findBoxById(spec.boxes, "b009_body");
    expect(b007?.font_size ?? 0).toBeGreaterThanOrEqual(CH6_B_KAMI_MIN_FONT);
    expect(b009?.font_size ?? 0).toBeGreaterThanOrEqual(CH6_B_CLOSING_MIN_FONT);
    const b007Rect = boxRect(b007);
    const b008Rect = boxRect(b008);
    const b009Rect = boxRect(b009);
    expect(b007Rect).toBeDefined();
    expect(b008Rect).toBeDefined();
    expect(b009Rect).toBeDefined();
    expect(verticalOverlapArea(b007Rect!, b008Rect!)).toBe(0);
    expect(verticalOverlapArea(b008Rect!, b009Rect!)).toBeLessThanOrEqual(300000);
  });

  test("Ch6_B b006_body gpt-oss-safeguard paragraph uses larger font with source-like line breaks", async () => {
    const spec = await loadCommittedTextSpec("Ch6_B");
    const entry = findBoxById(spec.boxes, "b006_body");
    expect(entry?.text ?? "").toContain("gpt-oss-safeguard");
    expect(entry?.font_size ?? 0).toBeGreaterThanOrEqual(CH6_B_SAFEGUARD_MIN_FONT);
    expect(entry?.text ?? "").toMatch(/gpt-oss-safeguard[\s\S]*\n[\s\S]*自帶政策/);
  });

  test("Overview committed spec omits img-0.jpeg source patch under editable title", async () => {
    const spec = await loadCommittedTextSpec("Overview");
    const patchList = [
      ...(Array.isArray(spec.image_patches) ? spec.image_patches : []),
      ...(Array.isArray(spec.image_patch_regions) ? spec.image_patch_regions : []),
    ] as { id?: string }[];
    const ids = patchList.map((row) => String(row.id ?? ""));
    expect(ids).not.toContain("img-0.jpeg");
    expect(patchList.length).toBeGreaterThanOrEqual(6);
  });
});



describe("pixel diagnostics", () => {
  test("estimateInkColor uses light ink on dark panel with white glyphs", () => {
    const family = registerFontForRole("body");
    expect(family).toBeTruthy();
    const canvas = createCanvas(140, 70);
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "rgb(22,25,30)";
    ctx.fillRect(0, 0, 140, 70);
    ctx.fillStyle = "rgb(248,248,248)";
    ctx.font = `24px ${family}`;
    ctx.fillText("白字測試", 12, 48);
    const ink = estimateInkColor(canvas, [0, 0, 140, 70]);
    expect(hexBrightness(ink)).toBeGreaterThanOrEqual(200);
  });

  test("compareRenderedPageToSource emits mask and threshold-bounded report", async () => {
    const { fixture } = await writeSyntheticDiffFixture({ pageId: "cmp_page", inkPatch: false });
    const renderedPath = join(fixture.root, "vector.png");
    const maskPath = join(fixture.root, "vector.mask.png");
    await renderPage(fixture.spec, renderedPath, "vector");
    const report = await compareRenderedPageToSource(
      renderedPath,
      fixture.sourcePath,
      fixture.spec,
      { maskPath, threshold: DIFF_THRESHOLD },
    );
    expect(await Bun.file(maskPath).exists()).toBe(true);
    expect(report.threshold).toBe(DIFF_THRESHOLD);
    expect(report.changed_pixels).toBeGreaterThan(0);
    expect(report.changed_ratio).toBeLessThan(1);
  });

  test("generateDiffMasks writes report JSON and mask files", async () => {
    const bundle = await writeSyntheticDiffFixture({ pageId: "Synthetic" });
    const result = await generateDiffMasks(bundle.textDir, bundle.outputRoot, bundle.reviewDir, {
      threshold: DIFF_THRESHOLD,
      pageOrder: [bundle.pageId],
      templateImageNames: { [bundle.pageId]: bundle.outName },
    });
    expect(await Bun.file(result.reportPath).exists()).toBe(true);
    expect(await Bun.file(join(bundle.reviewDir, "diffs", `${bundle.pageId}.mask.png`)).exists()).toBe(
      true,
    );
    const pageReport = result.pages[bundle.pageId] as { changed_pixels: number; threshold: number };
    expect(pageReport.changed_pixels).toBeGreaterThan(0);
    expect(pageReport.threshold).toBe(DIFF_THRESHOLD);
  });

  test("oracle_patch mode reduces changed pixels versus vector under scaled JPEG oracle", async () => {
    const scaleX = (400 / 200) * 1.00035;
    const scaleY = (300 / 150) * 1.00201;
    const { fixture } = await writeSyntheticDiffFixture({
      pageId: "patch_page",
      inkPatch: true,
      scaleX,
      scaleY,
    });
    const vectorPath = join(fixture.root, "vector.png");
    const patchPath = join(fixture.root, "patch.png");
    await renderPage(fixture.spec, vectorPath, "vector");
    await renderPage(fixture.spec, patchPath, "oracle_patch");
    const vectorReport = await compareRenderedPageToSource(
      vectorPath,
      fixture.sourcePath,
      fixture.spec,
      { threshold: DIFF_THRESHOLD },
    );
    const patchReport = await compareRenderedPageToSource(
      patchPath,
      fixture.sourcePath,
      fixture.spec,
      { threshold: DIFF_THRESHOLD },
    );
    const slack = Math.max(1, Math.floor(0.005 * vectorReport.total_pixels));
    expect(vectorReport.changed_pixels).toBeGreaterThan(slack);
    expect(patchReport.changed_pixels).toBeLessThanOrEqual(
      vectorReport.changed_pixels - slack,
    );
  });
});

describe("rendering contracts", () => {
  test("vector table render reserves left gutter for row icons", async () => {
    const bundle = await writeSyntheticTableFixture();
    const output = join(bundle.outputRoot, "TablePage.png");
    await renderPage(bundle.spec, output, "vector");
    const image = await loadCanvasFromFile(output);
    const [x0, y0, x1, y1] = TABLE_BOX;
    const gutterDark = countDarkPixels(image, [x0, y0, Math.min(x1, x0 + GUTTER_WIDTH), y1]);
    const bodyDark = countDarkPixels(image, [x0 + GUTTER_WIDTH, y0, x1, y1]);
    expect(gutterDark).toBe(0);
    expect(bodyDark).toBeGreaterThan(0);
  });

  test("verifyGeneratedAssets reports missing outputs and dimensions", async () => {
    const report = await verifyGeneratedAssets();
    expect(report).toHaveProperty("missing");
    expect(report).toHaveProperty("dimensions");
    expect(Array.isArray(report.missing)).toBe(true);
    expect(typeof report.dimensions).toBe("object");
  });
});

describe("generated zh-TW assets", () => {
  beforeAll(async () => {
    const oracleDir = await mkdtemp(join(tmpdir(), "zh-tw-desktop-oracle-"));
    DESKTOP_ZH_TW_ORACLE_ROOT = oracleDir;
    ORACLE_PAGE_PATHS = await renderPdfPages(CANONICAL_ZH_TW_EDITABLE_PDF, oracleDir, PDF_ORACLE_RENDER_DPI, "zh-TW-editable");
  }, 120_000);

  test("verifyGeneratedAssets reports no missing files and expected dimensions", async () => {
    const report = await verifyGeneratedAssets();
    expect(report.missing).toEqual([]);
    expect(report.dimensions["Overview.png"]).toEqual([5742, 8000]);
    expect(report.dimensions["Ch4_A.png"]).toEqual([5742, 8000]);
    expect(report.dimensions["web_footer.png"]).toEqual([5749, 479]);
    expect(report.dimensions["web/Overview.png"]).toEqual([5749, 8479]);
    expect(report.dimensions["web/Ch4.png"]).toEqual([5749, 15379]);
  });

  test("each generated page PNG pairs with Desktop oracle slide by portrait aspect ratio", async () => {
    for (let index = 0; index < PAGE_ORDER.length; index += 1) {
      const pageId = PAGE_ORDER[index]!;
      const generatedPath = join(OUTPUT_ROOT, `${pageId}.png`);
      const oraclePath = ORACLE_PAGE_PATHS[index];

      expect(existsSync(generatedPath)).toBe(true);
      expect(existsSync(oraclePath)).toBe(true);

      const generated = await loadImage(generatedPath);
      const oracle = await loadImage(oraclePath);

      expect(generated.height).toBeGreaterThan(generated.width);
      expect(oracle.height).toBeGreaterThan(oracle.width);

      const generatedAspect = generated.width / generated.height;
      const oracleAspect = oracle.width / oracle.height;
      expect(Math.abs(generatedAspect - oracleAspect)).toBeLessThanOrEqual(ORACLE_ASPECT_RATIO_TOLERANCE);
    }
  });
});

describe("zh-TW book cover outputs", () => {
  test("verifyGeneratedAssets reports missing book cover outputs when output root has no copies", async () => {
    const emptyOutputRoot = await mkdtemp(join(tmpdir(), "zh-tw-cover-verify-"));
    const report = await verifyGeneratedAssets(
      emptyOutputRoot,
      join(emptyOutputRoot, "text-specs-absent"),
    );
    expect(report.missing).toEqual(
      expect.arrayContaining([...ZH_TW_COVER_OUTPUT_NAMES]),
    );
  });

  test("verifyGeneratedAssets records dimensions for book cover outputs in comics-zh-TW", async () => {
    const report = await verifyGeneratedAssets();
    for (const name of ZH_TW_COVER_OUTPUT_NAMES) {
      expect(report.dimensions[name]).toBeDefined();
    }
  });

  test("verifyGeneratedAssets reports missing translated cover text specs and cover editable PPTX when absent", async () => {
    const emptyOutputRoot = await mkdtemp(join(tmpdir(), "zh-tw-cover-verify-specs-"));
    const absentTextDir = join(emptyOutputRoot, "text-specs-absent");
    const report = await verifyGeneratedAssets(emptyOutputRoot, absentTextDir);
    expect(report.missing).toEqual(
      expect.arrayContaining([
        ...ZH_TW_COVER_VERIFY_TEXT_SPECS,
        ZH_TW_COVER_EDITABLE_PPTX,
      ]),
    );
  });

  test("comics-zh-TW cover outputs exist and front cover is not a byte copy of English book-cover/Cover.png", async () => {
    const englishCoverPath = join(PROJECT_ROOT, "book-cover", "Cover.png");
    const englishCoverBytes = await readFile(englishCoverPath);

    for (const name of ZH_TW_COVER_OUTPUT_NAMES) {
      const outPath = join(OUTPUT_ROOT, name);
      expect(existsSync(outPath)).toBe(true);
    }

    const outputCoverBytes = await readFile(join(OUTPUT_ROOT, "Cover.png"));
    expect(outputCoverBytes).not.toEqual(englishCoverBytes);
  });
});

describe("zh-TW translated book cover", () => {
  test("loadCoverTextSpecs reads translated zh-TW cover copy from YAML", async () => {
    const loadCoverTextSpecs = requireLoadCoverTextSpecs();
    const COVER_PAGE_ORDER_EXPORT = requireCoverPipelineExport<typeof COVER_PAGE_ORDER>("COVER_PAGE_ORDER");
    expect(COVER_PAGE_ORDER_EXPORT).toEqual([...COVER_PAGE_ORDER]);

    const specs = await loadCoverTextSpecs();
    for (const pageId of COVER_PAGE_ORDER) {
      expect(specs[pageId]).toBeDefined();
      expect(String(specs[pageId]!.template_image ?? "")).toContain("book-cover/wordless");
    }

    const allText = collectCoverBoxText(specs);
    expect(allText).toContain("仁工智慧");
    expect(allText).toContain("讓 AI 與人類社群");
    expect(allText).toContain("共同繁榮");
    expect(allText).toContain("給出版社夥伴");
    expect(allText).toContain("唐鳳");
    expect(allText).toContain("簡介");

    for (const char of Array.from(COVER_SIMPLIFIED_TRAP_SENTINEL)) {
      expect(allText).not.toContain(char);
    }
  });

  test("renderCoverAssets renders translated covers instead of byte-copying English sources", async () => {
    const buildCoverTextSpecs = requireBuildCoverTextSpecs();
    const renderCoverAssets = requireRenderCoverAssets();

    const tempRoot = await mkdtemp(join(tmpdir(), "zh-tw-cover-render-"));
    const textDir = join(tempRoot, "cover-text");
    const outputRoot = join(tempRoot, "out");
    await buildCoverTextSpecs(undefined, textDir);
    await renderCoverAssets(textDir, outputRoot);

    const englishCoverPath = join(PROJECT_ROOT, "book-cover", "Cover.png");
    const englishCoverBytes = await readFile(englishCoverPath);

    for (const pageId of COVER_PAGE_ORDER) {
      const outName = COVER_PAGE_OUTPUT_PNG[pageId];
      const outPath = join(outputRoot, outName);
      expect(existsSync(outPath)).toBe(true);
      const dim = await readPngDimensions(outPath);
      if (pageId === "Cover") {
        expectCoverFrontDimensions(dim);
      } else if (pageId === "Cover_Back") {
        const canonicalCoverPdf = requireCoverPipelineExport<string>("DEFAULT_ZH_TW_COVER_CANONICAL_PDF");
        expect(dim).toEqual(existsSync(canonicalCoverPdf) ? { width: 5742, height: 8000 } : { width: 5748, height: 8000 });
      } else {
        const canonicalCoverPdf = requireCoverPipelineExport<string>("DEFAULT_ZH_TW_COVER_CANONICAL_PDF");
        expect(dim).toEqual(existsSync(canonicalCoverPdf) ? { width: 767, height: 8000 } : { width: 766, height: 8000 });
      }
    }

    const coverExtraImageNames = requireCoverPipelineExport<string[]>("COVER_EXTRA_IMAGE_NAMES");
    for (const extraName of coverExtraImageNames) {
      expect(existsSync(join(outputRoot, extraName))).toBe(true);
    }

    const renderedCoverBytes = await readFile(join(outputRoot, "Cover.png"));
    expect(renderedCoverBytes).not.toEqual(englishCoverBytes);
  }, 120_000);

  test("buildCoverEditablePptx writes editable cover slides with named text boxes", async () => {
    const buildCoverEditablePptx = requireBuildCoverEditablePptx();
    const COVER_PPTX_OUTPUT_PATH = requireCoverPipelineExport<string>("COVER_PPTX_OUTPUT_PATH");
    expect(COVER_PPTX_OUTPUT_PATH).toBe(join(OUTPUT_ROOT, "zh-TW-cover-editable.pptx"));

    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-cover-pptx-"));
    const outPath = join(tempDir, "zh-TW-cover-editable.pptx");
    const written = await buildCoverEditablePptx(outPath);
    expect(written).toBe(outPath);

    const entries = await listPptxZipEntryNames(outPath);
    const slideXmlPaths = entries.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
    expect(slideXmlPaths).toHaveLength(3);

    const slideXmlJoined = (
      await Promise.all(slideXmlPaths.map((path) => readPptxZipEntryText(outPath, path)))
    ).join("\n");
    expect(slideXmlJoined).toContain('name="front_title"');
    expect(slideXmlJoined).toContain('name="back_heading"');
    expect(slideXmlJoined).toContain('name="spine_title"');
    expect(slideXmlJoined).toContain("仁工智慧");
    expect(slideXmlJoined).toContain("唐鳳");
  }, 120_000);

  test("parseCoverEditablePptxToYaml maps Cover_Spine spine_authors after Desktop-style shape rename", async () => {
    const parseCoverEditablePptxToYaml = requireParseCoverEditablePptxToYaml();
    const buildCoverEditablePptx = requireBuildCoverEditablePptx();
    const ZH_TW_COVER_YAML_PATH = requireCoverPipelineExport<string>("ZH_TW_COVER_YAML_PATH");

    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-cover-parseback-"));
    const yamlPath = join(tempDir, "zh-TW-cover.yml");
    const pptxPath = join(tempDir, "zh-TW-cover-editable.pptx");
    await writeFile(yamlPath, await readFile(ZH_TW_COVER_YAML_PATH, "utf8"));
    const spineAuthorsBlock = extractYamlPageBoxBlock(
      await readFile(yamlPath, "utf8"),
      "Cover_Spine",
      "spine_authors",
    );
    const preMergeSpineAuthorsBlock = patchYamlBoxLiteralTextBlock(
      spineAuthorsBlock,
      COVER_SPINE_AUTHORS_PRE_MERGE_LITERAL,
    );
    await writeYamlPageBoxBlockReplacement(yamlPath, "Cover_Spine", "spine_authors", preMergeSpineAuthorsBlock);
    const beforeYaml = await readFile(yamlPath, "utf8");
    expect(readCompactYamlBox(beforeYaml, "Cover_Spine", "spine_authors").text?.trim()).toBe(
      COVER_SPINE_AUTHORS_PRE_MERGE_LITERAL,
    );

    await buildCoverEditablePptx(pptxPath, yamlPath);

    const spineSlideEntry = "ppt/slides/slide3.xml";
    const originalSlide = await readPptxZipEntryText(pptxPath, spineSlideEntry);
    const originalShape = extractPptxShapeXmlByName(originalSlide, "spine_authors");
    if (!originalShape) throw new Error("missing spine_authors shape on Cover_Spine slide");

    const desktopShape = simulateDesktopResavedCoverSpineAuthorsShape(originalShape);
    expect(desktopShape).toContain('name="唐"');
    expect(desktopShape).not.toContain('name="spine_authors"');
    expect(extractPptxShapeTextWithBreaks(desktopShape)).toBe(COVER_SPINE_AUTHORS_DESKTOP_PARSED_TEXT);

    await writePptxZipEntry(
      pptxPath,
      spineSlideEntry,
      replacePptxShapeXml(originalSlide, "spine_authors", desktopShape),
    );

    const result = await parseCoverEditablePptxToYaml(pptxPath, yamlPath);
    expect(result.yaml_path).toBe(yamlPath);
    expect(result.updated).toBeGreaterThan(0);
    expect(result.unknown_shapes).toEqual([]);

    const box = readCompactYamlBox(await readFile(yamlPath, "utf8"), "Cover_Spine", "spine_authors");
    expect(box.text?.trim()).toBe(COVER_SPINE_AUTHORS_DESKTOP_PARSED_TEXT);
  }, 120_000);

  test("parseCoverEditablePptxToYaml preserves Cover front_title center align when Desktop strips algn ctr", async () => {
    const parseCoverEditablePptxToYaml = requireParseCoverEditablePptxToYaml();
    const buildCoverEditablePptx = requireBuildCoverEditablePptx();
    const ZH_TW_COVER_YAML_PATH = requireCoverPipelineExport<string>("ZH_TW_COVER_YAML_PATH");

    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-cover-align-parseback-"));
    const yamlPath = join(tempDir, "zh-TW-cover.yml");
    const pptxPath = join(tempDir, "zh-TW-cover-editable.pptx");
    await writeFile(yamlPath, await readFile(ZH_TW_COVER_YAML_PATH, "utf8"));
    const frontTitleBlock = extractYamlPageBoxBlock(await readFile(yamlPath, "utf8"), "Cover", "front_title");
    const centeredFrontTitleBlock = ensureYamlBoxAlign(frontTitleBlock, "center");
    await writeYamlPageBoxBlockReplacement(yamlPath, "Cover", "front_title", centeredFrontTitleBlock);

    const beforeYaml = await readFile(yamlPath, "utf8");
    expect(readCompactYamlBox(beforeYaml, "Cover", "front_title").align).toBe("center");

    await buildCoverEditablePptx(pptxPath, yamlPath);

    const coverSlideEntry = "ppt/slides/slide1.xml";
    const originalSlide = await readPptxZipEntryText(pptxPath, coverSlideEntry);
    const originalShape = extractPptxShapeXmlByName(originalSlide, "front_title");
    if (!originalShape) throw new Error("missing front_title shape on Cover slide");
    expect(originalShape).toMatch(/<a:pPr[^>]*algn="ctr"/);

    const editedText = "仁工智慧（合本）";
    let desktopShape = originalShape.replace(/\s+algn="ctr"/g, "");
    expect(desktopShape).not.toMatch(/<a:pPr[^>]*algn="ctr"/);
    desktopShape = rewritePptxShapeTextRuns(desktopShape, [{ text: editedText, bold: false }]);

    await writePptxZipEntry(
      pptxPath,
      coverSlideEntry,
      replacePptxShapeXml(originalSlide, "front_title", desktopShape),
    );

    const result = await parseCoverEditablePptxToYaml(pptxPath, yamlPath);
    expect(result.yaml_path).toBe(yamlPath);
    expect(result.updated).toBeGreaterThan(0);
    expect(result.unknown_shapes).toEqual([]);

    const box = readCompactYamlBox(await readFile(yamlPath, "utf8"), "Cover", "front_title");
    expect(box.text?.trim()).toBe(editedText);
    expect(box.align).toBe("center");
  }, 120_000);
});


describe("zh-TW editable PPTX export", () => {
  test("PPTX_OUTPUT_PATH targets comics-zh-TW/zh-TW-editable.pptx", () => {
    expect(PPTX_OUTPUT_PATH).toBe(join(OUTPUT_ROOT, "zh-TW-editable.pptx"));
  });

  test("buildEditablePptx is an exported async builder", () => {
    expect(typeof buildEditablePptx).toBe("function");
    expect(buildEditablePptx.length).toBeLessThanOrEqual(1);
  });

  test("default buildEditablePptxFromYaml uses ingested YAML target_size for presentation slide width not legacy 5749", async () => {
    const buildEditablePptxFromYamlProd = requirePdfOracleExport<
      (options: { yamlPath: string; outputPath: string }) => Promise<string>
    >("buildEditablePptxFromYaml");
    const root = await mkdtemp(join(tmpdir(), "zh-tw-pptx-default-yaml-geometry-"));
    const yamlPath = join(root, "zh-TW.yml");
    const pptxPath = join(root, "zh-TW-editable.pptx");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH, "utf8"));
    const overviewYaml = Bun.YAML.parse(await readFile(yamlPath, "utf8")) as {
      pages?: { Overview?: { target_size?: number[] } };
    };
    expect(overviewYaml.pages?.Overview?.target_size).toEqual([
      INGESTED_ZH_TW_TARGET_WIDTH_PX,
      INGESTED_ZH_TW_TARGET_HEIGHT_PX,
    ]);
    await buildEditablePptxFromYamlProd({ yamlPath, outputPath: pptxPath });
    const ingestedEmu = pptxSlideSizeEmuFromInches(
      INGESTED_ZH_TW_PPTX_SLIDE_WIDTH_IN,
      INGESTED_ZH_TW_PPTX_SLIDE_HEIGHT_IN,
    );
    const legacyEmu = pptxSlideSizeEmuFromInches(
      LEGACY_PPTX_SLIDE_WIDTH_IN,
      INGESTED_ZH_TW_PPTX_SLIDE_HEIGHT_IN,
    );
    const slideEmu = await readPptxPresentationSlideSizeEmu(pptxPath);
    expect(slideEmu.cx).toBe(ingestedEmu.cx);
    expect(slideEmu.cy).toBe(ingestedEmu.cy);
    expect(slideEmu.cx).not.toBe(legacyEmu.cx);
  }, 120_000);

  test("buildEditablePptx writes thirteen slides with rels and background media images", async () => {
    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-pptx-"));
    const outPath = join(tempDir, "zh-TW-editable.pptx");
    const written = await buildEditablePptx(outPath);
    expect(written).toBe(outPath);

    const zipBytes = new Uint8Array(await readFile(outPath));
    expect(zipBytes[0]).toBe(0x50);
    expect(zipBytes[1]).toBe(0x4b);

    const entries = await listPptxZipEntryNames(outPath);
    for (let slide = 1; slide <= PAGE_ORDER.length; slide += 1) {
      expect(entries).toContain(`ppt/slides/slide${slide}.xml`);
      expect(entries).toContain(`ppt/slides/_rels/slide${slide}.xml.rels`);
    }

    const mediaImages = entries.filter((name) => /^ppt\/media\/.+\.(png|jpe?g)$/i.test(name));
    expect(mediaImages.length).toBeGreaterThanOrEqual(PAGE_ORDER.length);
    const firstBackground = await loadImage(Buffer.from(await readPptxZipEntryBytes(outPath, mediaImages[0]!)));
    expect(firstBackground.height).toBeLessThanOrEqual(2000);
    expect(firstBackground.width).toBeLessThan(5749);
  }, 120_000);

  test("Ch5_B chef caption appears as editable a:t text on slide11", async () => {
    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-pptx-ch5b-"));
    const outPath = join(tempDir, "zh-TW-editable.pptx");
    await buildEditablePptx(outPath);

    const ch5BSlideIndex = PAGE_ORDER.indexOf("Ch5_B") + 1;
    expect(ch5BSlideIndex).toBe(11);

    const slideXml = await readPptxZipEntryText(outPath, `ppt/slides/slide${ch5BSlideIndex}.xml`);
    expect(slideXml).toMatch(/<a:t>/);
    expect(slideXml).toContain(CH5_B_PPTX_CAPTION_SNIPPET);
    expect(slideXml).not.toMatch(/<p:pic[^>]*>[\s\S]*舉例來說/);
  }, 120_000);

  test("PPTX text boxes carry YAML box ids and bold emphasis runs", async () => {
    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-pptx-anchors-"));
    const outPath = join(tempDir, "zh-TW-editable.pptx");
    await buildEditablePptx(outPath);

    const ch2ASlideIndex = PAGE_ORDER.indexOf("Ch2_A") + 1;
    expect(ch2ASlideIndex).toBe(4);

    const slideXml = await readPptxZipEntryText(outPath, `ppt/slides/slide${ch2ASlideIndex}.xml`);
    expect(slideXml).toContain('name="b004_body"');
    expect(slideXml).toMatch(/<a:rPr\b[^>]*\bb="1"/);
  }, 120_000);
});

describe("zh-TW editable PPTX parseback", () => {
  test("parseEditablePptxToYaml is an exported async writer", () => {
    expect("parseEditablePptxToYaml" in zhTwPipeline).toBe(true);
    expect(typeof zhTwPipeline.parseEditablePptxToYaml).toBe("function");
  });

  test("traverseShapes resolves nested group transform coordinates", () => {
    const mockSlideNode = [
      {
        "p:grpSp": [
          {
            "p:grpSpPr": [
              {
                "a:xfrm": [
                  { "a:off": [], ":@": { "@_x": "1000", "@_y": "2000" } },
                  { "a:ext": [], ":@": { "@_cx": "4000", "@_cy": "6000" } },
                  { "a:chOff": [], ":@": { "@_x": "0", "@_y": "0" } },
                  { "a:chExt": [], ":@": { "@_cx": "2000", "@_cy": "3000" } },
                ],
              },
            ],
          },
          {
            "p:sp": [
              {
                "p:nvSpPr": [
                  { "p:cNvPr": [], ":@": { "@_id": "42", "@_name": "test_shape" } },
                ],
              },
              {
                "p:spPr": [
                  {
                    "a:xfrm": [
                      { "a:off": [], ":@": { "@_x": "500", "@_y": "600" } },
                      { "a:ext": [], ":@": { "@_cx": "1000", "@_cy": "1200" } },
                    ],
                  },
                ],
              },
              {
                "p:txBody": [
                  { "a:p": [{ "a:r": [{ "a:t": [{ "#text": "Hello" }] }] }] },
                ],
              },
            ],
          },
        ],
      },
    ] as OrderedXmlNode[];

    const results: ParsedPptxShape[] = [];
    zhTwPipeline.traverseShapes(mockSlideNode, null, results);

    expect(results).toHaveLength(1);
    expect(results[0].name).toBe("test_shape");
    expect(results[0].text).toBe("Hello");
    expect(results[0].box).toEqual([1, 2, 2, 4]);
  });

  test("matchLegacyShapes maps shapes using text and spatial overlap heuristics", () => {
    const yamlBoxes = [
      { id: "b001_title", text: "真的傾聽\n人們想說的話", box: [100, 100, 400, 200] },
      { id: "b002_body", text: "不是：只聽受歡迎的人", box: [100, 300, 400, 400] },
    ] as BoxSpec[];

    const deskShapes: ParsedPptxShape[] = [
      {
        name: "Shape 1",
        text: "真的傾聽\n人們想說的話",
        box: [105, 105, 405, 205],
        font: undefined,
        align: "left",
        emphasis: [],
      },
      {
        name: "Shape 2",
        text: "不是：只聽受歡迎的人",
        box: [110, 310, 410, 410],
        font: undefined,
        align: "left",
        emphasis: [],
      },
    ];

    const matched = zhTwPipeline.matchLegacyShapes(deskShapes, yamlBoxes);
    expect(matched.get("b001_title")?.name).toBe("Shape 1");
    expect(matched.get("b002_body")?.name).toBe("Shape 2");
  });

  test("parseEditablePptxToYaml leaves temp YAML byte-identical when generated PPTX has no edits", async () => {
    const parseEditablePptxToYaml = requireParseEditablePptxToYaml();
    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-parseback-noop-"));
    const yamlPath = join(tempDir, "zh-TW.yml");
    const pptxPath = join(tempDir, "zh-TW-editable.pptx");
    const sourceYaml = await readFile(ZH_TW_YAML_PATH);
    await writeFile(yamlPath, sourceYaml);
    await buildEditablePptx(pptxPath);

    const result = await parseEditablePptxToYaml(pptxPath, yamlPath);
    expect(result.yaml_path).toBe(yamlPath);
    expect(result.updated).toBe(0);
    expect(result.unknown_shapes).toEqual([]);
    expect(await readFile(yamlPath)).toEqual(sourceYaml);
  }, 120_000);

  test("parseEditablePptxToYaml maps Ch2_A b004_body PPTX edits into compact YAML fields", async () => {
    const parseEditablePptxToYaml = requireParseEditablePptxToYaml();
    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-parseback-"));
    const yamlPath = join(tempDir, "zh-TW.yml");
    const pptxPath = join(tempDir, "zh-TW-editable.pptx");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH));
    await buildEditablePptx(pptxPath);

    const ch2ASlideIndex = PAGE_ORDER.indexOf("Ch2_A") + 1;
    const slideEntry = `ppt/slides/slide${ch2ASlideIndex}.xml`;
    const originalSlide = await readPptxZipEntryText(pptxPath, slideEntry);
    const originalShape = extractPptxShapeXmlByName(originalSlide, "b004_body");
    if (!originalShape) throw new Error("missing b004_body shape");

    const editedText = "當沒有人負責，沒有人會負責。";
    const expectedBox = [900, 2100, 4800, 2600];
    const expectedFontPx = 300;
    const sz = Math.round((expectedFontPx * 72 * 100) / PARSEBACK_PPTX_DPI);
    const x0 = Math.round((expectedBox[0]! / PARSEBACK_PPTX_DPI) * PARSEBACK_EMU_PER_INCH);
    const y0 = Math.round((expectedBox[1]! / PARSEBACK_PPTX_DPI) * PARSEBACK_EMU_PER_INCH);
    const cx = Math.round(((expectedBox[2]! - expectedBox[0]!) / PARSEBACK_PPTX_DPI) * PARSEBACK_EMU_PER_INCH);
    const cy = Math.round(((expectedBox[3]! - expectedBox[1]!) / PARSEBACK_PPTX_DPI) * PARSEBACK_EMU_PER_INCH);

    let editedShape = originalShape;
    editedShape = editedShape.replace(/<a:off x="\d+" y="\d+"/, `<a:off x="${x0}" y="${y0}"`);
    editedShape = editedShape.replace(/<a:ext cx="\d+" cy="\d+"/, `<a:ext cx="${cx}" cy="${cy}"`);
    editedShape = editedShape.replace(/sz="\d+"/g, `sz="${sz}"`);
    editedShape = editedShape.replace(/<a:pPr algn="[a-z]+"/, `<a:pPr algn="ctr"`);
    editedShape = rewritePptxShapeTextRuns(editedShape, [
      { text: "當沒有", bold: false },
      { text: "人", bold: true, hex: "F09444" },
      { text: "負責，", bold: false },
      { text: "沒有人", bold: true, hex: "F09444" },
      { text: "會負責。", bold: false },
    ]);

    await writePptxZipEntry(
      pptxPath,
      slideEntry,
      replacePptxShapeXml(originalSlide, "b004_body", editedShape),
    );

    const result = await parseEditablePptxToYaml(pptxPath, yamlPath);
    expect(result.yaml_path).toBe(yamlPath);
    expect(result.updated).toBeGreaterThanOrEqual(1);
    expect(result.unknown_shapes).toEqual([]);

    const box = readCompactYamlBox(await readFile(yamlPath, "utf8"), "Ch2_A", "b004_body");
    expect(box.box).toEqual(expectedBox);
    expect(box.font).toBe(expectedFontPx);
    expect(box.align).toBe("center");
    expect(box.text?.trim()).toBe(editedText);
    const editedRuns: PptxTextRun[] = [
      { text: "當沒有", bold: false },
      { text: "人", bold: true, hex: "F09444" },
      { text: "負責，", bold: false },
      { text: "沒有人", bold: true, hex: "F09444" },
      { text: "會負責。", bold: false },
    ];
    expect(box.emphasis).toEqual(pptxRunsToUtf16Ranges(editedRuns));
  }, 120_000);

  test("parseEditablePptxToYaml reconstructs Ch2_A b004_body emphasis from PPTX bold runs, not prior YAML", async () => {
    const parseEditablePptxToYaml = requireParseEditablePptxToYaml();
    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-parseback-emphasis-"));
    const yamlPath = join(tempDir, "zh-TW.yml");
    const pptxPath = join(tempDir, "zh-TW-editable.pptx");
    let yamlText = await readFile(ZH_TW_YAML_PATH, "utf8");
    const b004Block = extractYamlPageBoxBlock(yamlText, "Ch2_A", "b004_body");
    const b004WithoutEmphasis = b004Block.replace(/\n        emphasis: \[[^\n]+\]/, "");
    yamlText = yamlText.replace(b004Block, b004WithoutEmphasis);
    await writeFile(yamlPath, yamlText);
    await buildEditablePptx(pptxPath);

    const slideEntry = `ppt/slides/slide${PAGE_ORDER.indexOf("Ch2_A") + 1}.xml`;
    const originalSlide = await readPptxZipEntryText(pptxPath, slideEntry);
    const originalShape = extractPptxShapeXmlByName(originalSlide, "b004_body");
    if (!originalShape) throw new Error("missing b004_body shape");

    const editedShape = rewritePptxShapeTextRuns(originalShape, [
      { text: "當沒有", bold: true },
      { text: "人", bold: false, hex: "F09444" },
      { text: "負責，就", bold: false },
      { text: "沒有人", bold: false, hex: "F09444" },
      { text: "負責。", bold: true },
    ]);
    await writePptxZipEntry(
      pptxPath,
      slideEntry,
      replacePptxShapeXml(originalSlide, "b004_body", editedShape),
    );

    await parseEditablePptxToYaml(pptxPath, yamlPath);
    const box = readCompactYamlBox(await readFile(yamlPath, "utf8"), "Ch2_A", "b004_body");
    expect(box.text?.trim()).toBe(CH2_A_PARSEBACK_NOBODY);
    const reconstructionRuns: PptxTextRun[] = [
      { text: "當沒有", bold: true },
      { text: "人", bold: false, hex: "F09444" },
      { text: "負責，就", bold: false },
      { text: "沒有人", bold: false, hex: "F09444" },
      { text: "負責。", bold: true },
    ];
    expect(box.emphasis).toEqual(pptxRunsToUtf16Ranges(reconstructionRuns));
    expect(box.emphasis).not.toEqual([
      { start: 3, end: 4, bold: true },
      { start: 9, end: 12, bold: true },
    ]);
  }, 120_000);

  test("parseEditablePptxToYaml preserves compact YAML formatting outside the edited Ch2_A b004_body block", async () => {
    const parseEditablePptxToYaml = requireParseEditablePptxToYaml();
    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-parseback-fmt-"));
    const yamlPath = join(tempDir, "zh-TW.yml");
    const pptxPath = join(tempDir, "zh-TW-editable.pptx");
    const sourceYaml = await readFile(ZH_TW_YAML_PATH, "utf8");
    await writeFile(yamlPath, sourceYaml);
    await buildEditablePptx(pptxPath);

    const beforeBlock = extractYamlPageBoxBlock(sourceYaml, "Ch2_A", "b004_body");
    const neighborBlock = extractYamlPageBoxBlock(sourceYaml, "Ch2_A", "b003_body");
    expect(beforeBlock).toContain("text: |-");

    const ch2ASlideIndex = PAGE_ORDER.indexOf("Ch2_A") + 1;
    const slideEntry = `ppt/slides/slide${ch2ASlideIndex}.xml`;
    const originalSlide = await readPptxZipEntryText(pptxPath, slideEntry);
    const originalShape = extractPptxShapeXmlByName(originalSlide, "b004_body");
    if (!originalShape) throw new Error("missing b004_body shape");
    const editedShape = rewritePptxShapeTextRuns(originalShape, [
      { text: "當沒有", bold: false },
      { text: "人", bold: true, hex: "F09444" },
      { text: "負責，", bold: false },
      { text: "沒有人", bold: true, hex: "F09444" },
      { text: "會負責。", bold: false },
    ]);
    await writePptxZipEntry(
      pptxPath,
      slideEntry,
      replacePptxShapeXml(originalSlide, "b004_body", editedShape),
    );

    await parseEditablePptxToYaml(pptxPath, yamlPath);
    const yamlAfter = await readFile(yamlPath, "utf8");

    expect(extractYamlPageBoxBlock(yamlAfter, "Ch2_A", "b003_body")).toBe(neighborBlock);
    const afterBlock = extractYamlPageBoxBlock(yamlAfter, "Ch2_A", "b004_body");
    expect(afterBlock).toContain("text: |-");
    expect(afterBlock).toContain("emphasis:");
    expect(afterBlock).not.toBe(beforeBlock);
    const box = readCompactYamlBox(yamlAfter, "Ch2_A", "b004_body");
    expect(box.text?.trim()).toBe("當沒有人負責，沒有人會負責。");
    const fmtRuns: PptxTextRun[] = [
      { text: "當沒有", bold: false },
      { text: "人", bold: true, hex: "F09444" },
      { text: "負責，", bold: false },
      { text: "沒有人", bold: true, hex: "F09444" },
      { text: "會負責。", bold: false },
    ];
    expect(box.emphasis).toEqual(pptxRunsToUtf16Ranges(fmtRuns));
  }, 120_000);

  test("parseEditablePptxToYaml preserves existing YAML emphasis when PPTX returns no emphasis runs", async () => {
    const parseEditablePptxToYaml = requireParseEditablePptxToYaml();
    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-parseback-preserve-emphasis-"));
    const yamlPath = join(tempDir, "zh-TW.yml");
    const pptxPath = join(tempDir, "zh-TW-editable.pptx");
    const sourceYaml = await readFile(ZH_TW_YAML_PATH, "utf8");
    await writeFile(yamlPath, sourceYaml);
    const expectedBefore = readCompactYamlBox(sourceYaml, "Ch2_A", "b004_body").emphasis;
    expect(expectedBefore).toBeDefined();
    expect(expectedBefore!.length).toBeGreaterThan(0);

    await buildEditablePptxFromYaml(yamlPath, pptxPath);

    const slideEntry = `ppt/slides/slide${PAGE_ORDER.indexOf("Ch2_A") + 1}.xml`;
    const originalSlide = await readPptxZipEntryText(pptxPath, slideEntry);
    const originalShape = extractPptxShapeXmlByName(originalSlide, "b004_body");
    if (!originalShape) throw new Error("missing b004_body shape");

    const strippedShape = rewritePptxShapeToPlainUnstyledRuns(originalShape);
    expect(extractPptxTextRunsFromShape(strippedShape).every((run) => !run.bold && !run.hex)).toBe(true);

    await writePptxZipEntry(
      pptxPath,
      slideEntry,
      replacePptxShapeXml(originalSlide, "b004_body", strippedShape),
    );

    await parseEditablePptxToYaml(pptxPath, yamlPath);
    const box = readCompactYamlBox(await readFile(yamlPath, "utf8"), "Ch2_A", "b004_body");
    expect(box.text?.trim()).toBe(CH2_A_PARSEBACK_NOBODY);
    expect(box.emphasis).toEqual(expectedBefore);
  }, 120_000);

  test("parseEditablePptxToYaml updates Ch5_B b004_body without reintroducing page md", async () => {
    const parseEditablePptxToYaml = requireParseEditablePptxToYaml();
    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-parseback-md-free-"));
    const yamlPath = join(tempDir, "zh-TW.yml");
    const pptxPath = join(tempDir, "zh-TW-editable.pptx");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH, "utf8"));

    await buildEditablePptxFromYaml(yamlPath, pptxPath);
    const slideEntry = `ppt/slides/slide${PAGE_ORDER.indexOf("Ch5_B") + 1}.xml`;
    const originalSlide = await readPptxZipEntryText(pptxPath, slideEntry);
    const originalShape = extractPptxShapeXmlByName(originalSlide, "b004_body");
    if (!originalShape) throw new Error("missing b004_body shape");
    const priorRuns = extractPptxTextRunsFromShape(originalShape);
    const editedRuns = priorRuns.map((run) => ({
      text: run.text.replace(CH5_B_PARSEBACK_SOLIDARITY_OLD, CH5_B_PARSEBACK_SOLIDARITY_NEW_MID),
      bold: run.bold,
      hex: run.hex,
    }));
    const editedShape = rewritePptxShapeTextRuns(originalShape, editedRuns);
    await writePptxZipEntry(
      pptxPath,
      slideEntry,
      replacePptxShapeXml(originalSlide, "b004_body", editedShape),
    );

    const mdBefore = (await loadYamlTextSpecs(yamlPath)).Ch5_B.ocr_markdown ?? "";
    expect(mdBefore).toBe("");

    await parseEditablePptxToYaml(pptxPath, yamlPath);

    const yamlAfter = await readFile(yamlPath, "utf8");
    const box = readCompactYamlBox(yamlAfter, "Ch5_B", "b004_body");
    expect(box.text?.trim()).toContain(CH5_B_PARSEBACK_SOLIDARITY_NEW_MID);
    expect(yamlAfter).not.toContain("    md: |-");
    const mdAfter = (await loadYamlTextSpecs(yamlPath)).Ch5_B.ocr_markdown ?? "";
    expect(mdAfter).toBe("");
  }, 120_000);

  test("parseEditablePptxToYaml rejects unknown shape ids with VerificationError", async () => {
    const parseEditablePptxToYaml = requireParseEditablePptxToYaml();
    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-parseback-unknown-"));
    const yamlPath = join(tempDir, "zh-TW.yml");
    const pptxPath = join(tempDir, "zh-TW-editable.pptx");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH));
    await buildEditablePptx(pptxPath);

    const slideEntry = `ppt/slides/slide${PAGE_ORDER.indexOf("Ch2_A") + 1}.xml`;
    const slideXml = await readPptxZipEntryText(pptxPath, slideEntry);
    const orphanShape =
      '<p:sp><p:nvSpPr><p:cNvPr id="99999" name="orphan_not_in_yaml"></p:cNvPr><p:cNvSpPr/><p:nvPr></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst></a:avLst></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1200"/><a:t>幽靈</a:t></a:r></a:p></p:txBody></p:sp>';
    await writePptxZipEntry(
      pptxPath,
      slideEntry,
      slideXml.replace("</p:spTree>", `${orphanShape}</p:spTree>`),
    );

    const before = await readFile(yamlPath, "utf8");
    await expect(parseEditablePptxToYaml(pptxPath, yamlPath)).rejects.toThrow(VerificationError);
    await expect(parseEditablePptxToYaml(pptxPath, yamlPath)).rejects.toThrow(/orphan_not_in_yaml/);
    expect(await readFile(yamlPath, "utf8")).toBe(before);
  }, 120_000);
});

describe("zh-TW emphasis color and Overview alignment", () => {
  test("loadYamlTextSpecs preserves emphasis color on color-only YAML span", async () => {
    const root = await mkdtemp(join(tmpdir(), "zh-tw-emphasis-color-yaml-"));
    const yamlPath = join(root, "zh-TW.yml");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH, "utf8"));
    await writeMinimalYamlPageBoxPatch(yamlPath, "Ch4_A", CH4_A_EMPHASIS_COLOR_BOX, {
      emphasis: [{ text: EMPHASIS_COLOR_COMPASS_SPAN, color: EMPHASIS_COLOR_COMPASS }],
    });
    const specs = await loadYamlTextSpecs(yamlPath);
    const box = findBoxById(specs.Ch4_A.boxes, CH4_A_EMPHASIS_COLOR_BOX);
    expect(box?.emphasis).toEqual([{ text: EMPHASIS_COLOR_COMPASS_SPAN, color: EMPHASIS_COLOR_COMPASS }]);
  });

  test("vector render paints YAML emphasis color span with green ink in box pixels", async () => {
    await probeCanvasRuntime(EMPHASIS_COLOR_COMPASS_SPAN);
    const root = await mkdtemp(join(tmpdir(), "zh-tw-emphasis-color-render-"));
    const { canvas, boxRect } = await renderSyntheticCh4AEmphasisColorVector(root);
    const greenHits = countPixelsCloseToRgb(canvas, boxRect, EMPHASIS_COLOR_COMPASS_RGB);
    expect(greenHits).toBeGreaterThanOrEqual(MIN_EMPHASIS_GREEN_PIXEL_CLUSTER);
  }, 60_000);

  test("loadYamlTextSpecs preserves bold and color on the same emphasis span", async () => {
    const root = await mkdtemp(join(tmpdir(), "zh-tw-emphasis-bold-color-yaml-"));
    const yamlPath = join(root, "zh-TW.yml");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH, "utf8"));
    await writeMinimalYamlPageBoxPatch(yamlPath, "Ch1_A", "b006_body", {
      emphasis: [
        { text: EMPHASIS_COLOR_CH1_BALLOT_SPAN, bold: true, color: EMPHASIS_COLOR_CH1_BALLOT },
      ],
    });
    const specs = await loadYamlTextSpecs(yamlPath);
    const box = findBoxById(specs.Ch1_A.boxes, "b006_body");
    expect(box?.emphasis).toEqual([
      { text: EMPHASIS_COLOR_CH1_BALLOT_SPAN, bold: true, color: EMPHASIS_COLOR_CH1_BALLOT },
    ]);
  });

  test("PPTX export emits a:srgbClr on color-emphasized span run", async () => {
    const root = await mkdtemp(join(tmpdir(), "zh-tw-emphasis-color-pptx-"));
    const yamlPath = join(root, "zh-TW.yml");
    const pptxPath = join(root, "zh-TW-editable.pptx");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH, "utf8"));
    await writeMinimalYamlPageBoxPatch(yamlPath, "Ch4_A", CH4_A_EMPHASIS_COLOR_BOX, {
      emphasis: [{ text: EMPHASIS_COLOR_COMPASS_SPAN, color: EMPHASIS_COLOR_COMPASS }],
    });
    await buildEditablePptx(pptxPath);
    const slideIndex = PAGE_ORDER.indexOf("Ch4_A") + 1;
    const slideXml = await readPptxZipEntryText(pptxPath, `ppt/slides/slide${slideIndex}.xml`);
    const shapeXml = extractPptxShapeXmlByName(slideXml, CH4_A_EMPHASIS_COLOR_BOX);
    expect(shapeXml).toBeDefined();
    const compassRun = extractPptxSrgbClrRuns(shapeXml!).find((run) => run.text.includes("羅盤"));
    expect(compassRun?.hex).toBe("4F8F00");
  }, 120_000);

  test("PPTX export keeps bold and a:srgbClr on the same emphasized run", async () => {
    const root = await mkdtemp(join(tmpdir(), "zh-tw-emphasis-bold-color-pptx-"));
    const yamlPath = join(root, "zh-TW.yml");
    const pptxPath = join(root, "zh-TW-editable.pptx");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH, "utf8"));
    await writeMinimalYamlPageBoxPatch(yamlPath, "Ch1_A", "b006_body", {
      emphasis: [
        { text: EMPHASIS_COLOR_CH1_BALLOT_SPAN, bold: true, color: EMPHASIS_COLOR_CH1_BALLOT },
      ],
    });
    await buildEditablePptx(pptxPath);
    const slideIndex = PAGE_ORDER.indexOf("Ch1_A") + 1;
    const slideXml = await readPptxZipEntryText(pptxPath, `ppt/slides/slide${slideIndex}.xml`);
    const shapeXml = extractPptxShapeXmlByName(slideXml, "b006_body");
    expect(shapeXml).toBeDefined();
    const ballotRun = extractPptxTextRunsFromShape(shapeXml!).find((run) => run.text.includes(EMPHASIS_COLOR_CH1_BALLOT_SPAN));
    expect(ballotRun?.bold).toBe(true);
    expect(ballotRun?.hex).toBe("DF5D46");
  }, 120_000);

  test("parseEditablePptxToYaml maps PPTX a:srgbClr run back to YAML emphasis color", async () => {
    const root = await mkdtemp(join(tmpdir(), "zh-tw-emphasis-color-parseback-"));
    const yamlPath = join(root, "zh-TW.yml");
    const pptxPath = join(root, "zh-TW-editable.pptx");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH, "utf8"));
    await buildEditablePptx(pptxPath);
    const slideEntry = `ppt/slides/slide${PAGE_ORDER.indexOf("Ch4_A") + 1}.xml`;
    const originalSlide = await readPptxZipEntryText(pptxPath, slideEntry);
    const originalShape = extractPptxShapeXmlByName(originalSlide, CH4_A_EMPHASIS_COLOR_BOX);
    if (!originalShape) throw new Error(`missing ${CH4_A_EMPHASIS_COLOR_BOX} shape`);
    const coloredShape = rewritePptxShapeColorRun(originalShape, EMPHASIS_COLOR_COMPASS_SPAN, EMPHASIS_COLOR_COMPASS);
    await writePptxZipEntry(
      pptxPath,
      slideEntry,
      replacePptxShapeXml(originalSlide, CH4_A_EMPHASIS_COLOR_BOX, coloredShape),
    );
    await parseEditablePptxToYaml(pptxPath, yamlPath);
    const box = readCompactYamlBox(await readFile(yamlPath, "utf8"), "Ch4_A", CH4_A_EMPHASIS_COLOR_BOX);
    expect(box.emphasis).toEqual(
      expect.arrayContaining([expect.objectContaining({ text: EMPHASIS_COLOR_COMPASS_SPAN, color: EMPHASIS_COLOR_COMPASS })]),
    );
  }, 120_000);

  test("parseEditablePptxToYaml maps PPTX bold+color run back to YAML emphasis bold and color", async () => {
    const root = await mkdtemp(join(tmpdir(), "zh-tw-emphasis-bold-color-parseback-"));
    const yamlPath = join(root, "zh-TW.yml");
    const pptxPath = join(root, "zh-TW-editable.pptx");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH, "utf8"));
    await buildEditablePptx(pptxPath);
    const slideEntry = `ppt/slides/slide${PAGE_ORDER.indexOf("Ch1_A") + 1}.xml`;
    const originalSlide = await readPptxZipEntryText(pptxPath, slideEntry);
    const originalShape = extractPptxShapeXmlByName(originalSlide, "b006_body");
    if (!originalShape) throw new Error("missing b006_body shape");
    const styledShape = rewritePptxShapeBoldColorRun(
      originalShape,
      EMPHASIS_COLOR_CH1_BALLOT_SPAN,
      EMPHASIS_COLOR_CH1_BALLOT,
      true,
    );
    await writePptxZipEntry(
      pptxPath,
      slideEntry,
      replacePptxShapeXml(originalSlide, "b006_body", styledShape),
    );
    await parseEditablePptxToYaml(pptxPath, yamlPath);
    const box = readCompactYamlBox(await readFile(yamlPath, "utf8"), "Ch1_A", "b006_body");
    expect(box.emphasis).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          text: EMPHASIS_COLOR_CH1_BALLOT_SPAN,
          bold: true,
          color: EMPHASIS_COLOR_CH1_BALLOT,
        }),
      ]),
    );
  }, 120_000);

  test("Overview alignment when YAML sets b001_title and b003_title to center", async () => {
    const root = await mkdtemp(join(tmpdir(), "zh-tw-overview-align-"));
    const yamlPath = join(root, "zh-TW.yml");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH, "utf8"));
    await writeMinimalYamlPageBoxPatch(yamlPath, "Overview", "b001_title", { align: "center" });
    await writeMinimalYamlPageBoxPatch(yamlPath, "Overview", "b003_title", { align: "center" });
    const specs = await loadYamlTextSpecs(yamlPath);
    expect(findBoxById(specs.Overview.boxes, "b001_title")?.align).toBe("center");
    expect(findBoxById(specs.Overview.boxes, "b003_title")?.align).toBe("center");
    const pptxPath = join(root, "zh-TW-editable.pptx");
    await buildEditablePptx(pptxPath);
    const slideXml = await readPptxZipEntryText(pptxPath, "ppt/slides/slide1.xml");
    const b001 = extractPptxShapeXmlByName(slideXml, "b001_title");
    const b003 = extractPptxShapeXmlByName(slideXml, "b003_title");
    expect(b001).toMatch(/<a:pPr[^>]*algn="ctr"/);
    expect(b003).toMatch(/<a:pPr[^>]*algn="ctr"/);
  }, 120_000);
});

describe("zh-TW PDF oracle (canonical Desktop export)", () => {
  test("inspectCanonicalPdf reports page count, pt size, 600dpi render size, and PDF-derived PPTX inches", async () => {
    if (!existsSync(CANONICAL_ZH_TW_EDITABLE_PDF)) {
      throw new Error(`missing canonical oracle PDF at ${CANONICAL_ZH_TW_EDITABLE_PDF}`);
    }
    const inspectCanonicalPdf = requirePdfOracleExport<
      (pdfPath?: string) => Promise<InspectCanonicalPdfResult>
    >("inspectCanonicalPdf");
    const report = await inspectCanonicalPdf(CANONICAL_ZH_TW_EDITABLE_PDF);
    expect(report.pageCount).toBe(PDF_ORACLE_PAGE_COUNT);
    expect(report.pageSizePt).toEqual([...PDF_ORACLE_PAGE_SIZE_PT]);
    expect(report.renderDpi).toBe(PDF_ORACLE_RENDER_DPI);
    expect(report.renderSizePx).toEqual([...PDF_ORACLE_RENDER_SIZE_PX]);
    expect(report.pptxSlideSizeIn[0]).toBeCloseTo(PDF_ORACLE_PPTX_SLIDE_WIDTH_IN, 5);
    expect(report.pptxSlideSizeIn[1]).toBeCloseTo(PDF_ORACLE_PPTX_SLIDE_HEIGHT_IN, 5);
  });

  test("extractPdfTextRuns on page 1 includes Overview title and body font roles", async () => {
    if (!existsSync(CANONICAL_ZH_TW_EDITABLE_PDF)) {
      throw new Error(`missing canonical oracle PDF at ${CANONICAL_ZH_TW_EDITABLE_PDF}`);
    }
    const extractPdfTextRuns = requirePdfOracleExport<
      (pdfPath: string, pageIndex: number) => Promise<PdfTextRun[]>
    >("extractPdfTextRuns");
    const runs = await extractPdfTextRuns(CANONICAL_ZH_TW_EDITABLE_PDF, 1);
    const joined = runs.map((r) => r.text).join("");
    expect(joined).toContain("真的傾聽");
    const titleRun = runs.find((r) => r.text.includes("真的傾聽"));
    expect(titleRun).toBeDefined();
    expect(titleRun?.color?.toUpperCase().replace("#", "")).toBe("191919");
    expect(titleRun?.fontFamily).toBeTruthy();
    expect(titleRun?.fontSize).toBeGreaterThan(0);
    const bodyRun = runs.find((r) => r.text.includes("不是：只聽"));
    expect(bodyRun).toBeDefined();
    expect(bodyRun?.color?.toUpperCase().replace("#", "")).toBe("686868");
    expect(bodyRun?.fontFamily).toBeTruthy();
    expect(bodyRun?.fontSize).toBeGreaterThan(0);
    expect(bodyRun?.fontSize).toBeLessThan(titleRun?.fontSize ?? Infinity);
  });

  test("buildEditablePptxFromYaml with pdfPageModel sets presentation slide size from PDF page box not legacy 5749 width", async () => {
    const buildEditablePptxFromYamlWithOptions = requirePdfOracleExport<
      (options: {
        yamlPath: string;
        outputPath: string;
        pdfPageModel?: { pageSizePt: [number, number]; renderDpi: number };
      }) => Promise<string>
    >("buildEditablePptxFromYaml");
    const root = await mkdtemp(join(tmpdir(), "zh-tw-pdf-oracle-pptx-"));
    const yamlPath = join(root, "zh-TW.yml");
    const pptxPath = join(root, "zh-TW-editable.pptx");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH, "utf8"));
    await buildEditablePptxFromYamlWithOptions({
      yamlPath,
      outputPath: pptxPath,
      pdfPageModel: {
        pageSizePt: [...PDF_ORACLE_PAGE_SIZE_PT],
        renderDpi: PDF_ORACLE_RENDER_DPI,
      },
    });
    const oracleEmu = pptxSlideSizeEmuFromInches(
      PDF_ORACLE_PPTX_SLIDE_WIDTH_IN,
      PDF_ORACLE_PPTX_SLIDE_HEIGHT_IN,
    );
    const legacyEmu = pptxSlideSizeEmuFromInches(LEGACY_PPTX_SLIDE_WIDTH_IN, PDF_ORACLE_PPTX_SLIDE_HEIGHT_IN);
    const slideEmu = await readPptxPresentationSlideSizeEmu(pptxPath);
    expect(slideEmu.cx).toBe(oracleEmu.cx);
    expect(slideEmu.cy).toBe(oracleEmu.cy);
    expect(slideEmu.cx).not.toBe(legacyEmu.cx);
  }, 120_000);

  test("comparePdfOraclePages rejects dimension mismatch before pixel comparison", async () => {
    const comparePdfOraclePages = requirePdfOracleExport<
      (leftPngPath: string, rightPngPath: string, options?: { renderDpi?: number }) => Promise<ComparePdfOraclePagesResult>
    >("comparePdfOraclePages");
    const root = await mkdtemp(join(tmpdir(), "zh-tw-pdf-oracle-diff-"));
    const leftPath = join(root, "oracle-page-1.png");
    const rightPath = join(root, "wrong-size.png");
    await writeWhitePng(leftPath, PDF_ORACLE_RENDER_SIZE_PX[0], PDF_ORACLE_RENDER_SIZE_PX[1]);
    await writeWhitePng(rightPath, PDF_ORACLE_RENDER_SIZE_PX[0] + 2, PDF_ORACLE_RENDER_SIZE_PX[1]);
    const result = await comparePdfOraclePages(leftPath, rightPath, { renderDpi: PDF_ORACLE_RENDER_DPI });
    expect(result.ok).toBe(false);
    expect(result.dimensionMismatch).toBe(true);
    expect(result.diffPixels ?? 0).toBe(0);
  });

  test("sortPdfRenderedPagePaths orders pdftoppm page suffixes numerically not lexicographically", () => {
    const sortPdfRenderedPagePaths = requirePdfOracleExport<(paths: string[]) => string[]>(
      "sortPdfRenderedPagePaths",
    );
    const shuffled = [
      "/tmp/page-10.png",
      "/tmp/page-2.png",
      "/tmp/page-1.png",
      "/tmp/page-11.png",
    ];
    expect(sortPdfRenderedPagePaths(shuffled)).toEqual([
      "/tmp/page-1.png",
      "/tmp/page-2.png",
      "/tmp/page-10.png",
      "/tmp/page-11.png",
    ]);
  });

  test("renderPdfPages clears stale rendered page rasters before collecting current pages", async () => {
    const renderPdfPages = requirePdfOracleExport<
      (pdfPath: string, outputDir: string, renderDpi: number, prefix: string) => Promise<string[]>
    >("renderPdfPages");
    const root = await mkdtemp(join(tmpdir(), "zh-tw-pdf-pages-stale-"));
    const stalePath = join(root, "page-99.png");
    await writeWhitePng(stalePath, 2, 2);

    const pages = await renderPdfPages(CANONICAL_ZH_TW_EDITABLE_PDF, root, PDF_ORACLE_RENDER_DPI, "page");

    expect(pages).toHaveLength(PAGE_ORDER.length);
    expect(pages).not.toContain(stalePath);
    expect(existsSync(stalePath)).toBe(false);
  }, 120_000);

  test("diffPdfOracleRenderedPageLists rejects canonical/generated page-count mismatch before pixel diff", async () => {
    const diffPdfOracleRenderedPageLists = requirePdfOracleExport<
      (
        canonicalPaths: string[],
        generatedPaths: string[],
        options?: { renderDpi?: number },
      ) => Promise<PdfOracleRoundTripReport>
    >("diffPdfOracleRenderedPageLists");
    const root = await mkdtemp(join(tmpdir(), "zh-tw-pdf-oracle-page-mismatch-"));
    const canonicalPaths: string[] = [];
    const generatedPaths: string[] = [];
    for (let page = 1; page <= 2; page += 1) {
      const canonicalPath = join(root, `canonical-page-${page}.png`);
      await writeWhitePng(canonicalPath, 8, 8);
      canonicalPaths.push(canonicalPath);
    }
    const generatedPath = join(root, "generated-page-1.png");
    await writeWhitePng(generatedPath, 8, 8);
    generatedPaths.push(generatedPath);

    const report = await diffPdfOracleRenderedPageLists(canonicalPaths, generatedPaths, {
      renderDpi: PDF_ORACLE_RENDER_DPI,
    });
    expect(report.ok).toBe(false);
    expect(report.renderer.kind).toBe("incomplete");
    expect(report.incompleteReason ?? "").toMatch(/page/i);
    expect(report.pageDiffs ?? []).toHaveLength(0);
  });

  test("diffPdfOracleRenderedPageLists rejects per-page dimension mismatch when page counts match", async () => {
    const diffPdfOracleRenderedPageLists = requirePdfOracleExport<
      (
        canonicalPaths: string[],
        generatedPaths: string[],
        options?: { renderDpi?: number },
      ) => Promise<PdfOracleRoundTripReport>
    >("diffPdfOracleRenderedPageLists");
    const root = await mkdtemp(join(tmpdir(), "zh-tw-pdf-oracle-dim-mismatch-"));
    const canonicalPath = join(root, "canonical-page-1.png");
    const generatedPath = join(root, "generated-page-1.png");
    await writeWhitePng(canonicalPath, 8, 8);
    await writeWhitePng(generatedPath, 10, 8);

    const report = await diffPdfOracleRenderedPageLists([canonicalPath], [generatedPath], {
      renderDpi: PDF_ORACLE_RENDER_DPI,
    });
    expect(report.ok).toBe(false);
    expect(report.pageDiffs ?? []).toHaveLength(1);
    expect(report.pageDiffs![0]!.pageId).toBe("Overview");
  });



  test("runPdfOracleRoundTrip requires pptx-pdf export path and 5742×8000 geometry before pixel diff", async () => {
    if (!existsSync(CANONICAL_ZH_TW_EDITABLE_PDF)) {
      throw new Error(`missing canonical oracle PDF at ${CANONICAL_ZH_TW_EDITABLE_PDF}`);
    }
    const runPdfOracleRoundTrip = requirePdfOracleExport<
      (options: {
        canonicalPdfPath: string;
        yamlPath: string;
        outputDir: string;
        renderDpi?: number;
      }) => Promise<PdfOracleRoundTripReport>
    >("runPdfOracleRoundTrip");

    const root = await mkdtemp(join(tmpdir(), "zh-tw-pdf-oracle-roundtrip-"));
    const yamlPath = join(root, "zh-TW.yml");
    const outputDir = join(root, "out");
    await mkdir(outputDir, { recursive: true });
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH, "utf8"));

    const report = await runPdfOracleRoundTrip({
      canonicalPdfPath: CANONICAL_ZH_TW_EDITABLE_PDF,
      yamlPath,
      outputDir,
      renderDpi: PDF_ORACLE_RENDER_DPI,
    });

    expect(report.canonicalRenderSizePx).toEqual([...PDF_ORACLE_RENDER_SIZE_PX]);

    if (report.pageRenderSizePx) {
      expect(report.pageRenderSizePx).toEqual([...PDF_ORACLE_RENDER_SIZE_PX]);
    }
    if (report.generatedRenderSizePx) {
      expect(report.generatedRenderSizePx).toEqual([...PDF_ORACLE_RENDER_SIZE_PX]);
    }

    expectPdfOracleRoundTripRendererContract(report);

    if (report.renderer.kind === "pptx-pdf") {
      expect(report.pageDiffs?.length ?? 0).toBeGreaterThan(0);
      for (const metric of report.pageDiffs ?? []) {
        expect(metric.pageId.length).toBeGreaterThan(0);
        expect(
          metric.diffPixels !== undefined || metric.mismatchRatio !== undefined,
        ).toBe(true);
      }
    }
  }, 600_000);


  test("ingestCanonicalPdfToYaml writes PDF-derived metadata and rescales target-space coordinates without changing src", async () => {
    if (!existsSync(CANONICAL_ZH_TW_EDITABLE_PDF)) {
      throw new Error(`missing canonical oracle PDF at ${CANONICAL_ZH_TW_EDITABLE_PDF}`);
    }

    const ingestCanonicalPdfToYaml = requirePdfOracleExport<
      (options: {
        canonicalPdfPath: string;
        yamlPath: string;
        outputYamlPath: string;
        renderDpi?: number;
        dryRun?: boolean;
      }) => Promise<IngestCanonicalPdfToYamlReport>
    >("ingestCanonicalPdfToYaml");

    const tempDir = await mkdtemp(join(tmpdir(), "zh-tw-pdf-ingest-yaml-"));
    const yamlPath = join(tempDir, "zh-TW-source.yml");
    const outputYamlPath = join(tempDir, "zh-TW-ingested.yml");
    await writeFile(yamlPath, await readFile(ZH_TW_YAML_PATH, "utf8"));

    const report = await ingestCanonicalPdfToYaml({
      canonicalPdfPath: CANONICAL_ZH_TW_EDITABLE_PDF,
      yamlPath,
      outputYamlPath,
      renderDpi: PDF_ORACLE_RENDER_DPI,
      dryRun: false,
    });

    expect(report.canonicalRenderSizePx).toEqual([...PDF_ORACLE_RENDER_SIZE_PX]);
    expect(report.updatedPagesCount).toBeGreaterThanOrEqual(PDF_ORACLE_PAGE_COUNT);
    expect(report.renderDpi).toBe(PDF_ORACLE_RENDER_DPI);
    expect(report.dryRun).toBe(false);
    expect(report.outputYamlPath).toBe(outputYamlPath);

    const outputYaml = await readFile(outputYamlPath, "utf8");
    const doc = Bun.YAML.parse(outputYaml) as {
      pages: Record<
        string,
        {
          target_size?: number[];
          pdf_page_box?: number[];
          render_dpi?: number;
          boxes?: CompactYamlBox[];
        }
      >;
    };

    expect(doc.pages.Overview?.target_size).toEqual([...PDF_ORACLE_RENDER_SIZE_PX]);
    expect(doc.pages.Overview?.pdf_page_box).toEqual([...PDF_ORACLE_PAGE_SIZE_PT]);
    expect(doc.pages.Overview?.render_dpi).toBe(PDF_ORACLE_RENDER_DPI);

    const scaleTargetX = (legacyX: number) =>
      Math.round((legacyX * PDF_ORACLE_RENDER_SIZE_PX[0]) / LEGACY_ZH_TW_TARGET_WIDTH_PX);

    const b001 = readCompactYamlBox(outputYaml, "Overview", "b001_title");
    expect(b001.src).toEqual([...OVERVIEW_B001_SRC]);
    expect(b001.box?.[0]).toBe(scaleTargetX(OVERVIEW_B001_TITLE_BOX_X_LEGACY));
    expect(b001.box?.[1]).toBe(284);
    expect(b001.tight?.[0]).toBe(scaleTargetX(OVERVIEW_B001_TIGHT_X_LEGACY));
    expect(b001.tight?.[1]).toBe(329);
  }, 120_000);
});
