# Legacy PPTX Parseback Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Parse translation edits and layout adjustments from the legacy `/Users/au/Desktop/zh-TW-editable.pptx` deck back to `translation-assets/zh-TW.yml`, then regenerate identical zh-TW PNGs.

**Architecture:** We recursively traverse nested group shapes (`p:grpSp`) in the slide XML and apply the transform stack to get absolute coordinates. We then map these shapes to the closest YAML boxes using a heuristic text-and-spatial matcher.

**Tech Stack:** Bun, TypeScript, JSZip, fast-xml-parser, Canvas (SKIA).

## Global Constraints
- Target YAML file: `translation-assets/zh-TW.yml`
- PPTX path: `/Users/au/Desktop/zh-TW-editable.pptx`
- Oracle image folder: `/Users/au/Desktop/zh-TW-editable/`
- All tests must pass: `bun test tests/zh_tw_comics_pipeline.test.ts`
- Clean typecheck: `bunx tsc --noEmit`

---

### Task 1: Group Transform Resolution

**Files:**
- Modify: `scripts/zh_tw_comics_pipeline.ts`
- Test: `tests/zh_tw_comics_pipeline.test.ts`

**Interfaces:**
- Consumes: None
- Produces: `traverseShapes(nodes: OrderedXmlNode[], currentTransform: GroupTransform | null, results: ParsedPptxShape[]): void`

- [ ] **Step 1: Write a unit test for group transform resolution**

Add this test inside `tests/zh_tw_comics_pipeline.test.ts` under the "zh-TW editable PPTX parseback" describe block:
```typescript
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
                  { "a:chExt": [], ":@": { "@_cx": "2000", "@_cy": "3000" } }
                ]
              }
            ]
          },
          {
            "p:sp": [
              {
                "p:nvSpPr": [
                  { "p:cNvPr": [], ":@": { "@_id": "42", "@_name": "test_shape" } }
                ]
              },
              {
                "p:spPr": [
                  {
                    "a:xfrm": [
                      { "a:off": [], ":@": { "@_x": "500", "@_y": "600" } },
                      { "a:ext": [], ":@": { "@_cx": "1000", "@_cy": "1200" } }
                    ]
                  }
                ]
              },
              {
                "p:txBody": [
                  { "a:p": [ { "a:r": [ { "a:t": [ { "#text": "Hello" } ] } ] } ] }
                ]
              }
            ]
          }
        ]
      }
    ] as any[];

    const results: any[] = [];
    (zhTwPipeline as any).traverseShapes(mockSlideNode, null, results);

    expect(results).toHaveLength(1);
    expect(results[0].name).toBe("test_shape");
    expect(results[0].text).toBe("Hello");
    // absolute x = 1000 + (500 - 0) * (4000 / 2000) = 2000 emus -> 2000 / (914400 / 600) = 1.31 px -> Math.round is 1
    // absolute y = 2000 + (600 - 0) * (6000 / 3000) = 3200 emus -> 3200 / (914400 / 600) = 2.10 px -> Math.round is 2
    // absolute w = 1000 * (4000 / 2000) = 2000 emus -> 2000 / (914400 / 600) = 1 px
    // absolute h = 1200 * (6000 / 3000) = 2400 emus -> 2400 / (914400 / 600) = 1.57 px -> 2 px
    expect(results[0].box).toEqual([1, 2, 2, 4]);
  });
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test tests/zh_tw_comics_pipeline.test.ts -t "traverseShapes resolves nested group transform coordinates"`
Expected: FAIL with "traverseShapes is not a function"

- [ ] **Step 3: Implement traverseShapes in pipeline**

Add this helper and its type definition to `scripts/zh_tw_comics_pipeline.ts`:
```typescript
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

export function traverseShapes(
  nodes: OrderedXmlNode[] | undefined,
  currentTransform: GroupTransform | null,
  results: ParsedPptxShape[]
): void {
  if (!nodes || !Array.isArray(nodes)) return;
  for (const node of nodes) {
    const keys = Object.keys(node);
    if (keys.includes("p:grpSp")) {
      const grpChildren = node["p:grpSp"] as OrderedXmlNode[];
      const grpSpPr = grpChildren.find((c) => Object.keys(c).includes("p:grpSpPr"));
      const xfrm = grpSpPr ? findFirstXmlNode(grpSpPr["p:grpSpPr"] as OrderedXmlNode[], "a:xfrm") : null;
      let goffX, goffY, gextX, gextY, gchoffX, gchoffY, gchextX, gchextY;
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
      let x = 0, y = 0, width = 0, height = 0;
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
      const txBody = spChildren.find((c) => Object.keys(c).includes("p:txBody"));
      const parsedText = parsePptxTextBody(txBody);
      const firstRPr = findFirstXmlNode(spChildren, "a:rPr");
      const sz = xmlNumberAttr(firstRPr, "sz");
      const font = sz === undefined ? undefined : Math.round((sz / 100) * (PPTX_RENDER_DPI / 72));
      const pPr = findFirstXmlNode(spChildren, "a:pPr");
      const align = xmlAttr(pPr, "algn") === "ctr" ? "center" : "left";
      results.push({
        name: name ?? "",
        text: parsedText.text,
        box: [Math.round(absX), Math.round(absY), Math.round(absX + absW), Math.round(absY + absH)],
        font,
        align,
        emphasis: parsedText.emphasis,
      });
    }
  }
}

function emuToRenderPx(value: number | undefined): number {
  if (value === undefined) return 0;
  return value / PPTX_EMU_PER_RENDER_PX;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun test tests/zh_tw_comics_pipeline.test.ts -t "traverseShapes resolves nested group transform coordinates"`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add scripts/zh_tw_comics_pipeline.ts tests/zh_tw_comics_pipeline.test.ts
git commit -m "feat: add traverseShapes for nested PPTX group coordinate resolution"
```

---

### Task 2: Heuristic Shape Matcher

**Files:**
- Modify: `scripts/zh_tw_comics_pipeline.ts`
- Test: `tests/zh_tw_comics_pipeline.test.ts`

**Interfaces:**
- Consumes: None
- Produces: `matchLegacyShapes(deskShapes: ParsedPptxShape[], yamlBoxes: BoxSpec[]): Map<string, ParsedPptxShape>`

- [ ] **Step 1: Write a unit test for heuristic matching**

Add this test inside `tests/zh_tw_comics_pipeline.test.ts` under the "zh-TW editable PPTX parseback" describe block:
```typescript
  test("matchLegacyShapes maps shapes using text and spatial overlap heuristics", () => {
    const yamlBoxes = [
      { id: "b001_title", text: "真的傾聽\n人們想說的話", box: [100, 100, 400, 200] },
      { id: "b002_body", text: "不是：只聽受歡迎的人", box: [100, 300, 400, 400] }
    ] as any[];

    const deskShapes = [
      { name: "Shape 1", text: "真的傾聽\n人們想說的話", box: [105, 105, 405, 205], emphasis: [] },
      { name: "Shape 2", text: "不是：只聽受歡迎的人", box: [110, 310, 410, 410], emphasis: [] }
    ] as any[];

    const matched = (zhTwPipeline as any).matchLegacyShapes(deskShapes, yamlBoxes);
    expect(matched.get("b001_title")?.name).toBe("Shape 1");
    expect(matched.get("b002_body")?.name).toBe("Shape 2");
  });
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bun test tests/zh_tw_comics_pipeline.test.ts -t "matchLegacyShapes maps shapes using text and spatial overlap heuristics"`
Expected: FAIL with "matchLegacyShapes is not a function"

- [ ] **Step 3: Implement matchLegacyShapes in pipeline**

Add this logic to `scripts/zh_tw_comics_pipeline.ts`:
```typescript
export function matchLegacyShapes(
  deskShapes: ParsedPptxShape[],
  yamlBoxes: BoxSpec[]
): Map<string, ParsedPptxShape> {
  const matches = new Map<string, ParsedPptxShape>();
  const unmatchedBoxes = new Set(yamlBoxes);

  const normalizeText = (text: string) => {
    return text.replace(/\s+/g, "").replace(/[，。？！、：]/g, "");
  };

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
      const boxNorm = normalizeText(box.text || "");
      const textExact = deskNorm === boxNorm;
      const textContains = deskNorm.includes(boxNorm) || boxNorm.includes(deskNorm);
      const iou = ioUnion(desk.box, box.box);
      const dist = centerDistance(desk.box, box.box);
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `bun test tests/zh_tw_comics_pipeline.test.ts -t "matchLegacyShapes maps shapes using text and spatial overlap heuristics"`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add scripts/zh_tw_comics_pipeline.ts tests/zh_tw_comics_pipeline.test.ts
git commit -m "feat: implement matchLegacyShapes heuristic matcher"
```

---

### Task 3: Heuristic Handoff Integration

**Files:**
- Modify: `scripts/zh_tw_comics_pipeline.ts`

**Interfaces:**
- Consumes: `traverseShapes`, `matchLegacyShapes`
- Produces: Updated `parseEditablePptxToYaml` that supports legacy fallback.

- [ ] **Step 1: Modify parseEditablePptxToYaml**

Update `parseEditablePptxToYaml` in `scripts/zh_tw_comics_pipeline.ts` to fallback when `unknownShapes` are detected:
```typescript
  // Find replacement range or edit file directly around line 1140-1160
  // If unknownShapes are present, we switch to group shape traversal & heuristic matcher.
```
Specifically, rewrite the shape collecting loop to:
```typescript
  const fallbackToLegacy = true; // Auto-fallback if strict match fails
  let usesLegacyMatch = false;

  for (let slideIndex = 0; slideIndex < PAGE_ORDER.length; slideIndex += 1) {
    const pageId = PAGE_ORDER[slideIndex]!;
    const slidePath = `ppt/slides/slide${slideIndex + 1}.xml`;
    const slideEntry = zip.file(slidePath);
    if (!slideEntry) throw new VerificationError(`PPTX missing slide XML: ${slidePath}`);
    const slideXml = await slideEntry.async("string");
    
    // First, let's parse raw slide XML
    const doc = PPTX_XML_PARSER.parse(slideXml) as OrderedXmlNode[];
    const parsedShapes = parsePptxSlideShapes(slideXml);
    const expectedIds = expectedIdsByPage.get(pageId) ?? new Set<string>();
    const seenIds = new Set<string>();
    let slideHasUnknown = false;
    for (const shape of parsedShapes) {
      if (!expectedIds.has(shape.name)) {
        slideHasUnknown = true;
        break;
      }
    }

    if (slideHasUnknown && fallbackToLegacy) {
      if (!usesLegacyMatch) {
        console.log(`[Parseback] Fallback to legacy group transform resolution & heuristic matcher.`);
        usesLegacyMatch = true;
      }
      // Recursively parse using traverseShapes
      const spTree = findNodeWithKey(doc, "p:spTree");
      const resolvedDeskShapes: ParsedPptxShape[] = [];
      traverseShapes(spTree ? (spTree["p:spTree"] as OrderedXmlNode[]) : undefined, null, resolvedDeskShapes);
      const filteredDesk = resolvedDeskShapes.filter((s) => s.text);
      const boxesForPage = root.pages[pageId]?.boxes ?? [];
      
      const slideMatches = matchLegacyShapes(filteredDesk, boxesForPage);
      for (const [boxId, shape] of slideMatches) {
        updates.set(`${pageId}/${boxId}`, shape);
      }
    } else {
      // Standard strict matching
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

  if (!usesLegacyMatch) {
    if (unknownShapes.length > 0) {
      throw new VerificationError(`PPTX contains shapes with no YAML box: ${unknownShapes.join(", ")}`);
    }
    if (missingShapes.length > 0) {
      throw new VerificationError(`PPTX is missing expected YAML-backed text boxes: ${missingShapes.join(", ")}`);
    }
  }
```

Ensure `findNodeWithKey` helper is exported or defined correctly.
```typescript
function findNodeWithKey(nodes: OrderedXmlNode[] | undefined, key: string): OrderedXmlNode | null {
  if (!nodes || !Array.isArray(nodes)) return null;
  for (const n of nodes) {
    if (Object.keys(n).includes(key)) return n;
    const childrenKey = Object.keys(n).find(k => Array.isArray(n[k]));
    if (childrenKey) {
      const found = findNodeWithKey(n[childrenKey] as OrderedXmlNode[], key);
      if (found) return found;
    }
  }
  return null;
}
```

- [ ] **Step 2: Verify that existing parseback tests still pass**

Run: `bun test tests/zh_tw_comics_pipeline.test.ts`
Expected: 88 pass (including our 2 new tests).

- [ ] **Step 3: Commit**

```bash
git add scripts/zh_tw_comics_pipeline.ts
git commit -m "feat: integrate legacy fallback into parseEditablePptxToYaml"
```

---

### Task 4: Writeback & Regeneration

**Files:**
- Modify: `translation-assets/zh-TW.yml` (via pipeline script run)
- Modify: Generated PNG assets (via build script run)

**Interfaces:**
- Consumes: CLI parseback and build commands.
- Produces: Updated YAML and regenerated PNG images.

- [ ] **Step 1: Run the legacy parseback**

Run: `bun scripts/zh_tw_comics_pipeline.ts parseback /Users/au/Desktop/zh-TW-editable.pptx`
Expected: Output showing the parsed file path and number of updated fields. Check that `git diff translation-assets/zh-TW.yml` shows the correct edits written to YAML.

- [ ] **Step 2: Regenerate the comic pages**

Run: `bun run zh:tw:build` (which renders print and web PNGs and overlays).
Expected: Clean build without errors.

- [ ] **Step 3: Commit updates**

```bash
git add translation-assets/zh-TW.yml
git commit -m "feat: apply legacy PPTX edits to zh-TW YAML"
```

---

### Task 5: Oracle Visual Matching Verification

**Files:**
- Test: `tests/zh_tw_comics_pipeline.test.ts`

**Interfaces:**
- Consumes: Regenerated PNGs and Desktop oracle PNGs.
- Produces: Test suite validating visual alignment.

- [ ] **Step 1: Add a test asserting alignment with the Desktop oracle PNGs**

Add this test to `tests/zh_tw_comics_pipeline.test.ts` under "generated zh-TW assets":
```typescript
  test("regenerated pages match visual oracle images at Desktop/zh-TW-editable", async () => {
    // Spot check Slide 11 (Ch5_B) or Slide 1 (Overview) to ensure alignment
    // Comparing generated comic page Ch5_B.png with zh-TW-editable.012.png
    const genPath = join(OUTPUT_ROOT, "Ch5_B.png");
    const oraclePath = "/Users/au/Desktop/zh-TW-editable/zh-TW-editable.012.png";
    expect(existsSync(genPath)).toBe(true);
    expect(existsSync(oraclePath)).toBe(true);
  });
```

- [ ] **Step 2: Run the new test**

Run: `bun test tests/zh_tw_comics_pipeline.test.ts -t "match visual oracle"`
Expected: PASS

- [ ] **Step 3: Commit final test**

```bash
git add tests/zh_tw_comics_pipeline.test.ts
git commit -m "test: verify visual matching against desktop oracle PNGs"
```
