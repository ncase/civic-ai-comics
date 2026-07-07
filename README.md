Comics by [Nicky Case](https://ncase.me/), for Audrey Tang & Caroline Green's [Civic.AI](https://civic.ai/)

Given freely to the public domain under [Creative Commons Zero](https://creativecommons.org/publicdomain/zero/1.0/deed.en) !

See `/design-resources` for the layout, fonts, and exact colour hex codes.

## How to translate!

For a simple manual translation:

1. Copy the folder `/comics-wordless` and rename your copy with [your language's ISO code](https://en.wikipedia.org/wiki/List_of_ISO_639_language_codes), eg `/comics-zh-tw`.
2. Fill in all the missing text. Don't forget to translate `web_footer.png`, too! Refer to the comics in `/comics-en` for the original, and refer to `/design-resources` to see & install the fonts you'll need.
3. Run the script `./convert_comics_to_web.sh comics-[your lang]`. You may need to install ImageMagick for this to work. (for macOS: `brew install imagemagick`) If you're stuck, ask Claude lol.

You now have the translated versions of the overview & all chapter comics,
in both print & web format!

### For the editable zh-TW pipeline in this repo

1. Install dependencies with `bun install`.
2. Edit `translation-assets/zh-TW.yml`. This is the compact source of truth. The JSON files under `translation-assets/zh-TW/text/` are generated artifacts; do not hand-edit them.
3. Run `bun run zh:tw:build` to regenerate all zh-TW assets:
   - text specs: `translation-assets/zh-TW/text/*.json`
   - print PNGs: `comics-zh-TW/Overview.png`, `comics-zh-TW/Ch1_A.png` ... `comics-zh-TW/Ch6_B.png`
   - web composites: `comics-zh-TW/web/Overview.png`, `comics-zh-TW/web/Ch1.png` ... `comics-zh-TW/web/Ch6.png`
   - editable PowerPoint: `comics-zh-TW/zh-TW-editable.pptx`
   - diff/review artifacts under `translation-assets/zh-TW/review/`
4. Run `bun run zh:tw:verify` after any build. It checks that the expected print, web, cover, footer, and generated text-spec files exist and have the expected dimensions.

Useful individual commands:

- `bun run zh:tw:specs` - rebuild only `translation-assets/zh-TW/text/*.json` from `translation-assets/zh-TW.yml`.
- `bun run zh:tw:render` - rebuild print PNGs and review overlays.
- `bun run zh:tw:web` - rebuild web composites.
- `bun run zh:tw:pptx` - rebuild `comics-zh-TW/zh-TW-editable.pptx`.
- `bun run zh:tw:diff` - rebuild diff masks against the source JPEGs.
- `bun run zh:tw:parseback [pptxPath] [yamlPath]` - read an edited PPTX and write supported edits back into YAML. Defaults are `comics-zh-TW/zh-TW-editable.pptx` and `translation-assets/zh-TW.yml`.

PowerPoint edit/parseback workflow:

1. Run `bun run zh:tw:pptx`.
2. Open `comics-zh-TW/zh-TW-editable.pptx` in PowerPoint.
3. Edit existing text boxes only. Supported round-tripped edits are text, box position, box size, font size, alignment, and bold spans.
4. Do not rename text boxes. The pipeline stores each YAML box id as the PowerPoint object name and uses that name as the stable parseback anchor.
5. Do not reorder slides. Slide order is fixed by the pipeline: `Overview`, then `Ch1_A`, `Ch1_B`, ... `Ch6_A`, `Ch6_B`.
6. Avoid adding or deleting text boxes. Unknown boxes cause parseback to stop instead of silently corrupting YAML; missing expected boxes also stop parseback.
7. Save the PPTX, then run `bun run zh:tw:parseback`.
8. Inspect the YAML diff. Parseback intentionally edits only the touched box fields: `text`, `box`, `font`, `align`, and `emphasis`. It does not re-dump or reformat the whole YAML file.
9. Run `bun run zh:tw:build`, then `bun run zh:tw:verify`.

Notes on PPTX fidelity:

- The editable PPTX uses screen-resolution slide backgrounds so the file stays practical to open and email. The print-quality PNGs remain separate under `comics-zh-TW/`.
- Use the installed PowerPoint font family names exactly: `jf蘭陽黑體 1.0` for titles, `jf蘭陽明體2.0` for body/footer text. The shorter names `jf蘭陽黑體` and `jf蘭陽明體` show missing-font warnings in PowerPoint.
- Bold parseback is stored as UTF-16 `start`/`end` ranges in `emphasis` so repeated text such as `人` / `沒有人` round-trips without ambiguity.
- The page-level `md` field in `translation-assets/zh-TW.yml` is kept for provenance/readability. PPTX parseback updates the renderable box fields, not the page markdown.
- Parseback is tested against generated PPTX XML and written to tolerate common PowerPoint rewrites such as attribute reordering, multiple paragraphs, `<a:br/>`, and `b="true"`/`b="1"` bold flags. If a specific PowerPoint save changes the XML differently, inspect the YAML diff and add a fixture before trusting the parseback.
