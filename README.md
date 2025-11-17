# PDF Structure Analyzer

A single-page web application that uses [Mozilla pdf.js](https://mozilla.github.io/pdf.js/) inside the browser to
inspect legal PDFs before they are ingested into an LLM knowledge base. The page never uploads your files to a server:
all parsing, heuristics, and rendering run on the client.

## Features

- **Document overview** – show the number of pages and highlight whether a table of contents or index section is detected.
- **Page grouping** – classify every page as mainly text, mainly table of contents, or mainly index based on the text layout, then summarize contiguous ranges ("Pages 4-8: Mainly TOC").
- **Structural cues** – for text-heavy pages the tool estimates repeating headers, footers, column counts, and invisible text items.
- **Manual overrides** – mark detected headers or footers as normal text when the heuristics latch on to body content instead.
- **Column raster preview** – enter a column number to overlay a translucent highlight on that section of the page canvas.
- **Live preview** – render any page with previous/next navigation so you can visually verify the heuristics.
- **Action shortcuts** – Analyze, Preview, and Extract buttons keep the workflow focused and export a Markdown snapshot of the findings.

## Usage

1. Clone or download this repository.
2. Open `index.html` in any modern browser (Chrome, Edge, Firefox, or Safari).
3. Upload a PDF using the file picker, then click **Analyze** to process it.
4. Use **Preview** to jump to the canvas viewer and page details, or **Extract** to download a Markdown summary of the analysis.
5. Browse the detected page groups, type a page number into the textbox selector, and inspect its metadata alongside the preview (including header/footer overrides and the column raster overlay).

Because the project is fully static there is no build step and `npm install` is not required. Everything—markup, styling,
and JavaScript—lives directly inside `index.html`, so you can email or host a single self-contained file. If you prefer to
serve the page locally over HTTP, any static server (for example `python -m http.server`) will work.

## Implementation notes

- The UI, styling, and browser logic all live in `index.html`, which imports pdf.js from a CDN at runtime.
- Table-of-contents pages are detected by searching for explicit headings (`"Table of Contents"`) and by spotting lines
  that end with page numbers or dotted leaders. Index pages rely on similar patterns (alphabetized entries followed by
  numbers).
- Header/footer detection is based on text lines that appear near the top/bottom of the page, while the column counter
  clusters x positions to determine if the body text is multi-column.
- Invisible text is approximated by looking for text runs whose width is nearly zero, which typically happens with OCR
  artifacts or hidden annotations.

Feel free to adapt the heuristics to the documents you work with—for example by adding rules for appendices or exhibits.
