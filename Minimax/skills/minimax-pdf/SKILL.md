---
name: minimax-pdf
description: HTML-first PDF production skill for reports, papers, and structured documents. Must be applied before generating PDF deliverables from HTML.
---

## A. Scope and Operating Contract

This skill governs **authoring and converting HTML into print-quality PDF**.

Primary output goals:
- Stable pagination and predictable layout on Linux runtime.
- Searchable/selectable text (no screenshot-based fallback).
- Professional, citation-safe long-form documents.

## B. Hard Constraints (Do Not Violate)

### B1. Conversion entrypoint

For HTML->PDF, use `html_to_pdf` only.

Forbidden:
- Screenshot/print hacks or manual browser printing
- Direct invocation of low-level local scripts for conversion

Reason: image-stitch paths degrade text quality and create pagination discontinuity.

### B2. Rendering safety rules

- Do not inject Paged.js manually. The runtime pipeline handles loading.
- Do not rely on CSS counters (`counter-reset`, `counter-increment`, `counter()`).
- Do not use runtime charting engines (ECharts, Chart.js, D3, Plotly, etc.).
- Charts should be pre-rendered as static images and prefer landscape aspect ratio.
- Decorative emoji/icon glyphs are disallowed unless explicitly requested.

## C Intent Parsing

Classify the task before execution:

| Intent | Typical user request | Pipeline |
|---|---|---|
| Build | "写一份报告并导出 PDF" | `build-pdf` |
| Transform | "把这篇内容翻译后做成 PDF" | `transform-pdf` |
| Existing PDF ops | "提取/合并/拆分 PDF" | `process-existing-pdf` |
| LaTeX explicit | "请用 LaTeX/.tex/Tectonic" | `latex-compile` |

Clarification policy:
- If request is clear, execute directly.
- If ambiguous, ask once with a compact checklist:
  - 文档类型/主题
  - 是否要封面
  - 字数或页数边界
  - 语言与格式偏好

Important clarification behavior:
- Ask at most one clarification round for intent.
- After that, execute with explicit assumptions rather than repeatedly asking.

## D. Content Governance

### D1. Language policy
- Chinese user query -> Chinese content
- English user query -> English content
- User-specified language -> obey exactly

### D2. Outline policy
- User provides outline -> preserve hierarchy/order, no silent restructuring
- No outline -> choose structure by document type and keep narrative flow consistent

### D3. Citation integrity
- Never invent references.
- Every citation must be verifiable (author/title/year/source).
- Reused source should keep the same citation index.

Recommended citation style for this skill: **IEEE numeric**.

Sample reference list:
```text
[1] R. Patel and L. Chen, "A comparative study on model routing," Journal of Applied AI, vol. 8, no. 3, pp. 44-58, 2025.
[2] M. Rivera, Systems Design Handbook, 2nd ed. New York, NY, USA: Northbridge Press, 2024.
[3] T. Huang, "Model evaluation checklist," Research Notes, https://example.org/eval (accessed Feb. 14, 2026).
```

## E. Conversion Fidelity Checklist

When transforming existing material (translation/rewrite/reformat), preserve source fidelity:

### E1. Links
- Keep original destination URL in `href`.
- Do not replace links with plain text.
- Ensure conversion uses `preserve_links=true`.

### E2. Images
Use a three-pass check:
1. Count extracted image assets
2. Count `<img>` tags in HTML
3. Validate post-conversion image statistics

### E3. Structure
- Preserve source section sequence and anchor semantics.
- Keep figure/table placement and numbering intent.
- Do not add synthetic cover if source had none.

## F. Implementation Blueprint

### F1. Forbidden patterns

| Pattern | Why unstable | Replacement |
|---|---|---|
| CSS content counters for numbering | pagination DOM shifts can break numbering | explicit labels in markup |
| Dynamic JS chart libraries at render-time | print pagination conflicts | pre-rendered static charts |
| Emoji/icon-heavy typography | Linux fallback inconsistency | plain text labels |

Chart image policy:
- Prefer landscape charts (`width > height`) to reduce page-break artifacts.

### F2. Overflow guards (required baseline)

```css
/* Keep printable blocks inside page width */
pre, table, figure, img, svg, .diagram, blockquote, .eq-block {
  max-inline-size: 100%;
  box-sizing: border-box;
}

pre {
  overflow-x: auto;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

figure img, figure svg {
  max-inline-size: 82%;
  max-block-size: 42vh;
  height: auto;
}

table { overflow-x: auto; }
.katex-display { overflow-x: auto; }
code { overflow-wrap: anywhere; }
a { overflow-wrap: anywhere; }
tr { break-inside: avoid; }

body {
  text-align: justify;
  text-align-last: start;
}
```

### F3. Page model setup

```css
@page {
  size: A4;
  margin: 2.4cm 1.9cm;
  @top-center { content: string(doc_title); }
  @bottom-center { content: counter(page); }
}

@page :first {
  @top-center { content: none; }
  @bottom-center { content: none; }
}

@page titlepage {
  @top-center { content: none; }
  @bottom-center { content: none; }
}

@page contents {
  @top-center { content: none; }
  @bottom-center { content: none; }
}

body { string-set: doc_title ""; }
h1 { string-set: doc_title content(); }
.cover-page { page: titlepage; }
.toc-sheet { page: contents; }
```

Pagination notes:
- Apply `break-inside: avoid` only to compact units (single figure, single row, callout box).
- Never apply it to large wrappers (chapter/section container).
- Use `thead { display: table-header-group; }` for multi-page table headers.

### F4. Visual direction

Default target is **print-academic**, not dashboard aesthetics.

Avoid:
- heavy card shells
- KPI tile walls
- dark decorative title bars
- oversized rounded/shadowed ornaments

Prefer:
- plain headings + thin dividers
- data-dense tables
- restrained grayscale palette
- simple, high-contrast typography

Type scale suggestion:
- Body: 11pt
- Subheading: 14pt
- Primary heading: 18-20pt
- Line height: 1.6-1.7

### F. Cover page rules

Full-bleed baseline:
```css
*,
*::before,
*::after { box-sizing: border-box; }

html, body {
  margin: 0;
  padding: 0;
}

@page :first {
  margin: 0;
}

.cover-page {
  inline-size: 210mm;
  block-size: 297mm;
  position: relative;
  display: grid;
  place-items: center;
  overflow: hidden;
  break-after: page;
}
```

Cover variants:
- **Minimal**: white background, centered title/meta, no decoration
- **Designed**: low-saturation geometry/gradient, keep center area clear for title

If using image background, do not use CSS `background-image`. Use absolute `<img>`:

```html
<section class="cover-page">
  <img class="cover-photo" src="cover.jpg" alt="">
  <div class="cover-layer">...</div>
</section>
```

```css
.cover-photo {
  position: absolute;
  inset: 0;
  inline-size: 100%;
  block-size: 100%;
  object-fit: cover;
  object-position: center;
  z-index: 0;
}

.cover-layer {
  position: absolute;
  inset-block-start: 50%;
  inset-inline-start: 50%;
  transform: translate(-50%, -50%);
  z-index: 1;
}
```

### F5. Numbering, references, TOC

Use explicit labels in markup, not CSS counters.

```html
<figure id="arch-overview">
  <img src="system-overview.png" alt="System overview">
  <figcaption data-caption="Figure 1">Architecture Overview</figcaption>
</figure>

<table id="latency-table">
  <caption data-caption="Table 1">Latency by Scenario</caption>
  ...
</table>

<div class="eq-block" data-eq="(1)">$$f(x)=x^2+1$$</div>
```

```css
figcaption::before {
  content: attr(data-caption) " ";
  font-weight: 700;
}

caption::before {
  content: attr(data-caption) " ";
  font-weight: 700;
}

.eq-block::after {
  content: attr(data-eq);
  float: right;
}
```

Anchor placement rule:
- Put `id` on the highest logical container (`figure`, `table`, section wrapper), not on inline caption text.

TOC example with computed page numbers:

```html
<nav class="toc-sheet" aria-label="Contents">
  <ul class="toc-list">
    <li><a href="#sec-intro">1 Intro</a></li>
    <li><a href="#sec-method">2 Method</a></li>
  </ul>
</nav>
```

```css
.toc-list {
  list-style: none;
  margin: 0;
  padding: 0;
}

.toc-list li { margin: 0.45em 0; }

.toc-list a {
  display: flex;
  gap: 0.5em;
  color: inherit;
  text-decoration: none;
}

.toc-list a::after {
  margin-inline-start: auto;
  content: target-counter(attr(href url), page);
}
```

Optional in-text page reference:
```css
a.page-ref::after {
  content: " (p." target-counter(attr(href url), page) ")";
  opacity: 0.72;
  font-size: 0.86em;
}
```

### F6. Formula and diagram policy

- For math, use KaTeX with auto-render.
- Keep formula color neutral and print-safe.
- For Mermaid, keep topology simple; if rendering becomes unstable, replace with static image.

### F7. Reusable layout patterns

Definition block:
```css
.definition {
  border-inline-start: 3px solid #475569;
  padding-inline-start: 1rem;
  margin: 1rem 0;
}

.definition-title { font-weight: 700; }
.definition-body { font-style: italic; }
```

Procedure block:
```css
.procedure {
  border: 1px solid #cbd5e1;
  padding: 0.75rem;
  background: #f8fafc;
}
```

Fixed-size centered badges must use flexbox:
```css
.badge-index {
  inline-size: 1.7em;
  block-size: 1.7em;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  border: 1px solid #cbd5e1;
  border-radius: 999px;
}
```

### F8. Citation anchors and footnotes

Reference anchor consistency:
- Every `<a href="#ref-n">[n]</a>` must map to one `<li id="ref-n">`.
- No dangling reference IDs.

```css
a.citation {
  color: #1f2937;
  text-decoration: none;
  vertical-align: super;
  font-size: 0.78em;
}

.ref-list li {
  padding-inline-start: 2em;
  text-indent: -2em;
}
```

Paged footnote pattern:
```css
.fn {
  float: footnote;
}

.fn::footnote-call {
  content: counter(footnote);
  vertical-align: super;
  font-size: 0.78em;
}

.fn::footnote-marker {
  content: counter(footnote) ". ";
}

@page {
  @footnote {
    margin-top: 1.1em;
    border-top: 1px solid #d1d5db;
    padding-top: 0.55em;
  }
}
```

### F9. Layout tuning guide

When page count or layout does not meet targets, adjust in this priority order:

| Symptom | First move | Second move | Avoid |
|---|---|---|---|
| Page count exceeds target | reduce heading sizes | reduce line-height slightly | aggressive body font shrink |
| Page count below target | increase line-height | increase page margins slightly | adding low-value filler text |
| Table overflows page width | reduce cell padding | allow word wrapping on long tokens | forcing fixed table widths |
| Figure breaks layout | reduce figure max-width/max-height | move figure near paragraph boundary | `avoid` on large parent containers |
| Text looks cramped | raise line-height | increase side margins slightly | oversized heading jumps |
| Resume too sparse/dense | tune margins first | then adjust heading scale | changing section order silently |

## G Fidelity Gates

Run all gates before final delivery.

### G1. Hyperlinks

- Ensure external links keep original `href`.
- In conversion call, enforce `preserve_links=true`.

### G2. Images (3-pass)

1. Source extraction count (baseline)
2. HTML `<img>` count and mapping
3. Post-conversion result check

### G3. Anchor integrity

- Cross-references must point to real IDs.
- Place `id` on top-level containers (`figure`, section wrapper), not inner text nodes.
- For TOC and page refs, use print-aware links with page target resolution (`target-counter`).

### G4. Structure parity (for transforms)

- Keep source section order unless user requests restructuring.
- Do not inject a cover page when source had none (unless user requests one).