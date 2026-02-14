# Best Practices: Creating Documents from Scratch

> **Scope**: This guide applies when creating a new document without a user-provided template or reference file. When the user supplies a template, outline, or explicit structural instructions, **defer to their input** — these recommendations do not override user intent.

---

## 1. Recommended Document Structure Elements

When building a formal document from scratch (proposals, reports, theses, contracts), the following elements elevate perceived quality:

| Element | When to Include | Implementation Notes |
|---------|----------------|---------------------|
| Cover page | Formal deliverables (proposals, reports, contracts, invitations) | Use designer-quality background images via `engine/facade.py`; text implemented in Word for editability |
| Back cover | Paired with cover page for polished deliverables | Lighter design than cover; contact info or branding |
| Table of Contents | Documents with 3+ major sections | Include TOC refresh hint in gray, smaller font |
| Header | Most multi-page documents | Typically: document title or company name |
| Footer with page numbers | Most multi-page documents | Field codes for automatic numbering |

## 2. Visual Standards

| Property | Recommended Value | Rationale |
|----------|------------------|-----------|
| Page margins | ≥72pt (1 inch) all sides | Sufficient white space for readability |
| Paragraph spacing (body) | After ≥200 twips (10pt) | Visual separation between paragraphs |
| Line spacing (body) | ≥1.5x (`Line="360"`, Auto) | CJK readability |
| Color palette | Low saturation, muted tones | Professional appearance; see `blueprints/Palettes.cs` |
| Font hierarchy | H1 > H2 > body, clear size contrast | Visual hierarchy guides the reader |

## 3. Scene Completeness Reference

Common document types have standard structural elements that readers expect. Use this as a checklist when creating from scratch:

| Document Type | Standard Elements |
|---------------|-------------------|
| Exam paper | Name/class/ID fields, point allocation, grading section |
| Contract | Signature areas for both parties, date, contract number, attachment list |
| Meeting minutes | Attendees, absentees, action items with owners, next meeting time |
| Proposal | Executive summary, scope, timeline, budget, terms |
| Academic paper | Abstract, keywords, references, author affiliations |
| Invoice | Invoice number, line items, totals, payment terms, bank details |

## 4. Default Outlines (When User Provides None)

| Document Category | Recommended Structure |
|-------------------|----------------------|
| Academic | Introduction → Literature Review → Methods → Results → Discussion → Conclusion |
| Business | Executive Summary → Analysis → Recommendations |
| Technical | Overview → Principles → Usage → Examples → FAQ |

## 5. Pagination Control

| Element | Property | Purpose |
|---------|----------|---------|
| Primary heading (H1) | `PageBreakBefore` + `KeepNext` | Chapter separation |
| Secondary heading (H2) | `KeepNext` | Bind heading with following content |
| Pre-table text | `KeepNext` | Keep introductory text with table |
| Body paragraphs | `WidowControl` | Prevent orphan/widow lines |
