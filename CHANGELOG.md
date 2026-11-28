# Changelog

## 3.0.0 — 2026-11-27

- Replaced custom Markdown and LaTeX parsing with Pandoc for CommonMark/GFM-style Markdown and native Word OMML math.
- Preserved the CLI flags from 2.x.
- Added contents preview plus dynamic Word TOC field, metadata cleanup, and per-chapter figure caption numbering.
- Added stress fixtures, rendered-PDF regression checks, and OOXML structural validation in CI.
- Removed Matplotlib from runtime dependencies; Pandoc binary is supplied by `pypandoc_binary` when not installed system-wide.
