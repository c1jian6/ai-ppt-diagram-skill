---
name: ppt-diagram
description: Draw flowcharts, architecture diagrams, decision trees and layered/panel figures as editable PowerPoint files. Use whenever the user asks for a diagram, a PPT figure, a slide illustration, or hands over a sketch/table to be redrawn.
---

# PPT diagram

Turn a short JSON spec into a `.pptx` whose boxes, arrows and labels are **real
PowerPoint shapes** (editable in PowerPoint), not a pasted picture.

All code is bundled in `scripts/`; the only external requirement is Graphviz
for topology diagrams.

## Workflow

1. **Environment** — run `python scripts/bootstrap.py --check`.
   - `status: ready` → continue.
   - Missing `graphviz` → run `python scripts/bootstrap.py --install`; if it
     still fails, stop and give the user the exact one-line install command.
     Never retry silently.
2. **Pick the layout family**
   | The figure is… | Use |
   |---|---|
   | flow / decision tree / network / anything with branches or loops | `direction` + `nodes` + `edges` (Graphviz) |
   | equal-width layers (input → method → fusion → output) | `"layout": "layers"` |
   | 3–5 tall columns holding their own content | `"layout": "panels"` |
3. **Write the spec** — read `references/spec.md`; copy the closest file in
   `examples/` and edit it. Keep labels short: ≤ 8 CJK chars or ≤ 16 latin
   chars per line; move detail into `sub` or `bullets`.
4. **Build** — `python scripts/draw.py spec.json -o out/figure.pptx`
   (defaults are `--mode native`, 16:9). Use `--mode image` only for very dense
   graphs where editability does not matter.
5. **Look at it** — render page 1 and inspect the image (the office-render
   skill, or any LibreOffice/`soffice` call). Check: text inside boxes, arrows
   not crossing labels, nothing outside the slide. Fix and rebuild if needed.
   If no renderer exists, say "visual check skipped" instead of claiming it looks fine.
6. **Answer** — give the `.pptx` path, one sentence on what was drawn, and offer
   the obvious knobs (theme, orientation, wording).

## Rules that prevent ugly output

- More than 5 nodes in one chain → split into two rows with `ranks`.
- `direction: "LR"` for pipelines, `"TB"` for decisions; don't chain 7+ nodes vertically.
- Use `clusters` for layers or subsystems, not for decoration.
- Semantic colours win: define a `tones` palette per figure (e.g. data / rules / core / output) instead of colouring node by node.
- A 16:9 slide fits ~10×5.5 in of drawing; if the builder warns that labels land
  below 9 pt, change the layout instead of ignoring it.
- Never deliver a file you have not opened/rendered at least once.

## Commands

| Goal | Command |
|---|---|
| check environment | `python scripts/bootstrap.py --check` |
| JSON → pptx | `python scripts/draw.py examples/pipeline.json -o out/pipeline.pptx` |
| new deck from spec | same command; each spec may hold a list of slides |
| rasterised fallback | `python scripts/draw.py spec.json --mode image --vector png` |
