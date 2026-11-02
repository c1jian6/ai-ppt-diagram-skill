# Spec reference

A spec is one JSON object (or a list of them / `{"slides": [...]}`) with one
diagram per slide. Only `title` is free-form; everything else maps onto
Graphviz attributes or the grid engines.

## Top level

| Key | Default | Meaning |
|---|---|---|
| `title` | — | slide title |
| `subtitle` | — | second line under the rule |
| `caption` | — | small note under the diagram |
| `kicker` | — | small label above the title |
| `theme` | `ocean` | `ocean` `forest` `sunset` `ink` `tech` |
| `tones` | — | per-figure palette (see below) |
| `notes` | — | speaker notes; when absent the DOT source is written there |
| `layout` | graph | `layers` or `panels` switch to the grid engines |
| `direction` | `TB` | `TB` `LR` `BT` `RL` (graph layout only) |
| `splines` | auto | `spline` (image), `ortho` (native), `polyline`, `curved` |
| `nodesep` / `ranksep` | 0.45 / 0.55 | spacing in inches |
| `node_fontsize` / `edge_fontsize` | 12.5 / 10.5 | points; scaled with the drawing |

## Node

| Key | Meaning |
|---|---|
| `label` | main text, `\n` for line breaks |
| `sub` | smaller second line (detail, counts, latency…) |
| `badge` | short bold prefix, e.g. `Phase 1` |
| `bullets` | list of lines rendered left-aligned as a card |
| `shape` | `box` `round` `stadium` `diamond` `oval` `cylinder` `hex` `io` `doc` `cloud` `proc` `plain` |
| `tone` | `primary` `accent` `success` `warn` `danger` `neutral` `ghost`, a custom name, or `#RRGGBB` |
| `width` / `height` | minimum size in inches |
| `group` | cluster label the node belongs to |
| `shadow` | `true` adds a soft shadow (use on one hero node only) |

## Edge

| Key | Meaning |
|---|---|
| `from` / `to` | node ids |
| `label` | edge text (keep it ≤ 8 CJK chars) |
| `style` | `dashed` `dotted` `bold` `invis` |
| `tone` | `normal` `muted` `strong` `success` `danger`, custom, or hex |
| `arrow` | `normal` `none` `open` `dot` `diamond` `tee` |
| `constraint` | `false` keeps the edge from changing the ranking (use for loops) |
| `minlen` / `weight` | rank distance / layout weight |
| `hidden` | invisible, for ordering only |

## Cluster / ranks

```json
"clusters": [{"label": "Services", "nodes": ["api", "auth"], "style": "rounded,dashed"}],
"ranks": [["a", "b"], ["c", "d"]]
```

## Custom palette

```json
"tones": {
  "data":   {"fill": "#E8F2FC", "stroke": "#2E75B6", "font": "#1B3A57"},
  "core":   {"fill": "#5B2D8E", "stroke": "#3E1D66", "font": "#FFFFFF"}
}
```

Given only `fill`, the stroke is derived (22 % darker) and the text colour flips
to white for dark fills.

## `layout: "layers"`

Equal-width columns, frame hugging its content, block arrows between columns.

```json
{"layout": "layers", "layers": [
  {"label": "Input", "items": ["Images", "Text"]},
  {"label": "Model", "items": [{"text": "Encoder", "tone": "primary"}]},
  {"label": "Output", "items": [{"text": "Scores", "sub": "top-5"}]}
]}
```

`items` accept a string or `{text, sub, tone}`.

## `layout: "panels"`

Tall panels with a coloured header band; content stacks inside.

```json
{"layout": "panels", "panels": [
  {"title": "Ingest", "tone": "in", "flow": true, "items": ["logs", "events"]},
  {"title": "Serve", "tone": "out", "items": [{"text": "Dashboard", "chip": true}]}
], "links": [{"from": 0, "to": 1}]}
```

- `flow: true` draws a down arrow between items.
- `chip: true` renders a pill; `span: true` renders a wide hero card (`height` in inches).
- `links`: adjacent panels (`|to-from| == 1`) get a block arrow in the gap;
  distant ones (`"route": "channel"`) run under the panels so nothing crosses a frame.
