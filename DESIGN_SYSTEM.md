# GoIndia Add-in — Design System

Status: **Download Data migrated; DataGPT and Query Data audited, no further changes planned.**
This document defines the shared visual language for `src/taskpane/taskpane.html` / `taskpane.js`
so every feature (current and future) looks and behaves like one product instead of three
bolted-together ones. It was written after auditing the three existing full-panel views —
**DataGPT** (`data-view="datagpt"`), **Download Data** (`data-view="download"`, the original
dropdown-based company/model downloader), and **Query Data** (`data-view="downloadtable"`, the
SQL-generation feature) — against each other.

Download Data now wraps its controls in `#downloadFeatureBox` (`.fb-box`/`.fb-header`/`.fb-body`/
`.fb-footer`), uses `.fb-field`/`.fb-label` for its two main fields, and its primary/secondary
buttons are `.fb-submit`/`.fb-btn-secondary` instead of bespoke `#refreshBtn`/`#helpBtn` CSS. In the
"download" ribbon view specifically it also goes edge-to-edge (no card border/shadow/margin, full
width) the same way Query Data does — but deliberately WITHOUT Query Data's flex:1/height:100%/
`.fb-body { overflow-y: auto }` treatment, since that would clip the company-search dropdown's
suggestion list (a descendant of `.fb-body`) instead of letting it float over the page. This same
markup also renders inside the unparameterized "main" view (see `RIBBON_VIEW_ALLOWLIST` in
`taskpane.js`), where the box keeps its normal card look (border/shadow/44px top clearance) rather
than going edge-to-edge — that treatment is scoped to `data-view="download"` only. Treat this as the target to migrate
toward, feature by feature, not a description of what's there today.

## 1. Design tokens

These are already the *de facto* palette — all three views independently converged on them. The
goal from here is to stop any NEW value from drifting away from this list, not to introduce new
colors.

### Color

| Token | Value | Use |
|---|---|---|
| `--brand-navy` | `#173760` | Primary brand color — headers, primary buttons, bold labels, titles |
| `--brand-red` | `#d80404` | Brand accent line only (header bottom-border, chat FAB gradient) — never a primary action color |
| `--accent-blue` | `#2f6fd0` | Interactive/focus state — input focus borders, links, selected-pill borders, gradients paired with navy |
| `--danger` | `#b80000` | Destructive/cancel actions, error text |
| `--danger-strong` | `#c0392b` | Solid-fill danger buttons (e.g. a confirmed "Delete") — one step more alarming than `--danger`'s outline use |
| `--danger-bg` | `#fff0f0` | Error/danger banner backgrounds |
| `--danger-border` | `#f3caca` | Error/danger banner borders |
| `--ok-bg` / `--ok-text` / `--ok-border` | `#eafaf0` / `#1a7a3d` / `#bfe8cf` | Success status banners |
| `--surface` | `#fff` | Card/panel background |
| `--surface-subtle` | `#f7f9fc` | Subtle container background (collapse headers, inset rows) |
| `--border` | `#e6eaf0` | Card/panel borders |
| `--border-input` | `#d6e5f5` | Input/textarea/button borders (the lighter blue-grey used everywhere a form control needs a border) |
| `--text-primary` | `#2c3744` | Body text in inputs/content |
| `--text-muted` | `#6b7686` / `#55606e` | Secondary text, descriptions, metadata (two near-identical greys exist today — standardize on `#6b7686` for new work) |

### Typography

- **Font family:** `Arial, sans-serif` everywhere. No exceptions for UI text.
- **Monospace** (`Consolas, Monaco, monospace`) only for literal data: cell addresses, SQL, grid
  previews — never for prose.
- **Scale:** headers 13–14px/700, field labels 11.5px/700, body/buttons 12.5–13px/400–700,
  metadata/captions 10–11.5px/400–600. Don't introduce a new size between two of these without a
  reason — the existing scale already covers "label vs. body vs. caption."

### Shape & spacing

- **Border radius:** 8px for inputs/buttons/fields, 10–12px for cards/panels, 999px (pill) for
  toggle-style chips and radio buttons. Nothing else.
- **Border weight:** 1.5px for form controls and card borders, 1px for subtle dividers/row borders,
  2px reserved for the single brand-accent underline (`.fb-header`'s `border-bottom`).
- **Shadow:** cards get a barely-there `0 1px 3px rgba(23,55,96,0.1)` lift, nothing heavier. Reserve
  stronger shadows (`0 12px 36px rgba(0,0,0,0.22)`) for things that float ABOVE the page — a modal,
  a dropdown menu, DataGPT's floating panel in its non-full-bleed form.

## 2. Core components

Every component below already exists in the codebase under these exact class names. New features
should reach for these first instead of writing new CSS.

### Card shell — `.fb-box` / `.fb-header` / `.fb-body` / `.fb-footer`

The base structural unit for any form-like feature. A ribbon view that wants the "fills the whole
taskpane" treatment (like Query Data and DataGPT) adds a `body[data-view="X"] .fb-box { flex: 1; ...
border: none; border-radius: 0; box-shadow: none; }` override — the box becomes the view, not a card
floating inside one. Download Data instead stays a card sitting inline in the page (it isn't its
own full-bleed view the way Query Data/DataGPT are).

`.fb-box`'s `overflow: hidden` clips its header's square corners to the box's rounded ones — fine
for a view with no floating content, but it will also clip anything that needs to visually extend
past the card (a dropdown's suggestion list, a tooltip). For that case, add `.fb-box-overflow-visible`
to the box and `.fb-header-self-rounded` to its header instead (see Download Data's
`#downloadFeatureBox`) — the header rounds its own top corners directly rather than relying on the
parent's clip.

```html
<div class="fb-box">
    <div class="fb-header">Feature Title</div>
    <div class="fb-body">
        <!-- .fb-field blocks go here -->
    </div>
    <div class="fb-footer">
        <button type="button" class="fb-submit">Primary Action</button>
    </div>
</div>
```

### Field + label — `.fb-field` / `.fb-label`

Every input gets a bold navy label above it, no exceptions. This is the single biggest
inconsistency found in the audit (see §3) — Download Data's "Data sources:" label is 10px/grey/400,
everything else is 11.5px/navy/700.

```html
<div class="fb-field">
    <label class="fb-label">Select Company</label>
    <input class="fb-input" ...>
</div>
```

### Inputs & textareas — `.fb-input` / `.fb-textarea`

1.5px `#d6e5f5` border, 8px radius, `#2f6fd0` on focus. Any bespoke input (the dropdown search box,
the cell-reference field) should still carry `.fb-input` alongside its own behavioral class, the way
`#dtCustomCell` already does (`class="fb-input dt-yt-item-loc-input"` pattern) — the specific class
adds behavior/sizing, `.fb-input` guarantees it still looks like every other input.

### Buttons

Five button types cover every case in the add-in. Don't invent a sixth without updating this
document.

| Class | Look | Use |
|---|---|---|
| `.fb-submit` | Solid navy, white text, full-width block | THE primary action in a form/footer (Download Data, Submit Feedback, Download, Write to sheet, Update) |
| `.fb-btn-secondary` | White fill, 1.5px `#d6e5f5` border, navy text | A secondary (non-destructive, non-cancel) action next to a `.fb-submit` — e.g. "Help" next to "Download Data" |
| `.dt-link` | White fill, 1.5px `#b80000` border/text | Cancel / dismiss — every "Cancel" in the add-in should be this, not a plain text link |
| `.dt-btn-danger-sm` | Solid `#c0392b`, white text, compact padding | A confirmed destructive action inside an alert banner (e.g. "Delete" inside the delete-confirm) |
| `.dt-icon-btn` (+ `.dt-btn-danger` modifier) | 24×24 transparent icon button | A row of compact per-item actions (Refresh/Modify/Delete on a `.dt-yt-item`) where text buttons would compete for width |

**Judgment call, applied as recommended below rather than as originally written:** the external
review suggested forcing *every* primary action app-wide — including DataGPT's composer send
button and the floating "Ask DataGPT" FAB — onto the `.fb-submit` shape. That's not what got
applied: a full-width navy block button doesn't fit inside a chat composer text field, and the FAB
is deliberately an attention-grabbing floating affordance, not a form submit. The actual
inconsistency worth fixing was **form-shaped views disagreeing with each other** (Download Data vs.
Query Data vs. Feedback), not chat needing to stop looking like chat — so `.fb-submit`/
`.fb-btn-secondary` were applied to Download Data's "Download Data"/"Help" buttons; DataGPT's
composer/FAB were left as their own purpose-built pattern, since nothing else in the add-in is a
floating chat panel for it to be inconsistent with.

### Pill radio — `.dt-radio`

Stylized native radio: hidden input, label is the clickable pill, `:has(input:checked)` (with a
`.selected` JS fallback) swaps border/background to the accent blue. Used today for the Query Data
destination picker; reusable anywhere a small set of mutually-exclusive options needs more visual
weight than a plain `<select>`.

### Toggle card — `.model-card` / `.mc-*`

A full-width card whose entire surface is the checkbox label, with a track-and-thumb switch
(`.mc-switch`/`.mc-track`) on the right. Currently only used for "AI Financial Model" in Download
Data — worth reusing anywhere else an optional, costed, hero feature needs to be toggled on.

### List/row card — `.dt-yt-item` pattern

The pattern established for "Your Tables": a bordered, lightly-shadowed card
(`border: 1px solid #e3eaf3; border-radius: 8px; box-shadow: 0 1px 3px rgba(23,55,96,0.1)`) with a
header row (title button + metadata badge + a compact row of `.dt-icon-btn` actions), a description
line, and conditionally-revealed inline forms/banners below. This is the **reference pattern for any
future list of items** — a history list, a saved-query list, anything with per-row actions — rather
than each new feature reinventing row chrome.

### Status banners — `.fb-status` (`.ok` / `.err`)

Single-line pill for a terse outcome message, colored green (success) or red (error). The dedicated
alert banner (`.dt-yt-delete-confirm`'s pattern: tinted background + border + inline actions) is a
heavier variant for a confirmation that needs its own visual weight — use it for "are you sure"
moments, `.fb-status` for everything else.

### Collapsible section — `.dt-collapse-header` / `.dt-collapse-body`

A full-width button (chevron + label, `#f7f9fc` background, top border) that toggles its sibling
body's `display`. Currently only in Query Data ("Download tables" / "Your Tables") — reusable
wherever a view grows enough sub-sections to need progressive disclosure.

## 3. Current state per view (audit)

- **DataGPT** — its own self-consistent system (gradient header bar, chat bubbles, floating
  composer). Structurally fine as-is; not a form, shouldn't be forced into form components. The one
  genuine inconsistency: `.chat-clear` and `.chat-mode-btn` are bespoke small buttons that don't map
  to any token above — low priority, but worth eventually expressing as a named "compact secondary
  button" token if a fourth view needs the same shape.
- **Download Data** — **migrated.** Now wrapped in `#downloadFeatureBox.fb-box` (`.fb-box-overflow-visible`
  so its dropdown suggestion list can still escape the card), with `.fb-field`/`.fb-label` for "Data
  sources"/"Select Company", `.fb-submit`/`.fb-btn-secondary` for "Download Data"/"Help", and
  `.dropdown-toggle`/`.sector-select` down to the standard 1.5px border. The old
  view-specific field-reorder (search-before-checkboxes in the ribbon view only) was dropped in
  favor of one consistent order everywhere, now that both fields carry their own label.
- **Query Data** — already built on `.fb-box`/`.fb-field`/`.fb-label`/`.fb-submit` and is the
  source of most of the newer component patterns above (`.dt-icon-btn`, `.dt-yt-item`,
  `.dt-collapse-*`, `.dt-radio`). Treat it as the closest thing to a reference implementation today.
  Its two near-duplicate muted greys were consolidated onto `#6b7686`.

## 4. Rules for new features going forward

1. Wrap any new form-shaped ribbon view in `.fb-box` / `.fb-header` / `.fb-body` / `.fb-footer` —
   don't place controls directly in `#addinUI`.
2. Every input/textarea gets a `.fb-label` above it. No bare `<label>` or inline-styled text
   standing in for one.
3. The view's single primary action is `.fb-submit`; Cancel/dismiss is `.dt-link`; a destructive
   confirm is `.dt-btn-danger-sm` inside a `.dt-yt-delete-confirm`-style banner.
4. A list of same-shaped items (rows the user can act on individually) uses the `.dt-yt-item` card
   pattern, not a plain divided list.
5. Don't introduce a new color, radius, or border width outside §1's token table without updating
   this document first — that drift is exactly how we ended up with three different dropdown border
   weights.
6. Chat-style / floating-overlay UI (anything conceptually closer to DataGPT than to a form) is
   exempt from the form-component rules above, but should still pull its palette from §1.

## 5. Migration status

1. ~~**Download Data** first (biggest gap, clearest win)~~ — **done.**
2. ~~Spot-check Query Data for leftover one-off values~~ — **done** (the two muted greys).
3. DataGPT left structurally alone, as planned; only touch the handful of bespoke small buttons
   noted in §3 if/when a second view needs that same "compact secondary button" shape.

None of this has been exercised in actual Excel yet (no live Office host in this environment) —
when you test, check both the "Download Data" ribbon view AND the plain unparameterized taskpane
("main", opened via the add-in's default icon rather than a ribbon button) since they share this
same markup; in particular confirm the company-search dropdown list still renders fully instead of
being clipped at the card's edge.
