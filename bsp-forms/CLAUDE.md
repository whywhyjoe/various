# BSP Forms — agent guide (CLAUDE.md)

JSON-configured forms app for SharePoint pages: one shared engine
(`bsp-forms.js` + `bsp-forms.css`) renders multi-page forms from a per-form
JSON config and writes submissions + attachments to a SharePoint list.
`README.md` is the human overview; `docs/CONFIG-REFERENCE.md` is the config
contract. Read both before changing behavior.

## Where the rules come from

- **UI = the BSP design system** (`whywhyjoe/bsp-design-system`, deployed as
  the sibling `bsp-design/`). This app is *employee-facing*, so BSP applies —
  NOT the DCS Workbench design system. Compose from BSP classes
  (`.field .input .stepper .msgbar .dropzone .tag .avatar .btn .spinner`);
  `bsp-forms.css` is an **additive layer** in the spirit of `editorial.css`:
  scoped to `.bspf`, token-built, never redefines `:root`, never restyles BSP
  classes. The choice-pill hues are feedback-tint pairs or `color-mix()` of
  chart-palette tokens — keep it that way; no free-standing hex.
- **Construction method = the DCS Workbench docs**
  (`whywhyjoe/dcs-workbench-tools/docs/`): classic-script engine, idempotent
  per mount, survives re-evaluation, edit-mode guard, host CSP nonce on every
  injected script, `?v=` cache-busting, paths derived from
  `document.currentScript`. The L1/L2 tier model does NOT apply here.
- **Buildless is non-negotiable.** What's authored is what runs: no modules,
  no bundler, no CDN at runtime. `dev/vendor/alpine.js` is dev-harness-only.

## File map

| Path | What |
| --- | --- |
| `bsp-forms.js` | The whole engine: settings/base resolution → asset loader → pnpjs v2 adapter (+mock seam) → config normalize/validate → rule engine → validators → markup builders → Alpine instance factory → doctor → boot/scan. Sections are banner-commented in that order. |
| `bsp-forms.css` | The `.bspf-*` layer: pills, combo, people picker, attachments, nav, done screen, edit note, doctor. Also owns the `[x-cloak]` rule (kept out of inline `<style>` for CSP). |
| `forms/example-it-request.json` | Reference config — exercises every field type and rule. Keep it exercising anything you add. |
| `webpart/bsp-forms.webpart.html` | The insert snippet users paste/point the web part at. |
| `dev/` | Harness (`index.html`), mock adapter (`mock-sp.js`, same method names as the real adapter), vendored Alpine. Never deployed. |

## Paid-for gotchas (don't relearn these)

- **No `<button>` inside the combo control.** The control is a
  `div[role="combobox"]` *because* selected multi-choice pills carry remove
  `<button>`s — HTML forbids nested buttons and the parser silently re-parents
  them, breaking layout. Keep it a div.
- **Errors clear live, not at blur.** `reval()` re-checks a field on input
  once it has an error. Without it, the error paragraph collapses at
  mousedown-on-Next and the button jumps out from under the click.
- **Alpine init:** markup is inserted first; if Alpine is already on the page,
  `Alpine.initTree(root)` is called (roots added after `Alpine.start()` are
  not picked up automatically). `_x_dataStack` presence = already initialized.
- **File objects stay out of Alpine state.** Reactive proxies break Blob
  method calls; raw `File`s live in the non-reactive `def.store.files`, only
  metadata (`filesMeta`) is reactive.
- **Date-only values are written as noon local** so the stored date can't
  shift a day across timezones. Date rule comparisons are calendar-day based.
- **Duplicate `column` mappings:** later visible field wins (documented
  contract — the payload builder just iterates config order).
- **Attachment failures never re-create the item** (`store.itemId` guard);
  retry uploads only what's still `pending`/`failed`.
- **`pnp.sp.setup` is global** — the adapter re-asserts `baseUrl` before its
  operations; prefer `pnp.Web(url)` when the bundle exposes it.
- The people API can't reliably filter disabled/room accounts; the adapter
  drops non-`User` principals and entries without an email. Best-effort by
  design — don't promise more in UI copy.

## Verifying a change

Serve the folder that contains BOTH this repo and `bsp-design-system`, then
open the harness (mock SP, vendored Alpine, writes logged to
`__BSPF_MOCK_WRITES__`):

```
python -m http.server 8000       # from the parent of both clones
http://localhost:8000/various/bsp-forms/dev/index.html      (?validate → doctor)
```

Walk all three pages of the example form; check the console is clean (the two
`about:invalid` photo errors are the mock exercising the initials fallback)
and inspect `__BSPF_MOCK_WRITES__` for the exact payload. Headless: drive that
page with Playwright against the pre-installed Chromium; prefer programmatic
`el.click()` for nav buttons — pointer-coordinate clicks flake when validation
messages shift layout mid-click.

## Scope guards

- Don't edit the design-system repos from here; extend via `bsp-forms.css`.
- Deployed artifacts are only `bsp-forms.js`, `bsp-forms.css`, `forms/*.json`.
- Next phases (not built, don't scaffold speculatively): branching, drafts,
  post-submit actions, builder UI.
