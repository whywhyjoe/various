# todo — personal scratch repo

Grab-bag of unrelated projects; treat each top-level folder as its own world.

- `devpad/` — **DCSPad**, a SharePoint-native JSFiddle-style developer workbench. The active project. Read `devpad/CLAUDE.md` before touching anything in it.
- `jsfiddle/` — unrelated to devpad despite the name: stdlib-only Python CLI scripts that fetch/push fiddles on jsfiddle.net, plus reverse-engineering notes. No packaging, no tests.
- `bilingual/` — dependency-free EN/FR string-swap system (keyed dictionary + `data-intl` markup + `intl.t()`) for SharePoint pages, with a demo page. See its README.
- `dcs-file-picker/` — **DCS File Broker**, the standard open/save component for DCS apps: one dialog over the local file system and SharePoint document libraries, with configurable metadata columns and file-type categories. Dependency-free ES modules, demo page, headless tests. Read `dcs-file-picker/CLAUDE.md` before touching it. Generalised from `sp-dcspad`, with no dependency on it.
- `halo-banner/` — browser tool that generates "halo banner" graphics for SharePoint pages, emitting either a pasteable HTML/CSS snippet or a standalone SVG file. No build step. Read `halo-banner/AGENTS.md` before touching it.
- `classic-referrer-redirects/` — referrer-based second-hop redirect script for one classic→modern SharePoint site migration: JSON mapping of old-site paths to new-site pages, plus a destination-page notice snippet. Dependency-free, ES5-safe, `node test.js`. See its README.

Don't create cross-folder dependencies.
