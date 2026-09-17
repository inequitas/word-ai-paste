# AI Paste

A Word add-in for pasting GPT-chat text into a proposal document so it comes
in matching the document's own styles — Heading 1/2/3, List Paragraph with
bullets, Grid Table 4 — instead of arriving as a pile of `##`/`**`/`|`
characters or mismatched fonts and colors that need to be cleaned up by hand.

Kevin's workflow: write with a company GPT chat tool, copy the reply, open
the AI Paste task pane in Word, paste, check the preview, click
**Insert at cursor**.

## What it does

- Detects whether the clipboard content is **Markdown** (`## Heading`,
  `**bold**`, `- bullet`, `| table |`) or **rendered/formatted HTML** from a
  chat window (or Word's own HTML), with a manual override if it guesses
  wrong.
- Parses headings, paragraphs, bold/italic/code/links, nested bullet and
  numbered lists, tables, blockquotes and code blocks into one internal
  model, regardless of which format it came from.
- Cleans the text up: promotes standalone bold lines to headings, strips
  citation markers, optionally strips manual heading numbers or emoji,
  collapses stray whitespace, and (on by default) adds the blank
  Normal-paragraph spacing Kevin's documents use between sections.
- Inserts the result at the cursor using **only the document's own
  styles** (`styleBuiltIn`, which is locale-independent — this also works
  correctly against a Dutch-language Word where the UI calls Heading 1
  "Kop 1"), with `font.bold`/`font.italic` always set explicitly on every
  run so nothing is inherited by accident, and separate numbered lists that
  each restart at 1.

## Privacy

**Your text never leaves your machine.** The task pane makes exactly two
kinds of network request: loading this page's own JS/CSS from GitHub Pages,
and loading `office.js` from Microsoft's CDN (required by every Word
add-in). There is no analytics, no telemetry, and nothing you paste or
insert is ever sent anywhere else. This matters for pasting client-
confidential proposal text, so it's worth being able to verify yourself:
open the browser dev tools' Network tab while using the pane and confirm
that's all you see.

## Install (Mac)

1. Clone or download this repo.
2. Run `./install.sh` — it copies `manifest.xml` into Word's sideload
   folder (`~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`).
   It does not touch certificates, the keychain, or any system/security
   settings.
3. Quit Word completely (Cmd+Q) and reopen it.
4. Open a document, then either:
   - Home tab → **AI Paste** button, or
   - Insert tab → Add-ins → My Add-ins → Developer Add-ins → AI Paste.

The task pane itself is hosted on GitHub Pages
(`https://inequitas.github.io/word-ai-paste/`) — Word just needs the
manifest sideloaded locally to know the add-in exists and point at that URL.

### Update

Pull the latest `manifest.xml` (it rarely changes) and rerun `./install.sh`
if it did. Everything else (the actual add-in code) updates automatically
the next time you open the task pane, since it's loaded fresh from GitHub
Pages — no reinstall needed for ordinary code changes.

### Uninstall

```
rm ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/word-ai-paste-manifest.xml
```

Then restart Word.

## Using it

1. Paste into the box (Cmd+V) — either Markdown or a formatted chat reply.
2. Check **Detected: Markdown / Formatted text** — click the other option if
   it guessed wrong.
3. Check the preview (heading outline + paragraph/list/table counts).
4. Open **Options** if you want to change the top heading level, table
   style, body style, or any of the cleanup toggles.
5. Click **Insert at cursor**.

If you paste rich HTML and then edit the text box by hand, it switches to
plain-text/Markdown mode for that paste (a small notice explains this) —
otherwise your manual edits wouldn't be reflected in the discarded HTML.

### Self-test

The **Run self-test** link at the bottom (only active inside Word) inserts a
built-in fixture exercising every feature — headings, a promoted bold
heading, mixed inline formatting, a citation marker, nested bullets, two
separate numbered lists, a table, a quote — then reads the result back via
the Word API and shows a pass/fail checklist. Use it after installing, or
after a Word update, to confirm everything still works in your copy of
Word.

## Options reference

| Option | Default | What it does |
| --- | --- | --- |
| Top heading level | Heading 1 | Shifts all headings so the input's highest heading maps to this level, preserving relative depth. Lets you paste a section under an existing heading. |
| Bold-only lines → headings | On | A standalone all-bold paragraph (≤ ~100 chars) becomes a heading one level below the nearest preceding real heading. |
| Remove citation markers | On | Strips `【…】`-style markers and ChatGPT's invisible citation tokens. Leaves normal `[1]` references alone. |
| Remove manual heading numbers | Off | Strips a leading `1.`, `2.3`, `2.3.1`, … from headings — turn this on for templates with auto-numbered headings. |
| Remove emoji | Off | Strips emoji from all text. |
| Blank lines between blocks | On | Kevin's house style: a blank Normal paragraph between most block transitions (body→body, list→body, before any heading, after a table) instead of paragraph-spacing settings. |
| Body style | Normal | Applied to ordinary paragraphs; pick a custom style from the document instead if you want. |
| Table style | Grid Table 4 | Header row + first column + banded rows on, matching Word's own default look for that style. Grid Table 1 Light and Plain Table 1 are also offered, plus any table style already in the document. |

## Architecture

- **Vite + TypeScript**, plain CSS, no UI framework.
- `src/model.ts` — the intermediate block model both parsers produce.
- `src/parse/markdown.ts` — Markdown via `marked`'s lexer (tokens only,
  never its HTML renderer).
- `src/parse/html.ts` — a DOMParser walk that ignores all class/style
  attributes except bold/italic, with extra handling for Word's own
  `Mso*`-class HTML export and for macOS's Cocoa HTML Writer output (what a
  plain RTF chat reply becomes once it's on the system clipboard — see the
  comment at the top of that file).
- `src/parse/detect.ts` — Markdown vs. HTML detection.
- `src/transform/normalize.ts` — the cleanup/options pipeline above.
- `src/word/adapter.ts` — a small interface over the slice of Office.js this
  add-in uses, so `src/word/insert.ts` can be unit-tested with a fake
  in-memory implementation (`tests/fakeWordAdapter.ts`) instead of a real
  Word runtime. `src/word/officeAdapter.ts` is the real implementation.
- `src/word/insert.ts` — turns blocks into Word calls inside one batch (one
  cursor read + one final sync). See the doc comment above `insertBlocks`
  for exactly how separate-list numbering restarts work and why.
- `src/word/selftest.ts` — the self-test described above.
- `src/taskpane/` — the task pane UI.

## Known limitations

- Table cells are inserted as plain text (no per-cell bold/italic/links) —
  acceptable for v1 since GPT-generated tables are almost always plain data.
- A table pasted as the very first block at an empty cursor position leaves
  one blank paragraph before it (the existing empty cursor paragraph isn't
  deleted, to avoid ever risking a paragraph that turns out to hold section
  properties).
- Nested *ordered* lists don't produce hierarchical numbering like "1.1,
  1.2" — each nesting level gets its own independent "1., 2., 3." counter.
  This matches what Kevin's own documents actually use (plain bulleted House
  style; nested ordered lists are rare in this workflow) but is worth
  knowing.
- The Cocoa-HTML-Writer heading heuristic (macOS RTF → HTML) ranks font
  sizes bigger than the most common (body) size as heading levels; a
  one-off oversized bold word inside an otherwise normal paragraph won't be
  misread as a heading, but a document generated with unusual formatting
  could confuse the ranking. Real `<h1>`-`<h6>` tags and Word's own
  `Mso*` classes always take precedence over this heuristic.
- `getStyles()` (used to populate the body/table style dropdowns with the
  document's own custom styles) needs WordApi 1.5; on an older host the
  dropdowns just fall back to the built-in choices.
- No image support — GPT chat replies are text/tables, and images pasted
  alongside are dropped rather than inserted.

## Development

```
npm install
npm run dev       # Vite dev server for UI-only iteration (Office absent)
npm test           # vitest
npm run typecheck   # tsc --noEmit
npm run build       # tsc --noEmit && vite build -> dist/
npm run validate-manifest
```

`npm run dev`/`vite preview` work without Word — the pane shows a banner
explaining it's not running inside Word, but pasting, detection, the
preview and the options panel are all fully testable that way. Only
**Insert at cursor** and the self-test need a real Word host.
