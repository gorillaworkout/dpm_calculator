# Signature fonts (Type name tab)

Every `*.ttf` / `*.otf` file in this folder is offered on the signing page
"Type name" tab; the partner picks one and the server renders the signature
PNG (and the signed PDF) with that font.

- File names: letters, digits, `-` and `_` only (e.g. `Bayu-Script-1.ttf`).
  Other names are ignored. The label shown is made from the name
  (`BayuScript-Regular.ttf` -> "Bayu Script").
- Order: alphabetical by file name; prefix `1-`, `2-`, `3-` to set the order
  (the prefix is not shown). The first font is the default.
- Maximum 6 fonts. No restart needed: the folder is read on each page view.
- Letters a font does not have are never drawn as empty boxes: the page greys
  out a style that cannot show the typed name (and pre-selects the first one
  that can), and the server refuses it with the same rule. If no style can
  show the name (e.g. most Vietnamese names with the current fonts), the
  partner is asked to draw instead.

Current fonts (chosen by Bayu, 2026-11-03):

| File | Label | Copyright (name table) | Coverage |
|---|---|---|---|
| `1-DarceyOliver-Regular.ttf` (default) | Darcey Oliver | (c) 2018 Bang Toyib | A-Z a-z 0-9, punctuation, some accents (á â ä é ë ó ô ö ú ü ă ...); no Vietnamese ê à ễ đ ơ ư ... |
| `2-Budaphest-Regular.ttf` | Budaphest | (c) 2018 Pollux of Geminorum | A-Z a-z, space, `.` `,` only; no digits, `-`, `'` or accents |
| `3-Augustia.ttf` | Augustia | (c) 2017 Maulana Creative | A-Z a-z 0-9, punctuation; no accents |

The fonts carry no licence text (name IDs 13/14 are absent, or only a
Creative Market URL for Augustia). Check the licence where they were obtained
before production use.
