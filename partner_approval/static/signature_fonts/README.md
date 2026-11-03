# Signature fonts (Type name tab)

Every `*.ttf` / `*.otf` file in this folder is offered on the signing page
"Type name" tab; the partner picks one and the server renders the signature
PNG (and the signed PDF) with that font.

- File names: letters, digits, `-` and `_` only (e.g. `Bayu-Script-1.ttf`).
  Other names are ignored. The label shown is made from the name
  (`BayuScript-Regular.ttf` -> "Bayu Script").
- Order: alphabetical by file name; prefix `1-`, `2-`, `3-` to set the order
  (the prefix is not shown). The first font is the default.
- Maximum 8 fonts. No restart needed: the folder is read on each page view.
- Letters a font does not have are never drawn as empty boxes: the page greys
  out a style that cannot show the typed name (and pre-selects the first one
  that can), and the server refuses it with the same rule. Styles 4-6 cover
  full Vietnamese, so every Latin/Vietnamese name has at least 3 styles; only
  other scripts (e.g. Chinese) end with "please draw".

Current fonts:

| File | Label | Licence | Coverage |
|---|---|---|---|
| `1-DarceyOliver-Regular.ttf` (default) | Darcey Oliver | (c) 2018 Bang Toyib, no licence text in the file (chosen by Bayu) | A-Z a-z 0-9, punctuation, a few accents (á â ä é ë ó ô ö ú ü ă ...); no full Vietnamese |
| `2-Budaphest-Regular.ttf` | Budaphest | (c) 2018 Pollux of Geminorum, no licence text (chosen by Bayu) | A-Z a-z, space, `.` `,` only |
| `3-Augustia.ttf` | Augustia | (c) 2017 Maulana Creative, licence URL creativemarket.com/maulanacreative (chosen by Bayu) | A-Z a-z 0-9, punctuation; no accents |
| `4-GreatVibes-Regular.ttf` | Great Vibes | SIL OFL 1.1, `GreatVibes-OFL.txt` (google/fonts `ofl/greatvibes`) | Latin + full Vietnamese |
| `5-DancingScript-Regular.ttf` | Dancing Script | SIL OFL 1.1, Reserved Font Name "Dancing Script", `DancingScript-OFL.txt` (google/fonts `ofl/dancingscript`, variable font, default weight 400) | Latin + full Vietnamese |
| `6-Allura-Regular.ttf` | Allura | SIL OFL 1.1, `Allura-OFL.txt` (google/fonts `ofl/allura`) | Latin + full Vietnamese |

Check the licence of fonts 1-3 where they were obtained before production use.
