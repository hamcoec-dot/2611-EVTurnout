# Party Disclaimer: new wording + prominent styling

## Context
Party-turnout disclaimer is small grey italic 12px text (`.party-disclaimer`, index.css:948) — easy to miss. Replace wording and restyle as a visible callout.

New text (exact):
> Party turnout data is reflective of the voter's last primary election voting history, if applicable.

## Changes

1. **Text replace** — old string `This data reflects the last primary voting history if applicable` -> new text, all occurrences:
   - `index.html` — 16 occurrences (lines 146, 235, 252, 264, 278, 301, 349, 457, 465, 492, 518, 544, 574, 601, 619, 649). Use Edit with `replace_all`.
   - `app.js:552` — inside a double-quoted JS string; apostrophe in "voter's" is safe there (no escaping needed).

2. **Restyle `.party-disclaimer`** in `index.css:948` (keep class name, so all 17 spots update at once):
   ```css
   .party-disclaimer {
     display: block;
     font-size: 13px !important;
     font-weight: 600;
     font-style: normal !important;
     color: #92400E !important;            /* dark amber text */
     background: #FEF3C7;                  /* light amber fill */
     border: 1px solid #FCD34D;
     border-left: 4px solid #F59E0B;
     border-radius: 6px;
     padding: 6px 10px;
     margin-top: 6px !important;
     margin-bottom: 0 !important;
     line-height: 1.4;
   }
   .party-disclaimer::before {
     content: "\24D8  ";                   /* (i) info glyph, ASCII-escaped */
     font-weight: 700;
   }
   ```
   - Keep CSS/HTML pure ASCII — verify.py fails on non-ASCII chars, hence the `\24D8` escape.
   - index.html:146 has inline `text-align:center; flex:1 1 100%` — fine with block callout; leave as is.

3. **No cache-bust** — do NOT touch `?v=` params on `app.js`/`index.css` (user is testing their URL rewriting).

## Verification
- `python verify.py` — all checks pass (esp. non-ASCII check).
- Grep old string -> 0 hits; new string -> 17 hits.
- Open index.html in browser pane: check callout in chart headers (narrow headers wrap OK), the section intros (301, 457), the party details panel rendered by app.js:552, and mobile width (375px).
