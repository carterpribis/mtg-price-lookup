# Magic card price lookup

A Google Sheet that looks up the price of a Magic: The Gathering card as soon as its name is typed in.

I made it for a program where kids were constantly asking staff what their cards were worth, which was taking up a lot of staff time. With this, they looked the values up themselves.

## How it works

Type a card in column A, and optionally a set code in column B and a variant in column C (`foil`, `borderless`, or both). An on-edit trigger sends the lookup to [Scryfall](https://scryfall.com/docs/api), a free public card database, and writes the price in column D and the time in column E.

| Column | Contents |
|---|---|
| A | Card name |
| B | Set code, optional (for example `otj`) |
| C | Variant, optional: blank, `nonfoil`, `foil`, `borderless`, `foil borderless` |
| D | Price in US dollars (written by the script) |
| E | When it was last updated (written by the script) |

- Ordinary cards are looked up by exact name, and by set if one is given.
- Borderless cards are found with a search for borderless printings, and the script picks the printing with the right finish.
- Foil cards get the foil price.
- The script installs its own on-edit trigger the first time the sheet is opened.
- Requests are spaced out to stay well inside Scryfall's rate limit.

## Setting it up

1. Create a sheet with a tab named `Sheet1` and the five column headings above in row 1.
2. Open **Extensions → Apps Script**, paste in `Code.gs`, and save.
3. Reload the sheet and approve the permissions when asked.

No API key is needed.

## Known gaps

- The **MTG Prices → Refresh All Prices** menu item points to a function, `refreshAllPrices`, that is not in the file.
- Prices are written to a 24-hour cache, but the line that reads from the cache is commented out, so every edit makes a fresh request.
