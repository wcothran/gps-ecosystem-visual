# Editing the GPS Ecosystem — a one-page guide for SCDE staff

You can update what shows up on the public **GPS Ecosystem** visual by editing **one Excel file** on SharePoint. No code, no special tools.

## The file you edit

`GPS-Ecosystem.xlsx` on the SCDE SharePoint site. Open it the way you'd open any Excel file — in Excel Online (your browser) or the Excel desktop app.

## Sheet 1: `Items` — one row per card

Each card you see in the visual is one row in this sheet.

| Column        | What to put in it                                                              | Required? |
|---------------|--------------------------------------------------------------------------------|-----------|
| `Section`     | Pick from the dropdown: Audiences, Tools & Apps, Data Services, Foundation, Governance, Outputs, Inputs | **Yes** |
| `Name`        | The full name (also shown as the title when someone clicks the card)           | **Yes**   |
| `Short`       | A shorter label for the card face (leave blank to use the full Name)           | No        |
| `Description` | The paragraph someone sees when they click the card                            | **Yes**   |
| `Future`      | `TRUE` if it's planned / not yet live; otherwise `FALSE`                       | No        |
| `Dashboards`  | A list of dashboard names separated by `;` (e.g. `Student View; Class View`)   | No        |
| `DocURL`      | A link to the full documentation for this item                                 | No        |
| `Order`       | Optional sort number (1, 2, 3…). Leave blank to put the item at the end of its section. | No |

## Sheet 2: `Settings` — page-wide text

The page title, the one-sentence definition under the title, and the subtitle text under each section header.

To change any of these, find the matching `Key` row and edit the `Value` cell.

---

## Common things you might want to do

### Fix a typo

1. Open `GPS-Ecosystem.xlsx` from SharePoint.
2. Find the row with the typo. Edit the cell.
3. Save (`Ctrl + S`, or just close — Excel Online auto-saves).
4. Wait for the next publish (or message the GPS X-Team if it's urgent — they're the ones who push changes live).

### Add a new item (e.g. a new Data Service)

1. Open `GPS-Ecosystem.xlsx`.
2. Go to the `Items` sheet.
3. Find the first empty row at the bottom.
4. Click the `Section` cell — a dropdown will appear. Pick `Data Services`.
5. Fill in `Name` and `Description`. Other columns are optional.
6. Save.

### Remove an item

1. Right-click the row number on the left → **Delete row**.
2. Save.

### Mark something as "planned" / "future"

1. Find the row. Set the `Future` cell to `TRUE`.
2. Save.

The visual will show a gold `*` next to the card and a "Planned / Future Capability" note in the detail panel.

### Rename a section header or change the definition

1. Go to the `Settings` sheet.
2. Find the right row (e.g. `Title`, `Definition`, `DataServicesSubtitle`).
3. Edit the `Value` cell. Save.

---

## Made a mistake? Roll back.

SharePoint keeps every version of this file automatically.

1. In SharePoint, find `GPS-Ecosystem.xlsx` in the file list.
2. Click the **⋯** menu next to the filename → **Version history**.
3. Pick an earlier version (you'll see the editor's name and timestamp on each one).
4. Click **Restore**.
5. Tell the X-Team to re-publish.

---

## How does this become public?

After you save, the X-Team picks up the latest copy of this file and uploads it to the public web host. There's no automatic sync today (that's a future improvement). If you need something out the door quickly, message the X-Team in the GPS channel.

---

## Things the visual won't let you change from the sheet

By design, the layout, colors, and section structure are fixed in the code — they reflect SCDE design system decisions, not editorial ones. If you want to:

- Add a brand-new section (something other than the seven listed above), or
- Change the color scheme, the page layout, or which section a card visually appears in,

…that's a developer change. Reach out to the X-Team to discuss.
