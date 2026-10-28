# Sales Spotlight: setting the fortnightly topic

The Sales Spotlight sits on the Product Portal page, above **What are you looking for?**. It shows one topic at a time: the top 10 customers (or products) for a product group or a set of item codes, this year to date.

The figures come from NetSuite every morning through the sales dashboard. Nobody pulls a report. You only choose the topic.

## Where the topics live

The **Spotlight** list on the **Product Portal** SharePoint site:
https://checkfireltd.sharepoint.com/sites/CheckFireProductPortal/Lists/Spotlight/AllItems.aspx

It was set up on 28 Oct 2026 with these columns:

| Column | Type | What goes in it |
|---|---|---|
| Title | (already there) | The headline, e.g. *Who's buying LFX this year* |
| Kind | Choice: `Top customers`, `Top items` | Rank customers, or rank products |
| Category | Single line of text | A product group exactly as NetSuite names it, e.g. `LFX`, `Fire Blankets`, `Cabinets/Trolleys/Stands`. Optional if you use Items |
| Items | Multiple lines of text (plain) | Optional. Item codes separated by `;`. End a code with `*` to take every code that starts that way, e.g. `SD-CFP-*; FP-COM-*; BX-FP-COM-*` |
| Start | Date only | First day it shows |
| EndDate | Date only | Last day it shows |
| Show | Yes/No (default Yes) | Untick to hide a row without deleting it |
| Tag | Single line of text | Optional. A short label on the card under Previous spotlights, e.g. `NPD range` |

If you fill in both Category and Items, a sale has to match both.

**Why Kind and EndDate:** SharePoint keeps the names "Type" and "End" for itself. If the list is ever rebuilt, the hub also accepts `StartDate`, `ProductGroup`, `ItemCodes`, `Active` and `Label` in place of Start, Category, Items, Show and Tag.

## Dates: month first

The Product Portal site uses US date settings, so dates are typed **month/day/year**:

- 26 October 2026 is `10/26/2026`
- 8 November 2026 is `11/8/2026`

A UK-style `26/10/2026` is rejected. The date picker avoids the problem entirely. The hub reads the dates as the right UK day.

## Every fortnight

Add a new row with a Start and EndDate. It goes live on its Start date, and the hub picks it up the next time someone opens the page. The previous rows stay on the Spotlight page under **Previous spotlights**. Opening one of those shows today's figures for that topic, not the figures shown at the time.

A row with no dates sits under Previous spotlights while a dated row is running. That is how the NPD range rows are set up.

If a fortnight ends with no new dated row, the first undated row in the list takes over as "this fortnight" (today that is *Who is buying the new LFX extinguishers*). So add the next fortnight's row before EndDate passes, or the NPD row will be on screen at the Monday meeting.

## Product groups you can use

Commander Extinguishers · CommanderEDGE Extinguishers · Contempo Extinguishers · Renovated Extinguishers · Other Extinguishers · LFX · Fire Blankets · Cabinets/Trolleys/Stands · Signs · Alarm & Escape Equipment · Hose Reels & Accessories · Service Products & Spares · Service Tools · Pins & Sealing Products · Refill Packs · First Aid Supplies · Miscellaneous

## Ready-made examples

| Title | Kind | Category | Items |
|---|---|---|---|
| Who's buying LFX this year | Top customers | LFX | |
| Commander flat-pack stands | Top customers | | `SD-CFP-*; FP-COM-*; BX-FP-COM-*` |
| Our best-selling fire blankets | Top items | Fire Blankets | |

## What is in the list now (28 Oct 2026)

| Title | Kind | Category / Items | Dates | Tag |
|---|---|---|---|---|
| Who's buying LFX this year | Top customers | Category `LFX` | 10/26/2026 – 11/8/2026 | |
| Who is buying the new LFX extinguishers | Top customers | `FX-LFX-003-ACF; FX-LFX-003-AFS; FX-LFX-006-ACF; FX-LFX-006-AFS; FX-LFX-009-ACF; FX-LFX-009-AFS` | none | NPD range |
| Who is buying LFX fire blankets | Top customers | `FB-CAR-001-BGH; FB-CAR-002-BGH; FB-CAR-003-BGH; FB-CAR-004-BGH; FB-CAR-101-BGH; FB-FKL-001-BGH; FB-FKL-002-BGH; FB-FKL-101-BGH; FB-LIB-001-BGH; FB-LIB-002-BGH; FB-EXT-001-BGH; FB-EXT-002-BGH; FB-EXT-003-BGH; CB-BGH-001-BLK` | none | NPD range |
| Who is buying flat-pack tub stands | Top customers | `SD-FPT-001-RED; SD-FPT-002-CHR; SD-FPT-002-RED` | none | NPD range |
| Contempo Gold: which models sell | Top items | `CX-CO2-002-DGD; CX-CO2-005-DGD; CX-FFF-006-DGD; CX-FFF-009-DGD; CX-PWA-006-DGD; CX-PWA-009-DGD; CX-PWD-006-DGD; CX-PWD-009-DGD` | none | NPD range |
| Who is buying Contempo Gold stands | Top customers | `CS32-PG; CS33-PG` | none | NPD range |
| Who is buying the 3L wet chemical | Top customers | `FX-WCH-003-ACF` | none | NPD range |

The NPD codes come from the old NPD board; tell David if the range has changed. Two Contempo Gold codes (CX-CO2-005-DGD and CX-PWD-009-DGD) had no sales this year when the list was set up.

## On the page

- Click anyone on the podium or in the ranks to see their breakdown: their months against last year and what they bought, for the spotlighted products only. ‹ › or the arrow keys step through the ranking; Esc closes.
- The year-so-far chart has two views: **Running total** and **Each month**. Hover a month for the figures.
- **Present** opens a full-screen top 10 for the meeting; click a row there for the same breakdown.

## Good to know

- **Who sees it:** only people who can open the Sales & Marketing SharePoint site, where the figures live. Everyone else sees the Product Portal as before, without the band.
- **The figures:** sales in £, net of credit notes, from 1 January to the latest morning NetSuite refresh. CheckFire only, not PJ Fire.
- **Units:** counted as NetSuite records them. Bulk pallet kit codes such as `BX-FP-COM-001-RED` count one per pallet, not one per stand.
- **To take the band down:** untick Show on every row. As long as the list has rows, the band only shows rows from it.
- **If the list is deleted:** the hub falls back to its own built-in copy of the LFX and NPD topics (in `config.js`).
