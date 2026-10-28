# Sales Spotlight: setting the fortnightly topic

The Sales Spotlight sits on the Product Portal page, above **What are you looking for?**. It shows one topic at a time: the top 10 customers (or products) for a product group or a set of item codes, this year to date.

The figures come from NetSuite every morning through the sales dashboard. Nobody pulls a report. You only choose the topic.

## One-off: create the list

On the **Product Portal** SharePoint site (`/sites/CheckFireProductPortal`), create a list called **Spotlight** with these columns:

| Column | Type | What goes in it |
|---|---|---|
| Title | (already there) | The headline, e.g. *Who's buying LFX this year* |
| Type | Choice: `Top customers`, `Top items` | Rank customers, or rank products |
| Category | Single line of text | A product group exactly as NetSuite names it, e.g. `LFX`, `Fire Blankets`, `Cabinets/Trolleys/Stands`. Optional if you use Items |
| Items | Single line of text | Optional. Item codes separated by `;`. End a code with `*` to take every code that starts that way, e.g. `SD-CFP-*; FP-COM-*; BX-FP-COM-*` |
| Start | Date | First day it shows |
| End | Date | Last day it shows |
| Show | Yes/No (default Yes) | Untick to hide a row without deleting it |
| Tag | Single line of text | Optional. A short label on the card under Previous spotlights, e.g. `NPD range` |

If you fill in both Category and Items, a sale has to match both.

## Every fortnight

Add a new row. It goes live on its Start date, and the hub picks it up the next time someone opens the page. The previous rows stay on the Spotlight page under **Previous spotlights**. Opening one of those shows today's figures for that topic, not the figures shown at the time.

If a fortnight ends with no new row, the last one stays up, so the Monday meeting never opens on an empty space.

## Product groups you can use

Commander Extinguishers · CommanderEDGE Extinguishers · Contempo Extinguishers · Renovated Extinguishers · Other Extinguishers · LFX · Fire Blankets · Cabinets/Trolleys/Stands · Signs · Alarm & Escape Equipment · Hose Reels & Accessories · Service Products & Spares · Service Tools · Pins & Sealing Products · Refill Packs · First Aid Supplies · Miscellaneous

## Ready-made examples

| Title | Type | Category | Items |
|---|---|---|---|
| Who's buying LFX this year | Top customers | LFX | |
| Commander flat-pack stands | Top customers | | `SD-CFP-*; FP-COM-*; BX-FP-COM-*` |
| Our best-selling fire blankets | Top items | Fire Blankets | |

## The NPD range

Until the list exists, the hub shows the NPD range as built-in topics under the current one:

| Title | Type | Items | Tag |
|---|---|---|---|
| Who is buying the new LFX extinguishers | Top customers | `FX-LFX-003-ACF; FX-LFX-003-AFS; FX-LFX-006-ACF; FX-LFX-006-AFS; FX-LFX-009-ACF; FX-LFX-009-AFS` | NPD range |
| Who is buying LFX fire blankets | Top customers | `FB-CAR-001-BGH; FB-CAR-002-BGH; FB-CAR-003-BGH; FB-CAR-004-BGH; FB-CAR-101-BGH; FB-FKL-001-BGH; FB-FKL-002-BGH; FB-FKL-101-BGH; FB-LIB-001-BGH; FB-LIB-002-BGH; FB-EXT-001-BGH; FB-EXT-002-BGH; FB-EXT-003-BGH; CB-BGH-001-BLK` | NPD range |
| Who is buying flat-pack tub stands | Top customers | `SD-FPT-001-RED; SD-FPT-002-CHR; SD-FPT-002-RED` | NPD range |
| Contempo Gold: which models sell | Top items | `CX-CO2-002-DGD; CX-CO2-005-DGD; CX-FFF-006-DGD; CX-FFF-009-DGD; CX-PWA-006-DGD; CX-PWA-009-DGD; CX-PWD-006-DGD; CX-PWD-009-DGD` | NPD range |
| Who is buying Contempo Gold stands | Top customers | `CS32-PG; CS33-PG` | NPD range |
| Who is buying the 3L wet chemical | Top customers | `FX-WCH-003-ACF` | NPD range |

When you create the list, copy these rows in (with dates) so nothing disappears. The codes come from the old NPD board; tell David if the range has changed.

## On the page

- Click anyone on the podium or in the ranks to see their breakdown: their months against last year, what they bought, and (for a customer) how big this topic is in their whole account. ‹ › or the arrow keys step through the ranking; Esc closes.
- The year-so-far chart has two views: **Running total** and **Each month**. Hover a month for the figures.
- **Present** opens a full-screen top 10 for the meeting; click a row there for the same breakdown.

## Good to know

- **Who sees it:** only people who can open the Sales & Marketing SharePoint site. Everyone else sees the Product Portal as before, without the band.
- **The figures:** sales in £, net of credit notes, from 1 January to the latest morning NetSuite refresh. CheckFire only, not PJ Fire.
- **Units:** counted as NetSuite records them. Bulk pallet kit codes such as `BX-FP-COM-001-RED` count one per pallet, not one per stand.
- **No list yet?** The hub shows *Who's buying LFX this year* until the list exists.
- **To take the band down:** untick Show on every row. Once the list exists, the band only shows rows from it.
