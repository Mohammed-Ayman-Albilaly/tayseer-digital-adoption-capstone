# Tayseer Digital Adoption: Capstone (SDA-DSC-112)

Data Visualization & Storytelling capstone, SDAIA Academy.

**Decision:** recommend where to direct the next SAR 40 million to lift lagging regions toward the 65% digital-adoption target.

**Recommendation:** a tiered allocation. SAR 32M, weighted by gap, to Najran, Northern Borders, Al-Baha and Jazan (87% of the shortfall), and SAR 2M each to Al-Jouf, Hail, Tabuk and Asir.

## What we did
- Built an interactive Tableau dashboard of regional digital adoption (Day 2): national KPI plus regional adoption against the 65% target.
- Interpreted it: national adoption is 66.2%, yet 8 of 13 regions are below 65% (12.2 points combined shortfall).
- Turned it into a 7-slide executive story: BLUF/Ask, Situation, Complication, Evidence, Options, Recommendation, Ask + Next Step.
- Wrote a 7-minute speaker script with timings and Q&A preparation.

## Dashboard
[Tayseer Digital Adoption on Tableau Public](https://public.tableau.com/views/Dvis_17912834975050/TayseerDigitalAdoption)

## Files
| File | Purpose |
|---|---|
| `Tayseer_Capstone_Deck.pptx` | The 7-slide presentation (speaker notes included) |
| `Speaker_Script.md` | 7-minute script, speaker split, Q&A prep |
| `build_deck.js` | Script that generates the deck from the dashboard figures |

## Tools
- Tableau Public (dashboard)
- pptxgenjs (deck generation)
- Claude Code (drafting assistance)

## Assumptions and limits
- Dashboard figures are current-period adoption only; no time trend was analysed.
- The SAR allocation is illustrative, derived from each region's gap to 65%. Cost per point of adoption must be validated with regional delivery teams.

## Team
نواف السياري
تركي العتيبي
محمد البلالي

SDAIA GitHub: (https://github.com/SDAIAAcademy)

## Rebuild the deck
```bash
npm install
node build_deck.js
```
