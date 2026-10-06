const pptxgen = require("pptxgenjs");
let applyTheme = null;
try {
  ({ applyTheme } = require("C:/Users/Hmood/AppData/Roaming/Claude/local-agent-mode-sessions/skills-plugin/e905f0ed-5655-4839-840e-bfec50083ece/85df399a-adf8-4c43-b580-84c44b85a506/skills/pptx/scripts/apply_theme.js"));
} catch (e) {
  console.warn("apply_theme.js not found; deck keeps Office default theme colors");
}

const THEME = {
  name: "Tayseer Capstone",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "0B1F3A", lt1: "FFFFFF", dk2: "0B2545", lt2: "EEF2F7",
    accent1: "0B7A75", accent2: "F2A541", accent3: "C0392B",
    accent4: "2E8B57", accent5: "5B6B82", accent6: "8FB8DE",
    hlink: "0B7A75", folHlink: "5B6B82",
  },
};
const HEX = THEME.colors;

// Data read from the Tableau dashboard "Tayseer - Digital Adoption"
const REGIONS = [
  ["Riyadh", 70.8], ["Makkah", 70.7], ["Eastern Province", 68.7], ["Madinah", 68.2],
  ["Qassim", 66.1], ["Al-Jouf", 64.8], ["Hail", 64.7], ["Tabuk", 64.5], ["Asir", 64.4],
  ["Jazan", 63.0], ["Al-Baha", 62.8], ["Northern Borders", 62.7], ["Najran", 60.9],
];
const TARGET = 65;
const BIG4 = ["Najran", "Northern Borders", "Al-Baha", "Jazan"];
const gaps = REGIONS.filter(([, v]) => v < TARGET)
  .map(([n, v]) => [n, +(TARGET - v).toFixed(1)])
  .sort((a, b) => b[1] - a[1]);
const totalGap = gaps.reduce((s, [, g]) => s + g, 0);
const bigGap = gaps.filter(([n]) => BIG4.includes(n)).reduce((s, [, g]) => s + g, 0);
const bigShare = Math.round((bigGap / totalGap) * 100);

const ALLOC = [
  ["Najran", 12.5], ["Northern Borders", 7.0], ["Al-Baha", 6.5], ["Jazan", 6.0],
  ["Al-Jouf", 2.0], ["Hail", 2.0], ["Tabuk", 2.0], ["Asir", 2.0],
];
if (ALLOC.reduce((s, [, v]) => s + v, 0) !== 40) throw new Error("Allocation must total 40");
console.log("total gap", totalGap.toFixed(1), "| big-4 gap", bigGap.toFixed(1), "| share", bigShare + "%");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Tayseer Digital Adoption - Capstone";
pres.author = "Mohammed Albilaly";
pres.company = "SDAIA Academy - SDA-DSC-112";
const C = pres.SchemeColor;
const W = 13.33;
const FOOTER = "Tayseer Digital Adoption  |  SDA-DSC-112 Capstone";

function titlePlaceholder(color) {
  return {
    placeholder: {
      options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 1.3, fontSize: 30, bold: true, color, valign: "top", margin: 0 },
      text: "",
    },
  };
}

pres.defineSlideMaster({
  title: "DARK",
  objects: [
    { rect: { x: 0, y: 0, w: W, h: 7.5, fill: { color: C.text2 } } },
    titlePlaceholder(C.background1),
    { text: { text: FOOTER, options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
  ],
  slideNumber: { x: 12.1, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent6 },
});

pres.defineSlideMaster({
  title: "LIGHT",
  objects: [
    { rect: { x: 0, y: 0, w: W, h: 7.5, fill: { color: C.background1 } } },
    titlePlaceholder(C.text1),
    { text: { text: FOOTER, options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.1, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent5 },
});

// ---------- helpers ----------
let n = 0;
const nm = (s) => `${s}-${++n}`;

function card(slide, x, y, w, h, fill, name) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { type: "none" }, objectName: name || nm("card") });
}
function badge(slide, x, y, d, label, fill, color) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: nm("badge") });
  slide.addText(label, { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: 16, bold: true, color, margin: 0, isTextBox: true, objectName: nm("badge-label") });
}
function stat(slide, x, y, w, h, value, label, valueColor, fill, labelColor) {
  card(slide, x, y, w, h, fill);
  slide.addText(value, { x: x + 0.3, y: y + 0.15, w: w - 0.6, h: h * 0.5, fontFace: "Cambria", fontSize: 36, bold: true, color: valueColor, valign: "middle", margin: 0, isTextBox: true, objectName: nm("stat-value") });
  slide.addText(label, { x: x + 0.3, y: y + h * 0.58, w: w - 0.6, h: h * 0.36, fontSize: 14, color: labelColor, valign: "top", margin: 0, isTextBox: true, objectName: nm("stat-label") });
}
const chartCommon = {
  barDir: "bar",
  catAxisOrientation: "maxMin",
  valAxisHidden: true,
  valAxisMinVal: 0,
  showLegend: false,
  showValue: true,
  dataLabelPosition: "outEnd",
  dataLabelFormatCode: "0.0",
  dataLabelFontSize: 12,
  dataLabelColor: HEX.dk1,
  dataLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 12,
  catAxisLabelColor: HEX.dk1,
  catAxisLabelFontFace: "+mn-lt",
  valGridLine: { style: "none" },
  catGridLine: { style: "none" },
  barGapWidthPct: 45,
};
function legendChip(slide, x, y, color, text, w) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.05, w: 0.2, h: 0.2, fill: { color }, line: { type: "none" }, objectName: nm("legend-chip") });
  slide.addText(text, { x: x + 0.3, y, w, h: 0.3, fontSize: 12, color: C.text1, valign: "middle", margin: 0, isTextBox: true, objectName: nm("legend-text") });
}

// ---------- Slide 1: BLUF / Ask ----------
{
  const s = pres.addSlide({ masterName: "DARK" });
  s.addText("Invest SAR 40M where the gap is widest: four regions need most of it", { placeholder: "title" });
  s.addText("SAR 40M", { x: 0.6, y: 2.1, w: 6.0, h: 1.5, fontFace: "Cambria", fontSize: 80, bold: true, color: C.accent2, margin: 0, valign: "middle", isTextBox: true, objectName: "hero-stat" });
  s.addText("to lift lagging regions toward the 65% digital-adoption target", { x: 0.6, y: 3.7, w: 5.8, h: 1.1, fontSize: 22, color: C.background1, margin: 0, valign: "top", isTextBox: true, objectName: "hero-caption" });
  card(s, 7.0, 1.9, 5.7, 4.5, C.text1, "ask-card");
  s.addText("THE ASK", { x: 7.4, y: 2.1, w: 4.9, h: 0.4, fontSize: 14, bold: true, color: C.accent2, margin: 0, isTextBox: true, objectName: "ask-label" });
  s.addText(
    [
      { text: "Approve a tiered allocation", options: { bold: true, fontSize: 22, breakLine: true } },
      { text: "SAR 32M to the four widest gaps: Najran, Northern Borders, Al-Baha, Jazan", options: { bullet: true, breakLine: true, paraSpaceAfter: 10 } },
      { text: "SAR 8M closing buffer for four near-miss regions: Al-Jouf, Hail, Tabuk, Asir", options: { bullet: true } },
    ],
    { x: 7.4, y: 2.6, w: 4.9, h: 3.5, fontSize: 18, color: C.background1, valign: "top", margin: 0, paraSpaceAfter: 10, isTextBox: true, objectName: "ask-body" }
  );
  s.addNotes("[0:00-0:30] THE ASK FIRST. 'Today we are asking you to approve SAR 40 million for digital adoption, tiered so that 80 percent, SAR 32 million, goes to the four regions with the widest gaps: Najran, Northern Borders, Al-Baha and Jazan, and SAR 8 million closes the gap in four regions that are nearly there. Here is why.' Stop at 30 seconds.");
}

// ---------- Slide 2: Situation ----------
{
  const s = pres.addSlide({ masterName: "LIGHT" });
  s.addText("Nationally Tayseer is on track: 66.2% adoption against a 65% target", { placeholder: "title" });
  card(s, 0.6, 1.9, 4.2, 4.6, C.accent1, "kpi-tile");
  s.addText("66.2%", { x: 0.9, y: 2.5, w: 3.6, h: 1.5, fontFace: "Cambria", fontSize: 66, bold: true, color: C.background1, margin: 0, valign: "middle", isTextBox: true, objectName: "kpi-value" });
  s.addText("national digital adoption", { x: 0.9, y: 4.1, w: 3.6, h: 0.6, fontSize: 18, color: C.background1, margin: 0, valign: "top", isTextBox: true, objectName: "kpi-label" });
  stat(s, 5.2, 1.9, 3.6, 2.2, "+1.2 pts", "above the 65% target", HEX.accent4, C.background2, C.text1);
  stat(s, 9.1, 1.9, 3.6, 2.2, "13", "regions tracked in the Tayseer dashboard", HEX.accent1, C.background2, C.text1);
  card(s, 5.2, 4.4, 7.5, 2.1, C.background2, "limit-card");
  badge(s, 5.5, 4.7, 0.5, "!", C.accent2, C.text1);
  s.addText(
    [
      { text: "Limit of this evidence", options: { bold: true, fontSize: 16, breakLine: true } },
      { text: "The dashboard shows current adoption only. We judge 'on track' against the 65% target today; a trend over time is outside this analysis.", options: { fontSize: 15 } },
    ],
    { x: 6.2, y: 4.6, w: 6.2, h: 1.7, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 6, isTextBox: true, objectName: "limit-text" }
  );
  s.addNotes("[0:30-1:15] SITUATION. 'Start with the good news. Nationally, Tayseer adoption is 66.2 percent, 1.2 points above our 65 percent target, across 13 regions. If we stopped at this number we would conclude nothing needs to be done.' Be upfront: the dashboard shows current adoption, not a trend, so 'on track' means 'at target today'.");
}

// ---------- Slide 3: Complication ----------
{
  const s = pres.addSlide({ masterName: "LIGHT" });
  s.addText("The national average hides it: 8 of 13 regions sit below 65%", { placeholder: "title" });
  legendChip(s, 0.6, 1.8, HEX.accent4, "At or above 65% target", 2.4);
  legendChip(s, 3.6, 1.8, HEX.accent3, "Below target", 1.6);
  s.addChart(
    pres.charts.BAR,
    [{ name: "Digital adoption %", labels: REGIONS.map(([r]) => r), values: REGIONS.map(([, v]) => v) }],
    { ...chartCommon, x: 0.6, y: 2.2, w: 7.4, h: 4.7, valAxisMaxVal: 80, chartColors: REGIONS.map(([, v]) => (v >= TARGET ? HEX.accent4 : HEX.accent3)) }
  );
  stat(s, 8.5, 1.9, 4.2, 1.5, "8 of 13", "regions are below the 65% target", HEX.accent3, C.background2, C.text1);
  stat(s, 8.5, 3.6, 4.2, 1.5, `${totalGap.toFixed(1)} pts`, "combined shortfall to target", HEX.accent3, C.background2, C.text1);
  stat(s, 8.5, 5.3, 4.2, 1.5, "4.1 pts", "Najran, the widest single gap", HEX.accent3, C.background2, C.text1);
  s.addNotes("[1:15-2:00] COMPLICATION. 'But averages hide gaps. Eight of our 13 regions are below 65 percent, a combined shortfall of 12.2 percentage points. Najran is furthest behind, 4.1 points short.' Orient the audience: green is at or above target, red is below.");
}

// ---------- Slide 4: Evidence ----------
{
  const s = pres.addSlide({ masterName: "LIGHT" });
  s.addText(`Four regions hold ${bigShare}% of the shortfall, led by Najran at 4.1 points`, { placeholder: "title" });
  legendChip(s, 0.6, 1.8, HEX.accent3, "Widest four gaps (2+ pts)", 2.8);
  legendChip(s, 4.0, 1.8, HEX.accent2, "Near-miss (under 1 pt)", 2.4);
  s.addChart(
    pres.charts.BAR,
    [{ name: "Gap to 65% (pts)", labels: gaps.map(([r]) => r), values: gaps.map(([, g]) => g) }],
    { ...chartCommon, x: 0.6, y: 2.2, w: 7.4, h: 4.2, valAxisMaxVal: 5, chartColors: gaps.map(([r]) => (BIG4.includes(r) ? HEX.accent3 : HEX.accent2)) }
  );
  s.addText("Gap = 65% target minus current adoption, in percentage points.", { x: 0.6, y: 6.5, w: 7.4, h: 0.3, fontSize: 12, color: C.accent5, margin: 0, isTextBox: true, objectName: "chart-footnote" });
  card(s, 8.5, 1.9, 4.2, 2.3, C.accent3, "evidence-card-1");
  s.addText(`${bigShare}%`, { x: 8.8, y: 2.0, w: 3.6, h: 1.2, fontFace: "Cambria", fontSize: 54, bold: true, color: C.background1, margin: 0, valign: "middle", isTextBox: true, objectName: "evidence-stat-1" });
  s.addText("of the combined shortfall sits in just four regions", { x: 8.8, y: 3.2, w: 3.6, h: 0.9, fontSize: 16, color: C.background1, margin: 0, valign: "top", isTextBox: true, objectName: "evidence-label-1" });
  card(s, 8.5, 4.5, 4.2, 2.3, C.background2, "evidence-card-2");
  s.addText("< 1 pt", { x: 8.8, y: 4.6, w: 3.6, h: 1.2, fontFace: "Cambria", fontSize: 54, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "evidence-stat-2" });
  s.addText("is all the other four regions need to reach target", { x: 8.8, y: 5.8, w: 3.6, h: 0.9, fontSize: 16, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "evidence-label-2" });
  s.addNotes("[2:00-4:00] EVIDENCE. Point at the chart before the conclusion. 'Najran 4.1, Northern Borders 2.3, Al-Baha 2.2, Jazan 2.0: together 10.6 of the 12.2 points, that is 87 percent of the shortfall. Al-Jouf, Hail, Tabuk and Asir are each within 0.6 points, nearly there.' So what: these two groups need different treatment.");
}

// ---------- Slide 5: Options ----------
{
  const s = pres.addSlide({ masterName: "LIGHT" });
  s.addText("Three ways to spend SAR 40M: only a tiered split matches money to need", { placeholder: "title" });
  const opts = [
    { tag: "A", name: "Spread evenly", spend: "SAR 5M to each of the 8 lagging regions", up: "Simple and visibly fair", risk: "Al-Jouf (0.2 pt gap) gets the same as Najran (4.1 pts): about 20x more money per point of gap", rec: false },
    { tag: "B", name: "Concentrate on the widest four", spend: "SAR 10M to each of the 4 widest gaps", up: "Maximum push where the gaps are largest", risk: "Four regions stay below 65%, and funding per point is uneven between the four", rec: false },
    { tag: "C", name: "Tiered by gap", spend: "SAR 32M weighted to the 4 widest, SAR 2M to each of the other 4", up: "Every lagging region funded; about SAR 3M per point of gap in the widest four", risk: "Needs regional cost-per-point validation", rec: true },
  ];
  opts.forEach((o, i) => {
    const x = 0.6 + i * 4.15;
    const fill = o.rec ? C.accent1 : C.background2;
    const fg = o.rec ? C.background1 : C.text1;
    card(s, x, 1.9, 3.85, 4.6, fill, `option-card-${o.tag}`);
    badge(s, x + 0.3, 2.15, 0.55, o.tag, o.rec ? C.accent2 : C.accent1, o.rec ? C.text1 : C.background1);
    if (o.rec) s.addText("RECOMMENDED", { x: x + 1.1, y: 2.2, w: 2.5, h: 0.45, fontSize: 12, bold: true, color: C.accent2, valign: "middle", margin: 0, isTextBox: true, objectName: "recommended-tag" });
    s.addText(
      [
        { text: o.name, options: { bold: true, fontSize: 20, breakLine: true } },
        { text: o.spend, options: { fontSize: 15, breakLine: true } },
        { text: "Upside", options: { bold: true, fontSize: 14, breakLine: true } },
        { text: o.up, options: { fontSize: 14, breakLine: true } },
        { text: "Risk", options: { bold: true, fontSize: 14, breakLine: true } },
        { text: o.risk, options: { fontSize: 14 } },
      ],
      { x: x + 0.3, y: 2.95, w: 3.3, h: 3.4, color: fg, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true, objectName: `option-text-${o.tag}` }
    );
  });
  s.addText("Allocations are illustrative and use only the dashboard gaps; regional cost per point of adoption must be validated.", { x: 0.6, y: 6.6, w: 12.1, h: 0.3, fontSize: 12, color: C.accent5, margin: 0, isTextBox: true, objectName: "options-footnote" });
  s.addNotes("[4:00-5:15] OPTIONS. 'We looked at three realistic ways. A: spread evenly, SAR 5M each: fair, but Al-Jouf, 0.2 points short, gets the same as Najran, 4.1 short, about 20 times more money per point. B: put everything into the widest four, SAR 10M each: leaves four regions below target. C: tiered, SAR 32M weighted to the widest four and SAR 2M to each of the rest.' Be explicit these are illustrative allocations built from dashboard gaps only.");
}

// ---------- Slide 6: Recommendation ----------
{
  const s = pres.addSlide({ masterName: "LIGHT" });
  s.addText("Recommend Option C: fund every lagging region in proportion to its gap", { placeholder: "title" });
  s.addText("SAR millions by region (total SAR 40M)", { x: 0.6, y: 1.85, w: 7.4, h: 0.3, fontSize: 14, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "alloc-subtitle" });
  s.addChart(
    pres.charts.BAR,
    [{ name: "SAR millions", labels: ALLOC.map(([r]) => r), values: ALLOC.map(([, v]) => v) }],
    { ...chartCommon, x: 0.6, y: 2.2, w: 7.4, h: 4.6, valAxisMaxVal: 15, chartColors: ALLOC.map(([r]) => (BIG4.includes(r) ? HEX.accent1 : HEX.accent6)) }
  );
  const why = [
    ["1", "Closes the widest gaps first", "SAR 32M (80%) goes to the four regions that hold 87% of the shortfall"],
    ["2", "No lagging region left behind", "SAR 2M each helps the four near-miss regions cross 65%"],
    ["3", "Money follows need", "About SAR 3M per point of gap in each of the widest four"],
  ];
  why.forEach(([num, head, body], i) => {
    const y = 1.9 + i * 1.65;
    card(s, 8.5, y, 4.2, 1.45, C.background2, `why-card-${num}`);
    badge(s, 8.75, y + 0.45, 0.55, num, C.accent1, C.background1);
    s.addText(
      [
        { text: head, options: { bold: true, fontSize: 16, breakLine: true } },
        { text: body, options: { fontSize: 14 } },
      ],
      { x: 9.5, y: y + 0.12, w: 3.0, h: 1.2, color: C.text1, valign: "middle", margin: 0, paraSpaceAfter: 4, isTextBox: true, objectName: `why-text-${num}` }
    );
  });
  s.addNotes("[5:15-6:15] RECOMMENDATION. 'We recommend Option C. SAR 12.5M to Najran, 7M to Northern Borders, 6.5M to Al-Baha, 6M to Jazan, and 2M each to Al-Jouf, Hail, Tabuk and Asir. That is about SAR 3 million per point of gap in the widest four. It closes the biggest gaps first and leaves nobody behind.' Tie back: every number here traces to the gap chart on slide 4.");
}

// ---------- Slide 7: Ask + next step ----------
{
  const s = pres.addSlide({ masterName: "DARK" });
  s.addText("Decision today: approve the SAR 40M tiered allocation", { placeholder: "title" });
  const steps = [
    ["1", "Approve today", "SAR 40M tiered allocation: SAR 32M to the four widest gaps, SAR 8M to the four near-miss regions"],
    ["2", "Validate within 30 days (proposed)", "Confirm cost per point of adoption with each region's delivery team before funds are released"],
    ["3", "Review at 6 months (proposed)", "Re-check the Tayseer dashboard: are all eight lagging regions at or above 65%?"],
  ];
  steps.forEach(([num, head, body], i) => {
    const x = 0.6 + i * 4.15;
    card(s, x, 1.9, 3.85, 2.9, C.text1, `step-card-${num}`);
    badge(s, x + 0.3, 2.15, 0.55, num, C.accent2, C.text1);
    s.addText(
      [
        { text: head, options: { bold: true, fontSize: 18, breakLine: true } },
        { text: body, options: { fontSize: 14 } },
      ],
      { x: x + 0.3, y: 2.9, w: 3.3, h: 1.8, color: C.background1, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true, objectName: `step-text-${num}` }
    );
  });
  card(s, 0.6, 5.1, 12.1, 1.5, C.accent1, "closing-banner");
  s.addText("The ask, once more: approve SAR 40M, with 80% to Najran, Northern Borders, Al-Baha and Jazan.", { x: 0.9, y: 5.2, w: 11.5, h: 1.3, fontSize: 22, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "closing-ask" });
  s.addNotes("[6:15-7:00] ASK + NEXT STEP. Restate the ask. 'Approve the SAR 40 million tiered allocation today. Within 30 days we validate cost per point with each region's delivery team, and at six months we re-check the dashboard against the 65 percent target.' Then stop and invite questions. Q&A: acknowledge uncertainty (cost per point is an assumption) and return to the Big Idea.");
}

(async () => {
  const out = "Tayseer_Capstone_Deck.pptx";
  await pres.writeFile({ fileName: out });
  if (applyTheme) await applyTheme(out, THEME);
  console.log("wrote", out);
})();
