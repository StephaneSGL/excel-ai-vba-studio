# Native Excel tables and charts

Excel AI & VBA Studio models worksheet tables and charts as first-class workbook
objects. They are persisted as real Excel `ListObject` and `ChartObject`
instances, not as a single worksheet AutoFilter or a static chart image.

Microsoft references: [`ListObjects.Add`](https://learn.microsoft.com/en-us/office/vba/api/excel.listobjects.add),
[`ChartObjects.Add`](https://learn.microsoft.com/en-us/office/vba/api/excel.chartobjects.add),
and the [`XlChartType` enumeration](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcharttype).

## Table choices

The table designer exposes:

- worksheet and local A1 range;
- stable Excel table name and display name;
- inventoried header-row and totals-row state (preserved, not toggled);
- first-column and last-column emphasis;
- row stripes and column stripes;
- all 60 built-in table styles: `TableStyleLight1`–`21`,
  `TableStyleMedium1`–`28`, and `TableStyleDark1`–`11`.

Several tables may use the same columns when their row ranges do not overlap.

Native creation requires `headerRow=true` and `totalsRow=false`. Headerless
creation and header-row transitions are refused because Excel synthesizes or
moves header cells. Microsoft Excel also inserts or removes worksheet cells when
the totals row is toggled and rewrites formulas that refer to those cells, so
totals-row transitions are refused. Existing header and totals states are
inventoried and preserved; a table with totals has its range locked, while its
name and style remain editable.
For example, `A1:C5`, `A20:C30`, and `A40:C50` are three independent tables.
An overlapping range, a duplicate name, or an invalid Excel header is rejected
before the source file is replaced.

Table names follow Excel's native rules: the first character may be a letter,
underscore, or backslash; the remaining characters may also include digits and
periods. Cell-reference-like names and workbook-wide duplicates are rejected.
The integrated grid stores the native table definition immediately; Excel's
exact built-in style rendering appears after the native save and reopen. The
preview does not paint fake cell styles over existing formatting.

## Chart-type inventory

The catalog follows the complete 103-value `XlChartType` enumeration published
by Microsoft. The designer groups the choices as follows:

| Family | Choices |
| --- | --- |
| Area | area, stacked, 100% stacked, 3D variants, and modern `Ex` variants |
| Bar | clustered, stacked, 100% stacked, 3D variants, and modern `Ex` variants |
| Column | clustered, stacked, 100% stacked, 3D variants, and modern `Ex` variants |
| Line | plain, markers, stacked, 100% stacked, 3D, and modern `Ex` variants |
| Pie | plain, exploded, 3D, exploded 3D, pie-of-pie, bar-of-pie, and modern `Ex` |
| Doughnut | plain, exploded, and modern `Ex` |
| Scatter | points, straight or smooth lines, with or without markers, and modern `Ex` |
| Bubble | plain, 3D effect, and modern `Ex` |
| Radar | plain, markers, and filled |
| Stock | HLC, OHLC, VHLC, and VOHLC |
| Surface | 3D, wireframe, top view, and top-view wireframe |
| Combination | custom, column/line, column/line secondary axis, stacked-area/column, and other combinations |
| Modern | treemap, histogram, waterfall, sunburst, box-and-whisker, Pareto, funnel, and region map |
| Recommended | Excel suggested chart |
| Legacy 3D | cone, cylinder, and pyramid bar/column variants, including clustered, stacked, and 100% stacked |

The exact constant names and numeric values are maintained in
`src/common/excelWorkbookObjects.ts`. Of the 103 catalog entries, 101 can be
selected for offline local creation. `xlSuggestedChart` remains visible but is
disabled because it is a recommendation request, not a stable persisted type.
`xlRegionMap` is also inventoried but disabled because Excel can transmit its
geographic data to Bing Maps while creating the chart. Other modern types
remain dependent on the locally installed Excel version; an unsupported type
fails the transaction instead of silently becoming another chart.

## Chart choices

### Object and data source

- worksheet, chart name, alternative text;
- local source range and plot-by-rows or plot-by-columns;
- position and size in points;
- chart style, rounded corners, gap width, and series overlap.

Only local, simple A1 ranges are accepted. Workbook-qualified, external, DDE,
and formula-based range input is rejected at the tool boundary. A single
table/chart range is limited to 1,000,000 cells and one transaction to
5,000,000 referenced cells so an accidental whole-sheet request cannot stall
Excel. A chart uses either one automatic source range or explicit series, never
both in the same write request.

### Series

- add, remove, select, reorder, show, or hide a series (using Excel's native
  series filter rather than a transparent-line approximation);
- fixed series name or a name-cell range containing exactly one cell;
- category, values, X-values, and bubble-size ranges;
- chart type per series for combination charts;
- primary or secondary axis;
- fill and line colours;
- line width and solid, dash, dot, or dash-dot style;
- automatic, circle, dash, diamond, dot, none, picture, plus, square, star,
  triangle, or X marker, including marker size;
- line smoothing;
- data-label value, category, series name, percentage, bubble size, and
  position.

When explicit series use different chart types, the editor canonicalizes the
top-level chart to Excel's custom combination type before native saving. A
homogeneous request is canonicalized back to its single concrete series type
instead of being persisted under a misleading `xlCombo` label. Bubble and non-bubble series cannot be
mixed: Excel silently promotes every series in that combination to Bubble, so
the extension rejects the request before opening Excel instead of changing its
meaning.

Excel accepts only certain explicit data-label positions for each native series
type. The designer filters the list accordingly; all other types keep Excel's
automatic label placement.

Every explicit label definition must include at least one `showValue`,
`showCategoryName`, `showSeriesName`, `showPercentage`, or `showBubbleSize`
boolean. A position alone is rejected because Excel's `ApplyDataLabels()` would
silently enable value labels. After labels are applied, omitted compatible show
flags are set to `false`, so the persisted result does not depend on Excel's
defaults. When every explicit show flag is `false`, the series' labels are
removed without calling `ApplyDataLabels()`; a position is therefore rejected
unless at least one show flag is `true`.

| Series type | Explicit positions |
| --- | --- |
| Clustered column/bar | center, inside base, inside end, outside end |
| Stacked and 100% stacked column/bar | center, inside base, inside end |
| 2D line, scatter, and bubble | above, below, center, left, right |
| Pie, exploded pie, 3D pie, pie-of-pie, and bar-of-pie | best fit, center, inside end, outside end |
| Doughnut, 3D column/bar/line, area, radar, and other types | automatic placement only |

### Axes and presentation

- chart title and category/value axis titles;
- primary and secondary category/value axes;
- automatic or explicit minimum, maximum, major unit, and minor unit;
- linear or logarithmic scale and reversed order;
- source-linked automatic or explicit number format. Clearing the format field
  restores Excel's `NumberFormatLinked` state instead of leaving the previous
  custom format in place;
- major and minor gridlines;
- legend visibility and bottom, corner, left, right, or top position. A custom
  position preserves an existing manual Excel layout during an update; the
  extension does not invent custom legend coordinates for a new chart.

Some combinations are invalid in Excel itself—for example, a logarithmic axis
with a non-positive bound, bubble options on a non-bubble series, or secondary
axes on 3D and other incompatible chart families. Gap width and overlap are
applied only to compatible chart groups. The extension validates what can be
decided statically, then reopens the working copy in Excel and verifies the
persisted result before committing it.

When an existing chart's anchor or modeled series are unchanged, the integrated
editor leaves those native objects untouched instead of rebuilding them. This
keeps their geometry and allows Excel to retain unmodeled native series features
such as trendlines, error bars, and per-point formatting during a title-only or
legend-only edit. Changing the built-in Excel chart style also keeps the native
series, trendlines, and error bars, while allowing Excel to reapply the
theme-controlled series and point formatting. An explicit series change
intentionally replaces the modeled series; unsupported series features should
therefore be edited in Excel itself.

## Ribbon and Copilot

The Insert ribbon opens the same chart model used for create and edit actions.
The model starts from the current cell selection, but its range and series can
be changed before saving. `#excelWorkbookDesign` exposes the corresponding
create, update, and delete operations for tables and charts to GitHub Copilot.

Opening a workbook remains a package read and does not start Microsoft Excel.
Excel COM is used only for a native object transaction or an explicit native
handoff. Every native transaction checks the source SHA-256, creates a persistent
backup, edits a working copy with macros disabled, verifies the supported target
properties and the inventory of untargeted objects after reopening, and either
replaces the workbook atomically or rolls back.

## Deliberate limits

- PivotCharts and PivotTables are not created or rewritten by this designer;
  detected PivotCharts remain native and are reported as non-editable.
- External chart data and cross-workbook formulas are not accepted.
- Availability of modern chart types is determined by the installed Excel
  build, not simulated by the extension.
- Region-map creation is disabled in automation because Excel may contact Bing
  Maps with workbook geography data. Existing map charts remain preserved.
- Classic `c:chartSpace` charts are inventoried and remain editable after
  reopening. Modern `chartEx` objects are detected and preserved through the
  native save path, but are reported as non-editable instead of being decoded
  as an incorrect classic chart.
- The extension never executes a macro to create or display a chart.
