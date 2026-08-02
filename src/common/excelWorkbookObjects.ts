export interface SheetTableStyle {
    name: string;
    showFirstColumn: boolean;
    showLastColumn: boolean;
    showRowStripes: boolean;
    showColumnStripes: boolean;
}

export interface SheetTableData {
    id: string;
    name: string;
    displayName: string;
    rangeRef: string;
    headerRow: boolean;
    totalsRow: boolean;
    style: SheetTableStyle;
}

export type ChartPlotBy = 'columns' | 'rows';
export type ChartAxisGroup = 'primary' | 'secondary';
export type ChartLegendPosition = 'bottom' | 'corner' | 'custom' | 'left' | 'right' | 'top';
export type ChartDataLabelPosition =
    | 'above'
    | 'below'
    | 'bestFit'
    | 'center'
    | 'insideBase'
    | 'insideEnd'
    | 'left'
    | 'outsideEnd'
    | 'right';
export type ChartMarkerStyle =
    | 'automatic'
    | 'circle'
    | 'dash'
    | 'diamond'
    | 'dot'
    | 'none'
    | 'picture'
    | 'plus'
    | 'square'
    | 'star'
    | 'triangle'
    | 'x';

export interface SheetChartSeriesData {
    id: string;
    name?: string;
    nameRange?: string;
    categoryRange?: string;
    valuesRange: string;
    xValuesRange?: string;
    bubbleSizesRange?: string;
    chartType?: number;
    axisGroup?: ChartAxisGroup;
    color?: string;
    lineColor?: string;
    lineWidth?: number;
    dashStyle?: 'solid' | 'dash' | 'dot' | 'dashDot';
    markerStyle?: ChartMarkerStyle;
    markerSize?: number;
    smooth?: boolean;
    visible?: boolean;
    dataLabels?: {
        showValue?: boolean;
        showCategoryName?: boolean;
        showSeriesName?: boolean;
        showPercentage?: boolean;
        showBubbleSize?: boolean;
        position?: ChartDataLabelPosition;
    };
}

export const CHART_DATA_LABEL_SHOW_KEYS = [
    'showValue',
    'showCategoryName',
    'showSeriesName',
    'showPercentage',
    'showBubbleSize',
] as const;

export function chartDataLabelsHaveExplicitShowOption(
    labels: SheetChartSeriesData['dataLabels'] | undefined,
): boolean {
    return labels !== undefined
        && CHART_DATA_LABEL_SHOW_KEYS.some(property => labels[property] !== undefined);
}

export function chartDataLabelsHaveEnabledShowOption(
    labels: SheetChartSeriesData['dataLabels'] | undefined,
): boolean {
    return labels !== undefined
        && CHART_DATA_LABEL_SHOW_KEYS.some(property => labels[property] === true);
}

export interface SheetChartAxisData {
    visible?: boolean;
    title?: string;
    minimumScale?: number | null;
    maximumScale?: number | null;
    majorUnit?: number | null;
    minorUnit?: number | null;
    logarithmic?: boolean;
    reverseOrder?: boolean;
    numberFormat?: string;
    majorGridlines?: boolean;
    minorGridlines?: boolean;
}

export interface SheetChartData {
    id: string;
    name: string;
    chartType: number;
    sourceRangeRef?: string;
    plotBy: ChartPlotBy;
    anchor: {
        left: number;
        top: number;
        width: number;
        height: number;
    };
    title?: {
        visible: boolean;
        text: string;
    };
    legend?: {
        visible: boolean;
        position: ChartLegendPosition;
    };
    categoryAxis?: SheetChartAxisData;
    valueAxis?: SheetChartAxisData;
    secondaryCategoryAxis?: SheetChartAxisData;
    secondaryValueAxis?: SheetChartAxisData;
    series?: SheetChartSeriesData[];
    style?: number;
    roundedCorners?: boolean;
    gapWidth?: number;
    overlap?: number;
    alternativeText?: string;
}

export interface ExcelChartTypeOption {
    id: string;
    constant: string;
    value: number;
    group: string;
    label: string;
    modern?: boolean;
}

/**
 * Complete XlChartType inventory published by Microsoft. Availability of the
 * modern types still depends on the locally installed Excel version.
 */
export const EXCEL_CHART_TYPES: readonly ExcelChartTypeOption[] = [
    { id: 'area', constant: 'xlArea', value: 1, group: 'Area', label: 'Area' },
    { id: 'area-stacked', constant: 'xlAreaStacked', value: 76, group: 'Area', label: 'Stacked Area' },
    { id: 'area-stacked-100', constant: 'xlAreaStacked100', value: 77, group: 'Area', label: '100% Stacked Area' },
    { id: '3d-area', constant: 'xl3DArea', value: -4098, group: 'Area', label: '3D Area' },
    { id: '3d-area-stacked', constant: 'xl3DAreaStacked', value: 78, group: 'Area', label: '3D Stacked Area' },
    { id: '3d-area-stacked-100', constant: 'xl3DAreaStacked100', value: 79, group: 'Area', label: '3D 100% Stacked Area' },
    { id: 'area-ex', constant: 'xlAreaEx', value: 135, group: 'Area', label: 'Area (modern)', modern: true },
    { id: 'area-stacked-ex', constant: 'xlAreaStackedEx', value: 136, group: 'Area', label: 'Stacked Area (modern)', modern: true },
    { id: 'area-stacked-100-ex', constant: 'xlAreaStacked100Ex', value: 137, group: 'Area', label: '100% Stacked Area (modern)', modern: true },
    { id: 'bar-clustered', constant: 'xlBarClustered', value: 57, group: 'Bar', label: 'Clustered Bar' },
    { id: 'bar-stacked', constant: 'xlBarStacked', value: 58, group: 'Bar', label: 'Stacked Bar' },
    { id: 'bar-stacked-100', constant: 'xlBarStacked100', value: 59, group: 'Bar', label: '100% Stacked Bar' },
    { id: '3d-bar-clustered', constant: 'xl3DBarClustered', value: 60, group: 'Bar', label: '3D Clustered Bar' },
    { id: '3d-bar-stacked', constant: 'xl3DBarStacked', value: 61, group: 'Bar', label: '3D Stacked Bar' },
    { id: '3d-bar-stacked-100', constant: 'xl3DBarStacked100', value: 62, group: 'Bar', label: '3D 100% Stacked Bar' },
    { id: 'bar-clustered-ex', constant: 'xlBarClusteredEx', value: 132, group: 'Bar', label: 'Clustered Bar (modern)', modern: true },
    { id: 'bar-stacked-ex', constant: 'xlBarStackedEx', value: 133, group: 'Bar', label: 'Stacked Bar (modern)', modern: true },
    { id: 'bar-stacked-100-ex', constant: 'xlBarStacked100Ex', value: 134, group: 'Bar', label: '100% Stacked Bar (modern)', modern: true },
    { id: 'column-clustered', constant: 'xlColumnClustered', value: 51, group: 'Column', label: 'Clustered Column' },
    { id: 'column-stacked', constant: 'xlColumnStacked', value: 52, group: 'Column', label: 'Stacked Column' },
    { id: 'column-stacked-100', constant: 'xlColumnStacked100', value: 53, group: 'Column', label: '100% Stacked Column' },
    { id: '3d-column', constant: 'xl3DColumn', value: -4100, group: 'Column', label: '3D Column' },
    { id: '3d-column-clustered', constant: 'xl3DColumnClustered', value: 54, group: 'Column', label: '3D Clustered Column' },
    { id: '3d-column-stacked', constant: 'xl3DColumnStacked', value: 55, group: 'Column', label: '3D Stacked Column' },
    { id: '3d-column-stacked-100', constant: 'xl3DColumnStacked100', value: 56, group: 'Column', label: '3D 100% Stacked Column' },
    { id: 'column-clustered-ex', constant: 'xlColumnClusteredEx', value: 124, group: 'Column', label: 'Clustered Column (modern)', modern: true },
    { id: 'column-stacked-ex', constant: 'xlColumnStackedEx', value: 125, group: 'Column', label: 'Stacked Column (modern)', modern: true },
    { id: 'column-stacked-100-ex', constant: 'xlColumnStacked100Ex', value: 126, group: 'Column', label: '100% Stacked Column (modern)', modern: true },
    { id: 'line', constant: 'xlLine', value: 4, group: 'Line', label: 'Line' },
    { id: 'line-markers', constant: 'xlLineMarkers', value: 65, group: 'Line', label: 'Line with Markers' },
    { id: 'line-stacked', constant: 'xlLineStacked', value: 63, group: 'Line', label: 'Stacked Line' },
    { id: 'line-stacked-markers', constant: 'xlLineMarkersStacked', value: 66, group: 'Line', label: 'Stacked Line with Markers' },
    { id: 'line-stacked-100', constant: 'xlLineStacked100', value: 64, group: 'Line', label: '100% Stacked Line' },
    { id: 'line-stacked-100-markers', constant: 'xlLineMarkersStacked100', value: 67, group: 'Line', label: '100% Stacked Line with Markers' },
    { id: '3d-line', constant: 'xl3DLine', value: -4101, group: 'Line', label: '3D Line' },
    { id: 'line-ex', constant: 'xlLineEx', value: 127, group: 'Line', label: 'Line (modern)', modern: true },
    { id: 'line-stacked-ex', constant: 'xlLineStackedEx', value: 128, group: 'Line', label: 'Stacked Line (modern)', modern: true },
    { id: 'line-stacked-100-ex', constant: 'xlLineStacked100Ex', value: 129, group: 'Line', label: '100% Stacked Line (modern)', modern: true },
    { id: 'pie', constant: 'xlPie', value: 5, group: 'Pie', label: 'Pie' },
    { id: 'pie-exploded', constant: 'xlPieExploded', value: 69, group: 'Pie', label: 'Exploded Pie' },
    { id: '3d-pie', constant: 'xl3DPie', value: -4102, group: 'Pie', label: '3D Pie' },
    { id: '3d-pie-exploded', constant: 'xl3DPieExploded', value: 70, group: 'Pie', label: 'Exploded 3D Pie' },
    { id: 'pie-of-pie', constant: 'xlPieOfPie', value: 68, group: 'Pie', label: 'Pie of Pie' },
    { id: 'bar-of-pie', constant: 'xlBarOfPie', value: 71, group: 'Pie', label: 'Bar of Pie' },
    { id: 'pie-ex', constant: 'xlPieEx', value: 130, group: 'Pie', label: 'Pie (modern)', modern: true },
    { id: 'doughnut', constant: 'xlDoughnut', value: -4120, group: 'Doughnut', label: 'Doughnut' },
    { id: 'doughnut-exploded', constant: 'xlDoughnutExploded', value: 80, group: 'Doughnut', label: 'Exploded Doughnut' },
    { id: 'doughnut-ex', constant: 'xlDoughnutEx', value: 131, group: 'Doughnut', label: 'Doughnut (modern)', modern: true },
    { id: 'xy-scatter', constant: 'xlXYScatter', value: -4169, group: 'Scatter', label: 'Scatter' },
    { id: 'xy-scatter-lines', constant: 'xlXYScatterLines', value: 74, group: 'Scatter', label: 'Scatter with Lines' },
    { id: 'xy-scatter-lines-no-markers', constant: 'xlXYScatterLinesNoMarkers', value: 75, group: 'Scatter', label: 'Scatter with Lines, No Markers' },
    { id: 'xy-scatter-smooth', constant: 'xlXYScatterSmooth', value: 72, group: 'Scatter', label: 'Scatter with Smoothed Lines' },
    { id: 'xy-scatter-smooth-no-markers', constant: 'xlXYScatterSmoothNoMarkers', value: 73, group: 'Scatter', label: 'Scatter with Smoothed Lines, No Markers' },
    { id: 'xy-scatter-ex', constant: 'xlXYScatterEx', value: 138, group: 'Scatter', label: 'Scatter (modern)', modern: true },
    { id: 'bubble', constant: 'xlBubble', value: 15, group: 'Bubble', label: 'Bubble' },
    { id: 'bubble-3d', constant: 'xlBubble3DEffect', value: 87, group: 'Bubble', label: 'Bubble with 3D Effects' },
    { id: 'bubble-ex', constant: 'xlBubbleEx', value: 139, group: 'Bubble', label: 'Bubble (modern)', modern: true },
    { id: 'radar', constant: 'xlRadar', value: -4151, group: 'Radar', label: 'Radar' },
    { id: 'radar-markers', constant: 'xlRadarMarkers', value: 81, group: 'Radar', label: 'Radar with Markers' },
    { id: 'radar-filled', constant: 'xlRadarFilled', value: 82, group: 'Radar', label: 'Filled Radar' },
    { id: 'stock-hlc', constant: 'xlStockHLC', value: 88, group: 'Stock', label: 'High-Low-Close' },
    { id: 'stock-ohlc', constant: 'xlStockOHLC', value: 89, group: 'Stock', label: 'Open-High-Low-Close' },
    { id: 'stock-vhlc', constant: 'xlStockVHLC', value: 90, group: 'Stock', label: 'Volume-High-Low-Close' },
    { id: 'stock-vohlc', constant: 'xlStockVOHLC', value: 91, group: 'Stock', label: 'Volume-Open-High-Low-Close' },
    { id: 'surface', constant: 'xlSurface', value: 83, group: 'Surface', label: '3D Surface' },
    { id: 'surface-wireframe', constant: 'xlSurfaceWireframe', value: 84, group: 'Surface', label: '3D Surface Wireframe' },
    { id: 'surface-top', constant: 'xlSurfaceTopView', value: 85, group: 'Surface', label: 'Surface Top View' },
    { id: 'surface-top-wireframe', constant: 'xlSurfaceTopViewWireframe', value: 86, group: 'Surface', label: 'Surface Top View Wireframe' },
    { id: 'combo', constant: 'xlCombo', value: -4152, group: 'Combo', label: 'Custom Combination' },
    { id: 'combo-column-line', constant: 'xlComboColumnClusteredLine', value: 113, group: 'Combo', label: 'Clustered Column and Line' },
    { id: 'combo-column-line-secondary', constant: 'xlComboColumnClusteredLineSecondaryAxis', value: 114, group: 'Combo', label: 'Column and Line on Secondary Axis' },
    { id: 'combo-area-column', constant: 'xlComboAreaStackedColumnClustered', value: 115, group: 'Combo', label: 'Stacked Area and Clustered Column' },
    { id: 'combo-other', constant: 'xlOtherCombinations', value: 116, group: 'Combo', label: 'Other Combination' },
    { id: 'treemap', constant: 'xlTreemap', value: 117, group: 'Modern', label: 'Treemap', modern: true },
    { id: 'histogram', constant: 'xlHistogram', value: 118, group: 'Modern', label: 'Histogram', modern: true },
    { id: 'waterfall', constant: 'xlWaterfall', value: 119, group: 'Modern', label: 'Waterfall', modern: true },
    { id: 'sunburst', constant: 'xlSunburst', value: 120, group: 'Modern', label: 'Sunburst', modern: true },
    { id: 'box-whisker', constant: 'xlBoxwhisker', value: 121, group: 'Modern', label: 'Box and Whisker', modern: true },
    { id: 'pareto', constant: 'xlPareto', value: 122, group: 'Modern', label: 'Pareto', modern: true },
    { id: 'funnel', constant: 'xlFunnel', value: 123, group: 'Modern', label: 'Funnel', modern: true },
    { id: 'region-map', constant: 'xlRegionMap', value: 140, group: 'Modern', label: 'Map', modern: true },
    { id: 'suggested', constant: 'xlSuggestedChart', value: -2, group: 'Recommended', label: 'Recommended by Excel', modern: true },
    { id: 'cone-bar-clustered', constant: 'xlConeBarClustered', value: 102, group: 'Legacy 3D', label: 'Clustered Cone Bar' },
    { id: 'cone-bar-stacked', constant: 'xlConeBarStacked', value: 103, group: 'Legacy 3D', label: 'Stacked Cone Bar' },
    { id: 'cone-bar-stacked-100', constant: 'xlConeBarStacked100', value: 104, group: 'Legacy 3D', label: '100% Stacked Cone Bar' },
    { id: 'cone-column', constant: 'xlConeCol', value: 105, group: 'Legacy 3D', label: '3D Cone Column' },
    { id: 'cone-column-clustered', constant: 'xlConeColClustered', value: 99, group: 'Legacy 3D', label: 'Clustered Cone Column' },
    { id: 'cone-column-stacked', constant: 'xlConeColStacked', value: 100, group: 'Legacy 3D', label: 'Stacked Cone Column' },
    { id: 'cone-column-stacked-100', constant: 'xlConeColStacked100', value: 101, group: 'Legacy 3D', label: '100% Stacked Cone Column' },
    { id: 'cylinder-bar-clustered', constant: 'xlCylinderBarClustered', value: 95, group: 'Legacy 3D', label: 'Clustered Cylinder Bar' },
    { id: 'cylinder-bar-stacked', constant: 'xlCylinderBarStacked', value: 96, group: 'Legacy 3D', label: 'Stacked Cylinder Bar' },
    { id: 'cylinder-bar-stacked-100', constant: 'xlCylinderBarStacked100', value: 97, group: 'Legacy 3D', label: '100% Stacked Cylinder Bar' },
    { id: 'cylinder-column', constant: 'xlCylinderCol', value: 98, group: 'Legacy 3D', label: '3D Cylinder Column' },
    { id: 'cylinder-column-clustered', constant: 'xlCylinderColClustered', value: 92, group: 'Legacy 3D', label: 'Clustered Cylinder Column' },
    { id: 'cylinder-column-stacked', constant: 'xlCylinderColStacked', value: 93, group: 'Legacy 3D', label: 'Stacked Cylinder Column' },
    { id: 'cylinder-column-stacked-100', constant: 'xlCylinderColStacked100', value: 94, group: 'Legacy 3D', label: '100% Stacked Cylinder Column' },
    { id: 'pyramid-bar-clustered', constant: 'xlPyramidBarClustered', value: 109, group: 'Legacy 3D', label: 'Clustered Pyramid Bar' },
    { id: 'pyramid-bar-stacked', constant: 'xlPyramidBarStacked', value: 110, group: 'Legacy 3D', label: 'Stacked Pyramid Bar' },
    { id: 'pyramid-bar-stacked-100', constant: 'xlPyramidBarStacked100', value: 111, group: 'Legacy 3D', label: '100% Stacked Pyramid Bar' },
    { id: 'pyramid-column', constant: 'xlPyramidCol', value: 112, group: 'Legacy 3D', label: '3D Pyramid Column' },
    { id: 'pyramid-column-clustered', constant: 'xlPyramidColClustered', value: 106, group: 'Legacy 3D', label: 'Clustered Pyramid Column' },
    { id: 'pyramid-column-stacked', constant: 'xlPyramidColStacked', value: 107, group: 'Legacy 3D', label: 'Stacked Pyramid Column' },
    { id: 'pyramid-column-stacked-100', constant: 'xlPyramidColStacked100', value: 108, group: 'Legacy 3D', label: '100% Stacked Pyramid Column' },
] as const;

const AXISLESS_CHART_TYPES = new Set([
    -4120, -4102,
    5, 68, 69, 70, 71, 80,
    117, 120, 123, 130, 131, 140,
]);
const OVERLAP_CHART_TYPES = new Set([
    51, 52, 53, 57, 58, 59,
    124, 125, 126, 132, 133, 134,
    -4152, 113, 114, 115, 116,
]);
const SECONDARY_AXIS_CHART_TYPES = new Set([
    1, 76, 77, 135, 136, 137,
    57, 58, 59, 132, 133, 134,
    51, 52, 53, 124, 125, 126,
    4, 63, 64, 65, 66, 67, 127, 128, 129,
    -4169, 72, 73, 74, 75, 138,
    15, 87, 139,
    -4152, 113, 114, 115, 116,
]);
const BUBBLE_CHART_TYPES = new Set([15, 87, 139]);

const DATA_LABEL_POSITIONS_BY_CHART_TYPE = new Map<number, readonly ChartDataLabelPosition[]>([
    // Excel COM only exposes the four end/base positions on clustered columns
    // and bars. Stacked variants reject outsideEnd.
    [51, ['center', 'insideBase', 'insideEnd', 'outsideEnd']],
    [57, ['center', 'insideBase', 'insideEnd', 'outsideEnd']],
    [52, ['center', 'insideBase', 'insideEnd']],
    [53, ['center', 'insideBase', 'insideEnd']],
    [58, ['center', 'insideBase', 'insideEnd']],
    [59, ['center', 'insideBase', 'insideEnd']],

    // Two-dimensional line and XY families expose point-relative positions.
    [4, ['above', 'below', 'center', 'left', 'right']],
    [63, ['above', 'below', 'center', 'left', 'right']],
    [64, ['above', 'below', 'center', 'left', 'right']],
    [65, ['above', 'below', 'center', 'left', 'right']],
    [66, ['above', 'below', 'center', 'left', 'right']],
    [67, ['above', 'below', 'center', 'left', 'right']],
    [-4169, ['above', 'below', 'center', 'left', 'right']],
    [72, ['above', 'below', 'center', 'left', 'right']],
    [73, ['above', 'below', 'center', 'left', 'right']],
    [74, ['above', 'below', 'center', 'left', 'right']],
    [75, ['above', 'below', 'center', 'left', 'right']],
    [15, ['above', 'below', 'center', 'left', 'right']],
    [87, ['above', 'below', 'center', 'left', 'right']],

    // Pie labels support fit/end positions, including the 3D and split-pie
    // variants. Doughnut and the remaining families expose an automatic
    // position through COM but reject attempts to assign DataLabels.Position.
    [5, ['bestFit', 'center', 'insideEnd', 'outsideEnd']],
    [68, ['bestFit', 'center', 'insideEnd', 'outsideEnd']],
    [69, ['bestFit', 'center', 'insideEnd', 'outsideEnd']],
    [70, ['bestFit', 'center', 'insideEnd', 'outsideEnd']],
    [71, ['bestFit', 'center', 'insideEnd', 'outsideEnd']],
    [-4102, ['bestFit', 'center', 'insideEnd', 'outsideEnd']],
]);

export function chartTypeSupportsAxes(chartType: number): boolean {
    return !AXISLESS_CHART_TYPES.has(chartType);
}

export function chartTypeSupportsSecondaryAxes(
    chartType: number,
    seriesChartTypes: readonly number[] = [],
): boolean {
    const candidates = seriesChartTypes.length ? seriesChartTypes : [chartType];
    return candidates.some(candidate => SECONDARY_AXIS_CHART_TYPES.has(candidate));
}

export function isChartTypeCreatable(chartType: number): boolean {
    // xlSuggestedChart asks Excel to recommend a chart; it is not a stable
    // concrete ChartType that can be persisted transactionally.
    // xlRegionMap can contact Bing Maps while Excel creates the chart. Native
    // automation remains offline-by-default and therefore inventories but does
    // not create that type without a future explicit network-consent workflow.
    return chartType !== -2 && chartType !== 140;
}

export function chartTypeSupportsGapWidth(
    chartType: number,
    seriesChartTypes: readonly number[] = [],
): boolean {
    const candidates = seriesChartTypes.length ? seriesChartTypes : [chartType];
    return candidates.some(candidate => {
        const group = EXCEL_CHART_TYPES.find(option => option.value === candidate)?.group;
        return group === 'Bar'
            || group === 'Column'
            || group === 'Combo'
            || [68, 71, 118, 122].includes(candidate);
    });
}

export function chartSeriesSupportsBubbleSizes(chartType: number): boolean {
    return BUBBLE_CHART_TYPES.has(chartType);
}

/** Excel silently promotes every series to Bubble when bubble and non-bubble types are mixed. */
export function chartSeriesTypesCanCoexist(chartTypes: readonly number[]): boolean {
    const bubbleSeriesCount = chartTypes.filter(chartSeriesSupportsBubbleSizes).length;
    return bubbleSeriesCount === 0 || bubbleSeriesCount === chartTypes.length;
}

/**
 * Excel derives a chart's top-level type from its concrete series. A single
 * effective type must therefore become the top-level type; only genuinely
 * heterogeneous series persist as an xlCombo chart.
 */
export function canonicalChartTypeForSeries(
    requestedChartType: number,
    seriesChartTypes: readonly number[],
): number {
    if (seriesChartTypes.length === 0) return requestedChartType;
    const distinctTypes = new Set(seriesChartTypes);
    return distinctTypes.size === 1
        ? seriesChartTypes[0]
        : -4152;
}

/** Series.Smooth is exposed by Excel only for line and XY-scatter series. */
export function chartSeriesSupportsSmooth(chartType: number): boolean {
    const group = EXCEL_CHART_TYPES.find(option => option.value === chartType)?.group;
    return group === 'Line' || group === 'Scatter';
}

/** Percentage data labels are meaningful only for pie and doughnut series. */
export function chartSeriesSupportsPercentageDataLabels(chartType: number): boolean {
    const group = EXCEL_CHART_TYPES.find(option => option.value === chartType)?.group;
    return group === 'Pie' || group === 'Doughnut';
}

/** Positions that Excel COM accepts for DataLabels.Position on this series type. */
export function chartSeriesSupportedDataLabelPositions(
    chartType: number,
): readonly ChartDataLabelPosition[] {
    return DATA_LABEL_POSITIONS_BY_CHART_TYPE.get(chartType) ?? [];
}

export function chartSeriesSupportsDataLabelPosition(
    chartType: number,
    position: ChartDataLabelPosition,
): boolean {
    return chartSeriesSupportedDataLabelPositions(chartType).includes(position);
}

export function chartAxisGroupSupportsCategoryScale(
    chartType: number,
    series: readonly SheetChartSeriesData[],
    axisGroup: ChartAxisGroup,
): boolean {
    if (series.length === 0) {
		const group = EXCEL_CHART_TYPES.find(option => option.value === chartType)?.group;
		return axisGroup === 'primary' && (group === 'Scatter' || group === 'Bubble');
    }
    return series
        .filter(item => (item.axisGroup ?? 'primary') === axisGroup)
        .some(item => {
            const effectiveType = item.chartType ?? chartType;
            const group = EXCEL_CHART_TYPES.find(option => option.value === effectiveType)?.group;
            return group === 'Scatter' || group === 'Bubble';
        });
}

export function chartTypeSupportsOverlap(
    chartType: number,
    seriesChartTypes: readonly number[] = [],
): boolean {
    const candidates = seriesChartTypes.length ? seriesChartTypes : [chartType];
    return candidates.some(candidate => OVERLAP_CHART_TYPES.has(candidate));
}

export function buildExcelTableStyleCatalog(): string[] {
    return [
        ...Array.from({ length: 21 }, (_, index) => `TableStyleLight${index + 1}`),
        ...Array.from({ length: 28 }, (_, index) => `TableStyleMedium${index + 1}`),
        ...Array.from({ length: 11 }, (_, index) => `TableStyleDark${index + 1}`),
    ];
}

/** A ListObject range always contains its internal header row and one data row. */
export function minimumExcelTableRangeRows(totalsRow: boolean): number {
    return totalsRow ? 3 : 2;
}

export const SIMPLE_A1_RANGE = /^\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6})?$/;
export const MAX_WORKBOOK_OBJECT_RANGE_CELLS = 1_000_000;
export const MAX_WORKBOOK_OBJECT_TRANSACTION_CELLS = 5_000_000;
const EXCEL_TABLE_NAME = /^[\p{L}_\\][\p{L}\p{N}._]{0,254}$/u;

export interface SimpleA1RangeBounds {
    startColumn: number;
    startRow: number;
    endColumn: number;
    endRow: number;
    cellCount: number;
}

export function normalizeA1Range(rangeRef: string): string {
    return rangeRef.replace(/\$/g, '').trim().toUpperCase();
}

function excelColumnNumber(letters: string): number {
    let result = 0;
    for (const letter of letters) result = result * 26 + letter.charCodeAt(0) - 64;
    return result;
}

/**
 * Excel table names follow the workbook-name identifier rules: a Unicode
 * letter, underscore or backslash first, then Unicode letters/numbers, dots or
 * underscores, with a 255-character limit. Actual A1/R1C1 cell references and
 * the reserved single-letter R/C shortcuts are not names.
 */
export function normalizeExcelTableName(value: string): string {
    return value.trim().normalize('NFC');
}

export function excelTableNameComparisonKey(value: string): string {
    return normalizeExcelTableName(value).toLocaleLowerCase('en-US');
}

export function isValidExcelTableName(value: string): boolean {
    const name = normalizeExcelTableName(value);
    if (!EXCEL_TABLE_NAME.test(name) || /^(?:R|C)$/i.test(name)) return false;

    const a1Match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/i.exec(name);
    if (
        a1Match
        && excelColumnNumber(a1Match[1].toUpperCase()) <= 16_384
        && Number(a1Match[2]) <= 1_048_576
    ) return false;

    const r1c1Match = /^R([1-9][0-9]{0,6})C([1-9][0-9]{0,4})$/i.exec(name);
    return !(
        r1c1Match
        && Number(r1c1Match[1]) <= 1_048_576
        && Number(r1c1Match[2]) <= 16_384
    );
}

export function parseSimpleA1Range(rangeRef: string): SimpleA1RangeBounds | null {
    const normalized = normalizeA1Range(rangeRef);
    const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/.exec(normalized);
    if (!match) return null;
    const startColumn = excelColumnNumber(match[1]);
    const startRow = Number(match[2]);
    const endColumn = excelColumnNumber(match[3] ?? match[1]);
    const endRow = Number(match[4] ?? match[2]);
    if (
        startColumn < 1 || endColumn > 16_384 || startColumn > endColumn
        || startRow < 1 || endRow > 1_048_576 || startRow > endRow
    ) return null;
    return {
        startColumn,
        startRow,
        endColumn,
        endRow,
        cellCount: (endColumn - startColumn + 1) * (endRow - startRow + 1),
    };
}

export function simpleA1RangesOverlap(first: string, second: string): boolean {
    const left = parseSimpleA1Range(first);
    const right = parseSimpleA1Range(second);
    if (!left || !right) return false;
    return left.startRow <= right.endRow
        && right.startRow <= left.endRow
        && left.startColumn <= right.endColumn
        && right.startColumn <= left.endColumn;
}
