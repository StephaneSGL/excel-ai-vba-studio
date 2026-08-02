import * as path from 'path';
import * as vscode from 'vscode';
import {
	EXCEL_CHART_TYPES,
	MAX_WORKBOOK_OBJECT_RANGE_CELLS,
	MAX_WORKBOOK_OBJECT_TRANSACTION_CELLS,
	SIMPLE_A1_RANGE,
	buildExcelTableStyleCatalog,
	canonicalChartTypeForSeries,
	chartAxisGroupSupportsCategoryScale,
	chartDataLabelsHaveEnabledShowOption,
	chartDataLabelsHaveExplicitShowOption,
	chartTypeSupportsAxes,
	chartTypeSupportsSecondaryAxes,
	chartTypeSupportsGapWidth,
	chartTypeSupportsOverlap,
	chartSeriesSupportsBubbleSizes,
	chartSeriesSupportsDataLabelPosition,
	chartSeriesSupportsPercentageDataLabels,
	chartSeriesSupportsSmooth,
	chartSeriesTypesCanCoexist,
	excelTableNameComparisonKey,
	isChartTypeCreatable,
	isValidExcelTableName,
	minimumExcelTableRangeRows,
	normalizeA1Range,
	normalizeExcelTableName,
	parseSimpleA1Range,
	simpleA1RangesOverlap,
	type ChartAxisGroup,
	type ChartLegendPosition,
	type ChartMarkerStyle,
	type SheetChartAxisData,
	type SheetChartData,
	type SheetChartSeriesData,
	type SheetTableData
} from '../common/excelWorkbookObjects';
import {
	EXCEL_AI_WORKBOOK_DESIGN_TOOL,
	type WorkbookObjectDesignOperation,
	type WorkbookObjectDesignToolInput
} from './types';
import type { ExcelAiVbaWorkbookService } from './workbookService';
import {
	assertNoReparsePointChain,
	canonicalizeWorkbookUri,
	pathIsInside
} from './security';

const MAX_OPERATIONS = 100;
const MAX_SERIES = 255;
const MAX_TEXT = 255;
const MAX_POINT = 100_000;
const COLOR = /^#[0-9a-f]{6}$/i;
const TABLE_STYLES = new Set(buildExcelTableStyleCatalog());
const CHART_TYPES = new Set(EXCEL_CHART_TYPES.map(option => option.value));
const LEGEND_POSITIONS = new Set<ChartLegendPosition>([
	'bottom', 'corner', 'custom', 'left', 'right', 'top'
]);
const MARKER_STYLES = new Set<ChartMarkerStyle>([
	'automatic', 'circle', 'dash', 'diamond', 'dot', 'none', 'picture',
	'plus', 'square', 'star', 'triangle', 'x'
]);

interface LanguageModelApi {
	registerTool?: (name: string, tool: unknown) => vscode.Disposable;
}

interface LanguageModelConstructors {
	LanguageModelToolResult?: new (parts: unknown[]) => unknown;
	LanguageModelTextPart?: new (value: string) => unknown;
}

function objectValue(value: unknown, label: string): Record<string, unknown> {
	if (!value || typeof value !== 'object' || Array.isArray(value)) {
		throw new Error(`${label} doit être un objet.`);
	}
	return value as Record<string, unknown>;
}

function onlyKeys(
	value: Record<string, unknown>,
	allowed: readonly string[],
	label: string
): void {
	const expected = new Set(allowed);
	const unexpected = Object.keys(value).find(key => !expected.has(key));
	if (unexpected) {
		throw new Error(`${label}.${unexpected} n’est pas une propriété autorisée.`);
	}
}

function textValue(
	value: unknown,
	label: string,
	options: { required?: boolean; max?: number } = {}
): string | undefined {
	if (value === undefined) {
		if (options.required) throw new Error(`${label} est obligatoire.`);
		return undefined;
	}
	if (typeof value !== 'string' || value.includes('\0')) {
		throw new Error(`${label} doit être une chaîne sans caractère NUL.`);
	}
	const trimmed = value.trim();
	if (options.required && !trimmed) throw new Error(`${label} ne peut pas être vide.`);
	if (trimmed.length > (options.max ?? MAX_TEXT)) {
		throw new Error(`${label} dépasse ${(options.max ?? MAX_TEXT)} caractères.`);
	}
	return trimmed;
}

function booleanValue(value: unknown, label: string, fallback: boolean): boolean {
	if (value === undefined) return fallback;
	if (typeof value !== 'boolean') throw new Error(`${label} doit être un booléen.`);
	return value;
}

function numberValue(
	value: unknown,
	label: string,
	minimum: number,
	maximum: number,
	options: { required?: boolean; integer?: boolean } = {}
): number | undefined {
	if (value === undefined && !options.required) return undefined;
	if (
		typeof value !== 'number' || !Number.isFinite(value) || value < minimum ||
		value > maximum || (options.integer && !Number.isInteger(value))
	) {
		throw new Error(`${label} doit être un nombre valide entre ${minimum} et ${maximum}.`);
	}
	return value;
}

function sheetName(value: unknown, label: string): string {
	const parsed = textValue(value, label, { required: true, max: 31 }) as string;
	if (/[:\\/?*\[\]]/.test(parsed) || /^'|'$/.test(parsed)) {
		throw new Error(`${label} n’est pas un nom de feuille Excel valide.`);
	}
	return parsed;
}

function objectName(value: unknown, label: string): string {
	return textValue(value, label, { required: true }) as string;
}

function tableName(value: unknown, label: string): string {
	const parsed = normalizeExcelTableName(
		textValue(value, label, { required: true }) as string
	);
	if (!isValidExcelTableName(parsed)) {
		throw new Error(`${label} n’est pas un nom de tableau Excel valide.`);
	}
	return parsed;
}

function rangeValue(value: unknown, label: string): string {
	const parsed = normalizeA1Range(textValue(value, label, { required: true, max: 32 }) as string);
	if (!SIMPLE_A1_RANGE.test(parsed) || !parseSimpleA1Range(parsed)) {
		throw new Error(`${label} doit être une plage A1 locale simple dans les limites d’Excel.`);
	}
	if ((parseSimpleA1Range(parsed)?.cellCount ?? Infinity) > MAX_WORKBOOK_OBJECT_RANGE_CELLS) {
		throw new Error(`${label} dépasse la limite de ${MAX_WORKBOOK_OBJECT_RANGE_CELLS.toLocaleString('fr-FR')} cellules.`);
	}
	return parsed;
}

function optionalRange(value: unknown, label: string): string | undefined {
	return value === undefined ? undefined : rangeValue(value, label);
}

function colorValue(value: unknown, label: string): string | undefined {
	if (value === undefined) return undefined;
	if (typeof value !== 'string' || !COLOR.test(value)) {
		throw new Error(`${label} doit être une couleur #RRGGBB.`);
	}
	return value.toLowerCase();
}

function parseTable(value: unknown, label: string): SheetTableData {
	const source = objectValue(value, label);
	onlyKeys(source, [
		'id', 'name', 'displayName', 'rangeRef', 'headerRow', 'totalsRow', 'style'
	], label);
	const name = tableName(source.name, `${label}.name`);
	const displayName = source.displayName === undefined
		? name
		: tableName(source.displayName, `${label}.displayName`);
	const styleSource = source.style === undefined
		? {}
		: objectValue(source.style, `${label}.style`);
	onlyKeys(styleSource, [
		'name', 'showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes'
	], `${label}.style`);
	const styleName = textValue(styleSource.name, `${label}.style.name`) ?? 'TableStyleMedium2';
	if (!TABLE_STYLES.has(styleName)) {
		throw new Error(`${label}.style.name n’est pas un style de tableau Excel intégré.`);
	}
	const rangeRef = rangeValue(source.rangeRef, `${label}.rangeRef`);
	const headerRow = booleanValue(source.headerRow, `${label}.headerRow`, true);
	const totalsRow = booleanValue(source.totalsRow, `${label}.totalsRow`, false);
	const rangeBounds = parseSimpleA1Range(rangeRef) as NonNullable<ReturnType<typeof parseSimpleA1Range>>;
	if (rangeBounds.endRow - rangeBounds.startRow + 1 < minimumExcelTableRangeRows(totalsRow)) {
		throw new Error(`${label}.rangeRef ne contient pas assez de lignes pour l’en-tête, les données${totalsRow ? ' et les totaux' : ''}.`);
	}
	return {
		id: textValue(source.id, `${label}.id`) ?? `table:${excelTableNameComparisonKey(name)}`,
		name,
		displayName,
		rangeRef,
		headerRow,
		totalsRow,
		style: {
			name: styleName,
			showFirstColumn: booleanValue(styleSource.showFirstColumn, `${label}.style.showFirstColumn`, false),
			showLastColumn: booleanValue(styleSource.showLastColumn, `${label}.style.showLastColumn`, false),
			showRowStripes: booleanValue(styleSource.showRowStripes, `${label}.style.showRowStripes`, true),
			showColumnStripes: booleanValue(styleSource.showColumnStripes, `${label}.style.showColumnStripes`, false)
		}
	};
}

function parseAxis(value: unknown, label: string): SheetChartAxisData | undefined {
	if (value === undefined) return undefined;
	const source = objectValue(value, label);
	onlyKeys(source, [
		'visible', 'title', 'minimumScale', 'maximumScale', 'majorUnit', 'minorUnit',
		'logarithmic', 'reverseOrder', 'numberFormat', 'majorGridlines', 'minorGridlines'
	], label);
	const scale = (property: 'minimumScale' | 'maximumScale' | 'majorUnit' | 'minorUnit') => {
		if (source[property] === null) return null;
		return numberValue(source[property], `${label}.${property}`, -1e307, 1e307);
	};
	const result: SheetChartAxisData = {};
	for (const property of [
		'visible', 'logarithmic', 'reverseOrder', 'majorGridlines', 'minorGridlines'
	] as const) {
		if (source[property] !== undefined) {
			result[property] = booleanValue(source[property], `${label}.${property}`, false);
		}
	}
	const title = textValue(source.title, `${label}.title`, { max: 1_000 });
	const numberFormat = textValue(source.numberFormat, `${label}.numberFormat`, { max: 255 });
	if (title !== undefined) result.title = title;
	if (numberFormat !== undefined) result.numberFormat = numberFormat;
	for (const property of ['minimumScale', 'maximumScale', 'majorUnit', 'minorUnit'] as const) {
		const parsed = scale(property);
		if (parsed !== undefined) result[property] = parsed;
	}
	if (
		typeof result.minimumScale === 'number' &&
		typeof result.maximumScale === 'number' &&
		result.minimumScale >= result.maximumScale
	) {
		throw new Error(`${label}.minimumScale doit être inférieur à maximumScale.`);
	}
	for (const property of ['majorUnit', 'minorUnit'] as const) {
		if (typeof result[property] === 'number' && result[property] <= 0) {
			throw new Error(`${label}.${property} doit être strictement positif.`);
		}
	}
	if (result.logarithmic && (
		(typeof result.minimumScale === 'number' && result.minimumScale <= 0) ||
		(typeof result.maximumScale === 'number' && result.maximumScale <= 0)
	)) {
		throw new Error(`${label} ne peut pas utiliser une échelle logarithmique avec une borne non positive.`);
	}
	return result;
}

function parseDataLabels(value: unknown, label: string): SheetChartSeriesData['dataLabels'] {
	if (value === undefined) return undefined;
	const source = objectValue(value, label);
	onlyKeys(source, [
		'showValue', 'showCategoryName', 'showSeriesName', 'showPercentage',
		'showBubbleSize', 'position'
	], label);
	if (!chartDataLabelsHaveExplicitShowOption(source as SheetChartSeriesData['dataLabels'])) {
		throw new Error(`${label} doit définir explicitement au moins une option d’affichage.`);
	}
	const positions = new Set([
		'above', 'below', 'bestFit', 'center', 'insideBase', 'insideEnd', 'left',
		'outsideEnd', 'right'
	]);
	const position = textValue(source.position, `${label}.position`);
	if (position !== undefined && !positions.has(position)) {
		throw new Error(`${label}.position n’est pas prise en charge.`);
	}
	const result: NonNullable<SheetChartSeriesData['dataLabels']> = {};
	for (const property of [
		'showValue', 'showCategoryName', 'showSeriesName', 'showPercentage', 'showBubbleSize'
	] as const) {
		if (source[property] !== undefined) {
			result[property] = booleanValue(source[property], `${label}.${property}`, false);
		}
	}
	if (position !== undefined) {
		result.position = position as NonNullable<SheetChartSeriesData['dataLabels']>['position'];
	}
	if (result.position !== undefined && !chartDataLabelsHaveEnabledShowOption(result)) {
		throw new Error(`${label}.position exige au moins une option d’affichage activée.`);
	}
	return result;
}

function parseSeries(value: unknown, label: string): SheetChartSeriesData {
	const source = objectValue(value, label);
	onlyKeys(source, [
		'id', 'name', 'nameRange', 'categoryRange', 'valuesRange', 'xValuesRange',
		'bubbleSizesRange', 'chartType', 'axisGroup', 'color', 'lineColor',
		'lineWidth', 'dashStyle', 'markerStyle', 'markerSize', 'smooth', 'visible',
		'dataLabels'
	], label);
	const chartType = source.chartType === undefined
		? undefined
		: numberValue(source.chartType, `${label}.chartType`, -10_000, 10_000, { required: true, integer: true }) as number;
	if (chartType !== undefined && !CHART_TYPES.has(chartType)) throw new Error(`${label}.chartType n’est pas un type XlChartType publié.`);
	if (chartType !== undefined && !isChartTypeCreatable(chartType)) {
		throw new Error(`${label}.chartType n’est pas autorisé pour une création locale sans accès réseau.`);
	}
	const axisGroup = textValue(source.axisGroup, `${label}.axisGroup`);
	if (axisGroup !== undefined && axisGroup !== 'primary' && axisGroup !== 'secondary') {
		throw new Error(`${label}.axisGroup doit être primary ou secondary.`);
	}
	const dashStyle = textValue(source.dashStyle, `${label}.dashStyle`);
	if (dashStyle && !['solid', 'dash', 'dot', 'dashDot'].includes(dashStyle)) {
		throw new Error(`${label}.dashStyle n’est pas pris en charge.`);
	}
	const markerStyle = textValue(source.markerStyle, `${label}.markerStyle`);
	if (markerStyle && !MARKER_STYLES.has(markerStyle as ChartMarkerStyle)) {
		throw new Error(`${label}.markerStyle n’est pas pris en charge.`);
	}
	const result: SheetChartSeriesData = {
		id: textValue(source.id, `${label}.id`) ?? `series:${label}`,
		valuesRange: rangeValue(source.valuesRange, `${label}.valuesRange`)
	};
	if (chartType !== undefined) result.chartType = chartType;
	if (axisGroup !== undefined) result.axisGroup = axisGroup as ChartAxisGroup;
	if (source.visible !== undefined) result.visible = booleanValue(source.visible, `${label}.visible`, true);
	const seriesName = textValue(source.name, `${label}.name`);
	if (seriesName && /^[=+\-@]/.test(seriesName.trim())) {
		throw new Error(`${label}.name doit être du texte littéral et ne peut pas être une formule.`);
	}
	if (seriesName !== undefined) result.name = seriesName;
	for (const property of ['nameRange', 'categoryRange', 'xValuesRange', 'bubbleSizesRange'] as const) {
		const parsed = optionalRange(source[property], `${label}.${property}`);
		if (parsed !== undefined) result[property] = parsed;
	}
	if (result.name && result.nameRange) {
		throw new Error(`${label} ne peut pas définir name et nameRange simultanément.`);
	}
	if (result.nameRange && parseSimpleA1Range(result.nameRange)?.cellCount !== 1) {
		throw new Error(`${label}.nameRange doit désigner exactement une cellule.`);
	}
	if (result.categoryRange && result.xValuesRange) {
		throw new Error(`${label} ne peut pas définir categoryRange et xValuesRange simultanément.`);
	}
	for (const property of ['color', 'lineColor'] as const) {
		const parsed = colorValue(source[property], `${label}.${property}`);
		if (parsed !== undefined) result[property] = parsed;
	}
	const lineWidth = numberValue(source.lineWidth, `${label}.lineWidth`, 0.1, 20);
	const markerSize = numberValue(source.markerSize, `${label}.markerSize`, 2, 72, { integer: true });
	if (lineWidth !== undefined) result.lineWidth = lineWidth;
	if (markerSize !== undefined) result.markerSize = markerSize;
	if (dashStyle) result.dashStyle = dashStyle as SheetChartSeriesData['dashStyle'];
	if (markerStyle) result.markerStyle = markerStyle as ChartMarkerStyle;
	if (source.smooth !== undefined) result.smooth = booleanValue(source.smooth, `${label}.smooth`, false);
	const dataLabels = parseDataLabels(source.dataLabels, `${label}.dataLabels`);
	if (dataLabels) result.dataLabels = dataLabels;
	return result;
}

function parseChart(value: unknown, label: string): SheetChartData {
	const source = objectValue(value, label);
	onlyKeys(source, [
		'id', 'name', 'chartType', 'sourceRangeRef', 'plotBy', 'anchor', 'title',
		'legend', 'categoryAxis', 'valueAxis', 'secondaryCategoryAxis',
		'secondaryValueAxis', 'series', 'style', 'roundedCorners', 'gapWidth',
		'overlap', 'alternativeText'
	], label);
	const name = objectName(source.name, `${label}.name`);
	const chartType = numberValue(source.chartType, `${label}.chartType`, -10_000, 10_000, {
		required: true,
		integer: true
	}) as number;
	if (!CHART_TYPES.has(chartType)) throw new Error(`${label}.chartType n’est pas un type XlChartType publié.`);
	if (!isChartTypeCreatable(chartType)) throw new Error(`${label}.chartType n’est pas autorisé pour une création locale sans accès réseau.`);
	const plotBy = textValue(source.plotBy, `${label}.plotBy`) ?? 'columns';
	if (plotBy !== 'columns' && plotBy !== 'rows') throw new Error(`${label}.plotBy doit être rows ou columns.`);
	const anchor = objectValue(source.anchor, `${label}.anchor`);
	onlyKeys(anchor, ['left', 'top', 'width', 'height'], `${label}.anchor`);
	const titleSource = source.title === undefined ? undefined : objectValue(source.title, `${label}.title`);
	if (titleSource) onlyKeys(titleSource, ['visible', 'text'], `${label}.title`);
	const legendSource = source.legend === undefined ? undefined : objectValue(source.legend, `${label}.legend`);
	if (legendSource) onlyKeys(legendSource, ['visible', 'position'], `${label}.legend`);
	const legendPosition = legendSource
		? textValue(legendSource.position, `${label}.legend.position`) ?? 'right'
		: 'right';
	if (!LEGEND_POSITIONS.has(legendPosition as ChartLegendPosition)) {
		throw new Error(`${label}.legend.position n’est pas prise en charge.`);
	}
	let series: SheetChartSeriesData[] | undefined;
	if (source.series !== undefined) {
		if (!Array.isArray(source.series) || source.series.length > MAX_SERIES) {
			throw new Error(`${label}.series doit contenir au maximum ${MAX_SERIES} séries.`);
		}
		series = source.series.map((item, index) => parseSeries(item, `${label}.series[${index}]`));
		const seriesIds = new Set<string>();
		for (const item of series) {
			if (seriesIds.has(item.id)) {
				throw new Error(`${label}.series contient plusieurs fois l’identifiant ${item.id}.`);
			}
			seriesIds.add(item.id);
		}
	}
	const sourceRangeRef = optionalRange(source.sourceRangeRef, `${label}.sourceRangeRef`);
	if (sourceRangeRef && series?.length) {
		throw new Error(`${label} ne peut pas définir sourceRangeRef et series simultanément.`);
	}
	if (!sourceRangeRef && (!series || series.length === 0)) {
		throw new Error(`${label} exige sourceRangeRef ou au moins une série.`);
	}
	const seriesTypes = series?.map(item => item.chartType ?? chartType) ?? [];
	if (!chartSeriesTypesCanCoexist(seriesTypes)) {
		throw new Error(`${label} ne peut pas mélanger des séries à bulles et non-bulles, car Excel les convertirait silencieusement en bulles.`);
	}
	const effectiveChartType = canonicalChartTypeForSeries(chartType, seriesTypes);
	if (effectiveChartType === -4152 && chartType !== -4152 && series) {
		series = series.map(item => ({ ...item, chartType: item.chartType ?? chartType }));
	}
	for (const [index, item] of (series ?? []).entries()) {
		const effectiveSeriesType = item.chartType ?? chartType;
		if (item.bubbleSizesRange && !chartSeriesSupportsBubbleSizes(effectiveSeriesType)) {
			throw new Error(`${label}.series[${index}].bubbleSizesRange exige une série de type bulle.`);
		}
		if (item.smooth !== undefined && !chartSeriesSupportsSmooth(effectiveSeriesType)) {
			throw new Error(`${label}.series[${index}].smooth exige une série ligne ou nuage de points.`);
		}
		if (item.dataLabels?.showBubbleSize === true && !chartSeriesSupportsBubbleSizes(effectiveSeriesType)) {
			throw new Error(`${label}.series[${index}].dataLabels.showBubbleSize exige une série de type bulle.`);
		}
		if (
			item.dataLabels?.showPercentage === true
			&& !chartSeriesSupportsPercentageDataLabels(effectiveSeriesType)
		) {
			throw new Error(`${label}.series[${index}].dataLabels.showPercentage exige une série de type secteur ou anneau.`);
		}
		if (
			item.dataLabels?.position !== undefined
			&& !chartSeriesSupportsDataLabelPosition(
				effectiveSeriesType,
				item.dataLabels.position,
			)
		) {
			throw new Error(`${label}.series[${index}].dataLabels.position n’est pas prise en charge par ce type de graphique.`);
		}
	}
	const result: SheetChartData = {
		id: textValue(source.id, `${label}.id`) ?? `chart:${name.toLocaleLowerCase('en-US')}`,
		name,
		chartType: effectiveChartType,
		plotBy,
		anchor: {
			left: numberValue(anchor.left, `${label}.anchor.left`, 0, MAX_POINT, { required: true }) as number,
			top: numberValue(anchor.top, `${label}.anchor.top`, 0, MAX_POINT, { required: true }) as number,
			width: numberValue(anchor.width, `${label}.anchor.width`, 20, MAX_POINT, { required: true }) as number,
			height: numberValue(anchor.height, `${label}.anchor.height`, 20, MAX_POINT, { required: true }) as number
		}
	};
	if (sourceRangeRef) result.sourceRangeRef = sourceRangeRef;
	if (titleSource) {
		result.title = {
			visible: booleanValue(titleSource.visible, `${label}.title.visible`, true),
			text: textValue(titleSource.text, `${label}.title.text`) ?? ''
		};
	}
	if (legendSource) {
		result.legend = {
			visible: booleanValue(legendSource.visible, `${label}.legend.visible`, true),
			position: legendPosition as ChartLegendPosition
		};
	}
	for (const property of ['categoryAxis', 'valueAxis', 'secondaryCategoryAxis', 'secondaryValueAxis'] as const) {
		const axis = parseAxis(source[property], `${label}.${property}`);
		if (axis) result[property] = axis;
	}
	if (!chartTypeSupportsAxes(effectiveChartType) && [
		result.categoryAxis,
		result.valueAxis,
		result.secondaryCategoryAxis,
		result.secondaryValueAxis
	].some(axis => axis !== undefined)) {
		throw new Error(`${label} ne peut pas définir d’axes pour ce type de graphique.`);
	}
	for (const property of ['categoryAxis', 'secondaryCategoryAxis'] as const) {
		const axis = result[property];
		const axisGroup: ChartAxisGroup = property === 'categoryAxis' ? 'primary' : 'secondary';
		if (!chartAxisGroupSupportsCategoryScale(effectiveChartType, series ?? [], axisGroup) && axis && [
			axis.minimumScale,
			axis.maximumScale,
			axis.majorUnit,
			axis.minorUnit,
			axis.logarithmic
		].some(value => value !== undefined)) {
			throw new Error(`${label}.${property} ne peut pas définir d’échelle de valeurs.`);
		}
	}
	for (const property of ['secondaryCategoryAxis', 'secondaryValueAxis'] as const) {
		const axis = result[property];
		if (axis && (axis.majorGridlines !== undefined || axis.minorGridlines !== undefined)) {
			throw new Error(`${label}.${property} ne peut pas définir de quadrillage secondaire.`);
		}
	}
	if (!chartTypeSupportsSecondaryAxes(effectiveChartType, seriesTypes) && [
		result.secondaryCategoryAxis,
		result.secondaryValueAxis
	].some(axis => axis !== undefined)) {
		throw new Error(`${label} ne peut pas définir d’axes secondaires pour ce type de graphique.`);
	}
	if (series?.some(item => (
		item.axisGroup === 'secondary'
		&& !chartTypeSupportsSecondaryAxes(item.chartType ?? chartType)
	))) {
		throw new Error(`${label} ne peut pas affecter de série à un axe secondaire pour ce type de graphique.`);
	}
	if (series?.length) result.series = series;
	if (effectiveChartType === -4152 && (
		!series?.length || series.some(item =>
			item.chartType === undefined ||
			item.chartType === -4152 ||
			!isChartTypeCreatable(item.chartType)
		) || new Set(series.map(item => item.chartType)).size < 2
	)) {
		throw new Error(`${label} exige au moins deux séries de types concrets distincts pour un graphique combiné personnalisé.`);
	}
	const style = numberValue(source.style, `${label}.style`, 1, 48, { integer: true });
	const gapWidth = numberValue(source.gapWidth, `${label}.gapWidth`, 0, 500, { integer: true });
	const overlap = numberValue(source.overlap, `${label}.overlap`, -100, 100, { integer: true });
	if (style !== undefined) result.style = style;
	if (gapWidth !== undefined) {
		if (!chartTypeSupportsGapWidth(effectiveChartType, seriesTypes)) throw new Error(`${label}.gapWidth n’est pas compatible avec ce type.`);
		result.gapWidth = gapWidth;
	}
	if (overlap !== undefined) {
		if (!chartTypeSupportsOverlap(effectiveChartType, seriesTypes)) throw new Error(`${label}.overlap n’est pas compatible avec ce type.`);
		result.overlap = overlap;
	}
	if (source.roundedCorners !== undefined) result.roundedCorners = booleanValue(source.roundedCorners, `${label}.roundedCorners`, false);
	const alternativeText = textValue(source.alternativeText, `${label}.alternativeText`, { max: 1_000 });
	if (alternativeText !== undefined) result.alternativeText = alternativeText;
	return result;
}

function parseOperation(value: unknown, index: number): WorkbookObjectDesignOperation {
	const label = `operations[${index}]`;
	const source = objectValue(value, label);
	const kind = textValue(source.kind, `${label}.kind`, { required: true });
	const sheet = sheetName(source.sheetName, `${label}.sheetName`);
	if (kind === 'createWorksheetTable') {
		onlyKeys(source, ['kind', 'sheetName', 'table'], label);
		const table = parseTable(source.table, `${label}.table`);
		if (!table.headerRow) {
			throw new Error(`${label}.table.headerRow=false est refusé à la création pour éviter que Microsoft Excel déplace les cellules.`);
		}
		if (table.totalsRow) {
			throw new Error(`${label}.table.totalsRow=true est refusé à la création car Microsoft Excel déplace les cellules et réécrit les références de formules.`);
		}
		return { kind, sheetName: sheet, table };
	}
	if (kind === 'updateWorksheetTable') {
		onlyKeys(source, ['kind', 'sheetName', 'name', 'table'], label);
		return { kind, sheetName: sheet, name: tableName(source.name, `${label}.name`), table: parseTable(source.table, `${label}.table`) };
	}
	if (kind === 'deleteWorksheetTable') {
		onlyKeys(source, ['kind', 'sheetName', 'name'], label);
		return { kind, sheetName: sheet, name: tableName(source.name, `${label}.name`) };
	}
	if (kind === 'createWorksheetChart') {
		onlyKeys(source, ['kind', 'sheetName', 'chart'], label);
		const chart = parseChart(source.chart, `${label}.chart`);
		if (chart.legend?.position === 'custom') {
			throw new Error(`${label}.chart.legend.position=custom peut seulement préserver une disposition manuelle existante lors d’une mise à jour.`);
		}
		return { kind, sheetName: sheet, chart };
	}
	if (kind === 'updateWorksheetChart') {
		onlyKeys(source, ['kind', 'sheetName', 'name', 'chart'], label);
		return { kind, sheetName: sheet, name: objectName(source.name, `${label}.name`), chart: parseChart(source.chart, `${label}.chart`) };
	}
	if (kind === 'deleteWorksheetChart') {
		onlyKeys(source, ['kind', 'sheetName', 'name'], label);
		return { kind, sheetName: sheet, name: objectName(source.name, `${label}.name`) };
	}
	throw new Error(`${label}.kind n’est pas une opération de tableau/graphique prise en charge.`);
}

export function parseWorkbookObjectToolInput(value: unknown): WorkbookObjectDesignToolInput {
	const source = objectValue(value, 'input');
	onlyKeys(source, ['workbookPath', 'operations'], 'input');
	const workbookPath = textValue(source.workbookPath, 'workbookPath', {
		required: true,
		max: 32_767
	}) as string;
	if (!/^file:/i.test(workbookPath) && !path.isAbsolute(workbookPath)) {
		throw new Error('workbookPath doit être un chemin absolu explicite pour cette opération avec effet de bord.');
	}
	if (!Array.isArray(source.operations) || source.operations.length < 1 || source.operations.length > MAX_OPERATIONS) {
		throw new Error(`operations doit contenir de 1 à ${MAX_OPERATIONS} opérations.`);
	}
	const operations = source.operations.map(parseOperation);
	let requestedCells = 0;
	for (const operation of operations) {
		if (operation.kind === 'createWorksheetTable' || operation.kind === 'updateWorksheetTable') {
			requestedCells += parseSimpleA1Range(operation.table.rangeRef)?.cellCount ?? 0;
		}
		if (operation.kind === 'createWorksheetChart' || operation.kind === 'updateWorksheetChart') {
			const ranges = [
				operation.chart.sourceRangeRef,
				...(operation.chart.series ?? []).flatMap(series => [
					series.nameRange,
					series.categoryRange,
					series.valuesRange,
					series.xValuesRange,
					series.bubbleSizesRange
				])
			].filter((range): range is string => Boolean(range));
			requestedCells += ranges.reduce(
				(sum, range) => sum + (parseSimpleA1Range(range)?.cellCount ?? 0),
				0
			);
		}
	}
	if (requestedCells > MAX_WORKBOOK_OBJECT_TRANSACTION_CELLS) {
		throw new Error(`La transaction référence ${requestedCells.toLocaleString('fr-FR')} cellules, au-delà de la limite de ${MAX_WORKBOOK_OBJECT_TRANSACTION_CELLS.toLocaleString('fr-FR')}.`);
	}
	const targets = new Set<string>();
	const desiredTableNames = new Set<string>();
	const desiredChartNames = new Set<string>();
	const desiredTableRangesBySheet = new Map<string, Array<{ name: string; rangeRef: string }>>();
	for (const operation of operations) {
		const objectType = operation.kind.includes('Table') ? 'table' : 'chart';
		const name = 'name' in operation
			? operation.name
			: operation.kind === 'createWorksheetTable'
				? operation.table.name
				: operation.chart.name;
		const comparisonName = objectType === 'table'
			? excelTableNameComparisonKey(name)
			: name.toLocaleLowerCase('en-US');
		const key = `${objectType}\0${operation.sheetName.toLocaleLowerCase('en-US')}\0${comparisonName}`;
		if (targets.has(key)) throw new Error(`Le même objet est ciblé plusieurs fois dans la transaction : ${operation.sheetName}.${name}.`);
		targets.add(key);
		if (operation.kind === 'createWorksheetTable' || operation.kind === 'updateWorksheetTable') {
			const tableNames = new Set([
				excelTableNameComparisonKey(operation.table.name),
				excelTableNameComparisonKey(operation.table.displayName)
			]);
			for (const desiredName of tableNames) {
				if (desiredTableNames.has(desiredName)) {
					throw new Error(`Le nom de tableau ${operation.table.name} est demandé plusieurs fois dans le classeur.`);
				}
				desiredTableNames.add(desiredName);
			}
			const sheetKey = operation.sheetName.toLocaleLowerCase('en-US');
			const ranges = desiredTableRangesBySheet.get(sheetKey) ?? [];
			const overlap = ranges.find(candidate => simpleA1RangesOverlap(
				candidate.rangeRef,
				operation.table.rangeRef
			));
			if (overlap) {
				throw new Error(`Les tableaux ${overlap.name} et ${operation.table.name} se chevauchent dans la feuille ${operation.sheetName}.`);
			}
			ranges.push({ name: operation.table.name, rangeRef: operation.table.rangeRef });
			desiredTableRangesBySheet.set(sheetKey, ranges);
		}
		if (operation.kind === 'createWorksheetChart' || operation.kind === 'updateWorksheetChart') {
			const desiredName = `${operation.sheetName.toLocaleLowerCase('en-US')}\0${operation.chart.name.toLocaleLowerCase('en-US')}`;
			if (desiredChartNames.has(desiredName)) {
				throw new Error(`Le nom de graphique ${operation.chart.name} est demandé plusieurs fois dans la feuille ${operation.sheetName}.`);
			}
			desiredChartNames.add(desiredName);
		}
	}
	return {
		workbookPath,
		operations
	};
}

export function registerWorkbookObjectLanguageModelTool(
	context: vscode.ExtensionContext,
	service: ExcelAiVbaWorkbookService
): void {
	const runtime = vscode as typeof vscode & LanguageModelConstructors & { lm?: LanguageModelApi };
	if (!runtime.lm?.registerTool) return;
	const tool = {
		async prepareInvocation(options: { input?: unknown }) {
			const input = parseWorkbookObjectToolInput(options?.input);
			const resolvedUri = await service.resolveToolWorkbookUri(input);
			if (!resolvedUri) {
				throw new Error('Aucun classeur XLSX ou XLSM local unique ne peut recevoir les objets demandés.');
			}
			await assertNoReparsePointChain(resolvedUri.fsPath);
			const canonicalUri = await canonicalizeWorkbookUri(resolvedUri);
			const canonicalPath = canonicalUri.fsPath;
			const insideWorkspace = (vscode.workspace.workspaceFolders ?? []).some(folder => (
				folder.uri.scheme === 'file'
				&& !folder.uri.authority
				&& pathIsInside(canonicalPath, folder.uri.fsPath)
			));
			const operationSummary = [...new Set(input.operations.map(operation => operation.kind))].join(', ');
			const scopeWarning = insideWorkspace
				? 'La cible se trouve dans l’espace de travail ouvert.'
				: 'Attention : la cible se trouve hors de l’espace de travail ouvert.';
			return {
				invocationMessage: `Modification transactionnelle des tableaux et graphiques dans ${path.basename(canonicalPath)}`,
				confirmationMessages: {
					title: `Modifier ${path.basename(canonicalPath)} ?`,
					message: [
						`${input.operations.length} opération(s) avec effet de bord : ${operationSummary}.`,
						`Chemin canonique complet : ${canonicalPath}`,
						scopeWarning,
						'Une sauvegarde persistante et une vérification transactionnelle seront créées avant remplacement.'
					].join('\n\n')
				}
			};
		},
		async invoke(options: { input?: unknown }, cancellationToken?: vscode.CancellationToken): Promise<unknown> {
			if (cancellationToken?.isCancellationRequested) throw new vscode.CancellationError();
			const input = parseWorkbookObjectToolInput(options?.input);
			const workbookUri = await service.resolveToolWorkbookUri(input);
			if (!workbookUri) throw new Error('Aucun classeur XLSX ou XLSM local unique ne peut recevoir les objets demandés.');
			await assertNoReparsePointChain(workbookUri.fsPath);
			const canonicalUri = await canonicalizeWorkbookUri(workbookUri);
			const result = await service.designWorkbookObjectsFromTool(canonicalUri, input.operations, cancellationToken);
			const Result = runtime.LanguageModelToolResult;
			const TextPart = runtime.LanguageModelTextPart;
			if (!Result || !TextPart) throw new Error('Les types Language Model Tool ne sont pas disponibles dans cette version de VS Code.');
			return new Result([new TextPart(JSON.stringify({ ok: true, ...result }))]);
		}
	};
	context.subscriptions.push(runtime.lm.registerTool(EXCEL_AI_WORKBOOK_DESIGN_TOOL, tool));
	service.getOutputChannel().appendLine(`[outil IA] ${EXCEL_AI_WORKBOOK_DESIGN_TOOL} enregistré pour les tableaux et graphiques natifs transactionnels.`);
}
