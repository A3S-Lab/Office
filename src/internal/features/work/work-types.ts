import type { Sheet } from '@fortune-sheet/core';
import type { WorkSpreadsheetChartLayout } from './work-spreadsheet-chart-layout';
import type {
  WorkSpreadsheetChart,
  WorkSpreadsheetChartAxes,
  WorkSpreadsheetChartAxis,
  WorkSpreadsheetChartSeriesStyle,
  WorkSpreadsheetErrorBars,
  WorkSpreadsheetImage,
  WorkSpreadsheetTrendline,
} from './work-spreadsheet-chart-types';

export * from './work-spreadsheet-chart-types';

export type WorkArtifactKind =
  | 'document'
  | 'markdown'
  | 'spreadsheet'
  | 'presentation'
  | 'pdf';
export type WorkLibraryView =
  | 'home'
  | 'recent'
  | 'favorites'
  | 'folder'
  | 'trash';
export type WorkSaveState = 'saved' | 'dirty' | 'saving' | 'error';
export type WorkStorageMode = 'server' | 'local';

export interface WorkDocumentMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export interface WorkDocumentColumnDefinition {
  widthPercent: number;
  spacing: number;
}

export interface WorkDocumentColumns {
  count: number;
  spacing: number;
  separator: boolean;
  custom?: WorkDocumentColumnDefinition[];
}

export type WorkDocumentPageChromeVariant = 'default' | 'first' | 'even';

export interface WorkDocumentPageChromeContent {
  headerHtml: string;
  footerHtml: string;
  showPageNumber: boolean;
}

export interface WorkDocumentPageChrome {
  differentFirstPage: boolean;
  differentOddEvenPages: boolean;
  default: WorkDocumentPageChromeContent;
  first: WorkDocumentPageChromeContent;
  even: WorkDocumentPageChromeContent;
}

export type WorkDocumentSectionBreakType =
  | 'nextPage'
  | 'continuous'
  | 'evenPage'
  | 'oddPage'
  | 'nextColumn';

export interface WorkDocumentSectionLayout {
  pageSize: 'a4' | 'letter';
  orientation: 'portrait' | 'landscape';
  margins: WorkDocumentMargins;
  columns: WorkDocumentColumns;
  breakAfter: WorkDocumentSectionBreakType;
  headerText?: string;
  footerText?: string;
  showPageNumbers?: boolean;
  pageNumberStart?: number;
  pageChrome?: WorkDocumentPageChrome;
}

export type WorkDocumentAttributeValue =
  | string
  | number
  | boolean
  | null
  | WorkDocumentAttributeValue[]
  | { [key: string]: WorkDocumentAttributeValue };

export interface WorkDocumentMark {
  type: string;
  attrs?: Record<string, WorkDocumentAttributeValue>;
}

export interface WorkDocumentNode {
  type: string;
  attrs?: Record<string, WorkDocumentAttributeValue>;
  content?: WorkDocumentNode[];
  marks?: WorkDocumentMark[];
  text?: string;
}

export interface WorkDocumentModel {
  schema: 'a3s.office.document';
  version: 1;
  revision: number;
  htmlFingerprint: string;
  root: WorkDocumentNode;
}

export interface WorkDocumentContent {
  type: 'document';
  html: string;
  model?: WorkDocumentModel;
  pageSize: 'a4' | 'letter';
  orientation?: 'portrait' | 'landscape';
  margins?: WorkDocumentMargins;
  columns?: WorkDocumentColumns;
  headerText?: string;
  footerText?: string;
  showPageNumbers?: boolean;
  pageNumberStart?: number;
  pageChrome?: WorkDocumentPageChrome;
  trackChanges?: boolean;
  comments?: WorkDocumentComment[];
  bibliography?: WorkDocumentBibliography;
}

export interface WorkMarkdownContent {
  type: 'markdown';
  markdown: string;
}

export type WorkDocumentCitationStyle = 'apa' | 'mla' | 'chicago' | 'ieee';

export interface WorkDocumentCitationPerson {
  first: string;
  middle?: string;
  last: string;
  suffix?: string;
}

export interface WorkDocumentCitationContributor {
  people?: WorkDocumentCitationPerson[];
  corporate?: string;
}

export interface WorkDocumentCitationSource {
  id: string;
  tag: string;
  sourceType: string;
  guid?: string;
  title: string;
  year?: string;
  contributors?: Record<string, WorkDocumentCitationContributor>;
  publisher?: string;
  city?: string;
  journalName?: string;
  volume?: string;
  issue?: string;
  pages?: string;
  url?: string;
  standardNumber?: string;
  conferenceName?: string;
  institution?: string;
  additionalFields?: Record<string, string>;
}

export interface WorkDocumentBibliography {
  style: WorkDocumentCitationStyle;
  styleName?: string;
  selectedStyle?: string;
  sources: WorkDocumentCitationSource[];
}

export interface WorkDocumentCommentReply {
  id: string;
  author: string;
  date: string;
  text: string;
}

export interface WorkDocumentComment {
  id: string;
  author: string;
  date: string;
  text: string;
  resolved: boolean;
  replies?: WorkDocumentCommentReply[];
}

export interface WorkSpreadsheetContent {
  type: 'spreadsheet';
  sheets: WorkSpreadsheetSheet[];
  calculation?: WorkSpreadsheetCalculationSettings;
  namedRanges?: WorkSpreadsheetNamedRange[];
  printAreas?: WorkSpreadsheetPrintArea[];
  printTitles?: WorkSpreadsheetPrintTitles[];
  pageBreaks?: WorkSpreadsheetPageBreaks[];
  pageSetups?: WorkSpreadsheetPageSetup[];
}

export type WorkSpreadsheetSheet = Omit<Sheet, 'images'> & {
  images?: WorkSpreadsheetImage[];
  charts?: WorkSpreadsheetChart[];
  pivotTables?: WorkSpreadsheetPivotTable[];
  formulaMetadata?: WorkSpreadsheetFormulaMetadata;
};

export type WorkSpreadsheetPivotAggregation =
  | 'sum'
  | 'count'
  | 'counta'
  | 'average'
  | 'max'
  | 'min'
  | 'product'
  | 'stdDev'
  | 'stdDevP'
  | 'var'
  | 'varP';

export interface WorkSpreadsheetPivotValue {
  fieldIndex: number;
  aggregation: WorkSpreadsheetPivotAggregation;
  caption?: string;
}

export type WorkSpreadsheetPivotFilterValue = string | number | boolean | null;

export interface WorkSpreadsheetPivotReportFilter {
  fieldIndex: number;
  selectedItem?: WorkSpreadsheetPivotFilterValue;
}

export interface WorkSpreadsheetPivotTable {
  id: string;
  name: string;
  sourceSheetId: string;
  sourceReference: string;
  anchor: string;
  rowFields: number[];
  columnFields: number[];
  reportFilters?: WorkSpreadsheetPivotReportFilter[];
  values: WorkSpreadsheetPivotValue[];
  rowGrandTotals: boolean;
  columnGrandTotals: boolean;
  styleName: string;
  refreshOnLoad: boolean;
  outputReference?: string;
}

export type WorkSpreadsheetCalculationMode =
  | 'automatic'
  | 'automatic-except-data-tables'
  | 'manual';

export interface WorkSpreadsheetCalculationSettings {
  mode: WorkSpreadsheetCalculationMode;
  fullCalculationOnLoad: boolean;
  forceFullCalculation: boolean;
  iterativeCalculation: boolean;
  maximumIterations: number;
  maximumChange: number;
  fullPrecision: boolean;
}

export type WorkSpreadsheetFormulaRangeType =
  | 'array'
  | 'dynamic-array'
  | 'data-table';

export interface WorkSpreadsheetDataTableOptions {
  input1Reference?: string;
  input2Reference?: string;
  twoDimensional?: boolean;
  rowOriented?: boolean;
  input1Deleted?: boolean;
  input2Deleted?: boolean;
  calculateOnLoad?: boolean;
}

export interface WorkSpreadsheetFormulaRange {
  type: WorkSpreadsheetFormulaRangeType;
  anchor: string;
  reference: string;
  formula?: string;
  dataTable?: WorkSpreadsheetDataTableOptions;
}

export interface WorkSpreadsheetFormulaMetadata {
  ranges?: WorkSpreadsheetFormulaRange[];
  sourceFormulas?: Record<string, string>;
  normalizedSharedFormulaGroups?: number;
  normalizedSharedFormulaCells?: number;
}

export interface WorkSpreadsheetNamedRange {
  id: string;
  name: string;
  reference: string;
  scopeSheetId?: string;
  comment?: string;
}

export interface WorkSpreadsheetPrintArea {
  sheetId: string;
  reference: string;
}

export interface WorkSpreadsheetPrintTitles {
  sheetId: string;
  rows?: string;
  columns?: string;
}

export interface WorkSpreadsheetPageBreaks {
  sheetId: string;
  rows?: number[];
  columns?: number[];
}

export interface WorkSpreadsheetPageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
}

export interface WorkSpreadsheetHeaderFooterSections {
  left?: string;
  center?: string;
  right?: string;
}

export type WorkSpreadsheetPaperSize =
  | 'a3'
  | 'a4'
  | 'a5'
  | 'letter'
  | 'legal'
  | 'tabloid';

export interface WorkSpreadsheetPageSetup {
  sheetId: string;
  paperSize?: WorkSpreadsheetPaperSize;
  orientation?: 'portrait' | 'landscape';
  scale?: number;
  fitToPage?: boolean;
  fitToWidth?: number;
  fitToHeight?: number;
  horizontalCentered?: boolean;
  verticalCentered?: boolean;
  header?: WorkSpreadsheetHeaderFooterSections;
  footer?: WorkSpreadsheetHeaderFooterSections;
  pageNumberStart?: number;
  pageOrder?: 'downThenOver' | 'overThenDown';
  scaleWithDocument?: boolean;
  alignWithMargins?: boolean;
  margins?: WorkSpreadsheetPageMargins;
}

export type WorkCompatibilitySeverity = 'info' | 'warning' | 'error';

export interface WorkCompatibilityIssue {
  code: string;
  severity: WorkCompatibilitySeverity;
  feature: string;
  message: string;
  location?: string;
}

export interface WorkCompatibilityReport {
  sourceFormat: string;
  sourceName: string;
  assessedAt: number;
  issues: WorkCompatibilityIssue[];
}

export type WorkSlideElementType =
  | 'text'
  | 'shape'
  | 'image'
  | 'table'
  | 'chart'
  | 'line';
export type WorkSlideTextAlign = 'left' | 'center' | 'right';
export type WorkSlideVerticalAlign = 'top' | 'middle' | 'bottom';
export type WorkSlideShapeType =
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'diamond'
  | 'line';

export interface WorkSlideTextRun {
  text: string;
  fontSize?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontFamily?: string;
  href?: string;
}

export interface WorkSlideImage {
  dataUrl: string;
  contentType: string;
  name: string;
}

export interface WorkSlideTable {
  rows: string[][];
  headerRows?: number;
}

export type WorkSlideChartErrorBars = WorkSpreadsheetErrorBars;
export type WorkSlideChartTrendline = WorkSpreadsheetTrendline;
export type WorkSlideChartSeriesStyle = WorkSpreadsheetChartSeriesStyle;

export interface WorkSlideChartSeries {
  name: string;
  values: number[];
  bubbleSizes?: number[];
  errorBars?: WorkSlideChartErrorBars[];
  trendlines?: WorkSlideChartTrendline[];
  style?: WorkSlideChartSeriesStyle;
}

export type WorkSlideChartType =
  | 'bar'
  | 'column'
  | 'line'
  | 'pie'
  | 'doughnut'
  | 'area'
  | 'radar'
  | 'scatter'
  | 'bubble';
export type WorkSlideRadarStyle = 'standard' | 'marker' | 'filled';
export type WorkSlideScatterStyle =
  | 'marker'
  | 'line'
  | 'lineMarker'
  | 'smooth'
  | 'smoothMarker';
export type WorkSlideBubbleSizeRepresents = 'area' | 'width';
export type WorkSlideChartAxis = WorkSpreadsheetChartAxis;
export type WorkSlideChartAxes = WorkSpreadsheetChartAxes;
export type WorkSlideChartLegendPosition =
  | 'right'
  | 'left'
  | 'top'
  | 'bottom'
  | 'topRight';
export type WorkSlideChartDataLabelPosition =
  | 'bestFit'
  | 'center'
  | 'insideBase'
  | 'insideEnd'
  | 'outsideEnd'
  | 'left'
  | 'right'
  | 'above'
  | 'below';

export interface WorkSlideChartDataLabels {
  showValue?: boolean;
  showCategoryName?: boolean;
  showSeriesName?: boolean;
  showPercentage?: boolean;
  showBubbleSize?: boolean;
  separator?: string;
  position?: WorkSlideChartDataLabelPosition;
}

export interface WorkSlideChart extends WorkSpreadsheetChartLayout {
  type: WorkSlideChartType;
  title?: string;
  categories: string[];
  series: WorkSlideChartSeries[];
  showLegend?: boolean;
  legendPosition?: WorkSlideChartLegendPosition;
  axes?: WorkSlideChartAxes;
  // Retained for persisted artifacts created before editable axis settings.
  categoryAxisTitle?: string;
  valueAxisTitle?: string;
  dataLabels?: WorkSlideChartDataLabels;
  doughnutHoleSize?: number;
  radarStyle?: WorkSlideRadarStyle;
  scatterStyle?: WorkSlideScatterStyle;
  bubbleScale?: number;
  showNegativeBubbles?: boolean;
  bubbleSizeRepresents?: WorkSlideBubbleSizeRepresents;
}

export interface WorkSlidePlaceholder {
  key: string;
  type: string;
  prompt?: string;
  inheritsGeometry?: boolean;
  inheritsStyle?: boolean;
}

export type WorkSlideTransitionType =
  | 'fade'
  | 'push'
  | 'wipe'
  | 'split'
  | 'cut';
export type WorkSlideTransitionSpeed = 'fast' | 'medium' | 'slow';
export type WorkSlideTransitionDirection =
  | 'left'
  | 'right'
  | 'up'
  | 'down'
  | 'in'
  | 'out';

export interface WorkSlideTransition {
  type: WorkSlideTransitionType;
  speed: WorkSlideTransitionSpeed;
  direction?: WorkSlideTransitionDirection;
  orientation?: 'horizontal' | 'vertical';
  advanceOnClick: boolean;
  advanceAfterMs?: number;
}

export interface WorkSlideElement {
  id: string;
  type: WorkSlideElementType;
  x: number;
  y: number;
  width: number;
  height: number;
  text: string;
  fontSize: number;
  color: string;
  fill: string;
  bold: boolean;
  align: WorkSlideTextAlign;
  radius?: number;
  shapeType?: WorkSlideShapeType;
  rotation?: number;
  opacity?: number;
  borderColor?: string;
  borderWidth?: number;
  fontFamily?: string;
  italic?: boolean;
  underline?: boolean;
  verticalAlign?: WorkSlideVerticalAlign;
  textRuns?: WorkSlideTextRun[];
  image?: WorkSlideImage;
  table?: WorkSlideTable;
  chart?: WorkSlideChart;
  href?: string;
  altText?: string;
  placeholder?: WorkSlidePlaceholder;
}

export interface WorkSlide {
  id: string;
  name: string;
  background: string;
  layoutId?: string;
  useLayoutBackground?: boolean;
  showMasterElements?: boolean;
  elements: WorkSlideElement[];
  notes?: string;
  comments?: WorkSlideComment[];
  transition?: WorkSlideTransition;
}

export interface WorkSlideComment {
  id: string;
  author: string;
  initials?: string;
  date: string;
  text: string;
  x: number;
  y: number;
}

export interface WorkPresentationMaster {
  id: string;
  name: string;
  background: string;
  elements: WorkSlideElement[];
}

export interface WorkPresentationLayout {
  id: string;
  name: string;
  masterId: string;
  background?: string;
  elements: WorkSlideElement[];
  showMasterElements?: boolean;
  sourceType?: string;
}

export interface WorkPresentationContent {
  type: 'presentation';
  slides: WorkSlide[];
  width?: number;
  height?: number;
  masters?: WorkPresentationMaster[];
  layouts?: WorkPresentationLayout[];
}

export type WorkPresentationPrintLayout =
  | 'slides'
  | 'notes'
  | 'handout-2'
  | 'handout-3'
  | 'handout-6';

export interface WorkPdfContent {
  type: 'pdf';
  pageCount?: number;
}

export type WorkArtifactContent =
  | WorkDocumentContent
  | WorkMarkdownContent
  | WorkSpreadsheetContent
  | WorkPresentationContent
  | WorkPdfContent;

export interface WorkArtifact {
  id: string;
  kind: WorkArtifactKind;
  title: string;
  favorite: boolean;
  createdAt: number;
  updatedAt: number;
  lastOpenedAt: number;
  revision: number;
  content: WorkArtifactContent;
  folderId?: string | null;
  trashedAt?: number | null;
  source?: WorkSourceFile | null;
  compatibility?: WorkCompatibilityReport | null;
}

export interface WorkSourceFile {
  name: string;
  contentType: string;
  size: number;
  updatedAt: number;
}

export interface WorkFolder {
  id: string;
  name: string;
  parentId?: string | null;
  createdAt: number;
  updatedAt: number;
  revision: number;
  trashedAt?: number | null;
}

export interface WorkStorageLimits {
  artifactBytes: number;
  sourceBytes: number;
  historyEntries: number;
}

export interface WorkLibrarySnapshot {
  artifacts: WorkArtifact[];
  folders: WorkFolder[];
  limits: WorkStorageLimits | null;
  storage: WorkStorageMode;
}

export interface WorkArtifactVersion {
  revision: number;
  updatedAt: number;
  current: boolean;
  artifact: WorkArtifact;
}

export interface WorkTemplate {
  id: string;
  kind: WorkArtifactKind;
  name: string;
  description: string;
  accent: string;
}

export function workArtifactExtension(kind: WorkArtifactKind): string {
  if (kind === 'document') return 'docx';
  if (kind === 'markdown') return 'md';
  if (kind === 'spreadsheet') return 'xlsx';
  if (kind === 'presentation') return 'pptx';
  return 'pdf';
}

export function workArtifactKindLabel(kind: WorkArtifactKind): string {
  if (kind === 'document') return '文字';
  if (kind === 'markdown') return 'Markdown';
  if (kind === 'spreadsheet') return '表格';
  if (kind === 'presentation') return '演示';
  return 'PDF';
}
