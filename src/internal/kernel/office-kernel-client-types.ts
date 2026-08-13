import type {
  OfficeKernelLayoutRequest,
  OfficeKernelLayoutResult,
  OfficeKernelPresentationGeometryRequest,
  OfficeKernelPresentationGeometryResult,
  OfficeKernelSpreadsheetCalculationRequest,
  OfficeKernelSpreadsheetCalculationResult,
  OfficeKernelSpreadsheetSessionCalculationResult,
  OfficeKernelSpreadsheetSessionCalculationScope,
  OfficeKernelSpreadsheetSessionUpdate,
  OfficeKernelTextLayoutRequest,
  OfficeKernelTextLayoutResult,
} from './office-kernel-protocol';

export interface OfficeKernelLayoutInput {
  revision: number;
  documentRevision: number;
  startPageIndex?: number;
  page: OfficeKernelLayoutRequest['page'];
  pageStyles?: OfficeKernelLayoutRequest['pageStyles'];
  blocks: OfficeKernelLayoutRequest['blocks'];
}

export interface OfficeKernelPresentationGeometryInput {
  revision: number;
  documentRevision: number;
  operation: OfficeKernelPresentationGeometryRequest['operation'];
  elements: OfficeKernelPresentationGeometryRequest['elements'];
}

export interface OfficeKernelTextLayoutInput {
  revision: number;
  documentRevision: number;
  paragraphs: OfficeKernelTextLayoutRequest['paragraphs'];
}

export interface OfficeKernelSpreadsheetCalculationInput {
  revision: number;
  documentRevision: number;
  sheets: OfficeKernelSpreadsheetCalculationRequest['sheets'];
  targets?: OfficeKernelSpreadsheetCalculationRequest['targets'];
}

export interface OfficeKernelSpreadsheetSessionCalculationInput {
  revision: number;
  documentRevision: number;
  update: OfficeKernelSpreadsheetSessionUpdate;
  calculation: OfficeKernelSpreadsheetSessionCalculationScope;
  fallbackSheets: OfficeKernelSpreadsheetCalculationRequest['sheets'];
}

export interface OfficeKernelClient {
  layout(
    input: OfficeKernelLayoutInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelLayoutResult>;
  presentationGeometry(
    input: OfficeKernelPresentationGeometryInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelPresentationGeometryResult>;
  textLayout(
    input: OfficeKernelTextLayoutInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelTextLayoutResult>;
  spreadsheetCalculation(
    input: OfficeKernelSpreadsheetCalculationInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelSpreadsheetCalculationResult>;
  spreadsheetSessionCalculation(
    input: OfficeKernelSpreadsheetSessionCalculationInput,
    signal?: AbortSignal,
  ): Promise<OfficeKernelSpreadsheetSessionCalculationResult>;
  dispose(): void;
}
