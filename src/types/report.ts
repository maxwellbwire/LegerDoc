export interface ColumnSummary {
  name: string;
  type: 'number' | 'text' | 'date' | 'mixed';
  nonEmptyCount: number;
  sum?: number;
  average?: number;
  min?: number;
  max?: number;
  uniqueCount?: number;
  topValues?: Array<{ value: string; count: number }>;
}

export interface ParsedSheet {
  name: string;
  headers: string[];
  rows: Record<string, string | number | boolean | null>[];
  rowCount: number;
  columnSummaries: ColumnSummary[];
}

export interface ParsedWorkbook {
  fileName: string;
  fileSize: number;
  sheetNames: string[];
  sheets: ParsedSheet[];
  parsedAt: string;
}

export type TemplateSourceKind = 'docx' | 'text' | 'preset';

export interface ParsedTemplate {
  name: string;
  sourceKind: TemplateSourceKind;
  rawText: string;
  htmlPreview: string;
  headings: string[];
  placeholders: string[];
  docxBase64?: string;
}

export interface ReportTable {
  caption: string;
  headers: string[];
  rows: string[][];
}

export interface ReportSection {
  id: string;
  heading: string;
  narrative: string;
  bullets: string[];
  table?: ReportTable;
}

export interface ReportMetric {
  label: string;
  value: string;
  context: string;
}

export interface PlaceholderMapping {
  key: string;
  value: string;
}

export interface GeneratedReport {
  reportTitle: string;
  subtitle: string;
  preparedDate: string;
  executiveSummary: string;
  keyMetrics: ReportMetric[];
  sections: ReportSection[];
  placeholderValues: PlaceholderMapping[];
  generatedAt: string;
}

export type ReportTone =
  | 'Executive Brief'
  | 'Operational Detail'
  | 'Client-Ready'
  | 'Board Memo';

export interface SavedReportHistoryItem {
  id: string;
  reportTitle: string;
  workbookName: string;
  templateName: string;
  generatedAt: string;
  report: GeneratedReport;
}
