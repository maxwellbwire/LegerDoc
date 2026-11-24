import { DragEvent, useEffect, useState } from 'react';
import {
  AlertCircle,
  Clock,
  FileSpreadsheet,
  FileText,
  Sparkles,
} from 'lucide-react';
import type {
  GeneratedReport,
  ParsedTemplate,
  ParsedWorkbook,
  ReportTone,
  SavedReportHistoryItem,
} from './types/report';
import { createSampleWorkbook, parseExcelFile } from './utils/excelParser';
import {
  parseDocxTemplate,
  parseTextTemplate,
  PRESET_TEMPLATES,
} from './utils/templateParser';
import { SpreadsheetInspector } from './components/SpreadsheetInspector';
import { TemplateEditor } from './components/TemplateEditor';
import { ReportDocumentView } from './components/ReportDocumentView';
import { ExportGuideView } from './components/ExportGuideView';

type ActiveView = 'report' | 'spreadsheet' | 'template' | 'guide';

const HISTORY_STORAGE_KEY = 'ledgerdoc_report_history_v1';

const TONE_OPTIONS: ReportTone[] = [
  'Executive Brief',
  'Operational Detail',
  'Client-Ready',
  'Board Memo',
];

export default function App() {
  const [activeView, setActiveView] = useState<ActiveView>('report');
  const [workbook, setWorkbook] = useState<ParsedWorkbook | null>(null);
  const [selectedSheets, setSelectedSheets] = useState<string[]>([]);
  const [template, setTemplate] = useState<ParsedTemplate>(() =>
    parseTextTemplate(
      PRESET_TEMPLATES[0].name,
      PRESET_TEMPLATES[0].content,
      'preset'
    )
  );
  const [tone, setTone] = useState<ReportTone>('Executive Brief');
  const [instructions, setInstructions] = useState<string>('');
  const [report, setReport] = useState<GeneratedReport | null>(null);
  const [isGenerating, setIsGenerating] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [excelDragOver, setExcelDragOver] = useState<boolean>(false);
  const [templateDragOver, setTemplateDragOver] = useState<boolean>(false);
  const [history, setHistory] = useState<SavedReportHistoryItem[]>(() => {
    try {
      const raw = localStorage.getItem(HISTORY_STORAGE_KEY);
      return raw ? (JSON.parse(raw) as SavedReportHistoryItem[]) : [];
    } catch {
      return [];
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(history.slice(0, 6)));
    } catch {
      // Ignore storage quota errors
    }
  }, [history]);

  const handleExcelFile = async (file: File) => {
    setErrorMessage(null);
    try {
      const parsed = await parseExcelFile(file);
      if (parsed.sheets.length === 0) {
        setErrorMessage('The uploaded Excel file does not contain any readable sheets.');
        return;
      }
      setWorkbook(parsed);
      setSelectedSheets(parsed.sheetNames);
    } catch (err: any) {
      setErrorMessage(err?.message || 'Failed to parse the Excel file. Verify it is a valid .xlsx, .xls, or .csv file.');
    }
  };

  const handleLoadSampleWorkbook = () => {
    setErrorMessage(null);
    const sample = createSampleWorkbook();
    setWorkbook(sample);
    setSelectedSheets(sample.sheetNames);
  };

  const handleTemplateFile = async (file: File) => {
    setErrorMessage(null);
    const lower = file.name.toLowerCase();
    try {
      if (lower.endsWith('.docx')) {
        const parsed = await parseDocxTemplate(file);
        setTemplate(parsed);
      } else if (lower.endsWith('.txt') || lower.endsWith('.md') || lower.endsWith('.html')) {
        const text = await file.text();
        setTemplate(parseTextTemplate(file.name, text, 'text'));
      } else {
        setErrorMessage('Unsupported template format. Upload a .docx (from Google Docs), .md, or .txt file.');
      }
    } catch (err: any) {
      setErrorMessage(err?.message || 'Could not read the uploaded template file.');
    }
  };

  const handleApplyTextTemplate = (name: string, text: string, kind: 'text' | 'preset') => {
    setErrorMessage(null);
    setTemplate(parseTextTemplate(name, text, kind));
  };

  const handleToggleSheet = (sheetName: string) => {
    setSelectedSheets((prev) =>
      prev.includes(sheetName)
        ? prev.length > 1
          ? prev.filter((s) => s !== sheetName)
          : prev
        : [...prev, sheetName]
    );
  };

  const handleGenerateReport = async () => {
    if (!workbook) {
      setErrorMessage('Upload an Excel workbook or load the sample workbook first.');
      return;
    }
    if (!template || (!template.rawText && template.headings.length === 0)) {
      setErrorMessage('Select or upload a document template before generating a report.');
      return;
    }

    setErrorMessage(null);
    setIsGenerating(true);
    setActiveView('report');

    try {
      const response = await fetch('/api/generate-report', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          workbook,
          selectedSheetNames: selectedSheets,
          template: {
            name: template.name,
            rawText: template.rawText,
            headings: template.headings,
            placeholders: template.placeholders,
          },
          tone,
          instructions: instructions.trim() || undefined,
        }),
      });

      const data = await response.json();
      if (!response.ok) {
        throw new Error(data?.error || `Request failed with status ${response.status}`);
      }

      const generated = data as GeneratedReport;
      setReport(generated);
      setHistory((prev) => [
        {
          id: `${Date.now()}`,
          reportTitle: generated.reportTitle,
          workbookName: workbook.fileName,
          templateName: template.name,
          generatedAt: generated.generatedAt,
          report: generated,
        },
        ...prev.slice(0, 5),
      ]);
    } catch (err: any) {
      setErrorMessage(err?.message || 'Failed to generate report. Please try again.');
    } finally {
      setIsGenerating(false);
    }
  };

  const handleRefineSection = async (sectionId: string, refinementPrompt: string) => {
    if (!report || !workbook) return;
    const target = report.sections.find((s) => s.id === sectionId);
    if (!target) return;

    setErrorMessage(null);
    try {
      const response = await fetch('/api/refine-section', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          workbook,
          selectedSheetNames: selectedSheets,
          sectionHeading: target.heading,
          currentNarrative: target.narrative,
          currentBullets: target.bullets,
          refinementPrompt,
          tone,
        }),
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data?.error || 'Failed to refine section.');
      }

      setReport({
        ...report,
        sections: report.sections.map((s) =>
          s.id === sectionId
            ? {
                ...s,
                heading: data.heading || s.heading,
                narrative: data.narrative ?? s.narrative,
                bullets: Array.isArray(data.bullets) ? data.bullets : s.bullets,
              }
            : s
        ),
      });
    } catch (err: any) {
      setErrorMessage(err?.message || 'Could not refine section.');
    }
  };

  const handleUpdatePlaceholder = (key: string, value: string) => {
    if (!report) return;
    const exists = report.placeholderValues.some((p) => p.key === key);
    const updatedList = exists
      ? report.placeholderValues.map((p) => (p.key === key ? { ...p, value } : p))
      : [...report.placeholderValues, { key, value }];
    setReport({
      ...report,
      placeholderValues: updatedList,
    });
  };

  const onExcelDrop = (e: DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    setExcelDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) handleExcelFile(file);
  };

  const onTemplateDrop = (e: DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    setTemplateDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) handleTemplateFile(file);
  };

  const totalRows = workbook?.sheets.reduce((acc, s) => acc + s.rowCount, 0) ?? 0;

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 text-slate-900">
      {/* Top Bar Contract: 3 zones separated by gap-8 */}
      <header className="no-print bg-white border-b border-slate-200 sticky top-0 z-30">
        <div className="max-w-[1440px] mx-auto px-6 h-16 flex items-center justify-between gap-8">
          <a
            href="#studio"
            onClick={(e) => {
              e.preventDefault();
              setActiveView('report');
            }}
            className="text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0"
          >
            LedgerDoc
          </a>

          <nav className="hidden md:flex items-center gap-6 text-sm font-medium">
            <button
              type="button"
              onClick={() => setActiveView('report')}
              className={`py-1 border-b-2 transition-colors whitespace-nowrap shrink-0 ${
                activeView === 'report'
                  ? 'border-blue-700 text-slate-900'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              Studio
            </button>
            <button
              type="button"
              onClick={() => setActiveView('spreadsheet')}
              className={`py-1 border-b-2 transition-colors whitespace-nowrap shrink-0 ${
                activeView === 'spreadsheet'
                  ? 'border-blue-700 text-slate-900'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              Spreadsheet
            </button>
            <button
              type="button"
              onClick={() => setActiveView('template')}
              className={`py-1 border-b-2 transition-colors whitespace-nowrap shrink-0 ${
                activeView === 'template'
                  ? 'border-blue-700 text-slate-900'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              Template
            </button>
            <button
              type="button"
              onClick={() => setActiveView('guide')}
              className={`py-1 border-b-2 transition-colors whitespace-nowrap shrink-0 ${
                activeView === 'guide'
                  ? 'border-blue-700 text-slate-900'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              Export Guide
            </button>
          </nav>

          <div className="flex items-center shrink-0">
            <button
              type="button"
              disabled={!workbook || isGenerating}
              onClick={handleGenerateReport}
              className="inline-flex items-center gap-2 px-4 py-2 text-xs font-medium text-white bg-blue-700 rounded-md hover:bg-blue-800 disabled:bg-slate-300 disabled:cursor-not-allowed transition-colors whitespace-nowrap shrink-0"
            >
              <Sparkles className="w-3.5 h-3.5" />
              <span>{isGenerating ? 'Generating...' : 'Generate Report'}</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 max-w-[1440px] w-full mx-auto px-6 py-6">
        {activeView === 'guide' ? (
          <ExportGuideView onReturnToStudio={() => setActiveView('report')} />
        ) : (
          <div className="flex flex-col lg:flex-row gap-6 items-start">
            {/* Left Workflow Rail (300px) */}
            <aside className="no-print w-full lg:w-[300px] shrink-0 space-y-5">
              {/* Step 1: Excel Source */}
              <section className="bg-white border border-slate-200 rounded-lg p-4 space-y-3">
                <div className="flex items-center justify-between">
                  <h2 className="text-xs font-semibold text-slate-900">01. Excel Workbook</h2>
                  {workbook && (
                    <button
                      type="button"
                      onClick={() => setActiveView('spreadsheet')}
                      className="text-xs text-blue-700 hover:underline whitespace-nowrap shrink-0"
                    >
                      Inspect
                    </button>
                  )}
                </div>

                <label
                  onDragOver={(e) => {
                    e.preventDefault();
                    setExcelDragOver(true);
                  }}
                  onDragLeave={() => setExcelDragOver(false)}
                  onDrop={onExcelDrop}
                  className={`block cursor-pointer border border-dashed rounded-md p-4 text-center transition-colors ${
                    excelDragOver
                      ? 'border-blue-600 bg-blue-50/40'
                      : 'border-slate-300 hover:border-slate-400 bg-slate-50/50'
                  }`}
                >
                  <FileSpreadsheet className="w-5 h-5 text-slate-500 mx-auto mb-1.5" />
                  <div className="text-xs font-medium text-slate-800 truncate">
                    {workbook ? workbook.fileName : 'Drop .xlsx, .xls, or .csv'}
                  </div>
                  <div className="text-[11px] text-slate-500 mt-0.5 tabular-nums">
                    {workbook
                      ? `${workbook.sheets.length} sheet${workbook.sheets.length === 1 ? '' : 's'} · ${totalRows} rows`
                      : 'or click to browse local files'}
                  </div>
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv,.ods"
                    className="sr-only"
                    onChange={(e) => {
                      const f = e.target.files?.[0];
                      if (f) handleExcelFile(f);
                      e.currentTarget.value = '';
                    }}
                  />
                </label>

                {!workbook && (
                  <button
                    type="button"
                    onClick={handleLoadSampleWorkbook}
                    className="w-full py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-200 rounded-md hover:bg-slate-50 transition-colors"
                  >
                    Load Sample Sales &amp; Ops Workbook
                  </button>
                )}

                {workbook && workbook.sheets.length > 0 && (
                  <div className="space-y-1.5 pt-1">
                    <div className="text-[11px] font-medium text-slate-500">Included Sheets</div>
                    <div className="space-y-1 max-h-28 overflow-y-auto">
                      {workbook.sheets.map((s) => (
                        <label
                          key={s.name}
                          className="flex items-center justify-between text-xs text-slate-700 cursor-pointer py-0.5"
                        >
                          <span className="inline-flex items-center gap-2 truncate">
                            <input
                              type="checkbox"
                              checked={selectedSheets.includes(s.name)}
                              onChange={() => handleToggleSheet(s.name)}
                              className="w-3.5 h-3.5 rounded border-slate-300 text-blue-700"
                            />
                            <span className="truncate">{s.name}</span>
                          </span>
                          <span className="text-[11px] font-mono text-slate-400 tabular-nums">
                            {s.rowCount}
                          </span>
                        </label>
                      ))}
                    </div>
                  </div>
                )}
              </section>

              {/* Step 2: Document Template */}
              <section className="bg-white border border-slate-200 rounded-lg p-4 space-y-3">
                <div className="flex items-center justify-between">
                  <h2 className="text-xs font-semibold text-slate-900">02. Document Template</h2>
                  <button
                    type="button"
                    onClick={() => setActiveView('template')}
                    className="text-xs text-blue-700 hover:underline whitespace-nowrap shrink-0"
                  >
                    Configure
                  </button>
                </div>

                <label
                  onDragOver={(e) => {
                    e.preventDefault();
                    setTemplateDragOver(true);
                  }}
                  onDragLeave={() => setTemplateDragOver(false)}
                  onDrop={onTemplateDrop}
                  className={`block cursor-pointer border border-dashed rounded-md p-4 text-center transition-colors ${
                    templateDragOver
                      ? 'border-blue-600 bg-blue-50/40'
                      : 'border-slate-300 hover:border-slate-400 bg-slate-50/50'
                  }`}
                >
                  <FileText className="w-5 h-5 text-slate-500 mx-auto mb-1.5" />
                  <div className="text-xs font-medium text-slate-800 truncate">{template.name}</div>
                  <div className="text-[11px] text-slate-500 mt-0.5 tabular-nums">
                    {template.headings.length} headings · {template.placeholders.length} placeholders
                  </div>
                  <input
                    type="file"
                    accept=".docx,.txt,.md,.html"
                    className="sr-only"
                    onChange={(e) => {
                      const f = e.target.files?.[0];
                      if (f) handleTemplateFile(f);
                      e.currentTarget.value = '';
                    }}
                  />
                </label>

                <div className="space-y-1">
                  <label className="block text-[11px] font-medium text-slate-500">
                    Or choose a template structure
                  </label>
                  <select
                    value={
                      PRESET_TEMPLATES.find((p) => p.name === template.name)?.id ?? 'custom'
                    }
                    onChange={(e) => {
                      const found = PRESET_TEMPLATES.find((p) => p.id === e.target.value);
                      if (found) {
                        handleApplyTextTemplate(found.name, found.content, 'preset');
                      }
                    }}
                    className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:border-blue-600"
                  >
                    {PRESET_TEMPLATES.map((p) => (
                      <option key={p.id} value={p.id}>
                        {p.name}
                      </option>
                    ))}
                    {!PRESET_TEMPLATES.some((p) => p.name === template.name) && (
                      <option value="custom">{template.name}</option>
                    )}
                  </select>
                </div>
              </section>

              {/* Step 3: Generation Parameters */}
              <section className="bg-white border border-slate-200 rounded-lg p-4 space-y-3">
                <h2 className="text-xs font-semibold text-slate-900">03. Report Formatting</h2>

                <div className="space-y-1">
                  <label className="block text-[11px] font-medium text-slate-500">
                    Narrative Tone
                  </label>
                  <select
                    value={tone}
                    onChange={(e) => setTone(e.target.value as ReportTone)}
                    className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-md focus:outline-none focus:border-blue-600"
                  >
                    {TONE_OPTIONS.map((opt) => (
                      <option key={opt} value={opt}>
                        {opt}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="space-y-1">
                  <label className="block text-[11px] font-medium text-slate-500">
                    Focus Instructions (optional)
                  </label>
                  <textarea
                    value={instructions}
                    onChange={(e) => setInstructions(e.target.value)}
                    rows={3}
                    placeholder="e.g., Flag regions under target and summarize cost variance by department..."
                    className="w-full px-2.5 py-2 text-xs border border-slate-300 rounded-md focus:outline-none focus:border-blue-600 placeholder:text-slate-400"
                  />
                </div>

                <button
                  type="button"
                  disabled={!workbook || isGenerating}
                  onClick={handleGenerateReport}
                  className="w-full inline-flex items-center justify-center gap-2 px-4 py-2 text-xs font-medium text-white bg-slate-900 rounded-md hover:bg-slate-800 disabled:bg-slate-300 disabled:cursor-not-allowed transition-colors whitespace-nowrap shrink-0"
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  <span>{isGenerating ? 'Formatting Report...' : 'Generate Formatted Report'}</span>
                </button>
              </section>

              {/* Recent Reports */}
              {history.length > 0 && (
                <section className="bg-white border border-slate-200 rounded-lg p-4 space-y-2.5">
                  <div className="flex items-center gap-1.5 text-xs font-semibold text-slate-900">
                    <Clock className="w-3.5 h-3.5 text-slate-500" />
                    <span>Recent Reports</span>
                  </div>
                  <div className="divide-y divide-slate-100">
                    {history.map((item) => (
                      <button
                        key={item.id}
                        type="button"
                        onClick={() => {
                          setReport(item.report);
                          setActiveView('report');
                        }}
                        className="w-full text-left py-2 hover:bg-slate-50 rounded transition-colors px-1"
                      >
                        <div className="text-xs font-medium text-slate-800 truncate">
                          {item.reportTitle}
                        </div>
                        <div className="text-[11px] text-slate-500 truncate mt-0.5">
                          {item.workbookName}
                          <span className="mx-1" aria-hidden="true">·</span>
                          {new Date(item.generatedAt).toLocaleDateString()}
                        </div>
                      </button>
                    ))}
                  </div>
                </section>
              )}
            </aside>

            {/* Primary Viewport */}
            <div className="flex-1 min-w-0 space-y-5">
              {/* Workspace Toolbar */}
              <div className="no-print flex flex-wrap items-center justify-between gap-4">
                <div className="flex items-center gap-1 p-1 bg-slate-200/70 rounded-md">
                  <button
                    type="button"
                    onClick={() => setActiveView('report')}
                    className={`px-3.5 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 ${
                      activeView === 'report'
                        ? 'bg-white text-slate-900 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Formatted Report
                  </button>
                  <button
                    type="button"
                    onClick={() => setActiveView('spreadsheet')}
                    className={`px-3.5 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 ${
                      activeView === 'spreadsheet'
                        ? 'bg-white text-slate-900 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Spreadsheet Data
                  </button>
                  <button
                    type="button"
                    onClick={() => setActiveView('template')}
                    className={`px-3.5 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 ${
                      activeView === 'template'
                        ? 'bg-white text-slate-900 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Template &amp; Placeholders
                  </button>
                </div>

                {workbook && (
                  <div className="text-xs text-slate-500 tabular-nums">
                    Source: <span className="font-medium text-slate-700">{workbook.fileName}</span>
                    <span className="mx-1.5" aria-hidden="true">·</span>
                    Template: <span className="font-medium text-slate-700">{template.name}</span>
                  </div>
                )}
              </div>

              {errorMessage && (
                <div className="no-print flex items-start gap-2.5 px-4 py-3 bg-red-50 border border-red-200 rounded-md text-xs text-red-800">
                  <AlertCircle className="w-4 h-4 text-red-600 shrink-0 mt-0.5" />
                  <span>{errorMessage}</span>
                </div>
              )}

              {activeView === 'report' && (
                <ReportDocumentView
                  report={report}
                  template={template}
                  isGenerating={isGenerating}
                  hasWorkbook={Boolean(workbook)}
                  onGenerate={handleGenerateReport}
                  onUpdateReport={setReport}
                  onRefineSection={handleRefineSection}
                />
              )}

              {activeView === 'spreadsheet' && (
                <SpreadsheetInspector
                  workbook={workbook}
                  selectedSheets={selectedSheets}
                  onToggleSheet={handleToggleSheet}
                  onUploadExcel={handleExcelFile}
                  onLoadSample={handleLoadSampleWorkbook}
                />
              )}

              {activeView === 'template' && (
                <TemplateEditor
                  template={template}
                  report={report}
                  onUploadTemplateFile={handleTemplateFile}
                  onApplyTextTemplate={handleApplyTextTemplate}
                  onUpdatePlaceholder={handleUpdatePlaceholder}
                />
              )}
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
