import { useState } from 'react';
import {
  Check,
  Copy,
  Download,
  Edit3,
  FileDown,
  Printer,
  RefreshCw,
  Sparkles,
} from 'lucide-react';
import type { GeneratedReport, ParsedTemplate, ReportSection } from '../types/report';
import {
  canFillUploadedDocxTemplate,
  copyRichReportForGoogleDocs,
  exportFilledDocxTemplate,
  exportMarkdownReport,
  exportStructuredDocx,
} from '../utils/docExporter';

interface ReportDocumentViewProps {
  report: GeneratedReport | null;
  template: ParsedTemplate | null;
  isGenerating: boolean;
  hasWorkbook: boolean;
  onGenerate: () => void;
  onUpdateReport: (updated: GeneratedReport) => void;
  onRefineSection: (sectionId: string, instruction: string) => Promise<void>;
}

export function ReportDocumentView({
  report,
  template,
  isGenerating,
  hasWorkbook,
  onGenerate,
  onUpdateReport,
  onRefineSection,
}: ReportDocumentViewProps) {
  const [copied, setCopied] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);
  const [editingTitle, setEditingTitle] = useState(false);
  const [editingSummary, setEditingSummary] = useState(false);
  const [refiningSectionId, setRefiningSectionId] = useState<string | null>(null);
  const [refinePrompt, setRefinePrompt] = useState('');
  const [isRefining, setIsRefining] = useState(false);
  const [editingSectionId, setEditingSectionId] = useState<string | null>(null);

  const handleCopyForGoogleDocs = async () => {
    if (!report) return;
    setExportError(null);
    try {
      await copyRichReportForGoogleDocs(report);
      setCopied(true);
      setTimeout(() => setCopied(false), 2600);
    } catch (err: any) {
      setExportError(err?.message || 'Clipboard copy failed. Please try downloading .docx instead.');
    }
  };

  const handleDownloadStructuredDocx = async () => {
    if (!report) return;
    setExportError(null);
    try {
      await exportStructuredDocx(report);
    } catch (err: any) {
      setExportError(err?.message || 'Could not generate .docx file.');
    }
  };

  const handleDownloadFilledTemplate = () => {
    if (!report || !template) return;
    setExportError(null);
    try {
      exportFilledDocxTemplate(template, report);
    } catch (err: any) {
      setExportError(err?.message || 'Failed to inject placeholders into uploaded .docx template.');
    }
  };

  const updateSection = (sectionId: string, updater: (sec: ReportSection) => ReportSection) => {
    if (!report) return;
    onUpdateReport({
      ...report,
      sections: report.sections.map((s) => (s.id === sectionId ? updater(s) : s)),
    });
  };

  const handleConfirmRefine = async (sectionId: string) => {
    if (!refinePrompt.trim()) return;
    setIsRefining(true);
    try {
      await onRefineSection(sectionId, refinePrompt.trim());
      setRefiningSectionId(null);
      setRefinePrompt('');
    } finally {
      setIsRefining(false);
    }
  };

  if (isGenerating) {
    return (
      <div className="bg-white border border-slate-200 rounded-lg p-10 max-w-[820px] mx-auto space-y-6">
        <div className="space-y-3">
          <div className="h-7 w-2/3 bg-slate-100 rounded animate-pulse" />
          <div className="h-4 w-1/3 bg-slate-100 rounded animate-pulse" />
        </div>
        <div className="h-px bg-slate-100" />
        <div className="space-y-2.5">
          <div className="h-4 w-full bg-slate-100 rounded animate-pulse" />
          <div className="h-4 w-11/12 bg-slate-100 rounded animate-pulse" />
          <div className="h-4 w-4/5 bg-slate-100 rounded animate-pulse" />
        </div>
        <div className="grid grid-cols-3 gap-4 pt-2">
          <div className="h-20 bg-slate-50 border border-slate-100 rounded-md animate-pulse" />
          <div className="h-20 bg-slate-50 border border-slate-100 rounded-md animate-pulse" />
          <div className="h-20 bg-slate-50 border border-slate-100 rounded-md animate-pulse" />
        </div>
        <p className="text-xs text-slate-500 text-center pt-2">
          Reading spreadsheet metrics and formatting report into your document template...
        </p>
      </div>
    );
  }

  if (!report) {
    return (
      <div className="bg-white border border-slate-200 rounded-lg p-12 max-w-[820px] mx-auto text-center mt-4">
        <h2 className="text-lg font-semibold text-slate-900 mb-2">Ready to Generate Report</h2>
        <p className="text-sm text-slate-600 max-w-md mx-auto mb-6 leading-relaxed">
          {hasWorkbook
            ? 'Your Excel workbook and document template are ready. Generate a formatted report with verified metrics, narrative analysis, and tables.'
            : 'Load your Excel workbook and select a document template in the left workflow panel to generate your formatted report.'}
        </p>
        <button
          type="button"
          disabled={!hasWorkbook || !template}
          onClick={onGenerate}
          className="inline-flex items-center gap-2 px-5 py-2.5 text-sm font-medium text-white bg-blue-700 rounded-md hover:bg-blue-800 disabled:bg-slate-300 disabled:cursor-not-allowed transition-colors whitespace-nowrap shrink-0"
        >
          <Sparkles className="w-4 h-4" />
          <span>Generate Report</span>
        </button>
      </div>
    );
  }

  const canFillTemplate = canFillUploadedDocxTemplate(template, report);

  return (
    <div className="space-y-5">
      {/* Export Toolbar */}
      <div className="no-print flex flex-wrap items-center justify-between gap-3 bg-white border border-slate-200 rounded-lg px-4 py-3 max-w-[820px] mx-auto">
        <div className="text-xs text-slate-500">
          Generated {new Date(report.generatedAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
          <span className="mx-1.5" aria-hidden="true">·</span>
          Click any text in the document below to edit
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={handleCopyForGoogleDocs}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-white bg-blue-700 rounded-md hover:bg-blue-800 transition-colors whitespace-nowrap shrink-0"
          >
            {copied ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
            <span>{copied ? 'Copied Rich Formatting' : 'Copy for Google Docs'}</span>
          </button>

          {canFillTemplate && (
            <button
              type="button"
              onClick={handleDownloadFilledTemplate}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-900 bg-slate-100 border border-slate-300 rounded-md hover:bg-slate-200 transition-colors whitespace-nowrap shrink-0"
            >
              <FileDown className="w-3.5 h-3.5" />
              <span>Download Filled Template (.docx)</span>
            </button>
          )}

          <button
            type="button"
            onClick={handleDownloadStructuredDocx}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0"
          >
            <Download className="w-3.5 h-3.5" />
            <span>Download .docx</span>
          </button>

          <button
            type="button"
            onClick={() => exportMarkdownReport(report)}
            className="px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0"
          >
            .md
          </button>

          <button
            type="button"
            onClick={() => window.print()}
            aria-label="Print or Save as PDF"
            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0"
          >
            <Printer className="w-3.5 h-3.5" />
            <span>PDF</span>
          </button>
        </div>
      </div>

      {exportError && (
        <div className="no-print max-w-[820px] mx-auto px-4 py-3 bg-red-50 border border-red-200 rounded-md text-xs text-red-700">
          {exportError}
        </div>
      )}

      {/* Editorial Document Sheet */}
      <article className="print-document-sheet bg-white border border-slate-200 rounded-lg max-w-[820px] mx-auto px-8 py-10 md:px-12 md:py-12 space-y-8 shadow-xs">
        {/* Document Header */}
        <header className="pb-6 border-b border-slate-200">
          {editingTitle ? (
            <div className="space-y-2">
              <input
                type="text"
                value={report.reportTitle}
                onChange={(e) => onUpdateReport({ ...report, reportTitle: e.target.value })}
                className="w-full text-2xl font-bold text-slate-900 border border-blue-500 rounded px-2 py-1 focus:outline-none"
              />
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  value={report.subtitle}
                  onChange={(e) => onUpdateReport({ ...report, subtitle: e.target.value })}
                  className="flex-1 text-sm text-slate-600 border border-slate-300 rounded px-2 py-1 focus:outline-none"
                />
                <button
                  type="button"
                  onClick={() => setEditingTitle(false)}
                  className="px-3 py-1 text-xs font-medium text-white bg-slate-900 rounded whitespace-nowrap shrink-0"
                >
                  Done
                </button>
              </div>
            </div>
          ) : (
            <div className="group relative">
              <div className="flex items-start justify-between gap-4">
                <h1 className="text-2xl md:text-3xl font-bold text-slate-900 tracking-tight text-balance">
                  {report.reportTitle}
                </h1>
                <button
                  type="button"
                  onClick={() => setEditingTitle(true)}
                  className="no-print opacity-0 group-hover:opacity-100 focus:opacity-100 p-1.5 text-slate-400 hover:text-slate-700 transition-opacity"
                  title="Edit title"
                >
                  <Edit3 className="w-4 h-4" />
                </button>
              </div>
              <p className="text-sm text-slate-500 mt-2">
                {report.subtitle}
                <span className="mx-2" aria-hidden="true">·</span>
                <span>{report.preparedDate}</span>
              </p>
            </div>
          )}
        </header>

        {/* Executive Summary */}
        <section className="space-y-2.5">
          <div className="flex items-center justify-between">
            <h2 className="text-base font-semibold text-slate-900">Executive Summary</h2>
            <button
              type="button"
              onClick={() => setEditingSummary(!editingSummary)}
              className="no-print text-xs text-slate-500 hover:text-slate-900 whitespace-nowrap shrink-0"
            >
              {editingSummary ? 'Done' : 'Edit'}
            </button>
          </div>
          {editingSummary ? (
            <textarea
              value={report.executiveSummary}
              onChange={(e) => onUpdateReport({ ...report, executiveSummary: e.target.value })}
              rows={4}
              className="w-full text-sm text-slate-700 leading-relaxed border border-blue-500 rounded-md p-3 focus:outline-none"
            />
          ) : (
            <p className="text-[15px] text-slate-700 leading-relaxed max-w-[70ch]">
              {report.executiveSummary}
            </p>
          )}
        </section>

        {/* Key Metrics Strip */}
        {report.keyMetrics.length > 0 && (
          <section className="space-y-3">
            <h2 className="text-base font-semibold text-slate-900">Key Metrics</h2>
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4 border-y border-slate-200 py-5">
              {report.keyMetrics.map((metric, idx) => (
                <div key={idx} className="space-y-1">
                  <div className="text-xs font-medium text-slate-500">{metric.label}</div>
                  <div className="text-xl font-bold text-slate-900 font-mono tabular-nums">
                    {metric.value}
                  </div>
                  <div className="text-xs text-slate-600">{metric.context}</div>
                </div>
              ))}
            </div>
          </section>
        )}

        {/* Ordered Report Sections */}
        <div className="space-y-8">
          {report.sections.map((section) => {
            const isEditing = editingSectionId === section.id;
            const isRefiningThis = refiningSectionId === section.id;

            return (
              <section key={section.id} className="space-y-3 pt-2">
                <div className="flex items-center justify-between gap-4">
                  {isEditing ? (
                    <input
                      type="text"
                      value={section.heading}
                      onChange={(e) =>
                        updateSection(section.id, (s) => ({ ...s, heading: e.target.value }))
                      }
                      className="flex-1 text-lg font-semibold text-slate-900 border border-blue-500 rounded px-2 py-1 focus:outline-none"
                    />
                  ) : (
                    <h2 className="text-lg font-semibold text-slate-900">{section.heading}</h2>
                  )}

                  <div className="no-print flex items-center gap-3 shrink-0">
                    <button
                      type="button"
                      onClick={() => {
                        setRefiningSectionId(isRefiningThis ? null : section.id);
                        setRefinePrompt('');
                      }}
                      className="inline-flex items-center gap-1 text-xs font-medium text-blue-700 hover:text-blue-800 whitespace-nowrap shrink-0"
                    >
                      <RefreshCw className="w-3 h-3" />
                      <span>Refine</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setEditingSectionId(isEditing ? null : section.id)}
                      className="text-xs font-medium text-slate-500 hover:text-slate-900 whitespace-nowrap shrink-0"
                    >
                      {isEditing ? 'Done' : 'Edit'}
                    </button>
                  </div>
                </div>

                {isRefiningThis && (
                  <div className="no-print p-3 bg-slate-50 border border-slate-200 rounded-md flex items-center gap-2">
                    <input
                      type="text"
                      value={refinePrompt}
                      onChange={(e) => setRefinePrompt(e.target.value)}
                      placeholder="e.g., Make this more concise or emphasize LATAM margin variance..."
                      className="flex-1 px-3 py-1.5 text-xs bg-white border border-slate-300 rounded focus:outline-none focus:border-blue-600"
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') handleConfirmRefine(section.id);
                      }}
                    />
                    <button
                      type="button"
                      disabled={isRefining || !refinePrompt.trim()}
                      onClick={() => handleConfirmRefine(section.id)}
                      className="px-3 py-1.5 text-xs font-medium text-white bg-blue-700 rounded hover:bg-blue-800 disabled:bg-slate-300 whitespace-nowrap shrink-0"
                    >
                      {isRefining ? 'Updating...' : 'Apply'}
                    </button>
                  </div>
                )}

                {isEditing ? (
                  <div className="space-y-3">
                    <textarea
                      value={section.narrative}
                      onChange={(e) =>
                        updateSection(section.id, (s) => ({ ...s, narrative: e.target.value }))
                      }
                      rows={4}
                      className="w-full text-sm text-slate-700 leading-relaxed border border-blue-500 rounded-md p-3 focus:outline-none"
                    />
                    <div className="space-y-1.5">
                      <div className="text-xs font-medium text-slate-500">Bullet Points (one per line)</div>
                      <textarea
                        value={section.bullets.join('\n')}
                        onChange={(e) =>
                          updateSection(section.id, (s) => ({
                            ...s,
                            bullets: e.target.value.split('\n').filter((line) => line.trim() !== ''),
                          }))
                        }
                        rows={3}
                        className="w-full text-sm text-slate-700 border border-slate-300 rounded-md p-2.5 focus:outline-none"
                      />
                    </div>
                  </div>
                ) : (
                  <>
                    {section.narrative && (
                      <div className="space-y-3 text-[15px] text-slate-700 leading-relaxed max-w-[70ch]">
                        {section.narrative.split(/\n{2,}/).map((para, pIdx) => (
                          <p key={pIdx}>{para.trim()}</p>
                        ))}
                      </div>
                    )}

                    {section.bullets.length > 0 && (
                      <ul className="list-disc pl-5 space-y-1.5 text-[15px] text-slate-700 max-w-[70ch]">
                        {section.bullets.map((b, bIdx) => (
                          <li key={bIdx}>{b}</li>
                        ))}
                      </ul>
                    )}
                  </>
                )}

                {section.table && section.table.headers.length > 0 && section.table.rows.length > 0 && (
                  <div className="pt-2 space-y-2">
                    {section.table.caption && (
                      <div className="text-xs text-slate-500">{section.table.caption}</div>
                    )}
                    <div className="overflow-x-auto border border-slate-200 rounded-md">
                      <table className="w-full border-collapse text-left">
                        <thead className="bg-slate-50 border-b border-slate-200">
                          <tr>
                            {section.table.headers.map((header, hIdx) => {
                              const sampleCell = section.table!.rows[0]?.[hIdx] ?? '';
                              const looksNumeric = /^[$€£\-+]?[\d,]+(\.\d+)?%?$/.test(sampleCell.trim());
                              return (
                                <th
                                  key={hIdx}
                                  className={`py-2.5 px-3 text-xs font-semibold text-slate-700 whitespace-nowrap ${
                                    looksNumeric ? 'text-right' : 'text-left'
                                  }`}
                                >
                                  {header}
                                </th>
                              );
                            })}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {section.table.rows.map((row, rIdx) => (
                            <tr key={rIdx} className="hover:bg-slate-50/70">
                              {section.table!.headers.map((_, cIdx) => {
                                const val = row[cIdx] ?? '';
                                const looksNumeric = /^[$€£\-+]?[\d,]+(\.\d+)?%?$/.test(val.trim());
                                return (
                                  <td
                                    key={cIdx}
                                    className={`py-2 px-3 text-sm text-slate-700 whitespace-nowrap ${
                                      looksNumeric ? 'text-right font-mono tabular-nums' : 'text-left'
                                    }`}
                                  >
                                    {val}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </section>
            );
          })}
        </div>
      </article>
    </div>
  );
}
