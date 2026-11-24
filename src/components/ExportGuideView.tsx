import { ArrowRight, Copy, FileSpreadsheet, FileText } from 'lucide-react';

interface ExportGuideViewProps {
  onReturnToStudio: () => void;
}

export function ExportGuideView({ onReturnToStudio }: ExportGuideViewProps) {
  return (
    <div className="max-w-4xl mx-auto space-y-8 py-4">
      <div className="border-b border-slate-200 pb-5">
        <h1 className="text-2xl font-bold text-slate-900">
          Excel to Google Docs Workflow Guide
        </h1>
        <p className="text-sm text-slate-600 mt-2 max-w-2xl leading-relaxed">
          Automate reading Excel spreadsheets into your Google Docs report template without
          connecting any external accounts. Bring your template in once, generate verified reports in
          seconds, and paste or import directly into Google Docs.
        </p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="bg-white border border-slate-200 rounded-lg p-6 space-y-3">
          <div className="w-9 h-9 rounded-md bg-slate-100 text-slate-800 flex items-center justify-center">
            <FileText className="w-4 h-4" />
          </div>
          <h2 className="text-base font-semibold text-slate-900">
            01. Bring Your Google Doc Template
          </h2>
          <p className="text-sm text-slate-600 leading-relaxed">
            In Google Docs, select <strong>File &rarr; Download &rarr; Microsoft Word (.docx)</strong>,
            or simply copy and paste your document headings into the Template tab. Add optional{' '}
            <code className="font-mono text-xs text-slate-800">{'{{placeholder}}'}</code> tags where
            exact metrics should land.
          </p>
        </div>

        <div className="bg-white border border-slate-200 rounded-lg p-6 space-y-3">
          <div className="w-9 h-9 rounded-md bg-slate-100 text-slate-800 flex items-center justify-center">
            <FileSpreadsheet className="w-4 h-4" />
          </div>
          <h2 className="text-base font-semibold text-slate-900">
            02. Upload Excel &amp; Generate
          </h2>
          <p className="text-sm text-slate-600 leading-relaxed">
            Drop any <code className="font-mono text-xs text-slate-800">.xlsx</code>,{' '}
            <code className="font-mono text-xs text-slate-800">.xls</code>, or{' '}
            <code className="font-mono text-xs text-slate-800">.csv</code> workbook. LedgerDoc
            parses every sheet, computes column totals and averages locally, and formats a complete
            report matching your template structure.
          </p>
        </div>

        <div className="bg-white border border-slate-200 rounded-lg p-6 space-y-3">
          <div className="w-9 h-9 rounded-md bg-slate-100 text-slate-800 flex items-center justify-center">
            <Copy className="w-4 h-4" />
          </div>
          <h2 className="text-base font-semibold text-slate-900">
            03. Transfer Back to Google Docs
          </h2>
          <p className="text-sm text-slate-600 leading-relaxed">
            Click <strong>Copy for Google Docs</strong> and paste directly into an open Google Doc
            to retain headings, tables, and bullet formatting, or click{' '}
            <strong>Download .docx</strong> to open in Google Docs with full layout preservation.
          </p>
        </div>
      </div>

      <div className="bg-white border border-slate-200 rounded-lg p-6 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
        <div className="space-y-1">
          <div className="text-sm font-semibold text-slate-900">
            Preserving Custom Headers, Logos &amp; Styles
          </div>
          <p className="text-xs text-slate-600 max-w-xl leading-relaxed">
            When you upload a <code className="font-mono">.docx</code> template containing{' '}
            <code className="font-mono">{'{{tags}}'}</code>, LedgerDoc enables{' '}
            <strong>Download Filled Template (.docx)</strong>, injecting computed values directly into
            your original document XML while keeping every font, header, and brand style intact.
          </p>
        </div>
        <button
          type="button"
          onClick={onReturnToStudio}
          className="inline-flex items-center gap-2 px-4 py-2 text-xs font-medium text-white bg-blue-700 rounded-md hover:bg-blue-800 transition-colors whitespace-nowrap shrink-0"
        >
          <span>Open Studio</span>
          <ArrowRight className="w-3.5 h-3.5" />
        </button>
      </div>
    </div>
  );
}
