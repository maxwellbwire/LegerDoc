import { useMemo, useState } from 'react';
import { FileSpreadsheet, Search, Upload } from 'lucide-react';
import type { ParsedWorkbook } from '../types/report';

interface SpreadsheetInspectorProps {
  workbook: ParsedWorkbook | null;
  selectedSheets: string[];
  onToggleSheet: (sheetName: string) => void;
  onUploadExcel: (file: File) => void;
  onLoadSample: () => void;
}

export function SpreadsheetInspector({
  workbook,
  selectedSheets,
  onToggleSheet,
  onUploadExcel,
  onLoadSample,
}: SpreadsheetInspectorProps) {
  const [activeSheetName, setActiveSheetName] = useState<string>('');
  const [searchQuery, setSearchQuery] = useState<string>('');

  const activeSheet = useMemo(() => {
    if (!workbook || workbook.sheets.length === 0) return null;
    return workbook.sheets.find((s) => s.name === activeSheetName) ?? workbook.sheets[0];
  }, [workbook, activeSheetName]);

  const filteredRows = useMemo(() => {
    if (!activeSheet) return [];
    const q = searchQuery.trim().toLowerCase();
    if (!q) return activeSheet.rows;
    return activeSheet.rows.filter((row) =>
      activeSheet.headers.some((h) => String(row[h] ?? '').toLowerCase().includes(q))
    );
  }, [activeSheet, searchQuery]);

  if (!workbook) {
    return (
      <div className="bg-white border border-slate-200 rounded-lg p-12 text-center max-w-2xl mx-auto mt-8">
        <div className="w-11 h-11 rounded-lg bg-slate-100 text-slate-700 flex items-center justify-center mx-auto mb-4">
          <FileSpreadsheet className="w-5 h-5" />
        </div>
        <h2 className="text-lg font-semibold text-slate-900 mb-2">No Excel Workbook Loaded</h2>
        <p className="text-sm text-slate-600 max-w-md mx-auto mb-6 leading-relaxed">
          Upload an Excel workbook (.xlsx, .xls, or .csv) to inspect sheets, compute column
          summaries, and verify every figure before generating your report.
        </p>
        <div className="flex items-center justify-center gap-3 flex-wrap">
          <label className="cursor-pointer inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-blue-700 rounded-md hover:bg-blue-800 transition-colors whitespace-nowrap shrink-0">
            <Upload className="w-4 h-4" />
            <span>Upload Excel File</span>
            <input
              type="file"
              accept=".xlsx,.xls,.csv,.ods"
              className="sr-only"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) onUploadExcel(f);
                e.currentTarget.value = '';
              }}
            />
          </label>
          <button
            type="button"
            onClick={onLoadSample}
            className="px-4 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0"
          >
            Load Sample Workbook
          </button>
        </div>
      </div>
    );
  }

  const numericColumns = activeSheet?.columnSummaries.filter((c) => c.type === 'number') ?? [];

  return (
    <div className="space-y-6">
      <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 pb-4 border-b border-slate-200">
        <div>
          <h2 className="text-lg font-semibold text-slate-900">{workbook.fileName}</h2>
          <p className="text-xs text-slate-500 mt-0.5 tabular-nums">
            {workbook.sheets.length} sheet{workbook.sheets.length === 1 ? '' : 's'}
            <span className="mx-1.5" aria-hidden="true">·</span>
            {(workbook.fileSize / 1024).toFixed(1)} KB
            <span className="mx-1.5" aria-hidden="true">·</span>
            {selectedSheets.length} selected for report
          </p>
        </div>

        <div className="flex items-center gap-3">
          <div className="relative">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Filter rows..."
              className="pl-9 pr-3 py-1.5 text-sm bg-white border border-slate-300 rounded-md text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-blue-600 w-56"
            />
          </div>
          <label className="cursor-pointer px-3.5 py-1.5 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0">
            Replace File
            <input
              type="file"
              accept=".xlsx,.xls,.csv,.ods"
              className="sr-only"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) onUploadExcel(f);
                e.currentTarget.value = '';
              }}
            />
          </label>
        </div>
      </div>

      <div className="flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-md">
          {workbook.sheets.map((sheet) => {
            const isActive = activeSheet?.name === sheet.name;
            return (
              <button
                key={sheet.name}
                type="button"
                onClick={() => setActiveSheetName(sheet.name)}
                className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 ${
                  isActive
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                {sheet.name} ({sheet.rowCount})
              </button>
            );
          })}
        </div>

        {activeSheet && (
          <label className="inline-flex items-center gap-2 text-xs text-slate-700 cursor-pointer select-none">
            <input
              type="checkbox"
              checked={selectedSheets.includes(activeSheet.name)}
              onChange={() => onToggleSheet(activeSheet.name)}
              className="w-4 h-4 rounded border-slate-300 text-blue-700 focus:ring-blue-600"
            />
            <span>Include "{activeSheet.name}" in AI report synthesis</span>
          </label>
        )}
      </div>

      {numericColumns.length > 0 && (
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          {numericColumns.slice(0, 4).map((col) => (
            <div key={col.name} className="bg-white border border-slate-200 rounded-lg p-4">
              <div className="text-xs text-slate-500 truncate" title={col.name}>
                {col.name}
              </div>
              <div className="text-lg font-semibold text-slate-900 font-mono tabular-nums mt-1">
                {col.sum !== undefined
                  ? col.name.includes('%')
                    ? `${(col.average ?? 0).toFixed(1)}% avg`
                    : col.sum.toLocaleString('en-US', { maximumFractionDigits: 1 })
                  : '—'}
              </div>
              <div className="text-xs text-slate-500 font-mono tabular-nums mt-1">
                Min {col.min?.toLocaleString('en-US', { maximumFractionDigits: 1 })}
                <span className="mx-1" aria-hidden="true">/</span>
                Max {col.max?.toLocaleString('en-US', { maximumFractionDigits: 1 })}
              </div>
            </div>
          ))}
        </div>
      )}

      {activeSheet && (
        <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          <div className="overflow-x-auto max-h-[520px]">
            <table className="w-full border-collapse text-left">
              <thead className="bg-slate-50 sticky top-0 z-10 border-b border-slate-200">
                <tr>
                  <th className="py-2.5 px-3 text-xs font-medium text-slate-500 w-12 font-mono tabular-nums">
                    #
                  </th>
                  {activeSheet.headers.map((header) => {
                    const summary = activeSheet.columnSummaries.find((c) => c.name === header);
                    const isNum = summary?.type === 'number';
                    return (
                      <th
                        key={header}
                        className={`py-2.5 px-3 text-xs font-semibold text-slate-700 whitespace-nowrap ${
                          isNum ? 'text-right' : 'text-left'
                        }`}
                      >
                        {header}
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredRows.length === 0 ? (
                  <tr>
                    <td
                      colSpan={activeSheet.headers.length + 1}
                      className="py-10 text-center text-sm text-slate-500"
                    >
                      No rows match "{searchQuery}".
                    </td>
                  </tr>
                ) : (
                  filteredRows.slice(0, 250).map((row, rIdx) => (
                    <tr key={rIdx} className="hover:bg-slate-50/80 transition-colors h-10">
                      <td className="py-2 px-3 text-xs text-slate-400 font-mono tabular-nums">
                        {rIdx + 1}
                      </td>
                      {activeSheet.headers.map((header) => {
                        const summary = activeSheet.columnSummaries.find((c) => c.name === header);
                        const isNum = summary?.type === 'number';
                        const rawVal = row[header];
                        const formatted =
                          typeof rawVal === 'number'
                            ? rawVal.toLocaleString('en-US', { maximumFractionDigits: 2 })
                            : rawVal === null || rawVal === undefined
                            ? ''
                            : String(rawVal);

                        return (
                          <td
                            key={header}
                            className={`py-2 px-3 text-sm text-slate-700 whitespace-nowrap ${
                              isNum ? 'text-right font-mono tabular-nums' : 'text-left'
                            }`}
                          >
                            {formatted}
                          </td>
                        );
                      })}
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
          {filteredRows.length > 250 && (
            <div className="px-4 py-2.5 bg-slate-50 border-t border-slate-200 text-xs text-slate-500 tabular-nums">
              Showing first 250 of {filteredRows.length} matching rows.
            </div>
          )}
        </div>
      )}
    </div>
  );
}
