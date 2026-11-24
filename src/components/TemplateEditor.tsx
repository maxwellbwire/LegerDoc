import { useState } from 'react';
import { FileText, Upload } from 'lucide-react';
import type { GeneratedReport, ParsedTemplate, PlaceholderMapping } from '../types/report';
import { PRESET_TEMPLATES } from '../utils/templateParser';

interface TemplateEditorProps {
  template: ParsedTemplate | null;
  report: GeneratedReport | null;
  onUploadTemplateFile: (file: File) => void;
  onApplyTextTemplate: (name: string, text: string, kind: 'text' | 'preset') => void;
  onUpdatePlaceholder: (key: string, value: string) => void;
}

export function TemplateEditor({
  template,
  report,
  onUploadTemplateFile,
  onApplyTextTemplate,
  onUpdatePlaceholder,
}: TemplateEditorProps) {
  const [draftName, setDraftName] = useState(template?.name ?? 'Custom Document Template');
  const [draftText, setDraftText] = useState(template?.rawText ?? PRESET_TEMPLATES[0].content);
  const [mode, setMode] = useState<'preview' | 'edit'>('preview');

  const placeholderMap = new Map<string, string>();
  if (report) {
    for (const p of report.placeholderValues) {
      placeholderMap.set(p.key, p.value);
    }
  }

  return (
    <div className="space-y-6">
      <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4 pb-4 border-b border-slate-200">
        <div>
          <h2 className="text-lg font-semibold text-slate-900">
            {template ? template.name : 'Document Template Configuration'}
          </h2>
          <p className="text-xs text-slate-500 mt-0.5">
            Upload a .docx exported from Google Docs, paste template text, or start from a standard structure.
          </p>
        </div>

        <div className="flex items-center gap-2">
          <div className="flex items-center gap-1 p-1 bg-slate-100 rounded-md">
            <button
              type="button"
              onClick={() => setMode('preview')}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 ${
                mode === 'preview'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Structure & Placeholders
            </button>
            <button
              type="button"
              onClick={() => {
                if (template) {
                  setDraftName(template.name);
                  setDraftText(template.rawText);
                }
                setMode('edit');
              }}
              className={`px-3 py-1.5 text-xs font-medium rounded transition-colors whitespace-nowrap shrink-0 ${
                mode === 'edit'
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Paste / Edit Text
            </button>
          </div>

          <label className="cursor-pointer inline-flex items-center gap-2 px-3.5 py-2 text-xs font-medium text-white bg-blue-700 rounded-md hover:bg-blue-800 transition-colors whitespace-nowrap shrink-0">
            <Upload className="w-3.5 h-3.5" />
            <span>Upload .docx / .md</span>
            <input
              type="file"
              accept=".docx,.txt,.md,.html"
              className="sr-only"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) {
                  onUploadTemplateFile(f);
                  setMode('preview');
                }
                e.currentTarget.value = '';
              }}
            />
          </label>
        </div>
      </div>

      {mode === 'edit' ? (
        <div className="bg-white border border-slate-200 rounded-lg p-6 space-y-4">
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1.5">Template Name</label>
            <input
              type="text"
              value={draftName}
              onChange={(e) => setDraftName(e.target.value)}
              className="w-full px-3 py-2 text-sm border border-slate-300 rounded-md focus:outline-none focus:border-blue-600"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-slate-700 mb-1.5">
              Template Content (supports Markdown headings and {'{{placeholder}}'} variables)
            </label>
            <textarea
              value={draftText}
              onChange={(e) => setDraftText(e.target.value)}
              rows={15}
              className="w-full px-3 py-2.5 text-sm font-mono border border-slate-300 rounded-md focus:outline-none focus:border-blue-600 leading-relaxed"
              placeholder="Paste your Google Doc template structure here..."
            />
          </div>
          <div className="flex items-center justify-between pt-2">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="text-xs text-slate-500">Load preset:</span>
              {PRESET_TEMPLATES.map((preset) => (
                <button
                  key={preset.id}
                  type="button"
                  onClick={() => {
                    setDraftName(preset.name);
                    setDraftText(preset.content);
                  }}
                  className="text-xs text-blue-700 hover:underline font-medium whitespace-nowrap shrink-0"
                >
                  {preset.name}
                </button>
              ))}
            </div>
            <button
              type="button"
              onClick={() => {
                onApplyTextTemplate(draftName.trim() || 'Custom Template', draftText, 'text');
                setMode('preview');
              }}
              className="px-4 py-2 text-xs font-medium text-white bg-slate-900 rounded-md hover:bg-slate-800 transition-colors whitespace-nowrap shrink-0"
            >
              Apply Template
            </button>
          </div>
        </div>
      ) : !template ? (
        <div className="bg-white border border-slate-200 rounded-lg p-12 text-center">
          <div className="w-11 h-11 rounded-lg bg-slate-100 text-slate-700 flex items-center justify-center mx-auto mb-4">
            <FileText className="w-5 h-5" />
          </div>
          <h3 className="text-base font-semibold text-slate-900 mb-2">Select or Upload a Document Template</h3>
          <p className="text-sm text-slate-600 max-w-lg mx-auto mb-6">
            In Google Docs, open your template and choose File &rarr; Download &rarr; Microsoft Word
            (.docx), or paste your template headings directly.
          </p>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 max-w-3xl mx-auto text-left">
            {PRESET_TEMPLATES.map((preset) => (
              <button
                key={preset.id}
                type="button"
                onClick={() => onApplyTextTemplate(preset.name, preset.content, 'preset')}
                className="p-4 border border-slate-200 rounded-lg hover:border-blue-600 hover:bg-slate-50/70 transition-colors text-left"
              >
                <div className="text-sm font-semibold text-slate-900">{preset.name}</div>
                <p className="text-xs text-slate-500 mt-1.5 leading-relaxed">{preset.description}</p>
              </button>
            ))}
          </div>
        </div>
      ) : (
        <div className="grid grid-cols-1 xl:grid-cols-12 gap-6 items-start">
          <div className="xl:col-span-7 bg-white border border-slate-200 rounded-lg p-6">
            <div className="flex items-center justify-between pb-3 mb-4 border-b border-slate-100">
              <span className="text-xs font-semibold text-slate-700">Template Content Preview</span>
              <span className="text-xs text-slate-500">
                {template.sourceKind === 'docx'
                  ? 'Original .docx formatting preserved for export'
                  : 'Structured text template'}
              </span>
            </div>
            <div
              className="prose prose-sm max-w-none text-slate-700 space-y-2 [&_h1]:text-xl [&_h1]:font-bold [&_h1]:text-slate-900 [&_h1]:mb-2 [&_h2]:text-base [&_h2]:font-semibold [&_h2]:text-slate-900 [&_h2]:mt-4 [&_h2]:mb-1.5 [&_h3]:text-sm [&_h3]:font-semibold [&_h3]:text-slate-800 [&_p]:text-sm [&_p]:leading-relaxed [&_table]:w-full [&_table]:border-collapse [&_td]:border [&_td]:border-slate-200 [&_td]:p-2 [&_th]:border [&_th]:border-slate-200 [&_th]:p-2 [&_th]:bg-slate-50"
              dangerouslySetInnerHTML={{ __html: template.htmlPreview }}
            />
          </div>

          <div className="xl:col-span-5 space-y-6">
            <div className="bg-white border border-slate-200 rounded-lg p-5">
              <h3 className="text-sm font-semibold text-slate-900 mb-3">
                Detected Section Headings ({template.headings.length})
              </h3>
              {template.headings.length === 0 ? (
                <p className="text-xs text-slate-500">
                  No explicit headings detected; Gemini will organize sections automatically from your Excel sheets.
                </p>
              ) : (
                <ol className="space-y-1.5 text-sm text-slate-700">
                  {template.headings.map((h, i) => (
                    <li key={`${h}-${i}`} className="flex items-baseline gap-2.5">
                      <span className="text-xs font-mono text-slate-400 tabular-nums w-5">
                        {String(i + 1).padStart(2, '0')}.
                      </span>
                      <span>{h}</span>
                    </li>
                  ))}
                </ol>
              )}
            </div>

            <div className="bg-white border border-slate-200 rounded-lg p-5">
              <h3 className="text-sm font-semibold text-slate-900 mb-1">
                Detected Placeholders ({template.placeholders.length})
              </h3>
              <p className="text-xs text-slate-500 mb-4">
                Variables formatted as <code className="font-mono text-slate-700">{'{{variable_name}}'}</code> are
                populated automatically from your workbook.
              </p>

              {template.placeholders.length === 0 ? (
                <p className="text-xs text-slate-500">
                  No {'{{placeholder}}'} tags found. When you generate a report, sections will be matched to the template headings above.
                </p>
              ) : (
                <div className="space-y-3">
                  {template.placeholders.map((key) => {
                    const mappedVal: PlaceholderMapping['value'] = placeholderMap.get(key) ?? '';
                    return (
                      <div key={key} className="space-y-1">
                        <label className="block text-xs font-mono text-slate-700">
                          {`{{${key}}}`}
                        </label>
                        <input
                          type="text"
                          value={mappedVal}
                          disabled={!report}
                          onChange={(e) => onUpdatePlaceholder(key, e.target.value)}
                          placeholder={report ? 'Enter value...' : 'Populated after report generation'}
                          className="w-full px-3 py-1.5 text-sm border border-slate-200 rounded-md bg-slate-50/50 focus:bg-white focus:outline-none focus:border-blue-600 disabled:text-slate-400"
                        />
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
