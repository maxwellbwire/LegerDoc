import 'dotenv/config';
import express from 'express';
import type { Request, Response } from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { GoogleGenAI, Type } from '@google/genai';
import type { ParsedSheet, ParsedWorkbook, ReportTone } from './src/types/report.ts';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT) || 3000;
const MODEL_NAME = 'gemini-4-hydrogen';

function getGeminiClient(): GoogleGenAI {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw Object.assign(
      new Error('GEMINI_API_KEY is not configured. Add your API key in Settings > Secrets.'),
      { statusCode: 400 }
    );
  }
  return new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
}

async function withRetry<T>(fn: () => Promise<T>, maxAttempts = 3): Promise<T> {
  let lastError: unknown;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await fn();
    } catch (error: any) {
      lastError = error;
      const status = error?.status ?? error?.statusCode;
      const isRetryable = status === 503 || status === 500 || status === 429;
      if (!isRetryable || attempt === maxAttempts) {
        throw error;
      }
      const delayMs = 600 * Math.pow(2, attempt - 1);
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }
  }
  throw lastError;
}

function buildWorkbookContext(workbook: ParsedWorkbook, selectedSheetNames?: string[]): string {
  const targetSheets: ParsedSheet[] =
    selectedSheetNames && selectedSheetNames.length > 0
      ? workbook.sheets.filter((s) => selectedSheetNames.includes(s.name))
      : workbook.sheets;

  const sheetsToSerialize = targetSheets.length > 0 ? targetSheets : workbook.sheets;

  return sheetsToSerialize
    .map((sheet) => {
      const summaryLines = sheet.columnSummaries.map((col) => {
        if (col.type === 'number') {
          return `- ${col.name} (numeric, count=${col.nonEmptyCount}, sum=${col.sum?.toLocaleString('en-US', { maximumFractionDigits: 2 })}, avg=${col.average?.toLocaleString('en-US', { maximumFractionDigits: 2 })}, min=${col.min}, max=${col.max})`;
        }
        const topVals = col.topValues
          ?.slice(0, 5)
          .map((t) => `${t.value} (${t.count})`)
          .join(', ');
        return `- ${col.name} (${col.type}, unique=${col.uniqueCount ?? 0}${topVals ? `, top: ${topVals}` : ''})`;
      });

      const maxRows = 120;
      const slicedRows = sheet.rows.slice(0, maxRows);
      const csvHeader = sheet.headers.join(' | ');
      const csvRows = slicedRows.map((row) =>
        sheet.headers
          .map((h) => {
            const v = row[h];
            return v === null || v === undefined ? '' : String(v).replace(/\s+/g, ' ').trim();
          })
          .join(' | ')
      );

      return [
        `### Sheet: "${sheet.name}" (${sheet.rowCount} total rows, ${sheet.headers.length} columns)`,
        `Column Statistics:`,
        ...summaryLines,
        ``,
        `Data Rows (showing ${slicedRows.length} of ${sheet.rowCount}):`,
        csvHeader,
        ...csvRows,
      ].join('\n');
    })
    .join('\n\n---\n\n');
}

function formatGeminiError(error: any): { status: number; message: string } {
  const status = Number(error?.status ?? error?.statusCode) || 500;
  const rawMessage = String(error?.message || 'Unexpected error while generating report.');

  if (status === 403 || rawMessage.includes('PERMISSION_DENIED')) {
    return {
      status: 403,
      message: 'API permission denied. Please verify your API key in the Settings > Secrets panel.',
    };
  }
  if (status === 400 && rawMessage.includes('API_KEY_INVALID')) {
    return {
      status: 400,
      message: 'Invalid API key. Please check your key in the Settings > Secrets panel.',
    };
  }
  if (status === 429 || rawMessage.includes('RESOURCE_EXHAUSTED')) {
    return {
      status: 429,
      message: 'Rate limit or quota reached. Wait a moment and retry, or check your quota in Settings > Secrets.',
    };
  }
  return { status, message: rawMessage };
}

interface GenerateReportRequestBody {
  workbook: ParsedWorkbook;
  selectedSheetNames?: string[];
  template: {
    name: string;
    rawText: string;
    headings: string[];
    placeholders: string[];
  };
  tone?: ReportTone;
  instructions?: string;
}

interface RefineSectionRequestBody {
  workbook: ParsedWorkbook;
  selectedSheetNames?: string[];
  sectionHeading: string;
  currentNarrative: string;
  currentBullets: string[];
  refinementPrompt: string;
  tone?: ReportTone;
}

async function startServer() {
  const app = express();
  app.use(express.json({ limit: '15mb' }));

  app.get('/api/health', (_req: Request, res: Response) => {
    res.json({ ok: true, model: MODEL_NAME });
  });

  app.post('/api/generate-report', async (req: Request, res: Response) => {
    try {
      const body = req.body as GenerateReportRequestBody;
      if (!body?.workbook || !Array.isArray(body.workbook.sheets) || body.workbook.sheets.length === 0) {
        return res.status(400).json({ error: 'Please provide a valid parsed Excel workbook.' });
      }
      if (!body?.template || (!body.template.rawText && body.template.headings.length === 0)) {
        return res.status(400).json({ error: 'Please provide a document template.' });
      }

      const ai = getGeminiClient();
      const workbookContext = buildWorkbookContext(body.workbook, body.selectedSheetNames);
      const tone = body.tone || 'Executive Brief';
      const placeholders = body.template.placeholders || [];
      const headings = body.template.headings || [];

      const systemInstruction = `You are a senior financial and operations analyst converting raw Excel workbook data into polished, publication-ready business reports formatted strictly to a user's document template.

Rules:
1. Ground every metric, percentage, currency figure, variance, and ranking strictly in the provided Excel workbook data and column statistics. Never invent or hallucinate numbers not supported by the spreadsheet.
2. Mirror the structure, headings, and style of the provided Document Template.
3. Write in a "${tone}" tone: concise, authoritative, clear prose without filler or hype.
4. Format numbers consistently (currency with $ or original unit, percentages to 1 decimal place, integers with thousands separators).
5. If the template contains {{placeholder}} tags, provide a well-formatted value in placeholderValues for EVERY placeholder listed. If a placeholder refers to a narrative section (e.g. {{executive_summary}} or {{recommendations}}), write a crisp paragraph or bulleted string suitable for direct insertion into the document.`;

      const prompt = [
        `## SOURCE EXCEL WORKBOOK: ${body.workbook.fileName}`,
        workbookContext,
        ``,
        `## TARGET DOCUMENT TEMPLATE: ${body.template.name}`,
        headings.length > 0 ? `Detected Template Headings:\n${headings.map((h) => `- ${h}`).join('\n')}` : '',
        placeholders.length > 0
          ? `Detected Template Placeholders (you MUST return a value for each key):\n${placeholders.map((p) => `- {{${p}}}`).join('\n')}`
          : 'No {{placeholder}} tags detected; generate structured sections following the template headings.',
        ``,
        `Full Template Content:`,
        body.template.rawText.slice(0, 12000),
        ``,
        body.instructions ? `## USER INSTRUCTIONS:\n${body.instructions}` : '',
        ``,
        `Generate the complete structured report JSON now.`,
      ]
        .filter(Boolean)
        .join('\n');

      const response = await withRetry(() =>
        ai.models.generateContent({
          model: MODEL_NAME,
          contents: prompt,
          config: {
            systemInstruction,
            responseMimeType: 'application/json',
            responseSchema: {
              type: Type.OBJECT,
              properties: {
                reportTitle: {
                  type: Type.STRING,
                  description: 'Clear, formal report title aligned with the template and workbook.',
                },
                subtitle: {
                  type: Type.STRING,
                  description: 'Concise subtitle summarizing reporting scope or period.',
                },
                preparedDate: {
                  type: Type.STRING,
                  description: 'Formatted report date (e.g., November 23, 2026).',
                },
                executiveSummary: {
                  type: Type.STRING,
                  description: '2-3 crisp sentences synthesizing the primary findings from the Excel data.',
                },
                keyMetrics: {
                  type: Type.ARRAY,
                  description: '3 to 5 headline quantitative metrics extracted directly from the workbook.',
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      label: { type: Type.STRING },
                      value: { type: Type.STRING },
                      context: { type: Type.STRING },
                    },
                    required: ['label', 'value', 'context'],
                  },
                },
                sections: {
                  type: Type.ARRAY,
                  description: 'Ordered report sections matching the template headings and structure.',
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      heading: { type: Type.STRING },
                      narrative: { type: Type.STRING },
                      bullets: {
                        type: Type.ARRAY,
                        items: { type: Type.STRING },
                      },
                      table: {
                        type: Type.OBJECT,
                        description: 'Optional summary data table derived from the workbook.',
                        properties: {
                          caption: { type: Type.STRING },
                          headers: {
                            type: Type.ARRAY,
                            items: { type: Type.STRING },
                          },
                          rows: {
                            type: Type.ARRAY,
                            items: {
                              type: Type.ARRAY,
                              items: { type: Type.STRING },
                            },
                          },
                        },
                        required: ['caption', 'headers', 'rows'],
                      },
                    },
                    required: ['heading', 'narrative', 'bullets'],
                  },
                },
                placeholderValues: {
                  type: Type.ARRAY,
                  description: 'Exact key-value pairs for every {{placeholder}} found in the template.',
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      key: { type: Type.STRING },
                      value: { type: Type.STRING },
                    },
                    required: ['key', 'value'],
                  },
                },
              },
              required: [
                'reportTitle',
                'subtitle',
                'preparedDate',
                'executiveSummary',
                'keyMetrics',
                'sections',
                'placeholderValues',
              ],
            },
          },
        })
      );

      const rawJson = (response.text ?? '{}').trim();
      const parsed = JSON.parse(rawJson);

      const normalizedSections = (Array.isArray(parsed.sections) ? parsed.sections : []).map(
        (sec: any, idx: number) => ({
          id: `sec-${idx + 1}`,
          heading: String(sec.heading || `Section ${idx + 1}`),
          narrative: String(sec.narrative || ''),
          bullets: Array.isArray(sec.bullets) ? sec.bullets.map(String) : [],
          table:
            sec.table && Array.isArray(sec.table.headers) && sec.table.headers.length > 0
              ? {
                  caption: String(sec.table.caption || ''),
                  headers: sec.table.headers.map(String),
                  rows: Array.isArray(sec.table.rows)
                    ? sec.table.rows.map((r: any[]) => (Array.isArray(r) ? r.map(String) : []))
                    : [],
                }
              : undefined,
        })
      );

      return res.json({
        reportTitle: String(parsed.reportTitle || 'Generated Performance Report'),
        subtitle: String(parsed.subtitle || body.workbook.fileName),
        preparedDate: String(parsed.preparedDate || new Date().toLocaleDateString('en-US', { dateStyle: 'long' })),
        executiveSummary: String(parsed.executiveSummary || ''),
        keyMetrics: Array.isArray(parsed.keyMetrics) ? parsed.keyMetrics : [],
        sections: normalizedSections,
        placeholderValues: Array.isArray(parsed.placeholderValues) ? parsed.placeholderValues : [],
        generatedAt: new Date().toISOString(),
      });
    } catch (error: any) {
      console.error('Error generating report:', error);
      const { status, message } = formatGeminiError(error);
      return res.status(status).json({ error: message });
    }
  });

  app.post('/api/refine-section', async (req: Request, res: Response) => {
    try {
      const body = req.body as RefineSectionRequestBody;
      if (!body?.workbook || !body?.sectionHeading || !body?.refinementPrompt) {
        return res.status(400).json({ error: 'Missing required section refinement fields.' });
      }

      const ai = getGeminiClient();
      const workbookContext = buildWorkbookContext(body.workbook, body.selectedSheetNames);
      const tone = body.tone || 'Executive Brief';

      const response = await withRetry(() =>
        ai.models.generateContent({
          model: MODEL_NAME,
          contents: [
            `## SOURCE WORKBOOK DATA:\n${workbookContext}`,
            `## CURRENT SECTION HEADING:\n${body.sectionHeading}`,
            `## CURRENT NARRATIVE:\n${body.currentNarrative}`,
            `## CURRENT BULLETS:\n${body.currentBullets.map((b) => `- ${b}`).join('\n')}`,
            `## REFINEMENT REQUEST:\n${body.refinementPrompt}`,
          ].join('\n\n'),
          config: {
            systemInstruction: `Refine this single report section in a "${tone}" tone using only factual figures from the provided Excel workbook. Return updated heading, narrative, and bullets.`,
            responseMimeType: 'application/json',
            responseSchema: {
              type: Type.OBJECT,
              properties: {
                heading: { type: Type.STRING },
                narrative: { type: Type.STRING },
                bullets: {
                  type: Type.ARRAY,
                  items: { type: Type.STRING },
                },
              },
              required: ['heading', 'narrative', 'bullets'],
            },
          },
        })
      );

      const parsed = JSON.parse((response.text ?? '{}').trim());
      return res.json({
        heading: String(parsed.heading || body.sectionHeading),
        narrative: String(parsed.narrative || body.currentNarrative),
        bullets: Array.isArray(parsed.bullets) ? parsed.bullets.map(String) : body.currentBullets,
      });
    } catch (error: any) {
      console.error('Error refining section:', error);
      const { status, message } = formatGeminiError(error);
      return res.status(status).json({ error: message });
    }
  });

  if (process.env.NODE_ENV === 'production') {
    const distPath = path.resolve(__dirname, 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req: Request, res: Response) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  } else {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`LedgerDoc server listening on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
