import mammoth from 'mammoth';
import PizZip from 'pizzip';
import type { ParsedTemplate } from '../types/report';

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

export function extractPlaceholders(text: string): string[] {
  const matches = new Set<string>();
  const regex = /\{\{\s*([a-zA-Z0-9_.\- ]+?)\s*\}\}/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(text)) !== null) {
    const key = match[1].trim();
    if (key) matches.add(key);
  }
  return Array.from(matches);
}

export function extractHeadingsFromText(text: string): string[] {
  const lines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  const headings: string[] = [];
  for (const line of lines) {
    if (/^#{1,4}\s+/.test(line)) {
      headings.push(line.replace(/^#{1,4}\s+/, '').trim());
    } else if (/^\d+\.\s+[A-Z]/.test(line) && line.length <= 75) {
      headings.push(line);
    } else if (
      line.length >= 4 &&
      line.length <= 58 &&
      !/[.!?]$/.test(line) &&
      !line.includes('{{') &&
      /^[A-Z0-9]/.test(line) &&
      line.split(' ').length <= 7
    ) {
      headings.push(line);
    }
  }

  return Array.from(new Set(headings)).slice(0, 12);
}

function extractPlaceholdersFromDocxXml(buffer: ArrayBuffer): string[] {
  try {
    const zip = new PizZip(buffer);
    const found = new Set<string>();
    const xmlFiles = Object.keys(zip.files).filter(
      (name) => name.startsWith('word/') && name.endsWith('.xml')
    );

    for (const fileName of xmlFiles) {
      const xml = zip.files[fileName]?.asText() ?? '';
      // Strip XML tags per paragraph so split Word runs ({{ + tag + }}) recombine cleanly
      const paragraphs = xml.split(/<\/w:p>/);
      for (const p of paragraphs) {
        const plain = p.replace(/<[^>]+>/g, '');
        for (const tag of extractPlaceholders(plain)) {
          found.add(tag);
        }
      }
    }
    return Array.from(found);
  } catch {
    return [];
  }
}

export async function parseDocxTemplate(file: File): Promise<ParsedTemplate> {
  const buffer = await file.arrayBuffer();
  const [htmlResult, textResult] = await Promise.all([
    mammoth.convertToHtml({ arrayBuffer: buffer }),
    mammoth.extractRawText({ arrayBuffer: buffer }),
  ]);

  const rawText = textResult.value.trim();
  const htmlPreview = htmlResult.value;

  const docDom = new DOMParser().parseFromString(htmlPreview, 'text/html');
  const htmlHeadings = Array.from(docDom.querySelectorAll('h1, h2, h3, h4'))
    .map((el) => el.textContent?.trim() ?? '')
    .filter(Boolean);

  const xmlPlaceholders = extractPlaceholdersFromDocxXml(buffer);
  const textPlaceholders = extractPlaceholders(rawText);
  const placeholders = Array.from(new Set([...xmlPlaceholders, ...textPlaceholders]));

  const headings =
    htmlHeadings.length > 0 ? Array.from(new Set(htmlHeadings)) : extractHeadingsFromText(rawText);

  return {
    name: file.name,
    sourceKind: 'docx',
    rawText,
    htmlPreview,
    headings,
    placeholders,
    docxBase64: arrayBufferToBase64(buffer),
  };
}

export function parseTextTemplate(name: string, rawText: string, sourceKind: 'text' | 'preset' = 'text'): ParsedTemplate {
  const trimmed = rawText.trim();
  const placeholders = extractPlaceholders(trimmed);
  const headings = extractHeadingsFromText(trimmed);

  const htmlPreview = trimmed
    .split(/\r?\n/)
    .map((line) => {
      const t = line.trim();
      if (!t) return '<div class="h-3"></div>';
      if (/^#\s+/.test(t)) return `<h1>${escapeHtml(t.replace(/^#\s+/, ''))}</h1>`;
      if (/^##\s+/.test(t)) return `<h2>${escapeHtml(t.replace(/^##\s+/, ''))}</h2>`;
      if (/^###\s+/.test(t)) return `<h3>${escapeHtml(t.replace(/^###\s+/, ''))}</h3>`;
      return `<p>${escapeHtml(t)}</p>`;
    })
    .join('\n');

  return {
    name,
    sourceKind,
    rawText: trimmed,
    htmlPreview,
    headings,
    placeholders,
  };
}

function escapeHtml(str: string): string {
  return str
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

export interface PresetTemplateOption {
  id: string;
  name: string;
  description: string;
  content: string;
}

export const PRESET_TEMPLATES: PresetTemplateOption[] = [
  {
    id: 'exec-quarterly',
    name: 'Executive Performance Report',
    description: 'Standard board-ready structure with executive summary, regional variance, cost discipline, and action items.',
    content: `# {{report_title}}
Prepared: {{prepared_date}}
Reporting Period: {{reporting_period}}

## 1. Executive Summary
{{executive_summary}}

## 2. Revenue & Target Attainment
Analyze performance by region and segment against targets. Highlight top over-performing markets and underperforming areas requiring intervention.
- Total Revenue: {{total_revenue}}
- Target Attainment: {{target_attainment}}
- Average Gross Margin: {{average_margin}}

## 3. Operating Cost & SLA Discipline
Summarize departmental spend versus budget, net budget variance, and SLA reliability across operational teams.

## 4. Risks & Strategic Recommendations
Provide prioritized operational and commercial recommendations grounded in the spreadsheet variance and churn metrics.`,
  },
  {
    id: 'monthly-ops',
    name: 'Monthly Operations & Financial Brief',
    description: 'Compact operational memo designed for leadership standups and monthly stakeholder distribution.',
    content: `# Monthly Operations & Financial Brief
Date: {{prepared_date}}

## Overview
{{executive_summary}}

## Key Performance Indicators
Summarize headline totals, averages, and period highlights in a structured table.

## Variance & Exception Analysis
Call out departments, regions, or categories exceeding budget or missing targets.

## Next Month Priorities
List concrete next steps tied directly to the exceptions observed in the workbook.`,
  },
  {
    id: 'client-summary',
    name: 'Client Delivery & Account Review',
    description: 'Clean client-facing review template emphasizing outcomes, SLA attainment, and growth momentum.',
    content: `# Account Performance Review
Prepared for: {{client_or_stakeholder}}
Date: {{prepared_date}}

## Executive Highlights
{{executive_summary}}

## Delivery & Volume Metrics
Present key operational volumes, margins, and service-level attainment clearly.

## Segment Breakdown
Provide a clear breakdown across regions, categories, or product lines.

## Recommended Next Steps
Outline 3–4 focused optimization opportunities for the upcoming period.`,
  },
];
