import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import type { GeneratedReport, ParsedTemplate } from '../types/report';

function base64ToUint8Array(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function triggerDownload(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function safeSlug(title: string): string {
  return (
    title
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '')
      .slice(0, 60) || 'generated_report'
  );
}

export function canFillUploadedDocxTemplate(
  template: ParsedTemplate | null,
  report: GeneratedReport | null
): boolean {
  return Boolean(
    template?.sourceKind === 'docx' &&
      template.docxBase64 &&
      template.placeholders.length > 0 &&
      report
  );
}

export function exportFilledDocxTemplate(
  template: ParsedTemplate,
  report: GeneratedReport
): void {
  if (!template.docxBase64) {
    throw new Error('No source .docx binary available.');
  }

  const zip = new PizZip(base64ToUint8Array(template.docxBase64));
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '{{', end: '}}' },
  });

  const dataMap: Record<string, string> = {
    report_title: report.reportTitle,
    subtitle: report.subtitle,
    prepared_date: report.preparedDate,
    executive_summary: report.executiveSummary,
  };

  for (const item of report.placeholderValues) {
    dataMap[item.key] = item.value;
  }

  doc.render(dataMap);

  const blob = doc.getZip().generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });

  triggerDownload(blob, `${safeSlug(report.reportTitle)}_template_filled.docx`);
}

export async function exportStructuredDocx(report: GeneratedReport): Promise<void> {
  const children: (Paragraph | Table)[] = [];

  children.push(
    new Paragraph({
      text: report.reportTitle,
      heading: HeadingLevel.TITLE,
      spacing: { after: 100 },
    })
  );

  children.push(
    new Paragraph({
      children: [
        new TextRun({
          text: `${report.subtitle}  ·  ${report.preparedDate}`,
          color: '475569',
          size: 21,
        }),
      ],
      spacing: { after: 280 },
    })
  );

  children.push(
    new Paragraph({
      text: 'Executive Summary',
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 120, after: 120 },
    })
  );

  children.push(
    new Paragraph({
      children: [new TextRun({ text: report.executiveSummary, size: 22 })],
      spacing: { after: 240 },
    })
  );

  if (report.keyMetrics.length > 0) {
    children.push(
      new Paragraph({
        text: 'Key Metrics Snapshot',
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 120, after: 120 },
      })
    );

    const metricTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          tableHeader: true,
          children: ['Metric', 'Value', 'Context'].map(
            (header) =>
              new TableCell({
                shading: { type: ShadingType.CLEAR, fill: 'F1F5F9' },
                borders: {
                  top: { style: BorderStyle.SINGLE, size: 1, color: 'CBD5E1' },
                  bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CBD5E1' },
                  left: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
                  right: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
                },
                children: [
                  new Paragraph({
                    children: [new TextRun({ text: header, bold: true, size: 20, color: '0F172A' })],
                  }),
                ],
              })
          ),
        }),
        ...report.keyMetrics.map(
          (m) =>
            new TableRow({
              children: [m.label, m.value, m.context].map(
                (val, idx) =>
                  new TableCell({
                    borders: {
                      top: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
                      bottom: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
                      left: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
                      right: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
                    },
                    children: [
                      new Paragraph({
                        alignment: idx === 1 ? AlignmentType.RIGHT : AlignmentType.LEFT,
                        children: [
                          new TextRun({
                            text: val,
                            bold: idx === 1,
                            size: 20,
                            color: '1E293B',
                          }),
                        ],
                      }),
                    ],
                  })
              ),
            })
        ),
      ],
    });

    children.push(metricTable);
    children.push(new Paragraph({ text: '', spacing: { after: 200 } }));
  }

  for (const section of report.sections) {
    children.push(
      new Paragraph({
        text: section.heading,
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 240, after: 120 },
      })
    );

    if (section.narrative) {
      const paragraphs = section.narrative.split(/\n{2,}/);
      for (const p of paragraphs) {
        children.push(
          new Paragraph({
            children: [new TextRun({ text: p.trim(), size: 22 })],
            spacing: { after: 140 },
          })
        );
      }
    }

    for (const bullet of section.bullets) {
      children.push(
        new Paragraph({
          text: bullet,
          bullet: { level: 0 },
          spacing: { after: 80 },
        })
      );
    }

    if (section.table && section.table.headers.length > 0 && section.table.rows.length > 0) {
      if (section.table.caption) {
        children.push(
          new Paragraph({
            children: [
              new TextRun({ text: section.table.caption, italics: true, size: 19, color: '475569' }),
            ],
            spacing: { before: 120, after: 80 },
          })
        );
      }

      const dataTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            tableHeader: true,
            children: section.table.headers.map(
              (h) =>
                new TableCell({
                  shading: { type: ShadingType.CLEAR, fill: 'F1F5F9' },
                  borders: {
                    top: { style: BorderStyle.SINGLE, size: 1, color: 'CBD5E1' },
                    bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CBD5E1' },
                    left: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
                    right: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
                  },
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: h, bold: true, size: 19, color: '0F172A' })],
                    }),
                  ],
                })
            ),
          }),
          ...section.table.rows.map(
            (row) =>
              new TableRow({
                children: section.table!.headers.map((_, colIdx) => {
                  const cellVal = row[colIdx] ?? '';
                  const looksNumeric = /^[$€£\-+]?[\d,]+(\.\d+)?%?$/.test(cellVal.trim());
                  return new TableCell({
                    borders: {
                      top: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
                      bottom: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
                      left: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
                      right: { style: BorderStyle.SINGLE, size: 1, color: 'E2E8F0' },
                    },
                    children: [
                      new Paragraph({
                        alignment: looksNumeric ? AlignmentType.RIGHT : AlignmentType.LEFT,
                        children: [new TextRun({ text: cellVal, size: 19, color: '1E293B' })],
                      }),
                    ],
                  });
                }),
              })
          ),
        ],
      });

      children.push(dataTable);
      children.push(new Paragraph({ text: '', spacing: { after: 180 } }));
    }
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
          },
        },
        children,
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  triggerDownload(blob, `${safeSlug(report.reportTitle)}.docx`);
}

export function buildMarkdownReport(report: GeneratedReport): string {
  const lines: string[] = [
    `# ${report.reportTitle}`,
    `${report.subtitle} · ${report.preparedDate}`,
    '',
    `## Executive Summary`,
    report.executiveSummary,
    '',
  ];

  if (report.keyMetrics.length > 0) {
    lines.push('## Key Metrics');
    lines.push('| Metric | Value | Context |');
    lines.push('| --- | --- | --- |');
    for (const m of report.keyMetrics) {
      lines.push(`| ${m.label} | ${m.value} | ${m.context} |`);
    }
    lines.push('');
  }

  for (const sec of report.sections) {
    lines.push(`## ${sec.heading}`);
    if (sec.narrative) {
      lines.push(sec.narrative);
      lines.push('');
    }
    for (const b of sec.bullets) {
      lines.push(`- ${b}`);
    }
    if (sec.bullets.length > 0) lines.push('');
    if (sec.table && sec.table.headers.length > 0) {
      if (sec.table.caption) lines.push(`*${sec.table.caption}*`);
      lines.push(`| ${sec.table.headers.join(' | ')} |`);
      lines.push(`| ${sec.table.headers.map(() => '---').join(' | ')} |`);
      for (const r of sec.table.rows) {
        lines.push(`| ${sec.table.headers.map((_, i) => r[i] ?? '').join(' | ')} |`);
      }
      lines.push('');
    }
  }

  return lines.join('\n');
}

export function exportMarkdownReport(report: GeneratedReport): void {
  const md = buildMarkdownReport(report);
  const blob = new Blob([md], { type: 'text/markdown;charset=utf-8' });
  triggerDownload(blob, `${safeSlug(report.reportTitle)}.md`);
}

function escapeHtml(str: string): string {
  return str
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

export async function copyRichReportForGoogleDocs(report: GeneratedReport): Promise<void> {
  const htmlParts: string[] = [
    `<div>`,
    `<h1>${escapeHtml(report.reportTitle)}</h1>`,
    `<p><em>${escapeHtml(report.subtitle)} · ${escapeHtml(report.preparedDate)}</em></p>`,
    `<h2>Executive Summary</h2>`,
    `<p>${escapeHtml(report.executiveSummary)}</p>`,
  ];

  if (report.keyMetrics.length > 0) {
    htmlParts.push(`<h3>Key Metrics</h3>`);
    htmlParts.push(
      `<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;width:100%;"><thead><tr><th>Metric</th><th>Value</th><th>Context</th></tr></thead><tbody>`
    );
    for (const m of report.keyMetrics) {
      htmlParts.push(
        `<tr><td><strong>${escapeHtml(m.label)}</strong></td><td>${escapeHtml(m.value)}</td><td>${escapeHtml(m.context)}</td></tr>`
      );
    }
    htmlParts.push(`</tbody></table>`);
  }

  for (const sec of report.sections) {
    htmlParts.push(`<h2>${escapeHtml(sec.heading)}</h2>`);
    if (sec.narrative) {
      for (const p of sec.narrative.split(/\n{2,}/)) {
        htmlParts.push(`<p>${escapeHtml(p.trim())}</p>`);
      }
    }
    if (sec.bullets.length > 0) {
      htmlParts.push(`<ul>${sec.bullets.map((b) => `<li>${escapeHtml(b)}</li>`).join('')}</ul>`);
    }
    if (sec.table && sec.table.headers.length > 0) {
      if (sec.table.caption) {
        htmlParts.push(`<p><em>${escapeHtml(sec.table.caption)}</em></p>`);
      }
      htmlParts.push(
        `<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;width:100%;"><thead><tr>${sec.table.headers
          .map((h) => `<th>${escapeHtml(h)}</th>`)
          .join('')}</tr></thead><tbody>`
      );
      for (const r of sec.table.rows) {
        htmlParts.push(
          `<tr>${sec.table.headers.map((_, i) => `<td>${escapeHtml(r[i] ?? '')}</td>`).join('')}</tr>`
        );
      }
      htmlParts.push(`</tbody></table>`);
    }
  }

  htmlParts.push(`</div>`);

  const html = htmlParts.join('');
  const plain = buildMarkdownReport(report);

  if (typeof ClipboardItem !== 'undefined' && navigator.clipboard?.write) {
    const item = new ClipboardItem({
      'text/html': new Blob([html], { type: 'text/html' }),
      'text/plain': new Blob([plain], { type: 'text/plain' }),
    });
    await navigator.clipboard.write([item]);
  } else {
    await navigator.clipboard.writeText(plain);
  }
}
