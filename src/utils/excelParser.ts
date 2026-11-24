import * as XLSX from 'xlsx';
import type { ColumnSummary, ParsedSheet, ParsedWorkbook } from '../types/report';

function computeColumnSummaries(
  headers: string[],
  rows: Record<string, string | number | boolean | null>[]
): ColumnSummary[] {
  return headers.map((header) => {
    const values = rows
      .map((r) => r[header])
      .filter((v) => v !== null && v !== undefined && String(v).trim() !== '');

    const numericValues: number[] = [];
    for (const v of values) {
      if (typeof v === 'number' && Number.isFinite(v)) {
        numericValues.push(v);
      } else if (typeof v === 'string') {
        const cleaned = v.replace(/[$,%\s,]/g, '');
        if (cleaned !== '' && !Number.isNaN(Number(cleaned))) {
          numericValues.push(Number(cleaned));
        }
      }
    }

    const isNumeric = values.length > 0 && numericValues.length / values.length >= 0.8;

    if (isNumeric && numericValues.length > 0) {
      const sum = numericValues.reduce((acc, n) => acc + n, 0);
      const min = Math.min(...numericValues);
      const max = Math.max(...numericValues);
      return {
        name: header,
        type: 'number',
        nonEmptyCount: values.length,
        sum,
        average: sum / numericValues.length,
        min,
        max,
      };
    }

    const freq = new Map<string, number>();
    for (const v of values) {
      const key = String(v).trim();
      freq.set(key, (freq.get(key) ?? 0) + 1);
    }

    const topValues = Array.from(freq.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 6)
      .map(([value, count]) => ({ value, count }));

    return {
      name: header,
      type: 'text',
      nonEmptyCount: values.length,
      uniqueCount: freq.size,
      topValues,
    };
  });
}

export async function parseExcelFile(file: File): Promise<ParsedWorkbook> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

  const sheets: ParsedSheet[] = workbook.SheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json<(string | number | boolean | Date | null)[]>(
      worksheet,
      {
        header: 1,
        defval: null,
        raw: true,
      }
    );

    const nonEmptyMatrix = matrix.filter((row) =>
      row.some((cell) => cell !== null && String(cell).trim() !== '')
    );

    if (nonEmptyMatrix.length === 0) {
      return {
        name: sheetName,
        headers: [],
        rows: [],
        rowCount: 0,
        columnSummaries: [],
      };
    }

    const rawHeaders = nonEmptyMatrix[0];
    const seenHeaders = new Map<string, number>();
    const headers = rawHeaders.map((cell, idx) => {
      const base = cell !== null && String(cell).trim() !== '' ? String(cell).trim() : `Column ${idx + 1}`;
      const count = seenHeaders.get(base) ?? 0;
      seenHeaders.set(base, count + 1);
      return count === 0 ? base : `${base} (${count + 1})`;
    });

    const dataRows = nonEmptyMatrix.slice(1).map((row) => {
      const record: Record<string, string | number | boolean | null> = {};
      headers.forEach((header, colIdx) => {
        const val = row[colIdx] ?? null;
        if (val instanceof Date) {
          record[header] = val.toISOString().slice(0, 10);
        } else {
          record[header] = val;
        }
      });
      return record;
    });

    return {
      name: sheetName,
      headers,
      rows: dataRows,
      rowCount: dataRows.length,
      columnSummaries: computeColumnSummaries(headers, dataRows),
    };
  });

  return {
    fileName: file.name,
    fileSize: file.size,
    sheetNames: workbook.SheetNames,
    sheets,
    parsedAt: new Date().toISOString(),
  };
}

export function createSampleWorkbook(): ParsedWorkbook {
  const revenueRows = [
    { Region: 'North America', Segment: 'Enterprise', 'Q3 Revenue ($)': 482000, 'Q3 Target ($)': 450000, 'Gross Margin (%)': 72.4, 'New Accounts': 14, 'Churn (%)': 1.8 },
    { Region: 'North America', Segment: 'Mid-Market', 'Q3 Revenue ($)': 215400, 'Q3 Target ($)': 225000, 'Gross Margin (%)': 68.1, 'New Accounts': 29, 'Churn (%)': 3.2 },
    { Region: 'EMEA', Segment: 'Enterprise', 'Q3 Revenue ($)': 364800, 'Q3 Target ($)': 340000, 'Gross Margin (%)': 70.6, 'New Accounts': 11, 'Churn (%)': 2.1 },
    { Region: 'EMEA', Segment: 'Mid-Market', 'Q3 Revenue ($)': 168500, 'Q3 Target ($)': 180000, 'Gross Margin (%)': 66.4, 'New Accounts': 22, 'Churn (%)': 3.9 },
    { Region: 'APAC', Segment: 'Enterprise', 'Q3 Revenue ($)': 241000, 'Q3 Target ($)': 210000, 'Gross Margin (%)': 71.2, 'New Accounts': 9, 'Churn (%)': 1.5 },
    { Region: 'APAC', Segment: 'Mid-Market', 'Q3 Revenue ($)': 119200, 'Q3 Target ($)': 115000, 'Gross Margin (%)': 65.8, 'New Accounts': 18, 'Churn (%)': 2.7 },
    { Region: 'LATAM', Segment: 'Enterprise', 'Q3 Revenue ($)': 98400, 'Q3 Target ($)': 110000, 'Gross Margin (%)': 64.9, 'New Accounts': 4, 'Churn (%)': 4.4 },
    { Region: 'LATAM', Segment: 'Mid-Market', 'Q3 Revenue ($)': 62100, 'Q3 Target ($)': 70000, 'Gross Margin (%)': 62.3, 'New Accounts': 9, 'Churn (%)': 5.1 },
  ];

  const opsRows = [
    { Department: 'Cloud Infrastructure', 'Budget ($)': 185000, 'Actual Spend ($)': 179400, 'Variance ($)': -5600, 'SLA Attainment (%)': 99.96 },
    { Department: 'Customer Support', 'Budget ($)': 142000, 'Actual Spend ($)': 148200, 'Variance ($)': 6200, 'SLA Attainment (%)': 97.4 },
    { Department: 'Sales Operations', 'Budget ($)': 210000, 'Actual Spend ($)': 204800, 'Variance ($)': -5200, 'SLA Attainment (%)': 98.8 },
    { Department: 'Product Engineering', 'Budget ($)': 390000, 'Actual Spend ($)': 382500, 'Variance ($)': -7500, 'SLA Attainment (%)': 99.2 },
  ];

  const revHeaders = Object.keys(revenueRows[0]);
  const opsHeaders = Object.keys(opsRows[0]);

  return {
    fileName: 'Q3_2026_Regional_Performance_Workbook.xlsx',
    fileSize: 48620,
    sheetNames: ['Regional Revenue', 'Operating Costs'],
    sheets: [
      {
        name: 'Regional Revenue',
        headers: revHeaders,
        rows: revenueRows,
        rowCount: revenueRows.length,
        columnSummaries: computeColumnSummaries(revHeaders, revenueRows),
      },
      {
        name: 'Operating Costs',
        headers: opsHeaders,
        rows: opsRows,
        rowCount: opsRows.length,
        columnSummaries: computeColumnSummaries(opsHeaders, opsRows),
      },
    ],
    parsedAt: new Date().toISOString(),
  };
}
