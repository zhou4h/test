import ExcelJS from 'exceljs';

interface TextSegment {
  text: string;
  bold?: boolean;
  italic?: boolean;
}

const MAX_COLUMN_WIDTH = 45;

const parseRichText = (text: string): TextSegment[] => {
  const segments: TextSegment[] = [];
  let remaining = text
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/\\n/g, '\n');

  const formatRegex = /(\*\*(.*?)\*\*|__(.*?)__|\*(.*?)\*|_(.*?)_|\\\*|\\_)/g;

  let lastIndex = 0;
  let match;

  while ((match = formatRegex.exec(remaining)) !== null) {
    if (match.index > lastIndex) {
      segments.push({
        text: remaining.slice(lastIndex, match.index)
      });
    }

    const [full, bold1, boldContent1, boldContent2, italic1, italicContent1, italicContent2, escape] = match;

    if (escape) {
      segments.push({ text: full[1] });
    } else if (bold1 || boldContent2) {
      segments.push({
        text: boldContent1 || boldContent2,
        bold: true
      });
    } else if (italic1 || italicContent2) {
      segments.push({
        text: italicContent1 || italicContent2,
        italic: true
      });
    }

    lastIndex = match.index + full.length;
  }

  if (lastIndex < remaining.length) {
    segments.push({
      text: remaining.slice(lastIndex)
    });
  }

  return segments;
};

const calculateRichTextWidth = (segments: TextSegment[]): number => {
  let width = 0;
  for (const segment of segments) {
    width += segment.text.split('').reduce((acc, char) =>
      acc + (char.charCodeAt(0) > 255 ? 2 : 1), 0
    );
  }
  return Math.min(width, MAX_COLUMN_WIDTH);
};

const downloadExcel = (buffer: ArrayBuffer, fileName: string) => {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();

  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
};

export const convertMarkdownReportToExcel = async (
  markdown: string,
  fileName: string = 'ecommerce-project-report.xlsx'
) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Project Report');

  const rows = markdown.split('\n').filter(row => {
    const trimmed = row.trim();
    return trimmed && !trimmed.startsWith('|---') && !trimmed.startsWith('===');
  });

  const parsedTableData: TextSegment[][][] = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    // Handle headings
    if (row.startsWith('# ')) {
      const segments = parseRichText(row.replace(/^#+\s*/, ''));
      const titleRow = worksheet.addRow([{
        richText: segments.map(s => ({ text: s.text, font: { bold: s.bold } }))
      }]);
      worksheet.mergeCells(`A${titleRow.number}:C${titleRow.number}`);
      continue;
    }

    // Handle code blocks
    if (row.startsWith('```')) {
      i++;
      const codeLines = [];
      while (i < rows.length && !rows[i].startsWith('```')) {
        codeLines.push(rows[i]);
        i++;
      }
      const codeRow = worksheet.addRow([codeLines.join('\n')]);
      codeRow.alignment = { wrapText: true };
      worksheet.mergeCells(`A${codeRow.number}:C${codeRow.number}`);
      continue;
    }

    // Handle normal text
    if (!row.startsWith('|')) {
      const segments = parseRichText(row);
      const textRow = worksheet.addRow([{
        richText: segments.map(s => ({ text: s.text, font: { bold: s.bold, italic: s.italic } }))
      }]);
      textRow.alignment = { wrapText: true };
      worksheet.mergeCells(`A${textRow.number}:C${textRow.number}`);
      continue;
    }

    // Handle table rows
    const cells = row.trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());
    parsedTableData.push(
      cells.map(cell => parseRichText(cell))
    );
  }

  // Calculate column widths
  const colWidths: number[] = [];
  parsedTableData.forEach(tableRow => {
    tableRow.forEach((cellSegments, colIndex) => {
      const width = calculateRichTextWidth(cellSegments);
      colWidths[colIndex] = Math.max(colWidths[colIndex] || 0, width);
    });
  });

  // Set column widths
  if (colWidths.length > 0) {
    worksheet.columns = colWidths.map(width => ({
      width: Math.min(width + 2, MAX_COLUMN_WIDTH)
    }));
  }

  // Add table data
  parsedTableData.forEach(tableRow => {
    const row = worksheet.addRow(
      tableRow.map(cellSegments => ({
        richText: cellSegments.map(segment => ({
          text: segment.text,
          font: { bold: segment.bold, italic: segment.italic }
        }))
      }))
    );

    row.eachCell(cell => {
      cell.alignment = { wrapText: true };
    });
  });


  const buffer = await workbook.xlsx.writeBuffer();
  downloadExcel(buffer, fileName);
};
