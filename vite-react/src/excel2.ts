// utils/markdownToExcel.ts
import ExcelJS from 'exceljs';

interface CellData {
  value: string;
  bold?: boolean;
  italic?: boolean;
}

const MAX_COLUMN_WIDTH = 45; // Excel column width units ≈ 7px, 320px ≈ 45 units

// Parse Markdown text (supports bold, italic, and line breaks)
const parseText = (text: string): CellData => {
  // Handle line breaks (supports <br> and \n)
  const withNewLines = text
    .replace(/<br\s*\/?>/gi, '\n')  // Replace <br> with newline
    .replace(/\\n/g, '\n');         // Replace \n with newline

  // Detect bold and italic formatting
  const bold = /(\*\*[^*]+\*\*)|(__[^_]+__)/.test(withNewLines);
  const italic = /(\*[^*]+\*)|(_[^_]+_)/.test(withNewLines);

  // Clean formatting marks
  const cleanValue = withNewLines
    .replace(/(\*\*|__)(.*?)\1/g, '$2')  // Remove bold marks
    .replace(/(\*|_)(.*?)\1/g, '$2')     // Remove italic marks
    .replace(/(\*|_)/g, '');             // Remove unmatched marks

  return { value: cleanValue, bold, italic };
};

// Parse Markdown table row
const parseMarkdownRow = (row: string): CellData[] => {
  // Remove leading/trailing pipes and split cells
  const cells = row.trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());

  return cells.map(cell => parseText(cell));
};

// Calculate column width (considering CJK character width)
const calculateColumnWidth = (content: string): number => {
  let width = 0;
  for (const char of content) {
    width += char.charCodeAt(0) > 255 ? 2 : 1; // CJK characters count as 2
  }
  return Math.min(width, MAX_COLUMN_WIDTH);
};

// Download Excel file using native browser functionality
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

  // Clean up
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
};

// Main conversion function
export const convertMarkdownReportToExcel2 = async (
  markdown: string,
  fileName: string = 'ecommerce-project-report.xlsx'
) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Project Report');

  // Split and filter Markdown lines
  const rows = markdown.split('\n').filter(row => {
    const trimmed = row.trim();
    // Skip table separators and empty lines
    return trimmed && !trimmed.startsWith('|---') &&
      !trimmed.startsWith('|--') &&
      !trimmed.startsWith('===');
  });

  // Parsed data storage
  const parsedData: CellData[][] = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    // Handle heading lines
    if (row.startsWith('# ')) {
      const title = parseText(row.replace(/^#+\s*/, ''));
      const titleRow = worksheet.addRow([title.value]);
      titleRow.font = { bold: true, size: 14 };
      titleRow.height = 25;
      worksheet.mergeCells(`A${titleRow.number}:C${titleRow.number}`);
      continue;
    }

    // Handle subheadings
    if (row.startsWith('## ') || row.startsWith('### ')) {
      const title = parseText(row.replace(/^#+\s*/, ''));
      const titleRow = worksheet.addRow([title.value]);
      titleRow.font = { bold: true };
      titleRow.height = 22;
      worksheet.mergeCells(`A${titleRow.number}:C${titleRow.number}`);
      continue;
    }

    // Handle code blocks
    if (row.startsWith('```')) {
      i++; // Skip code block start marker
      const codeLines = [];
      while (i < rows.length && !rows[i].startsWith('```')) {
        codeLines.push(rows[i]);
        i++;
      }

      const codeRow = worksheet.addRow([codeLines.join('\n')]);
      codeRow.font = { name: 'Courier New' };
      codeRow.alignment = { wrapText: true }; // Enable text wrapping
      worksheet.mergeCells(`A${codeRow.number}:C${codeRow.number}`);
      continue;
    }

    // Handle regular text
    if (!row.startsWith('|')) {
      const parsed = parseText(row);
      const textRow = worksheet.addRow([parsed.value]);
      textRow.font = {
        bold: parsed.bold,
        italic: parsed.italic
      };
      textRow.alignment = { wrapText: true }; // Enable text wrapping
      worksheet.mergeCells(`A${textRow.number}:C${textRow.number}`);
      continue;
    }

    // Handle table rows
    parsedData.push(parseMarkdownRow(row));
  }

  // Calculate table column widths
  const colWidths: number[] = [];
  parsedData.forEach(row => {
    row.forEach((cell, colIndex) => {
      const width = calculateColumnWidth(cell.value);
      if (!colWidths[colIndex] || width > colWidths[colIndex]) {
        colWidths[colIndex] = width;
      }
    });
  });

  // Set column widths (with padding)
  if (colWidths.length > 0) {
    worksheet.addRow([]); // Add spacer row before table
    worksheet.columns = colWidths.map(width => ({
      width: Math.min(width + 2, MAX_COLUMN_WIDTH)
    }));
  }

  // Add table rows with styling
  parsedData.forEach(rowData => {
    const row = worksheet.addRow(rowData.map(cell => cell.value));
    row.eachCell((cell, colNumber) => {
      const cellData = rowData[colNumber - 1];
      if (cellData) {
        // Apply font styles
        cell.font = {
          bold: cellData.bold,
          italic: cellData.italic
        };

        // Enable text wrapping if needed
        if (cellData.value.includes('\n')) {
          cell.alignment = { wrapText: true };
        }
      }
    });
  });

  // Set global styles
  worksheet.properties.defaultRowHeight = 20;

  // Generate and download file
  const buffer = await workbook.xlsx.writeBuffer();
  downloadExcel(buffer, fileName);
};
