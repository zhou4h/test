import {Document, HeadingLevel, Paragraph, Table, TableRow, TableCell, TextRun, Packer} from "docx";

/**
 * 解析文本样式（支持粗体、斜体）
 */
const parseInlineStyles = (text: string): TextRun[] => {
  const segments = text.split(/(\*\*.*?\*\*|\*.*?\*)/g);
  return segments
    .filter(segment => segment.length > 0)
    .map(segment => {
      if (segment.startsWith("**") && segment.endsWith("**")) {
        return new TextRun({
          text: segment.slice(2, -2),
          bold: true
        });
      }
      if (segment.startsWith("*") && segment.endsWith("*")) {
        return new TextRun({
          text: segment.slice(1, -1),
          // @ts-ignore
          italic: true
        });
      }
      return new TextRun(segment);
    });
};

/**
 * 将 Markdown 导出为真正的 Word 文件 (.docx)
 * @param markdown Markdown 格式的字符串
 * @param filename 文件名（不含扩展名）
 */
export const exportMarkdownToWord = async (markdown: string, filename: string = "document") => {
  const documentChildren = [];
  const lines = markdown.split("\n");
  let tableRows: TableRow[] = [];
  let inTable = false;

  for (const line of lines) {
    const trimmed = line.trim();

    // 处理表格分隔符
    if (/^[-:| ]+$/.test(trimmed.replace(/\|/g, ""))) {
      continue;
    }

    // 处理标题
    const headingMatch = trimmed.match(/^(#+)\s+(.+)/);
    if (headingMatch) {
      const level = Math.min(headingMatch[1].length, 6);
      documentChildren.push(
        new Paragraph({
          heading: HeadingLevel[`HEADING_${level}` as keyof typeof HeadingLevel],
          children: parseInlineStyles(headingMatch[2])
        })
      );
      continue;
    }

    // 处理表格
    if (trimmed.startsWith("|")) {
      const cells = trimmed
        .split("|")
        .slice(1, -1)
        .map(c => c.trim());

      tableRows.push(
        new TableRow({
          children: cells.map(content => new TableCell({
            children: [new Paragraph({
              children: parseInlineStyles(content)
            })]
          }))
        })
      );
      inTable = true;
      continue;
    }

    // 表格结束处理
    if (inTable && !trimmed.startsWith("|")) {
      documentChildren.push(
        new Table({
          rows: tableRows,
        })
      );
      tableRows = [];
      inTable = false;
    }

    // 处理普通段落
    if (trimmed) {
      documentChildren.push(
        new Paragraph({
          children: parseInlineStyles(trimmed)
        })
      );
    }
  }

  // 生成 Word 文档
  const doc = new Document({
    sections: [{
      properties: {},
      children: documentChildren
    }]
  });

  // 创建并下载文件
  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${filename}.docx`;
  document.body.appendChild(link);
  link.click();
  setTimeout(() => {
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }, 100);
};
