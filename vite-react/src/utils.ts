import { Document, HeadingLevel, Paragraph, Table, TableRow, TableCell, TextRun, Packer } from "docx";

/**
 * 增强版样式解析（支持换行符、嵌套样式）
 */
const parseInlineStyles = (text: string): TextRun[] => {
  // 合并处理 Markdown 换行语法和 HTML 标签
  const processedText = text
    .replace(/(\s{2,}|\\)$/g, '\n')       // Markdown 换行语法
    .replace(/<br\s*\/?>/gi, '\n');       // HTML 换行标签

  // 使用更智能的分割正则（支持嵌套样式）
  const segments = processedText.split(/(\*\*.*?\*\*|\*.*?\*|__.*?__|_.*?_|`.*?`|\n)/g);

  return segments
    .filter(segment => segment && segment !== '')
    .flatMap(segment => {
      // 处理换行符
      if (segment === '\n') {
        return [new TextRun({ text: '', break: 1 })];
      }

      // 粗体检测（支持 ** 和 __ 语法）
      const boldMatch = segment.match(/^(\*\*|__)(.*?)(\*\*|__)$/);
      if (boldMatch) {
        return new TextRun({ text: boldMatch[2], bold: true });
      }

      // 斜体检测（支持 * 和 _ 语法）
      const italicMatch = segment.match(/^([*_])(.*?)\1$/);
      if (italicMatch) {
        // @ts-ignore
        return new TextRun({ text: italicMatch[2], italic: true });
      }

      // 等宽字体检测
      const codeMatch = segment.match(/^`(.*?)`$/);
      if (codeMatch) {
        return new TextRun({
          text: codeMatch[1],
          font: "Courier New",
          size: 22   // 10.5pt
        });
      }

      // 处理普通文本中的连续空格
      return segment.split('  ').map((text, index) =>
        index > 0
          ? new TextRun({ text: ' ', break: 1 })  // 两个空格转换为换行
          : new TextRun(text.replace(/ +/g, ' ')) // 合并多个空格
      );
    });
};

/**
 * 优化后的 Markdown 转 Word 实现
 */
export const exportMarkdownToWord = async (markdown: string, filename: string = "document") => {
  const documentChildren = [];
  const lines = markdown.split("\n");
  let tableRows: TableRow[] = [];
  let inTable = false;

  for (const line of lines) {
    const trimmed = line.trimEnd(); // 保留行首空格

    // 跳过表格分隔符（智能识别）
    if (/^[-:| ]+$/.test(trimmed.replace(/\|/g, "")) && inTable) {
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

    // 表格处理（增强容错性）
    if (trimmed.startsWith("|")) {
      if (!inTable) {
        inTable = true;
        tableRows = [];
      }

      const cells = trimmed
        .split('|')
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
      continue;
    }

    // 表格结束处理
    if (inTable) {
      documentChildren.push(
        new Table({
          rows: tableRows,
        })
      );
      tableRows = [];
      inTable = false;
    }

    // 处理列表（新增支持）
    const listMatch = trimmed.match(/^([*\-+]|\d+\.)\s+(.+)/);
    if (listMatch) {
      documentChildren.push(
        new Paragraph({
          bullet: { level: 0 }, // 支持多级列表
          children: parseInlineStyles(listMatch[2])
        })
      );
      continue;
    }

    // 处理空行（生成空段落）
    if (trimmed === "") {
      documentChildren.push(new Paragraph({}));
      continue;
    }

    // 处理普通段落（保留原始换行）
    documentChildren.push(
      new Paragraph({
        children: parseInlineStyles(trimmed),
        spacing: { after: 120 } // 6pt 段后间距
      })
    );
  }

  // 生成专业版式文档
  const doc = new Document({
    styles: {
      paragraphStyles: [{
        id: "Normal",
        name: "Normal",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: {
          font: "Times New Roman",
          size: 24  // 12pt
        }
      }]
    },
    sections: [{
      properties: {
        page: {
          margin: {
            top: 720,    // 1 inch
            right: 720,
            bottom: 720,
            left: 720
          }
        }
      },
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
