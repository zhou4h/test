/**
 * 将 Markdown 导出为 Excel 文件 (.xlsx)
 * @param markdown Markdown 格式的字符串
 * @param filename 文件名（不含扩展名）
 */
export const exportMarkdownToExcel = (markdown: string, filename: string = 'document') => {
  // Markdown 格式转 HTML
  const replaceMarkdownFormatting = (text: string): string => {
    return text
      .replace(/(\*\*|__)(.*?)\1/g, '<strong>$2</strong>')
      .replace(/(\*|_)(.*?)\1/g, '<em>$2</em>');
  };

  // 解析 Markdown 内容
  const parseMarkdown = (): string => {
    const lines = markdown.split('\n');
    let html = '<table>';
    let inTable = false;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();

      // 跳过表格分隔符行 (|---|)
      if (line.match(/^\|(\s*[-:]+\s*\|)+\s*$/)) continue;

      // 处理标题
      const headerMatch = line.match(/^(#+)\s(.*)/);
      if (headerMatch) {
        const level = headerMatch[1].length;
        const content = replaceMarkdownFormatting(headerMatch[2]);
        const fontSize = 24 - (level - 1) * 2;
        html += `<tr><td colspan="10" style="max-width: 320px; font-size:${fontSize}px;font-weight:bold;padding:10px 0">${content}</td></tr>`;
        continue;
      }

      // 处理表格
      if (line.startsWith('|')) {
        if (!inTable) {
          html += '<tr>';
          inTable = true;
        }

        const cells = line
          .split('|')
          .slice(1, -1)
          .map(cell => replaceMarkdownFormatting(cell.trim().replace(/\n/g, '<br>')));

        cells.forEach(cell => {
          html += `<td style="border:1px solid #ddd;padding:5px;white-space:pre-wrap;word-wrap:break-word;max-width:320px">${cell}</td>`;
        });

        html += '</tr>';
        continue;
      }

      // 处理表格分隔符
      if (line.match(/^\|?(\s*[-:]+\s*\|)+\s*$/)) {
        continue;
      }

      // 处理普通文本
      if (line) {
        if (inTable) {
          inTable = false;
          html += '</table><table>';
        }
        const formattedLine = replaceMarkdownFormatting(line).replace(/\n/g, '<br>');
        html += `<tr><td colspan="10" style="padding:5px 0;white-space:pre-wrap">${formattedLine}</td></tr>`;
      }
    }

    return html + '</table>';
  };

  // 创建 Excel HTML 结构
  const excelHtml = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" 
          xmlns:x="urn:schemas-microsoft-com:office:excel"
          xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="UTF-8">
      <title>${filename}</title>
      <style>
        td, th, tr, h1, h2, h3 {
        max-width: 320px;;
        }
        td { max-width: 320px; word-wrap: break-word; }
        table { table-layout: auto; width: 100%; }
      </style>
    </head>
    <body>
      ${parseMarkdown()}
    </body>
    </html>
  `;

  // 创建 Blob 并下载
  const blob = new Blob([excelHtml], { type: 'application/vnd.ms-excel;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${filename}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

/**
 * 将 Markdown 导出为 Word 文件 (.doc)
 * @param markdown Markdown 格式的字符串
 * @param filename 文件名（不含扩展名）
 */
export const exportMarkdownToWord = (markdown: string, filename: string = 'document') => {
  // Markdown 格式转 HTML
  const replaceMarkdownFormatting = (text: string): string => {
    return text
      .replace(/(\*\*|__)(.*?)\1/g, '<strong>$2</strong>')
      .replace(/(\*|_)(.*?)\1/g, '<em>$2</em>');
  };

  // 转换 Markdown 为 HTML
  const convertToHtml = (): string => {
    let html = '';
    const lines = markdown.split('\n');
    let inTable = false;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();

      // 跳过表格分隔符行 (|---|)
      if (line.match(/^\|(\s*[-:]+\s*\|)+\s*$/)) continue;

      // 处理标题
      const headerMatch = line.match(/^(#+)\s(.*)/);
      if (headerMatch) {
        const level = Math.min(headerMatch[1].length, 6);
        const content = replaceMarkdownFormatting(headerMatch[2]);
        html += `<h${level}>${content}</h${level}>`;
        continue;
      }

      // 处理表格
      if (line.startsWith('|')) {
        if (!inTable) {
          html += '<table border="1" style="border-collapse:collapse;width:100%;max-width:100%;margin:15px 0">';
          inTable = true;
        }

        const cells = line
          .split('|')
          .slice(1, -1)
          .map(cell => replaceMarkdownFormatting(cell.trim().replace(/\n/g, '<br>')));

        html += '<tr>';
        cells.forEach((cell) => {
          const isHeader = i > 0 && lines[i - 1].includes('|--');
          const tag = isHeader ? 'th' : 'td';
          html += `<${tag} style="border:1px solid #ddd;padding:8px;max-width:320px;word-wrap:break-word">${cell}</${tag}>`;
        });
        html += '</tr>';
        continue;
      }

      // 处理表格分隔符
      if (line.match(/^\|?(\s*[-:]+\s*\|)+\s*$/)) {
        continue;
      }

      // 处理普通文本
      if (line) {
        if (inTable) {
          inTable = false;
          html += '</table>';
        }
        const formattedLine = replaceMarkdownFormatting(line).replace(/\n/g, '<br>');
        html += `<p style="margin:10px 0">${formattedLine}</p>`;
      }
    }

    if (inTable) html += '</table>';
    return html;
  };

  // 创建 Word 文档
  const fullHtml = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" 
          xmlns:w="urn:schemas-microsoft-com:office:word"
          xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="utf-8">
      <title>${filename}</title>
      <style>
        td, th, tr, h1, h2, h3 {
        max-width: 320px;;
        }
        table { border-collapse: collapse; width: 100%; margin: 10px 0; }
        td { border: 1px solid #ddd; padding: 8px; }
        h1 { font-size:24px; border-bottom:2px solid #eee }
        h2 { font-size:22px }
        h3 { font-size:20px }
        td, th { max-width:320px; word-wrap:break-word }
      </style>
    </head>
    <body>
      ${convertToHtml()}
    </body>
    </html>
  `;

  // 创建 Blob 并下载
  const blob = new Blob([fullHtml], { type: 'application/msword;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${filename}.doc`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};
