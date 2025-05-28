package com.test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.commonmark.node.Node;
import org.commonmark.parser.Parser;
import org.commonmark.renderer.html.HtmlRenderer;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.ByteArrayOutputStream;
import java.math.BigInteger;
import java.util.List;
import java.util.Map;

// FileConversionController.java
@RestController
@RequestMapping("/api")
public class FileConversionController {

    @PostMapping("/convert")
    public ResponseEntity<byte[]> convertMarkdown(
            @RequestBody Map<String, String> payload,
            @RequestParam String format) {

        String markdown = payload.get("markdown");

        try {
            byte[] fileContent;
            String contentType;
            String filename;

            if ("docx".equalsIgnoreCase(format)) {
                fileContent = convertToWord(markdown);
                contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                filename = "converted.docx";
            } else if ("xlsx".equalsIgnoreCase(format)) {
                fileContent = convertToExcel(markdown);
                contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                filename = "converted.xlsx";
            } else {
                throw new IllegalArgumentException("不支持的格式: " + format);
            }

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + filename + "\"")
                    .contentType(MediaType.parseMediaType(contentType))
                    .body(fileContent);

        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }

    private byte[] convertToWord(String markdown) throws Exception {
        try (XWPFDocument document = new XWPFDocument()) {
            // 添加标题样式
            XWPFParagraph titlePara = document.createParagraph();
            titlePara.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = titlePara.createRun();
            titleRun.setText("Markdown转换文档");
            titleRun.setFontSize(20);
            titleRun.setBold(true);

            // 转换Markdown为HTML
            String html = convertMarkdownToHtml(markdown);

            // 添加HTML内容
            XWPFParagraph contentPara = document.createParagraph();
            XWPFRun contentRun = contentPara.createRun();
            contentRun.setText(html);
            contentRun.setFontSize(12);

            // 设置精美样式
            setDocumentStyles(document);

            // 写入字节数组
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            document.write(out);
            return out.toByteArray();
        }
    }

    private byte[] convertToExcel(String markdown) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Markdown内容");

            // 创建标题样式
            CellStyle headerStyle = createHeaderStyle(workbook);

            // 转换Markdown表格
            List<List<String>> tableData = parseMarkdownTables(markdown);

            // 填充Excel
            int rowNum = 0;
            for (List<String> rowData : tableData) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                for (String cellData : rowData) {
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(cellData);
                    if (rowNum == 1) cell.setCellStyle(headerStyle);
                }
            }

            // 自动调整列宽
            for (int i = 0; i < tableData.get(0).size(); i++) {
                sheet.autoSizeColumn(i);
            }

            // 添加精美样式
            setExcelStyles(workbook, sheet);

            // 写入字节数组
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private String convertMarkdownToHtml(String markdown) {
        Parser parser = Parser.builder().build();
        Node document = parser.parse(markdown);
        HtmlRenderer renderer = HtmlRenderer.builder().build();
        return renderer.render(document);
    }

    private List<List<String>> parseMarkdownTables(String markdown) {
        // 实现Markdown表格解析逻辑
        // 返回二维列表: [[header1, header2], [row1col1, row1col2], ...]
    }

    private void setDocumentStyles(XWPFDocument document) {
        // 设置文档全局样式
        document.getDocument().getBody().setSectPr(createSectionProperties());

        // 设置段落样式
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        CTPageMar pageMar = sectPr.addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(720));  // 1英寸边距
        pageMar.setRight(BigInteger.valueOf(720));
        pageMar.setTop(BigInteger.valueOf(720));
        pageMar.setBottom(BigInteger.valueOf(720));
    }

    private CellStyle createHeaderStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        Font font = workbook.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFont(font);

        return style;
    }

    private void setExcelStyles(XSSFWorkbook workbook, Sheet sheet) {
        // 设置交替行颜色
        CellStyle evenRowStyle = workbook.createCellStyle();
        evenRowStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        evenRowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            if (i % 2 == 0) {
                Row row = sheet.getRow(i);
                for (Cell cell : row) {
                    cell.setCellStyle(evenRowStyle);
                }
            }
        }
    }
}
