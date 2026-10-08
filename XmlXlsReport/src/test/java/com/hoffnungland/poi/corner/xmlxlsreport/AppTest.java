package com.hoffnungland.poi.corner.xmlxlsreport;

import static org.junit.jupiter.api.Assertions.*;
import java.util.Map;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class AppTest {
    @Test void loadsSparseHeadersAndRebuildsMapping() throws Exception {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Template");
            var row = sheet.createRow(0);
            row.createCell(0).setCellValue("name");
            row.createCell(2).setCellValue("");
            row.createCell(4).setCellValue("amount");
            var node = new NodeSheet(sheet);
            node.loadHeader();
            assertEquals(Map.of("name", 0, "amount", 4), node.mapOfHeader);
            row.getCell(0).setCellValue("renamed");
            node.loadHeader();
            assertEquals(Map.of("renamed", 0, "amount", 4), node.mapOfHeader);
        }
    }
}
