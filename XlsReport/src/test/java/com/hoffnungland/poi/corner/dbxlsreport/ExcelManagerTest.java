package com.hoffnungland.poi.corner.dbxlsreport;

import static org.junit.jupiter.api.Assertions.*;
import java.io.File;
import java.nio.file.Path;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class ExcelManagerTest {
    @TempDir Path directory;

    @Test void exportsNestedJsonAndReopensWorkbook() throws Exception {
        var manager = new ExcelManager("report");
        try {
            assertTrue(manager.isWbEmpty());
            assertEquals("report", manager.getName());
            manager.getJsonResult("Report", "Data", "{\"items\":[\"alpha\",\"beta\"]}");
            assertFalse(manager.isWbEmpty());
            manager.finalWrite(directory + File.separator);
            try (var workbook = new XSSFWorkbook(directory.resolve("report.xlsx").toFile())) {
                var sheet = workbook.getSheet("Data");
                assertEquals("Report", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("items", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals("#0", sheet.getRow(1).getCell(1).getStringCellValue());
                assertEquals("alpha", sheet.getRow(1).getCell(2).getStringCellValue());
                assertEquals("beta", sheet.getRow(2).getCell(2).getStringCellValue());
                assertEquals(1, sheet.getNumMergedRegions());
                assertEquals(1, sheet.getPaneInformation().getHorizontalSplitPosition());
            }
        } finally {
            if (manager.swb != null) manager.swb.close();
        }
    }

    @Test void rejectsOverlongSheetNamesBeforeCreatingSheet() throws Exception {
        var manager = new ExcelManager("report");
        try {
            assertThrows(XlsWrkSheetException.class,
                    () -> manager.getJsonResult(null, "x".repeat(32), "{}"));
            assertTrue(manager.isWbEmpty());
            manager.getJsonResult(null, "x".repeat(31), "[]");
            assertFalse(manager.isWbEmpty());
        } finally {
            manager.swb.close();
        }
    }
}
