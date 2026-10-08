package com.hoffnungland.poi.corner.h2xlsreport;

import static org.junit.jupiter.api.Assertions.*;
import java.io.File;
import java.nio.file.Path;
import java.sql.DriverManager;
import com.hoffnungland.poi.corner.dbxlsreport.ExcelManager;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class QueryExportTest {
    @TempDir Path directory;

    @Test void exportsJdbcStringsNumbersAndNulls() throws Exception {
        try (var connection = DriverManager.getConnection("jdbc:h2:mem:export");
             var statement = connection.prepareStatement(
                     "SELECT 'Alice' AS NAME, CAST(42 AS INTEGER) AS AMOUNT, CAST(NULL AS VARCHAR) AS OPTIONAL")) {
            assertTrue(statement.execute());
            var manager = new ExcelManager("query");
            manager.getQueryResult("Results", statement);
            manager.finalWrite(directory + File.separator);
            try (var workbook = new XSSFWorkbook(directory.resolve("query.xlsx").toFile())) {
                var sheet = workbook.getSheet("Results");
                assertEquals("NAME", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("Alice", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals(42.0, sheet.getRow(1).getCell(1).getNumericCellValue());
                assertEquals(CellType.BLANK, sheet.getRow(1).getCell(2).getCellType());
                assertEquals(2, sheet.getPhysicalNumberOfRows());
            }
        }
    }
}
