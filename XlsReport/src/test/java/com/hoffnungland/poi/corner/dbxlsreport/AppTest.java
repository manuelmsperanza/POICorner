package com.hoffnungland.poi.corner.dbxlsreport;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class AppTest {

    @Test
    void greetingReturnsExpectedMessage() {
        assertEquals("Hello World!", App.greeting());
    }

    @Test
    void worksheetExceptionStoresMessage() {
        var exception = new XlsWrkSheetException("sheet missing");
        assertEquals("sheet missing", exception.getMessage());
    }
}
