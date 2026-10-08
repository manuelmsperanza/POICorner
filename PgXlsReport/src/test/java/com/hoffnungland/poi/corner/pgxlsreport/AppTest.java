package com.hoffnungland.poi.corner.pgxlsreport;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import org.junit.jupiter.api.Test;

class AppTest {

    @Test
    void rejectsIncompleteArgumentsWithoutConnecting() {
        for (int count = 0; count < 4; count++) {
            String[] args = new String[count];
            assertDoesNotThrow(() -> App.main(args));
        }
    }
}
