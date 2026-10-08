package com.hoffnungland.sdc;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import com.hoffnungland.poi.corner.orcxlsloader.App;

import org.junit.jupiter.api.Test;

class AppTest {

    @Test
    void rejectsIncompleteArgumentsWithoutConnecting() {
        for (int count = 0; count < 3; count++) {
            String[] args = new String[count];
            assertDoesNotThrow(() -> App.main(args));
        }
    }
}
