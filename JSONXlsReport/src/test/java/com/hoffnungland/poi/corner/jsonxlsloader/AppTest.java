package com.hoffnungland.poi.corner.jsonxlsloader;

import static org.junit.jupiter.api.Assertions.*;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class AppTest {
    @TempDir Path directory;

    @Test void acceptsDirectoriesAndJsonExtensions() {
        var filter = new JsonFilter();
        assertTrue(filter.accept(directory.toFile()));
        assertTrue(filter.accept(directory.resolve("report.json").toFile()));
        assertTrue(filter.accept(directory.resolve("report.JSON").toFile()));
        assertEquals("JSON *.json", filter.getDescription());
    }

    @Test void rejectsOtherExtensionsAndMissingExtensions() {
        var filter = new JsonFilter();
        for (String name : new String[]{"report.xlsx", "report.json.bak", "report", "report.", ".json"}) {
            assertFalse(filter.accept(directory.resolve(name).toFile()), name);
        }
    }
}
