package com.hoffnungland.poi.corner.extractshapesvsdx;

import static org.junit.jupiter.api.Assertions.assertEquals;
import org.junit.jupiter.api.Test;

class AppTest {
    @Test void normalizesQuotesWhitespaceAndIdentifiers() {
        assertEquals("ABC123\t'hello' \"world\"", App.normalizeShapeText("  ABC123:  ‘hello’\n“world”  "));
        assertEquals("plain text", App.normalizeShapeText(" plain\t text "));
        assertEquals("", App.normalizeShapeText("  "));
    }
}
