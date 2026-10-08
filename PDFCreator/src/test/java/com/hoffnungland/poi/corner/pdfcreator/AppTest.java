package com.hoffnungland.poi.corner.pdfcreator;

import static org.junit.jupiter.api.Assertions.*;
import java.nio.file.Path;
import com.itextpdf.kernel.pdf.*;
import com.itextpdf.kernel.pdf.canvas.parser.PdfTextExtractor;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class AppTest {
    @TempDir Path directory;
    @Test void writesReadableTaggedPdf() throws Exception {
        var output = directory.resolve("greeting.pdf");
        App.createPdf(output.toString());
        try (var pdf = new PdfDocument(new PdfReader(output.toString()))) {
            assertEquals(1, pdf.getNumberOfPages());
            assertTrue(pdf.isTagged());
            assertEquals(PdfVersion.PDF_2_0, pdf.getPdfVersion());
            assertEquals("Hello world!", PdfTextExtractor.getTextFromPage(pdf.getPage(1)).trim());
        }
    }
}
