package com.hoffnungland.poi.corner.pdfcreator;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfVersion;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.pdf.WriterProperties;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;

/** Creates a tagged PDF 2.0 document containing a greeting. */
public class App 
{
	private static final Logger logger = LogManager.getLogger(App.class);
	
	/**
     * Runs the module's command-line application.
     * @param args unused; Test.pdf is written to the current directory
     */
    public static void main( String[] args )
	{
		logger.traceEntry();
		try {
			createPdf("Test.pdf");
		} catch (FileNotFoundException e) {
			logger.error(e.getMessage(), e);
		} catch (IOException e) {
			logger.error(e.getMessage(), e);
		}
		logger.traceExit();
    	
    }
    /**
     * Writes a tagged greeting document to the supplied file.
     * @param destination output PDF filename
     * @throws IOException if the document cannot be written
     */
    public static void createPdf(String destination) throws IOException {
        PdfWriter writer = new PdfWriter(destination,
                new WriterProperties().setFullCompressionMode(true).setPdfVersion(PdfVersion.PDF_2_0));
        PdfDocument pdfDocument = new PdfDocument(writer);
        pdfDocument.setTagged();
        try (Document document = new Document(pdfDocument)) {
            document.add(new Paragraph("Hello world!"));
        }
    }
}
