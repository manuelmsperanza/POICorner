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

public class App 
{
	private static final Logger logger = LogManager.getLogger(App.class);
	
	public static void main( String[] args )
	{
		logger.traceEntry();
		try {
			PdfWriter writer = new PdfWriter("Test.pdf", new WriterProperties().setPdfVersion(PdfVersion.PDF_2_0));
			PdfDocument pdfDocument = new PdfDocument(writer);
			pdfDocument.setTagged();
			Document document = new Document(pdfDocument);
			document.add(new Paragraph("Hello world!"));
			document.close();
		} catch (FileNotFoundException e) {
			logger.error(e.getMessage(), e);
		} catch (IOException e) {
			logger.error(e.getMessage(), e);
		}
		logger.traceExit();
    	
    }
}
