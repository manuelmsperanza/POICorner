package com.hoffnungland.poi.corner.pdfcreator;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.pdf.ReaderProperties;
import com.itextpdf.kernel.pdf.WriterProperties;

import java.io.File;
import java.io.IOException;

public class RemovePdfPassword {
	private static final Logger logger = LogManager.getLogger(RemovePdfPassword.class);
	
	 public static void main( String[] args )
	    {
	    	logger.traceEntry();
	    	
	    
	    	        String src = args[0];  // Path to the encrypted PDF
	    	        String dest = args[1]; // Path for the unlocked PDF
	    	        String password = args[2];    // Password of the secured PDF

	    	        try {
	    	            // Creating the reader and passing the password
	    	            PdfReader reader = new PdfReader(src, new ReaderProperties().setPassword(password.getBytes()));

	    	            // Creating a new writer for the unlocked PDF
	    	            PdfWriter writer = new PdfWriter(dest);

	    	            // Creating the PdfDocument to manipulate the file
	    	            PdfDocument pdfDoc = new PdfDocument(reader, writer);

	    	            // Closing the document to write changes
	    	            pdfDoc.close();

	    	            System.out.println("Password removed successfully.");
	    	        } catch (IOException e) {
	    	            e.printStackTrace();
	    	        }
	    

	    	
	    	logger.traceExit();
	    	
	    }

}
