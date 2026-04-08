package com.hoffnungland.poi.corner.xmlxlsreport;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XmlReportWorkbook extends XSSFWorkbook {
	private static final Logger logger = LogManager.getLogger(XmlReportWorkbook.class);
	
	/**
	 * Index of {@link NodeSheet} instances keyed by worksheet name.
	 */
	public Map<String, NodeSheet> mapOfNodesSheets;
	
	/**
	 * Creates an empty workbook instance.
	 */
	public XmlReportWorkbook() {
		super();
	}

	/**
	 * Creates a workbook from the supplied template file.
	 *
	 * @param file template file to load.
	 * @throws IOException if the file cannot be read.
	 * @throws InvalidFormatException if the file is not a valid workbook format.
	 */
	public XmlReportWorkbook(File file) throws IOException, InvalidFormatException {
		super(file);
	}
	
	/**
	 * Creates a workbook from the provided input stream.
	 *
	 * @param inS stream that contains workbook data.
	 * @throws IOException if the workbook cannot be read.
	 */
	public XmlReportWorkbook(InputStream inS) throws IOException {
		super(inS);
	}
	
	/**
	 * Creates a workbook from the path of an existing template file.
	 *
	 * @param path absolute or relative path of the workbook template.
	 * @throws IOException if the workbook cannot be read.
	 */
	public XmlReportWorkbook(String path) throws IOException {
		super(path);
		
	}

	/**
	 * Initializes {@link #mapOfNodesSheets} by scanning every worksheet and loading
	 * the first-row header map for each sheet.
	 */
	public void initSheets(){
		logger.traceEntry();
		this.mapOfNodesSheets = new HashMap<String, NodeSheet>();
		
		//For every worksheet
		Iterator<Sheet> iterSheet = this.sheetIterator();
		while(iterSheet.hasNext()){
			
			//create a new NodeSheet and assign the current worksheet
			NodeSheet iterNodeSheet = new NodeSheet((XSSFSheet) iterSheet.next());
			iterNodeSheet.loadHeader();
			
			this.mapOfNodesSheets.put(iterNodeSheet.sheet.getSheetName(), iterNodeSheet);
		}
		
		logger.traceExit();
	}
}
