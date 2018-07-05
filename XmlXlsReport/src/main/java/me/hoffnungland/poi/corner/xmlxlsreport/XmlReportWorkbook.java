package me.hoffnungland.poi.corner.xmlxlsreport;

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
	
	//Create list of NodeSheet index by sheet name
	public Map<String, NodeSheet> mapOfNodesSheets;
	
	public XmlReportWorkbook() {
		super();
	}

	public XmlReportWorkbook(File file) throws IOException, InvalidFormatException {
		super(file);
	}
	
	public XmlReportWorkbook(InputStream inS) throws IOException {
		super(inS);
	}
	
	public XmlReportWorkbook(String path) throws IOException {
		super(path);
		
	}

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
