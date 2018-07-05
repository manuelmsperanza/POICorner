package me.hoffnungland.poi.corner.xmlxlsreport;

import java.util.HashMap;
import java.util.Map;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class NodeSheet {
	
	private static final Logger logger = LogManager.getLogger(NodeSheet.class);
	
	public Map<String, Integer> mapOfHeader;
	public XSSFSheet sheet;
	public int workingRow = 0;
	
	public NodeSheet(XSSFSheet sheet) {
		this.sheet = sheet;
	}

	public void loadHeader(){
		logger.traceEntry();
		this.mapOfHeader = new HashMap<String, Integer>();
		
		XSSFRow headerRow = this.sheet.getRow(this.sheet.getFirstRowNum());
		
		for(int headerIdx = headerRow.getFirstCellNum(); headerIdx <= headerRow.getLastCellNum(); headerIdx++){
			
			XSSFCell headerCell = headerRow.getCell(headerIdx);
			if(headerCell != null){
				String fieldName = headerCell.getStringCellValue();
				if(fieldName != null && !"".equals(fieldName)){
					this.mapOfHeader.put(fieldName, headerIdx);
				}
			}
			
		}
		
		logger.traceExit();
		
	}
	
}
