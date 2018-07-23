package me.hoffnungland.poi.corner.orcxlsloader;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.StringReader;
import java.sql.CallableStatement;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.DateUtil;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import me.hoffnungland.db.corner.oracleconn.OrclConnectionManager;
import me.hoffnungland.poi.corner.orcxlsreport.ExcelManager;
import me.hoffnungland.poi.corner.orcxlsreport.XlsWrkSheetException;
import me.hoffnungland.db.corner.dbconn.StatementCached;

public class ExcelLoader {

	private static final Logger logger = LogManager.getLogger(ExcelLoader.class);
	private static String dateMask = "dd/MM/yyyy HH:mm:ss";
	private static String fileDateMask = "yyyyMMddHHmmss";
	private String sourcePath;
	private String excelName;
	private String connectionName;
	
	private org.apache.poi.ss.usermodel.Workbook wb;
	private DocumentBuilder docBuilder;	
	
	public ExcelLoader(String sourcePath, String inExcelName, String connectionName) {
		this.sourcePath = sourcePath;
		this.excelName = inExcelName;
		this.connectionName = connectionName;
	}
	
	/**
	 * Utility method to manage the XML Document Builder creation
	 * @return the private field Document Builder
	 * @throws ParserConfigurationException
	 * @author ***REMOVED***
	 * @since 12-04-2018
	 */
	public DocumentBuilder getXmlDocumentBuilder() throws ParserConfigurationException{
		logger.traceEntry();
		if (this.docBuilder == null ){
			this.docBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
		}
		return logger.traceExit(this.docBuilder);
	}
	
	/**
	 * 
	 * @param dbManager
	 * @throws IOException
	 * @throws SAXException
	 * @throws ParserConfigurationException
	 * @throws SQLException
	 * @throws XlsWrkSheetException
	 * @author ***REMOVED***
	 * @since 12-04-2018
	 */
	public void loadWb(OrclConnectionManager dbManager) throws IOException, SAXException, ParserConfigurationException, SQLException, XlsWrkSheetException{
		
		logger.traceEntry();
		DateFormat df = new SimpleDateFormat(dateMask);
		DateFormat dfFile = new SimpleDateFormat(fileDateMask);
		
		logger.info("Loading " + this.excelName);
		this.wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(this.sourcePath + this.excelName);
		
		
		ExcelManager xlsMng = null;
		String xlsMngName = this.excelName.substring(0, this.excelName.indexOf(".xls")) + "_" + dfFile.format(new Date());
		Iterator<org.apache.poi.ss.usermodel.Sheet> sheetIter = this.wb.sheetIterator();
		while(sheetIter.hasNext()){
			org.apache.poi.ss.usermodel.Sheet workSheet = sheetIter.next();
			logger.info("Working " + workSheet.getSheetName());
			
			FileInputStream connectionFile = new FileInputStream("./etc/" + workSheet.getSheetName() + "/datamapping." + connectionName + ".properties");
			Properties connectionPropsFile = new Properties();
			connectionPropsFile.load(connectionFile);
			connectionFile.close();
			
			
			String backupFlag = connectionPropsFile.getProperty("TABLE.backup", "false");
			if("true".equals(backupFlag)) {
				if(xlsMng == null) {
					logger.info("Initialize the backup excel " + xlsMngName);
					xlsMng = new ExcelManager(xlsMngName);
				}
				
				StatementCached<PreparedStatement> prepStm = dbManager.executeFullTableQuery(workSheet.getSheetName(), workSheet.getSheetName());
				
				logger.info("Put query result into the excel file");
				xlsMng.getQueryResult(prepStm);						
			}
					
		}
		
		if(xlsMng != null) {
			xlsMng.createSummaryPage();
			logger.info("Closing excel file");
			xlsMng.finalWrite(this.sourcePath);
		}
		
		sheetIter = this.wb.sheetIterator();
		
		while(sheetIter.hasNext()){
			
			org.apache.poi.ss.usermodel.Sheet workSheet = sheetIter.next();
			logger.info("Working " + workSheet.getSheetName());
			
			FileInputStream connectionFile = new FileInputStream("./etc/" + workSheet.getSheetName() + "/datamapping." + connectionName + ".properties");
			Properties connectionPropsFile = new Properties();
			connectionPropsFile.load(connectionFile);
			connectionFile.close();
						
			List<String> columnList = new ArrayList<String>();
			Iterator<org.apache.poi.ss.usermodel.Row> rowIter = workSheet.rowIterator();
			if(rowIter.hasNext()){
				org.apache.poi.ss.usermodel.Row headerRow = rowIter.next();
				Iterator<org.apache.poi.ss.usermodel.Cell> cellIter = headerRow.cellIterator();
				while(cellIter.hasNext()){
					org.apache.poi.ss.usermodel.Cell headerCell = cellIter.next();
					
					String dbColumnName = connectionPropsFile.getProperty(headerCell.getStringCellValue(), null);
					logger.trace(headerCell.getStringCellValue() + " --> " + dbColumnName);
					columnList.add(dbColumnName);
				}
				
			}
			
			String xmlString = "<ROWSET/>";
			Document doc = this.getXmlDocumentBuilder().parse(new InputSource(new StringReader(xmlString)));
			Node root = doc.getDocumentElement();
			
			while(rowIter.hasNext()){
				Element rowEl = doc.createElement("ROW");
				
				org.apache.poi.ss.usermodel.Row contentRow = rowIter.next();
				Iterator<org.apache.poi.ss.usermodel.Cell> cellIter = contentRow.cellIterator();
				while(cellIter.hasNext()){
					org.apache.poi.ss.usermodel.Cell contentCell = cellIter.next();
					if(contentCell != null){
						int colIdx = contentCell.getColumnIndex();
						String columnName = columnList.get(colIdx);
						Element fieldEl = doc.createElement(columnName);
						String fieldValue = null;
						if(contentCell.getCellTypeEnum().equals(org.apache.poi.ss.usermodel.CellType.STRING)){
							logger.trace(contentCell.getStringCellValue());
							fieldValue = contentCell.getStringCellValue();
						} else if(contentCell.getCellTypeEnum().equals(org.apache.poi.ss.usermodel.CellType.NUMERIC)){
							if(DateUtil.isCellDateFormatted(contentCell)){
								logger.trace(contentCell.getDateCellValue());
								fieldValue = df.format(contentCell.getDateCellValue());
							} else {
								logger.trace(contentCell.getNumericCellValue());
								double d = contentCell.getNumericCellValue();
								if(d == (long) d) {
									fieldValue = String.format("%d",(long)d);
								} else {
									fieldValue = String.format("%s",d);
								}
								
							}
							
						}
						
						if(fieldValue != null && !"".equals(fieldValue)) {
							fieldEl.setTextContent(fieldValue);
							rowEl.appendChild(fieldEl);
						}
					}
				}
				
				root.appendChild(rowEl);
			}
			
			logger.debug("Dataset ready");

			String tableName = connectionPropsFile.getProperty("TABLE_NAME", null);
			
			String cleanFlag = connectionPropsFile.getProperty("TABLE.clean", "false");
			if("true".equals(cleanFlag)) {
				CallableStatement replyStm = dbManager.getCallableStm("DELETE " + tableName);
				
				logger.info("Cleaning " + tableName);
				replyStm.execute();
			}
			
			logger.info("Saving into " + tableName);
			dbManager.xmlSave(doc, tableName);
			
			dbManager.commit();
			logger.info("Loading to " + tableName + " is completed");
			
			String postExecCall = connectionPropsFile.getProperty("EXEC_POST_LOAD", null);
			if(postExecCall != null){
				logger.info("Execute post-loading " + postExecCall);
				CallableStatement postExecStm = dbManager.getCallableStm(postExecCall);
				postExecStm.execute();
				logger.info("Post-loading " + postExecCall + " is done");
				dbManager.commit();
			}
			
			
			String postExecSchedule = connectionPropsFile.getProperty("SCHEDULE_PROCEDURE_POST_LOAD", null); 
			if(postExecSchedule != null){
				logger.info("Schedule post-loading " + postExecSchedule);
				StatementCached<CallableStatement> postExecStm = dbManager.getCallableStatement("./sql/DBMS_SCHEDULER.CREATE_JOB.sql");
				CallableStatement clbStm  = postExecStm.getStm();
				
				clbStm.setString(1, postExecSchedule);
				
				clbStm.execute();
				
				logger.info("Post-loading " + postExecSchedule + " is scheduled");
				dbManager.commit();
			}
			
		}
		logger.traceExit();
	}
	
}