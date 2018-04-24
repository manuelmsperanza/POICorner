package net.dtdns.hoffunungland.poi.corner.orcxlsloader;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.StringReader;
import java.io.StringWriter;
import java.sql.CallableStatement;
import java.sql.Clob;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.DateUtil;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import net.dtdns.hoffunungland.db.corner.oracleconn.OrclConnectionManager;

import net.dtdns.hoffunungland.db.corner.dbconn.StatementCached;

public class ExcelLoader {

	

	private static final Logger logger = LogManager.getLogger(ExcelLoader.class);
	private static String dateMask = "dd/MM/yyyy HH:mm:ss";
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
	
	public void loadWb(OrclConnectionManager dbManager) throws IOException, SAXException, ParserConfigurationException, SQLException, TransformerException{
		
		logger.traceEntry();
		DateFormat df = new SimpleDateFormat(dateMask);
		
		this.wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(this.sourcePath + this.excelName);
		
		Iterator<org.apache.poi.ss.usermodel.Sheet> sheetIter = this.wb.sheetIterator();
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
						
						if(contentCell.getCellTypeEnum().equals(org.apache.poi.ss.usermodel.CellType.STRING)){
							logger.trace(contentCell.getStringCellValue());
							fieldEl.setTextContent(contentCell.getStringCellValue());
						} else if(contentCell.getCellTypeEnum().equals(org.apache.poi.ss.usermodel.CellType.NUMERIC)){
							if(DateUtil.isCellDateFormatted(contentCell)){
								logger.trace(contentCell.getDateCellValue());
								fieldEl.setTextContent(df.format(contentCell.getDateCellValue()));
							} else {
								logger.trace(contentCell.getNumericCellValue());
								double d = contentCell.getNumericCellValue();
								if(d == (long) d) {
									fieldEl.setTextContent(String.format("%d",(long)d));
								} else {
							    	fieldEl.setTextContent(String.format("%s",d));
								}
								
							}
							
						}
						
						rowEl.appendChild(fieldEl);
					}
				}
				
				root.appendChild(rowEl);
			}
			
			logger.debug("Dataset ready");
			TransformerFactory tf = TransformerFactory.newInstance();
			Transformer transformer = tf.newTransformer();
			transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
			StringWriter writer = new StringWriter();
			transformer.transform(new DOMSource(doc), new StreamResult(writer));
			String output = writer.getBuffer().toString().replaceAll("\n|\r", "");
			
			if(logger.isTraceEnabled()){
				logger.trace(output);
			}
			String tableName = connectionPropsFile.getProperty("TABLE_NAME", null);
			
			StatementCached<CallableStatement> reply = dbManager.getCallableStatement("./sql/DBMS_XMLSAVE.INSERTXML.sql");
			CallableStatement replyStm = reply.getStm();
			
			replyStm.setString(1, tableName);
			logger.info("Loading to " + tableName);
			Clob payLoad = dbManager.getClob();
			payLoad.setString(1, output);
			replyStm.setClob(2, payLoad);
			
			logger.info("Saving into " + tableName);
			replyStm.execute();
			
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
			
		}
		logger.traceExit();
	}
	
}
