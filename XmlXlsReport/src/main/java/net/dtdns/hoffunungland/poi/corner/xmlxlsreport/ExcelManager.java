package net.dtdns.hoffunungland.poi.corner.xmlxlsreport;

import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Clob;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Time;
import java.sql.Timestamp;
import java.sql.Types;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;



/**
 * Manage the work-sheet data read and write.
 * @version 0.7
 * @author ***REMOVED***
 * @since 31-08-2016
 */

public class ExcelManager {

	private static final Logger logger = LogManager.getLogger(ExcelManager.class);
	private String name;
	private XmlReportWorkbook wb;
	private org.apache.poi.xssf.usermodel.XSSFCellStyle headerCellStyle;
	private org.apache.poi.xssf.usermodel.XSSFCellStyle defaultCellStyle;
	private org.apache.poi.ss.usermodel.CreationHelper createHelper;

	private DocumentBuilder docBuilder;

	/**
	 * Constructor with input name string. Define also the styles.
	 * @param name The target excel file name prefix
	 * @author ***REMOVED***
	 * @since 31-08-2016
	 */

	public ExcelManager(String name){

		this.name = name;
	}

	public void loadTemplate(String fileName) throws IOException, ParserConfigurationException{

		logger.traceEntry();
		if(this.docBuilder == null){
			this.docBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
		}

		//Load the workbook specified
		this.wb = new XmlReportWorkbook(fileName);
		this.wb.initSheets();

		if(this.createHelper == null){
			this.createHelper = this.wb.getCreationHelper();

			this.headerCellStyle = (XSSFCellStyle) this.wb.createCellStyle();
			XSSFColor foreGroundcolor = new XSSFColor(new java.awt.Color(255,204,153));
			this.headerCellStyle.setFillForegroundColor(foreGroundcolor );
			this.headerCellStyle.setFillPattern(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND);
			this.headerCellStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
			this.headerCellStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
			this.headerCellStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
			this.headerCellStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);

			this.defaultCellStyle = (XSSFCellStyle) this.wb.createCellStyle();
			this.defaultCellStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
			this.defaultCellStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
			this.defaultCellStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
			this.defaultCellStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		}



		logger.traceExit();
	}


	public void insertResultSet(ResultSet resRs, String rootNodeName) throws SQLException, IOException, SAXException{

		logger.traceEntry();

		ResultSetMetaData rsmd = resRs.getMetaData();

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
		Calendar tsCal = Calendar.getInstance();

		NodeSheet tableSheet = this.wb.mapOfNodesSheets.get("Table");
		tableSheet.workingRow = tableSheet.sheet.getLastRowNum() + 1;
		org.apache.poi.ss.usermodel.Row bodyRow = tableSheet.sheet.createRow(tableSheet.workingRow);
		
		for(int colIdx = 0; colIdx < rsmd.getColumnCount(); colIdx++){

			if(resRs.getObject(colIdx + 1) != null){

				String columName = rsmd.getColumnLabel(colIdx + 1);
				
				if(rootNodeName.equals(columName)){

					Clob content = resRs.getClob(colIdx + 1);

					BufferedReader reader = new BufferedReader(content.getCharacterStream());

					Document doc = docBuilder.parse(new InputSource(reader));
					reader.close();

					Node root = doc.getDocumentElement();

					this.nodeToCellValue(root, tableSheet);

				}else {
					
					int columnType = rsmd.getColumnType(colIdx + 1);
					String columnTypeName = rsmd.getColumnTypeName(colIdx + 1);
					int columnPosition = tableSheet.mapOfHeader.get(columName).intValue();
					
					org.apache.poi.ss.usermodel.Cell contentCell = bodyRow.createCell(columnPosition);

					if(columnType == Types.VARCHAR || columnType == Types.CHAR){
						contentCell.setCellValue(resRs.getString(colIdx + 1));

					} else if(columnType == Types.LONGVARCHAR){

						//TODO: Fix the retrieve of long values 
						//(as per https://netbeans.org/bugzilla/show_bug.cgi?id=179959 all streams are closed by readmetadata)

						//					InputStream inputStream = resRs.getAsciiStream(colIdx + 1);
						//					if (inputStream != null && inputStream.available() > 0){
						//						
						//						InputStreamReader inStreamRead = new InputStreamReader(inputStream);
						//						BufferedReader reader = new BufferedReader(inStreamRead);
						//						
						//						
						//						String         line = null;
						//						StringBuilder  stringBuilder = new StringBuilder();
						//	
						//						while( ( line = reader.readLine() ) != null ) {
						//							stringBuilder.append( line );
						//							stringBuilder.append( ExcelManager.ls );
						//						}
						//	
						//						reader.close();
						//	
						//						contentCell.setCellValue(stringBuilder.toString());
						//					}

					} else if(columnType == Types.CLOB){

						Clob content = resRs.getClob(colIdx + 1);

						if (content != null){
							BufferedReader reader = new BufferedReader(content.getCharacterStream());
							String         line = null;
							StringBuilder  stringBuilder = new StringBuilder();

							while( ( line = reader.readLine() ) != null ) {
								stringBuilder.append( line );
								stringBuilder.append( "\n" );
							}

							reader.close();
							content.free();

							if (stringBuilder.length() > 32767){
								contentCell.setCellValue(stringBuilder.substring(0, 32766));

								CreationHelper factory = wb.getCreationHelper();
								Drawing<?> drawing = tableSheet.sheet.createDrawingPatriarch();
								// When the comment box is visible, have it show in a 1x3 space
								ClientAnchor anchor = factory.createClientAnchor();
								anchor.setCol1(contentCell.getColumnIndex());
								anchor.setCol2(contentCell.getColumnIndex()+1);
								anchor.setRow1(contentCell.getRowIndex());
								anchor.setRow2(contentCell.getRowIndex()+3);

								// Create the comment and set the text+author
								Comment comment = drawing.createCellComment(anchor);
								RichTextString str = factory.createRichTextString("DB value length " + stringBuilder.length());
								comment.setString(str);
								comment.setAuthor(System.getProperty("user.name"));

								// Assign the comment to the cell
								contentCell.setCellComment(comment);

							}else {
								contentCell.setCellValue(stringBuilder.toString());
							}
						}

					} else if(columnType == Types.DATE){
						java.sql.Date dateVal = resRs.getDate(colIdx + 1);
						if(dateVal != null){
							tsCal.setTime(dateVal);
							contentCell.setCellValue(df.format(tsCal.getTime()));
						}
					} else if(columnType == Types.TIME){

						Time timeVal = resRs.getTime(colIdx + 1);
						if(timeVal != null){
							tsCal.setTime(timeVal);
							contentCell.setCellValue(df.format(tsCal.getTime()));
						}

					} else if(columnType == Types.TIMESTAMP || columnTypeName.equals("TIMESTAMP WITH LOCAL TIME ZONE")){
						Timestamp tsVal = resRs.getTimestamp(colIdx + 1);
						if(tsVal != null){
							tsCal.setTimeInMillis(tsVal.getTime());
							contentCell.setCellValue(df.format(tsCal.getTime()));
						}
					} else {
						contentCell.setCellValue(resRs.getLong(colIdx + 1));
					}

					contentCell.setCellStyle(this.defaultCellStyle);
				}
			}
		}

		logger.traceExit();

	}

	private void nodeToCellValue(Node xmlNode, NodeSheet nodeSheet){
		int columnPosition = nodeSheet.mapOfHeader.get(xmlNode.getNodeName()).intValue();
		
		org.apache.poi.ss.usermodel.Cell contentCell = null;
		logger.debug("Scanning for the first free cell in " + nodeSheet.sheet.getSheetName() + " starting from the current working " + nodeSheet.workingRow + " to the last creted row: " + nodeSheet.sheet.getLastRowNum());
		
		for(int newCellPos = nodeSheet.workingRow; newCellPos <= nodeSheet.sheet.getLastRowNum(); newCellPos++){
			
			org.apache.poi.ss.usermodel.Row nodeWorkingRow = nodeSheet.sheet.getRow(newCellPos);
			org.apache.poi.ss.usermodel.Cell workingCell = nodeWorkingRow.getCell(columnPosition);
			
			if(workingCell == null){
				logger.trace("Create cell #" + columnPosition + " at row " + newCellPos);
				contentCell = nodeWorkingRow.createCell(columnPosition);
				break;
			} else if(workingCell.getCellTypeEnum() == org.apache.poi.ss.usermodel.CellType._NONE){
				logger.trace("Cell #" + columnPosition + " at row " + newCellPos + " is of NONE type");
				contentCell = nodeWorkingRow.createCell(columnPosition);
				break;
			} else if(workingCell.getCellTypeEnum() == org.apache.poi.ss.usermodel.CellType.BLANK){
				logger.trace("Cell #" + columnPosition + " at row " + newCellPos + " is blank");
				contentCell = workingCell;
				break;
			} else if(workingCell.getStringCellValue() == null || "".equals(workingCell.getStringCellValue())){
				logger.trace("Cell #" + columnPosition + " at row " + newCellPos + " is empty");
				contentCell = workingCell;
				break;
			}
		}
		if(contentCell == null){
			int newRowId = nodeSheet.sheet.getLastRowNum() + 1;
			logger.debug("Create new row " + newRowId + " into " + nodeSheet.sheet.getSheetName());
			contentCell = nodeSheet.sheet.createRow(newRowId).createCell(columnPosition);
		}
					
		switch(xmlNode.getNodeType()){
		case org.w3c.dom.Node.ELEMENT_NODE:
			logger.trace("Child is an ELEMENT_NODE: " + xmlNode.getNodeName());
			
			if(xmlNode.hasChildNodes()){
				if(xmlNode.getChildNodes().item(0).getNodeType() == org.w3c.dom.Node.TEXT_NODE){
					logger.debug("But it has a text value");
					contentCell.setCellValue(xmlNode.getChildNodes().item(0).getNodeValue());
				} else if(!this.wb.mapOfNodesSheets.containsKey(xmlNode.getNodeName())) {
					logger.warn("For row " + nodeSheet.workingRow + " the sheet " + xmlNode.getNodeName() + " does not exist");
	
				} else {
					
					NodeSheet childNodeSheet = this.wb.mapOfNodesSheets.get(xmlNode.getNodeName());
					childNodeSheet.workingRow = childNodeSheet.sheet.getLastRowNum() + 1;
					childNodeSheet.sheet.createRow(childNodeSheet.workingRow);
					
					logger.debug("Create hyperlink to " + xmlNode.getNodeName() + " sheet " + childNodeSheet.workingRow);
					contentCell.setCellValue(xmlNode.getNodeName() + "!A" + (childNodeSheet.workingRow + 1));
	
					org.apache.poi.ss.usermodel.Hyperlink tableNodeHl = this.createHelper.createHyperlink(org.apache.poi.common.usermodel.HyperlinkType.DOCUMENT);
					tableNodeHl.setAddress("'" + xmlNode.getNodeName() + "'!A" + (childNodeSheet.workingRow + 1));
					contentCell.setHyperlink(tableNodeHl);
	
					crawlNodes(xmlNode, childNodeSheet);
				}
			}
			break;
		case org.w3c.dom.Node.TEXT_NODE:
			logger.trace("Child is an TEXT_NODE: " + xmlNode.getNodeValue());
			contentCell.setCellValue(xmlNode.getNodeValue());
			break;
		default:
			logger.debug("Type code is " + xmlNode.getNodeType());

		}

		contentCell.setCellStyle(this.defaultCellStyle);
	}
	
	private void crawlNodes(Node xmlNode, NodeSheet nodeSheet){

		logger.traceEntry();

		if(xmlNode.hasChildNodes()){

			NodeList children = xmlNode.getChildNodes();
			for(int i = 0; i < children.getLength(); i++){

				Node childNode = children.item(i);
				
				if(!nodeSheet.mapOfHeader.containsKey(childNode.getNodeName())){
					logger.warn("For row " + nodeSheet.workingRow + " the column " + childNode.getNodeName() + " does not exist in sheet " + nodeSheet.sheet.getSheetName());
					
				} else {
					this.nodeToCellValue(childNode, nodeSheet);
					
				}
			}
		}


		logger.traceExit();
	}

	/**
	 * Loop all the workbook and create a summary page with the hyperlink toward the pages with at least one record
	 * @author ***REMOVED***
	 * @since 21-09-2016
	 */
	public void createSummaryPage(){
		logger.traceEntry();

		org.apache.poi.ss.usermodel.Sheet summarySheet = this.wb.createSheet("Summary");
		this.wb.setSheetOrder("Summary", 0);

		int rowIdx = 0;

		org.apache.poi.ss.usermodel.Row summaryHeaderRow = summarySheet.createRow(rowIdx++);
		org.apache.poi.ss.usermodel.Cell summaryHeadeeNameCell = summaryHeaderRow.createCell(0);
		summaryHeadeeNameCell.setCellValue("Sheet name");
		summaryHeadeeNameCell.setCellStyle(this.headerCellStyle);

		org.apache.poi.ss.usermodel.Cell summaryHeadeeCountCell = summaryHeaderRow.createCell(1);
		summaryHeadeeCountCell.setCellValue("Count");
		summaryHeadeeCountCell.setCellStyle(this.headerCellStyle);

		for (org.apache.poi.ss.usermodel.Sheet curWorkSheet : this.wb){
			int rowCount = curWorkSheet.getPhysicalNumberOfRows()-1;
			if(rowCount > 0){

				org.apache.poi.ss.usermodel.Row summaryRow = summarySheet.createRow(rowIdx++);
				org.apache.poi.ss.usermodel.Cell tableNameCell = summaryRow.createCell(0);
				tableNameCell.setCellValue(curWorkSheet.getSheetName());


				org.apache.poi.ss.usermodel.Hyperlink tableNameHl = this.createHelper.createHyperlink(org.apache.poi.common.usermodel.HyperlinkType.DOCUMENT);
				tableNameHl.setAddress("'" + curWorkSheet.getSheetName() + "'!A1");
				tableNameCell.setHyperlink(tableNameHl);
				org.apache.poi.ss.usermodel.Cell tableCountCell = summaryRow.createCell(1);
				tableCountCell.setCellValue(rowCount);

			}
		}

		if(System.getProperty("os.name").startsWith("***REMOVED***ows")){
			summarySheet.autoSizeColumn(0);
			summarySheet.autoSizeColumn(1);
		}

		summarySheet.setZoom(85);
		this.wb.setActiveSheet(0);

		logger.traceExit();

	}

	/**
	 * Loop all the workbook and remove the page without record
	 * @author ***REMOVED***
	 * @since 22-09-2016
	 */
	public void cleanNoRecordSheets(){
		logger.traceEntry();

		for(int workSheetIdx = this.wb.getNumberOfSheets() -1; workSheetIdx >= 0; workSheetIdx--){

			org.apache.poi.ss.usermodel.Sheet workSheet = this.wb.getSheetAt(workSheetIdx);
			logger.debug(workSheet.getSheetName() + " has " + workSheet.getPhysicalNumberOfRows() + " row(s)");
			if(workSheet.getPhysicalNumberOfRows() <= 1){
				logger.debug("Removing " + workSheet.getSheetName());
				this.wb.removeSheetAt(workSheetIdx);
			}
		}

		logger.traceExit();
	}

	/**
	 * Flush the workbook data into the file and close the workbook.
	 * @author ***REMOVED***
	 * @since 31-08-2016
	 */
	public void finalWrite(String targetPath)
	{
		logger.traceEntry();
		try {


			String xlsFilename = this.name + ".xlsx";
			if (targetPath != null & !"".equals(targetPath)){
				xlsFilename = targetPath + xlsFilename;
			}
			FileOutputStream fileOut = new FileOutputStream(xlsFilename);
			
			this.wb.write(fileOut);
			fileOut.close();
			this.wb.close();

		} catch (FileNotFoundException e) {
			logger.error(e.getMessage(), e);
		} catch (IOException e) {
			logger.error(e.getMessage(), e);
		} finally {
			this.wb = null;
			logger.traceExit();
		}
	}
	/**
	 * Check if the workbook contains worksheets 
	 * @return true if there is not sheet
	 * @since 03-11-2016
	 */
	public boolean isWbEmpty(){
		return (this.wb.getNumberOfSheets() == 0);
	}
	/**
	 * @return The object name
	 * @since 03-11-2016
	 */
	public String getName() {
		return logger.traceExit(name);
	}

}
