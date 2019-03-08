package me.hoffnungland.poi.corner.orcxlsreport;

import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Clob;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Time;
import java.sql.Timestamp;
import java.sql.Types;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import me.hoffnungland.db.corner.dbconn.StatementCached;



/**
 * Manage the work-sheet data read and write.
 * @version 0.8
 * @author ***REMOVED***
 * @since 31-08-2016
 */

public class ExcelManager {

	private static final Logger logger = LogManager.getLogger(ExcelManager.class);
	//private static String ls = System.getProperty("line.separator");
	protected String name;
	protected org.apache.poi.xssf.usermodel.XSSFWorkbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
	protected org.apache.poi.xssf.usermodel.XSSFCellStyle metadataHeaderCellStyle;
	protected org.apache.poi.xssf.usermodel.XSSFCellStyle headerCellStyle;
	protected org.apache.poi.xssf.usermodel.XSSFCellStyle defaultCellStyle;
	protected org.apache.poi.xssf.usermodel.XSSFCellStyle dateCellStyle;
	protected org.apache.poi.xssf.usermodel.XSSFCreationHelper createHelper = this.wb.getCreationHelper();

	/**
	 * Constructor with input name string. Define also the styles.
	 * @param name The target excel file name prefix
	 * @author ***REMOVED***
	 * @since 31-08-2016
	 */
	
	public ExcelManager(String name){
		this.name = name;
		
		this.headerCellStyle = this.wb.createCellStyle();
		//XSSFColor foreGroundcolor = new XSSFColor(new java.awt.Color(255,204,153));
		byte[] rgb = new byte[3];
		rgb[0] = (byte) 255; // red
		rgb[1] = (byte) 204; // green
		rgb[2] = (byte) 153; // blue
		XSSFColor foreGroundcolor = new XSSFColor(rgb, new org.apache.poi.xssf.usermodel.DefaultIndexedColorMap()); // #f2dcdb
		this.headerCellStyle.setFillForegroundColor(foreGroundcolor );
		this.headerCellStyle.setFillPattern(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND);
		this.headerCellStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.headerCellStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.headerCellStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.headerCellStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		
		this.metadataHeaderCellStyle = this.wb.createCellStyle();
		this.metadataHeaderCellStyle.setFillForegroundColor(foreGroundcolor);
		this.metadataHeaderCellStyle.setFillPattern(org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND);
		this.metadataHeaderCellStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.metadataHeaderCellStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.metadataHeaderCellStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.metadataHeaderCellStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.metadataHeaderCellStyle.setAlignment(HorizontalAlignment.CENTER);
		XSSFFont defaultFont= this.wb.createFont();
		defaultFont.setBold(true);
		this.metadataHeaderCellStyle.setFont(defaultFont);

		this.defaultCellStyle = this.wb.createCellStyle();
		this.defaultCellStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.defaultCellStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.defaultCellStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.defaultCellStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		
		this.dateCellStyle = this.wb.createCellStyle();
		this.dateCellStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.dateCellStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.dateCellStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.dateCellStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.dateCellStyle.setDataFormat(this.createHelper.createDataFormat().getFormat("d/m/yyyy h:mm"));
			
	}

	/**
	 * Get the information within the ResultSet of the input query and fill a new work-sheet.
	 * @param query the executed query having a valid ResultSet.
	 * @throws SQLException
	 * @throws IOException
	 * @throws XlsWrkSheetException raised in case of syntactic errors of the work-sheet data 
	 * @author ***REMOVED***
	 * @since 31-08-2016
	 */
	public void getQueryResult(StatementCached<PreparedStatement> prepStm) throws SQLException, IOException, XlsWrkSheetException {
		logger.traceEntry();
		String sheetName = prepStm.getName();
		if(sheetName.length() > 30){
			throw new XlsWrkSheetException("Work-sheet " + sheetName + " [" + sheetName.length() + "] must not have more than 30 characters");
		}

		org.apache.poi.xssf.usermodel.XSSFSheet workSheet = this.wb.createSheet(sheetName);
		this.createSheetHeader(workSheet, prepStm, 0, 0);
		this.createSheetContent(workSheet, prepStm, 1, 0, true);		
		workSheet.createFreezePane(0, 1);
		workSheet.setZoom(85);
		
		logger.traceExit();
	}
	
	/**
	 * Get the information within the ResultSet of the input query containing metadata and fill a new work-sheet.
	 * @param query the executed query having a valid ResultSet.
	 * @throws SQLException
	 * @throws IOException
	 * @throws XlsWrkSheetException raised in case of syntactic errors of the work-sheet data 
	 * @author ***REMOVED***
	 * @since 22-10-2018
	 */
	public void getMetadataResult(StatementCached<PreparedStatement> prepStm) throws SQLException, IOException, XlsWrkSheetException {
		logger.traceEntry();
		String sheetName = prepStm.getName();
		if(sheetName.length() > 30){
			throw new XlsWrkSheetException("Work-sheet " + sheetName + " [" + sheetName.length() + "] must not have more than 30 characters");
		}

		org.apache.poi.xssf.usermodel.XSSFSheet workSheet = this.wb.createSheet(sheetName);
		this.createMetadataHeader(workSheet, prepStm, 0, 0);
		this.createSheetHeader(workSheet, prepStm, 1, 0);
		this.createSheetContent(workSheet, prepStm, 2, 0, true);		
		workSheet.createFreezePane(0, 2);
		workSheet.setZoom(85);
		
		logger.traceExit();
	}

	/**
	 * Add the top row of the work-sheet. Get the information from the ResultSetMetaData of query's ResultSet.
	 * @param workSheet the working work-sheet
	 * @param query the executed query having a valid ResultSet.
	 * @param inRowId starting write row id (0 based)
	 * @param inColId starting write column id (0 based)
	 * @throws SQLException
	 * @author ***REMOVED***
	 * @since 31-08-2016 
	 */
	protected void createSheetHeader(org.apache.poi.xssf.usermodel.XSSFSheet workSheet, StatementCached<PreparedStatement> prepStm, int inRowId, int inColId) throws SQLException{
		logger.traceEntry();
		org.apache.poi.xssf.usermodel.XSSFRow headerRow = workSheet.createRow(inRowId);
		ResultSet resRs = prepStm.getStm().getResultSet();

		ResultSetMetaData rsmd = resRs.getMetaData();
		for(int headerIdx = 0; headerIdx < rsmd.getColumnCount(); headerIdx++){
			org.apache.poi.xssf.usermodel.XSSFCell columnNameCell = headerRow.createCell(headerIdx + inColId);
			String columnaName = rsmd.getColumnName(headerIdx + 1);
			columnNameCell.setCellValue(columnaName);
			logger.debug(columnaName + " of type " + rsmd.getColumnTypeName(headerIdx + 1) + " ("+ rsmd.getColumnType(headerIdx + 1) + ")" );
			columnNameCell.setCellStyle(this.headerCellStyle);
			if(System.getProperty("os.name").startsWith("***REMOVED***ows")){
				workSheet.autoSizeColumn(headerIdx);
			}
		}

		logger.traceExit();
	}
	
	/**
	 * Add the top row of the work-sheet. Get the information from the ResultSetMetaData of query's ResultSet.
	 * @param workSheet the working work-sheet
	 * @param query the executed query having a valid ResultSet.
	 * @param inRowId starting write row id (0 based)
	 * @param inColId starting write column id (0 based)
	 * @throws SQLException
	 * @author ***REMOVED***
	 * @since 22-10-2018 
	 */
	protected void createMetadataHeader(org.apache.poi.xssf.usermodel.XSSFSheet workSheet, StatementCached<PreparedStatement> prepStm, int inRowId, int inColId) throws SQLException{
		logger.traceEntry();
		org.apache.poi.xssf.usermodel.XSSFRow headerRow = workSheet.createRow(inRowId);
		ResultSet resRs = prepStm.getStm().getResultSet();

		ResultSetMetaData rsmd = resRs.getMetaData();
		org.apache.poi.xssf.usermodel.XSSFCell columnNameCell = headerRow.createCell(inColId);
		columnNameCell.setCellValue(workSheet.getSheetName());
		columnNameCell.setCellStyle(this.metadataHeaderCellStyle);
		if(rsmd.getColumnCount() > 1) {
			workSheet.addMergedRegion(new CellRangeAddress(inRowId, inRowId, inColId, inColId + rsmd.getColumnCount()-1));
		}
		logger.traceExit();
	}
	
	/**
	 * Add a row for each record within the query's ResultSet.
	 * It manage the following Oracle data type: VARCHAR, CHAR, CLOB, DATE, TIME, TIMESTAMP and NUMBER
	 * @param workSheet the working work-sheet
	 * @param query the executed query having a valid ResultSet.
	 * @param inRowId starting write row id (0 based)
	 * @param inColId starting write column id (0 based)
	 * @param applyDefaultStyle true to apply default style
	 * @throws SQLException
	 * @throws IOException
	 * @author ***REMOVED***
	 * @since 31-08-2016
	 */
	protected void createSheetContent(org.apache.poi.xssf.usermodel.XSSFSheet workSheet, StatementCached<PreparedStatement> prepStm, int inRowId, int inColId, boolean applyDefaultStyle) throws SQLException, IOException{
		logger.traceEntry();
		int rowId = inRowId;
		ResultSet resRs = prepStm.getStm().getResultSet();
		ResultSetMetaData rsmd = resRs.getMetaData();
		
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
		Calendar tsCal = Calendar.getInstance();
		logger.trace("Column count {}", rsmd.getColumnCount());
		while (resRs.next()) {
			
			org.apache.poi.xssf.usermodel.XSSFRow bodyRow = workSheet.getRow(rowId);
			if(bodyRow == null) {
				bodyRow = workSheet.createRow(rowId);
			}
			
			logger.trace("Working row #{}", rowId);
			for(int colIdx = inColId; colIdx < rsmd.getColumnCount(); colIdx++){
				org.apache.poi.xssf.usermodel.XSSFCell contentCell = bodyRow.getCell(colIdx);
				
				if(contentCell == null) {
					contentCell = bodyRow.createCell(colIdx);
				}
				
				if(applyDefaultStyle) {
					contentCell.setCellStyle(this.defaultCellStyle);
				}
				
				if(resRs.getObject(colIdx + 1) != null  && !resRs.wasNull()){
					
					int columnType = rsmd.getColumnType(colIdx + 1);
					String columnTypeName = rsmd.getColumnTypeName(colIdx + 1);
					
					logger.trace("Working col #{} of type {}", (colIdx + 1), columnTypeName);
					
					if(columnType == Types.VARCHAR || columnType == Types.CHAR){
						logger.trace("columnType VARCHAR or CHAR");
						String value = resRs.getString(colIdx + 1);
						logger.trace("Col value: {}", value);
						contentCell.setCellValue(resRs.getString(colIdx + 1));
						
					} else if(columnType == Types.LONGVARCHAR){
						logger.trace("columnType LONGVARCHAR");
						//TODO: Fix the retrieve of long values 
						//(as per https://netbeans.org/bugzilla/show_bug.cgi?id=179959 all streams are closed by readmetadata)
						
//						InputStream inputStream = resRs.getAsciiStream(colIdx + 1);
//						if (inputStream != null && inputStream.available() > 0){
//							
//							InputStreamReader inStreamRead = new InputStreamReader(inputStream);
//							BufferedReader reader = new BufferedReader(inStreamRead);
//							
//							
//							String         line = null;
//							StringBuilder  stringBuilder = new StringBuilder();
//		
//							while( ( line = reader.readLine() ) != null ) {
//								stringBuilder.append( line );
//								stringBuilder.append( ExcelManager.ls );
//							}
//		
//							reader.close();
//		
//							contentCell.setCellValue(stringBuilder.toString());
//						}
						
					} else if(columnType == Types.CLOB){
						logger.trace("columnType CLOB");
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
								Drawing<?> drawing = workSheet.createDrawingPatriarch();
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
						logger.trace("columnType DATE");
						java.sql.Date dateVal = resRs.getDate(colIdx + 1);
						
						if(dateVal != null){
							tsCal.setTime(dateVal);
							logger.trace("Col value: {}", df.format(tsCal.getTime()));
							contentCell.setCellValue(tsCal.getTime());
							if(applyDefaultStyle) {
								contentCell.setCellStyle(this.dateCellStyle);
							}
						}
					} else if(columnType == Types.TIME){
						logger.trace("columnType TIME");
						Time timeVal = resRs.getTime(colIdx + 1);
						if(timeVal != null){
							tsCal.setTime(timeVal);
							logger.trace("Col value: {}", df.format(tsCal.getTime()));
							contentCell.setCellValue(tsCal.getTime());
							if(applyDefaultStyle) {
								contentCell.setCellStyle(this.dateCellStyle);
							}
						}
						
					} else if(columnType == Types.TIMESTAMP || columnTypeName.equals("TIMESTAMP WITH LOCAL TIME ZONE")){
						logger.trace("columnType TIMESTAMP");
						Timestamp tsVal = resRs.getTimestamp(colIdx + 1);
						if(tsVal != null){
							tsCal.setTimeInMillis(tsVal.getTime());
							logger.trace("Col value: {}", df.format(tsCal.getTime()));
							contentCell.setCellValue(tsCal.getTime());
							if(applyDefaultStyle) {
								contentCell.setCellStyle(this.dateCellStyle);
							}
						}
					} else if(columnType == Types.BLOB || columnType == Types.NULL || columnType == Types.OTHER){
						logger.trace("columnType BLOB, NULL or OTHER");
					} else if(columnType == Types.INTEGER){
						logger.trace("columnType INTEGER");
						contentCell.setCellValue(resRs.getInt(colIdx + 1));
					} else if(columnType == Types.NUMERIC){
						logger.trace("columnType NUMERIC");
						long value = resRs.getLong(colIdx + 1);
						logger.trace("Col value: {}", value);
						contentCell.setCellValue(value);
					} else if(columnType == Types.DECIMAL){
						logger.trace("columnType DECIMAL");
						contentCell.setCellValue(resRs.getDouble(colIdx + 1));
					} else if(columnType == Types.DOUBLE){
						logger.trace("columnType DOUBLE");
						contentCell.setCellValue(resRs.getDouble(colIdx + 1));
					} else if(columnType == Types.FLOAT){
						logger.trace("columnType FLOAT");
						contentCell.setCellValue(resRs.getFloat(colIdx + 1));
					} else {
						logger.trace("columnType else {}", columnType);
						contentCell.setCellValue(resRs.getLong(colIdx + 1));
					}
				}
				
			}

			rowId++;
		}

		logger.traceExit();
	}
	
	/**
	 * Loop all the workbook and create a summary page with the hyperlink toward the pages with at least one record
	 * @author ***REMOVED***
	 * @since 21-09-2016
	 */
	public void createSummaryPage(int excludeRows){
		logger.traceEntry();
		
		org.apache.poi.xssf.usermodel.XSSFSheet selectedSheet = this.wb.getSheetAt(this.wb.getActiveSheetIndex());
		org.apache.poi.xssf.usermodel.XSSFSheet summarySheet = this.wb.createSheet("Summary");
		
		this.wb.setSheetOrder("Summary", 0);
		
		summarySheet.setSelected(true);
		selectedSheet.setSelected(false);

		this.wb.setActiveSheet(0);
		
		int rowIdx = 0;
		
		org.apache.poi.xssf.usermodel.XSSFRow summaryHeaderRow = summarySheet.createRow(rowIdx++);
		org.apache.poi.xssf.usermodel.XSSFCell summaryHeadeeNameCell = summaryHeaderRow.createCell(0);
		summaryHeadeeNameCell.setCellValue("Sheet name");
		summaryHeadeeNameCell.setCellStyle(this.headerCellStyle);
		
		org.apache.poi.xssf.usermodel.XSSFCell summaryHeadeeCountCell = summaryHeaderRow.createCell(1);
		summaryHeadeeCountCell.setCellValue("Count");
		summaryHeadeeCountCell.setCellStyle(this.headerCellStyle);
		
		for (org.apache.poi.ss.usermodel.Sheet curWorkSheet : this.wb){
			int rowCount = curWorkSheet.getPhysicalNumberOfRows() - excludeRows;
				
			org.apache.poi.xssf.usermodel.XSSFRow summaryRow = summarySheet.createRow(rowIdx++);
			org.apache.poi.xssf.usermodel.XSSFCell tableNameCell = summaryRow.createCell(0);
			tableNameCell.setCellValue(curWorkSheet.getSheetName());
			
			org.apache.poi.xssf.usermodel.XSSFHyperlink tableNameHl = this.createHelper.createHyperlink(org.apache.poi.common.usermodel.HyperlinkType.DOCUMENT);
			tableNameHl.setAddress("'" + curWorkSheet.getSheetName() + "'!A1");
			tableNameCell.setHyperlink(tableNameHl);
			org.apache.poi.xssf.usermodel.XSSFCell tableCountCell = summaryRow.createCell(1);
			tableCountCell.setCellValue(rowCount);
			
		}
		
		if(System.getProperty("os.name").startsWith("***REMOVED***ows")){
			summarySheet.autoSizeColumn(0);
			summarySheet.autoSizeColumn(1);
		}
		
		summarySheet.setZoom(85);
		
		
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
			
			org.apache.poi.xssf.usermodel.XSSFSheet workSheet = this.wb.getSheetAt(workSheetIdx);
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

			POIXMLProperties props = this.wb.getProperties();
			POIXMLProperties.CoreProperties coreProp = props.getCoreProperties();
	        coreProp.setCreator(System.getProperty("user.name"));
	        
			String xlsFilename = this.name + ".xlsx";
			if (targetPath != null && !"".equals(targetPath)){
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
