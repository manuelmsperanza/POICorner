package com.hoffnungland.poi.corner.dbxlsreport;

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
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONArray;
import org.json.JSONObject;

import com.hoffnungland.db.corner.dbconn.StatementCached;



/**
 * Manage the work-sheet data read and write.
 * @version 0.8
 * @author manuel.m.speranza
 * @since 31-08-2016
 */

public class ExcelManager {

	private static final Logger logger = LogManager.getLogger(ExcelManager.class);
	//private static String ls = System.getProperty("line.separator");
	protected String name;
	protected org.apache.poi.xssf.usermodel.XSSFWorkbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
	protected org.apache.poi.xssf.streaming.SXSSFWorkbook swb = null;
	protected org.apache.poi.xssf.usermodel.XSSFCellStyle metadataHeaderCellStyle;
	protected org.apache.poi.xssf.usermodel.XSSFCellStyle headerCellStyle;
	protected org.apache.poi.xssf.usermodel.XSSFCellStyle defaultCellStyle;
	protected org.apache.poi.xssf.usermodel.XSSFCellStyle dateCellStyle;
	protected org.apache.poi.ss.usermodel.CreationHelper createHelper = this.wb.getCreationHelper();

	/**
	 * Constructor with input name string. Define also the styles.
	 * @param name The target excel file name prefix
	 * @author manuel.m.speranza
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
		org.apache.poi.xssf.usermodel.XSSFColor foreGroundcolor = new org.apache.poi.xssf.usermodel.XSSFColor(rgb, new org.apache.poi.xssf.usermodel.DefaultIndexedColorMap()); // #f2dcdb
		this.headerCellStyle.setFillForegroundColor(foreGroundcolor);
		
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
		this.metadataHeaderCellStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
		org.apache.poi.ss.usermodel.Font defaultFont= this.wb.createFont();
		defaultFont.setBold(true);
		this.metadataHeaderCellStyle.setFont(defaultFont);

		this.defaultCellStyle = this.wb.createCellStyle();
		this.defaultCellStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.defaultCellStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.defaultCellStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.defaultCellStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.defaultCellStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.TOP);
		
		this.dateCellStyle = this.wb.createCellStyle();
		this.dateCellStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.dateCellStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.dateCellStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.dateCellStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
		this.dateCellStyle.setDataFormat(this.createHelper.createDataFormat().getFormat("d/m/yyyy h:mm"));
		
		this.swb = new org.apache.poi.xssf.streaming.SXSSFWorkbook(this.wb, 10000, true, true);

	}

	/**
	 * Get the information within the ResultSet of the input query and fill a new work-sheet.
	 * @param prepStm the executed query having a valid ResultSet.
	 * @throws SQLException
	 * @throws IOException
	 * @throws XlsWrkSheetException raised in case of syntactic errors of the work-sheet data 
	 * @author manuel.m.speranza
	 * @since 31-08-2016
	 */
	public void getQueryResult(StatementCached<PreparedStatement> prepStm) throws SQLException, IOException, XlsWrkSheetException {
		logger.traceEntry();
		this.getQueryResult(prepStm.getName(), prepStm.getStm());
			
		logger.traceExit();
	}
	public void getQueryResult(String sheetName, PreparedStatement prepStm) throws SQLException, IOException, XlsWrkSheetException {
		logger.traceEntry();
		
		if(sheetName.length() > 30){
			throw new XlsWrkSheetException("Work-sheet " + sheetName + " [" + sheetName.length() + "] must not have more than 30 characters");
		}

		int sheetCounter = 0;
		int rowCount = 0;
		do {
			String sheetNewName = sheetName;
			if(sheetCounter > 0) {
				String sheetSuffix = "(" + sheetCounter + ")";
				if(sheetName.length() + sheetSuffix.length() > 30) {
					sheetNewName = sheetName.substring(0, 30+(30-(sheetName.length() + sheetSuffix.length()))) + sheetSuffix;
				} else {
					sheetNewName = sheetName + sheetSuffix;
				}
			}
			logger.debug("Creating new Worksheet {}", sheetNewName);
			org.apache.poi.xssf.streaming.SXSSFSheet workSheet = this.swb.createSheet(sheetNewName);
			ResultSet resRs = prepStm.getResultSet();
			ResultSetMetaData rsmd = resRs.getMetaData();
			int columnsWidth[] = new int[rsmd.getColumnCount()];
			this.createSheetHeader(workSheet, columnsWidth, prepStm, 0, 0);
			logger.debug("columnsWidth {}", Arrays.toString(columnsWidth));
			rowCount = this.createSheetContent(workSheet, columnsWidth, prepStm, 1, 0, true);
			logger.debug("columnsWidth {}", Arrays.toString(columnsWidth));
			for(int colIdx = 0; colIdx < columnsWidth.length; colIdx++){
				//width = Truncate([{Number of Visible Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
				//workSheet.setColumnWidth(colIdx, (columnsWidth[colIdx] * 256));
				int width = (int) Math.ceil((columnsWidth[colIdx] + 1) * 7.2 / 7 * 256) + 5;
				workSheet.setColumnWidth(colIdx, width > 256*256 ? 256*256 : width);
			}
			
			workSheet.createFreezePane(0, 1);
			workSheet.setZoom(85);
			sheetCounter++;
			

		}while(!prepStm.getResultSet().isAfterLast() && rowCount > 0);
		
		logger.traceExit();
	}
	/**
	 * Get the information within the ResultSet of the input query containing metadata and fill a new work-sheet.
	 * @param prepStm the executed query having a valid ResultSet.
	 * @throws SQLException
	 * @throws IOException
	 * @throws XlsWrkSheetException raised in case of syntactic errors of the work-sheet data 
	 * @author manuel.m.speranza
	 * @since 22-10-2018
	 */
	public void getMetadataResult(StatementCached<PreparedStatement> prepStm) throws SQLException, IOException, XlsWrkSheetException {
		logger.traceEntry();
		this.getMetadataResult(prepStm.getName(), prepStm.getStm());
		
		
		logger.traceExit();
	}
	
	public void getMetadataResult(String sheetName, PreparedStatement prepStm) throws SQLException, IOException, XlsWrkSheetException {
		logger.traceEntry();
		if(sheetName.length() > 30){
			throw new XlsWrkSheetException("Work-sheet " + sheetName + " [" + sheetName.length() + "] must not have more than 30 characters");
		}
		int sheetCounter = 0;
		int rowCount = 0;
		do {
			String sheetNewName = sheetName;
			if(sheetCounter > 0) {
				String sheetSuffix = "(" + sheetNewName + ")";
				if(sheetName.length() + sheetSuffix.length() > 30) {
					sheetNewName = sheetName.substring(0, 30+(30-(sheetName.length() + sheetSuffix.length()))) + sheetSuffix;
				} else {
					sheetNewName = sheetName + sheetSuffix;
				}
			}
			logger.debug("Creating new Worksheet {}", sheetNewName);
			org.apache.poi.xssf.streaming.SXSSFSheet workSheet = this.swb.createSheet(sheetNewName);
			this.createMetadataHeader(workSheet, prepStm, 0, 0);
			
			ResultSet resRs = prepStm.getResultSet();
			ResultSetMetaData rsmd = resRs.getMetaData();
			int columnsWidth[] = new int[rsmd.getColumnCount()];
			
			this.createSheetHeader(workSheet, columnsWidth, prepStm, 1, 0);
			logger.debug("columnsWidth {}", Arrays.toString(columnsWidth));
			rowCount = this.createSheetContent(workSheet, columnsWidth, prepStm, 2, 0, true);
			logger.debug("columnsWidth {}", Arrays.toString(columnsWidth));
			for(int colIdx = 0; colIdx < columnsWidth.length; colIdx++){
				//width = Truncate([{Number of Visible Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
				//workSheet.setColumnWidth(colIdx, (columnsWidth[colIdx] * 256));
				int width = (int) Math.ceil((columnsWidth[colIdx] + 1) * 7.2 / 7 * 256) + 5;
				workSheet.setColumnWidth(colIdx, width > 256*256 ? 256*256 : width);
			}
			
			workSheet.createFreezePane(0, 2);
			workSheet.setZoom(85);
			sheetCounter++;
		}while(!prepStm.getResultSet().isAfterLast() && rowCount > 0);
		
		logger.traceExit();
	}

	/**
	 * Add the top row of the work-sheet. Get the information from the ResultSetMetaData of query's ResultSet.
	 * @param workSheet the working work-sheet
	 * @param prepStm the executed query having a valid ResultSet.
	 * @param inRowIdx starting write row id (0 based)
	 * @param inColIdx starting write column id (0 based)
	 * @throws SQLException
	 * @author manuel.m.speranza
	 * @since 31-08-2016 
	 */
	protected void createSheetHeader(org.apache.poi.xssf.streaming.SXSSFSheet workSheet, int[] columnsWidth, PreparedStatement prepStm, int inRowIdx, int inColIdx) throws SQLException{
		logger.traceEntry();
		org.apache.poi.xssf.streaming.SXSSFRow headerRow = workSheet.createRow(inRowIdx);
		ResultSet resRs = prepStm.getResultSet();

		ResultSetMetaData rsmd = resRs.getMetaData();
		for(int headerIdx = 0; headerIdx < rsmd.getColumnCount(); headerIdx++){
			org.apache.poi.xssf.streaming.SXSSFCell columnNameCell = headerRow.createCell(headerIdx + inColIdx);
			String columnName = rsmd.getColumnName(headerIdx + 1);
			columnNameCell.setCellValue(columnName);
			columnsWidth[headerIdx] = columnName.length();
			logger.debug(columnName + " of type " + rsmd.getColumnTypeName(headerIdx + 1) + " ("+ rsmd.getColumnType(headerIdx + 1) + ")" );
			columnNameCell.setCellStyle(this.headerCellStyle);
			/*if(System.getProperty("os.name").startsWith("Windows")){
				workSheet.autoSizeColumn(headerIdx);
			}*/
		}

		logger.traceExit();
	}
	
	/**
	 * Add the top row of the work-sheet. Get the information from the ResultSetMetaData of query's ResultSet.
	 * @param workSheet the working work-sheet
	 * @param prepStm the executed query having a valid ResultSet.
	 * @param inRowIdx starting write row id (0 based)
	 * @param inColIdx starting write column id (0 based)
	 * @throws SQLException
	 * @author manuel.m.speranza
	 * @since 22-10-2018 
	 */
	protected void createMetadataHeader(org.apache.poi.xssf.streaming.SXSSFSheet workSheet, PreparedStatement prepStm, int inRowIdx, int inColIdx) throws SQLException{
		logger.traceEntry();
		org.apache.poi.xssf.streaming.SXSSFRow headerRow = workSheet.createRow(inRowIdx);
		ResultSet resRs = prepStm.getResultSet();

		ResultSetMetaData rsmd = resRs.getMetaData();
		org.apache.poi.xssf.streaming.SXSSFCell columnNameCell = headerRow.createCell(inColIdx);
		columnNameCell.setCellValue(workSheet.getSheetName());
		columnNameCell.setCellStyle(this.metadataHeaderCellStyle);
		if(rsmd.getColumnCount() > 1) {
			workSheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(inRowIdx, inRowIdx, inColIdx, inColIdx + rsmd.getColumnCount()-1));
		}
		logger.traceExit();
	}
	
	/**
	 * Add a row for each record within the query's ResultSet.
	 * It manage the following Oracle data type: VARCHAR, CHAR, CLOB, DATE, TIME, TIMESTAMP and NUMBER
	 * @param workSheet the working work-sheet
	 * @param prepStm the executed query having a valid ResultSet.
	 * @param inRowIdx starting write row id (0 based)
	 * @param inColIdx starting write column id (0 based)
	 * @param applyDefaultStyle true to apply default style
	 * @throws SQLException
	 * @throws IOException
	 * @author manuel.m.speranza
	 * @since 31-08-2016
	 */
	protected int createSheetContent(org.apache.poi.xssf.streaming.SXSSFSheet workSheet, int[] columnsWidth, PreparedStatement prepStm, int inRowIdx, int inColIdx, boolean applyDefaultStyle) throws SQLException, IOException{
		logger.traceEntry();
		int rowIdx = inRowIdx;
		ResultSet resRs = prepStm.getResultSet();
		ResultSetMetaData rsmd = resRs.getMetaData();	
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
		DecimalFormat format = (DecimalFormat) DecimalFormat.getInstance();
		DecimalFormatSymbols symbols = format.getDecimalFormatSymbols();
		char sep = symbols.getDecimalSeparator();
		DateFormat dfTs = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss" + sep + "S");
		DateFormat dfTsTz = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss" + sep + "S XXX");
		Calendar tsCal = Calendar.getInstance();
		logger.trace("Column count {}", rsmd.getColumnCount());
		int availableRows = org.apache.poi.ss.SpreadsheetVersion.EXCEL2007.getMaxRows() - inRowIdx;
		int rowCount = 0;
		while (resRs.next() && availableRows > 0) {
			org.apache.poi.xssf.streaming.SXSSFRow bodyRow = workSheet.getRow(rowIdx);
			if(bodyRow == null) {
				bodyRow = workSheet.createRow(rowIdx);
			}
			rowCount++;
			availableRows--;
			logger.trace("Working row #{}", rowIdx);
			for(int colIdx = inColIdx; colIdx < rsmd.getColumnCount(); colIdx++){
				org.apache.poi.xssf.streaming.SXSSFCell contentCell = bodyRow.getCell(colIdx);
				
				int lineLength = 0;
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
						lineLength = value.length();
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
								if(line != null && line.length() > lineLength && line.length() < 256) {
									lineLength = line.length();	
								}
								
								stringBuilder.append( line );
								stringBuilder.append( "\n" );
							}
		
							reader.close();
							content.free();
		
							if (stringBuilder.length() > 32767){
								contentCell.setCellValue(stringBuilder.substring(0, 32766));
								
								org.apache.poi.ss.usermodel.CreationHelper factory = this.swb.getCreationHelper();
								org.apache.poi.ss.usermodel.Drawing<?> drawing = workSheet.createDrawingPatriarch();
								// When the comment box is visible, have it show in a 1x3 space
								org.apache.poi.ss.usermodel.ClientAnchor anchor = factory.createClientAnchor();
							    anchor.setCol1(contentCell.getColumnIndex());
							    anchor.setCol2(contentCell.getColumnIndex()+1);
							    anchor.setRow1(contentCell.getRowIndex());
							    anchor.setRow2(contentCell.getRowIndex()+3);

							    // Create the comment and set the text+author
							    org.apache.poi.ss.usermodel.Comment comment = drawing.createCellComment(anchor);
							    org.apache.poi.ss.usermodel.RichTextString str = factory.createRichTextString("DB value length " + stringBuilder.length());
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
							String value = df.format(tsCal.getTime());
							logger.trace("Col value: {}", value);
							lineLength = value.length();
							
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
							String value = df.format(tsCal.getTime());
							logger.trace("Col value: {}", value);
							lineLength = value.length();
							
							contentCell.setCellValue(tsCal.getTime());
							if(applyDefaultStyle) {
								contentCell.setCellStyle(this.dateCellStyle);
							}
						}
						
					} else if(columnType == Types.TIMESTAMP) {
						logger.trace("columnType TIMESTAMP");
						Timestamp tsVal = resRs.getTimestamp(colIdx + 1);
						if(tsVal != null){
							tsCal.setTimeInMillis(tsVal.getTime());
							String value = dfTs.format(tsCal.getTime());
							logger.trace("Col value: {}", value);
							lineLength = value.length();
							
							contentCell.setCellValue(tsCal.getTime());
							if(applyDefaultStyle) {
								contentCell.setCellStyle(this.dateCellStyle);
							}
						}
					} else if(columnTypeName.equals("TIMESTAMP WITH LOCAL TIME ZONE")){
						
						logger.trace("columnType TIMESTAMP");
						Timestamp tsVal = resRs.getTimestamp(colIdx + 1);
						if(tsVal != null){
							tsCal.setTimeInMillis(tsVal.getTime());
							
							String value = dfTsTz.format(tsCal.getTime());
							logger.trace("Col value: {}", value);
							lineLength = value.length();
							
							contentCell.setCellValue(tsCal.getTime());
							if(applyDefaultStyle) {
								contentCell.setCellStyle(this.dateCellStyle);
							}
						}
					} else if(columnType == Types.BLOB || columnType == Types.NULL || columnType == Types.OTHER || columnType == Types.VARBINARY || columnType == -101 || columnType == -104){
						logger.trace("columnType BLOB, NULL or OTHER");
					} else if(columnType == Types.INTEGER){
						logger.trace("columnType INTEGER");
						int intValue = resRs.getInt(colIdx + 1);
						lineLength = String.valueOf(intValue).length();
						contentCell.setCellValue(intValue);
					} else if(columnType == Types.NUMERIC){
						logger.trace("columnType NUMERIC");
						long value = resRs.getLong(colIdx + 1);
						logger.trace("Col value: {}", value);
						lineLength = String.valueOf(value).length();
						contentCell.setCellValue(value);
					} else if(columnType == Types.DECIMAL){
						logger.trace("columnType DECIMAL");
						double doubleValue = resRs.getDouble(colIdx + 1);
						lineLength = String.valueOf(doubleValue).length();
						contentCell.setCellValue(doubleValue);
					} else if(columnType == Types.DOUBLE){
						logger.trace("columnType DOUBLE");
						double doubleValue = resRs.getDouble(colIdx + 1);
						lineLength = String.valueOf(doubleValue).length();
						contentCell.setCellValue(doubleValue);
					} else if(columnType == Types.FLOAT){
						logger.trace("columnType FLOAT");
						float floatValue = resRs.getFloat(colIdx + 1);
						lineLength = String.valueOf(floatValue).length();
						contentCell.setCellValue(floatValue);
					} else if(columnType == Types.ROWID){	
						logger.trace("columnType ROWID");
						String value = resRs.getRowId(colIdx + 1).toString();
						lineLength = value.length();
						contentCell.setCellValue(value);
					} else {
						logger.trace("columnType else {}", columnType);
						String value = resRs.getObject(colIdx + 1).toString();
						lineLength = value.length();
						contentCell.setCellValue(value);
					}
				}
				
				if(lineLength > columnsWidth[colIdx - inColIdx] && lineLength < 256) {
					columnsWidth[colIdx - inColIdx] = lineLength;
				}
			}

			rowIdx++;
		}

		return logger.traceExit(rowCount);
	}
	
	/**
	 * Loop all the workbook and create a summary page with the hyperlink toward the pages with at least one record
	 * @author manuel.m.speranza
	 * @since 21-09-2016
	 */
	public void createSummaryPage(int excludeRows){
		logger.traceEntry();
		
		org.apache.poi.xssf.streaming.SXSSFSheet selectedSheet = this.swb.getSheetAt(this.swb.getActiveSheetIndex());
		org.apache.poi.xssf.streaming.SXSSFSheet summarySheet = this.swb.createSheet("Summary");
		
		this.swb.setSheetOrder("Summary", 0);
		
		summarySheet.setSelected(true);
		selectedSheet.setSelected(false);

		this.swb.setActiveSheet(0);
		
		int rowIdx = 0;
		
		org.apache.poi.xssf.streaming.SXSSFRow summaryHeaderRow = summarySheet.createRow(rowIdx++);
		org.apache.poi.xssf.streaming.SXSSFCell summaryHeadeeNameCell = summaryHeaderRow.createCell(0);
		summaryHeadeeNameCell.setCellValue("Sheet name");
		summaryHeadeeNameCell.setCellStyle(this.headerCellStyle);
		
		org.apache.poi.xssf.streaming.SXSSFCell summaryHeadeeCountCell = summaryHeaderRow.createCell(1);
		summaryHeadeeCountCell.setCellValue("Count");
		summaryHeadeeCountCell.setCellStyle(this.headerCellStyle);
		
		for (org.apache.poi.ss.usermodel.Sheet curWorkSheet : this.swb){
			int rowCount = curWorkSheet.getPhysicalNumberOfRows() - excludeRows;
				
			org.apache.poi.xssf.streaming.SXSSFRow summaryRow = summarySheet.createRow(rowIdx++);
			org.apache.poi.xssf.streaming.SXSSFCell tableNameCell = summaryRow.createCell(0);
			tableNameCell.setCellValue(curWorkSheet.getSheetName());
			
			org.apache.poi.ss.usermodel.Hyperlink tableNameHl = this.createHelper.createHyperlink(org.apache.poi.common.usermodel.HyperlinkType.DOCUMENT);
			tableNameHl.setAddress("'" + curWorkSheet.getSheetName() + "'!A1");
			tableNameCell.setHyperlink(tableNameHl);
			org.apache.poi.xssf.streaming.SXSSFCell tableCountCell = summaryRow.createCell(1);
			tableCountCell.setCellValue(rowCount);
			
		}
		
		/*if(System.getProperty("os.name").startsWith("Windows")){
			summarySheet.autoSizeColumn(0);
			summarySheet.autoSizeColumn(1);
		}*/
		
		summarySheet.setZoom(85);
		
		
		logger.traceExit();
		
	}
	
	/**
	 * Loop all the workbook and remove the page without record
	 * @author manuel.m.speranza
	 * @since 22-09-2016
	 */
	public void cleanNoRecordSheets(){
		logger.traceEntry();
		
		for(int workSheetIdx = this.swb.getNumberOfSheets() -1; workSheetIdx >= 0; workSheetIdx--){
			
			org.apache.poi.xssf.streaming.SXSSFSheet workSheet = this.swb.getSheetAt(workSheetIdx);
			logger.debug(workSheet.getSheetName() + " has " + workSheet.getPhysicalNumberOfRows() + " row(s)");
			if(workSheet.getPhysicalNumberOfRows() <= 1){
				logger.debug("Removing " + workSheet.getSheetName());
				this.swb.removeSheetAt(workSheetIdx);
			}
		}
		
		logger.traceExit();
	}
	
	/**
	 * Flush the workbook data into the file and close the workbook.
	 * @author manuel.m.speranza
	 * @since 31-08-2016
	 */
	public void finalWrite(String targetPath)
	{
		logger.traceEntry();
		try {
			
			org.apache.poi.ooxml.POIXMLProperties props = this.wb.getProperties();
			org.apache.poi.ooxml.POIXMLProperties.CoreProperties coreProp = props.getCoreProperties();
	        coreProp.setCreator(System.getProperty("user.name"));
	        
			String xlsFilename = this.name + ".xlsx";
			if (targetPath != null && !"".equals(targetPath)){
				xlsFilename = targetPath + xlsFilename;
			}
			logger.trace("Writing " + xlsFilename);
			FileOutputStream fileOut = new FileOutputStream(xlsFilename);
			this.swb.write(fileOut);
			fileOut.close();
			this.swb.close();
			
		} catch (FileNotFoundException e) {
			logger.error(e.getMessage(), e);
		} catch (IOException e) {
			logger.error(e.getMessage(), e);
		} finally {
			this.wb = null;
			this.swb = null;
			logger.traceExit();
		}
	}
	/**
	 * Check if the workbook contains worksheets 
	 * @return true if there is not sheet
	 * @since 03-11-2016
	 */
	public boolean isWbEmpty(){
		return (this.swb.getNumberOfSheets() == 0);
	}
	/**
	 * @return The object name
	 * @since 03-11-2016
	 */
	public String getName() {
		return logger.traceExit(name);
	}
	
	
	/**
	 * Tabulate a JSON string into an excel sheet
	 * @author manuel.m.speranza
	 * @throws XlsWrkSheetException 
	 * @since 11-03-2022
	 */
	
	public void getJsonResult(String sheetHeader, String sheetName, String jsonStr) throws XlsWrkSheetException {
		logger.traceEntry();
		
		org.apache.poi.xssf.streaming.SXSSFSheet workSheet = null;
		if(sheetName != null && !"".equals(sheetName)) {
			if(sheetName.length() > 31){
				throw new XlsWrkSheetException("Work-sheet " + sheetName + " [" + sheetName.length() + "] must not have more than 31 characters");
			}
			workSheet = this.swb.createSheet(sheetName);
		} else {
			workSheet = this.swb.createSheet();
		}
		
		int rowIdx = 0;
		if(sheetHeader != null && !"".equals(sheetHeader)) {
			org.apache.poi.xssf.streaming.SXSSFRow headerRow = workSheet.createRow(rowIdx);
			org.apache.poi.xssf.streaming.SXSSFCell columnNameCell = headerRow.createCell(0);
			columnNameCell.setCellValue(sheetHeader);
			columnNameCell.setCellStyle(this.headerCellStyle);
			workSheet.createFreezePane(0, 1);
			rowIdx++;
		}
		Map<Integer, Integer> columnsWidth = new HashMap<Integer, Integer>();
		if(jsonStr.startsWith("{")) {
			this.writeJsonObject(workSheet, columnsWidth, new JSONObject(jsonStr), rowIdx, 0);
		} else if(jsonStr.startsWith("[")) {
			this.writeJsonArray(workSheet, columnsWidth, new JSONArray(jsonStr), rowIdx, 0);
		}
		logger.debug("columnsWidth {}", columnsWidth);
		
		
				
		for(Entry<Integer, Integer> curEntry : columnsWidth.entrySet()){
			//width = Truncate([{Number of Visible Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
			//workSheet.setColumnWidth(colIdx, (columnsWidth.get(colIdx) * 256));
			int width = (int) Math.ceil((curEntry.getValue() + 1) * 7.2 / 7 * 256) + 5;
			workSheet.setColumnWidth(curEntry.getKey(), width > 256*256 ? 256*256 : width);
		}
		
		workSheet.setZoom(85);
		
		logger.traceExit();
	}
	
	/**
	 * Manage the conversion of a JSON Object to an excel value
	 * @author manuel.m.speranza 
	 * @since 11-03-2022
	 */
	public int writeJsonObject(org.apache.poi.xssf.streaming.SXSSFSheet workSheet, Map<Integer, Integer> columnsWidth, JSONObject jsonObj, int startRowIdx, int startColIdx) {
		logger.traceEntry();
		int rowIdx = startRowIdx;
		
		Iterator<String> jsonObjKeyIter = jsonObj.keys();
		while(jsonObjKeyIter.hasNext()) {
			String jsonKey = jsonObjKeyIter.next();
			logger.debug("jsonKey " + jsonKey + " @" + rowIdx + ", " + startColIdx);
			
			org.apache.poi.xssf.streaming.SXSSFRow contentRow = workSheet.getRow(rowIdx);
			if(contentRow == null) {
				contentRow = workSheet.createRow(rowIdx);
			}
			
			org.apache.poi.xssf.streaming.SXSSFCell contentCell = contentRow.getCell(startColIdx);
			if(contentCell == null) {
				contentCell = contentRow.createCell(startColIdx);
			}
			contentCell.setCellValue(jsonKey);
			int cellValueLength = jsonKey.length();
			if(!columnsWidth.containsKey(startColIdx) || cellValueLength > columnsWidth.get(startColIdx) && cellValueLength < 256){
				columnsWidth.put(startColIdx, cellValueLength);
			}
			contentCell.setCellStyle(this.defaultCellStyle);
			
			rowIdx = this.manageJsonValue(workSheet, columnsWidth, contentRow, jsonObj.get(jsonKey), rowIdx, startColIdx);
		}
		
		return logger.traceExit(rowIdx);
	}
	
	/**
	 * Manage the conversion of a JSON Array to an excel value
	 * @author manuel.m.speranza 
	 * @since 11-03-2022
	 */
	public int writeJsonArray(org.apache.poi.xssf.streaming.SXSSFSheet workSheet, Map<Integer, Integer> columnsWidth, JSONArray jsonArray, int startRowIdx, int startColIdx) {
		logger.traceEntry();
		int rowIdx = startRowIdx;
		int arrayItemIdx = 0;
		Iterator<Object> jsonArrayItr = jsonArray.iterator();
		while(jsonArrayItr.hasNext()) {
			logger.debug("arrayItemIdx #" + arrayItemIdx + " @" + rowIdx + ", " + startColIdx);
			Object curObject = jsonArrayItr.next();
			
			org.apache.poi.xssf.streaming.SXSSFRow contentRow = workSheet.getRow(rowIdx);
			if(contentRow == null) {
				contentRow = workSheet.createRow(rowIdx);
			}
			
			org.apache.poi.xssf.streaming.SXSSFCell contentCell = contentRow.getCell(startColIdx);
			if(contentCell == null) {
				contentCell = contentRow.createCell(startColIdx);
			}
			contentCell.setCellValue("#" + arrayItemIdx);
			int cellValueLength = String.valueOf(arrayItemIdx).length() + 1;
			if(!columnsWidth.containsKey(startColIdx) || cellValueLength > columnsWidth.get(startColIdx) && cellValueLength < 256){
				columnsWidth.put(startColIdx, cellValueLength);
			}
			
			contentCell.setCellStyle(this.defaultCellStyle);
			
			rowIdx = this.manageJsonValue(workSheet, columnsWidth, contentRow, curObject, rowIdx, startColIdx);
			arrayItemIdx++;
		}
		
		return logger.traceExit(rowIdx);
	}
	
	/**
	 * Manage the JSON value conversion
	 * @author manuel.m.speranza 
	 * @since 11-03-2022
	 */
	public int manageJsonValue(org.apache.poi.xssf.streaming.SXSSFSheet workSheet, Map<Integer, Integer> columnsWidth, org.apache.poi.xssf.streaming.SXSSFRow contentRow, Object value, int startRowIdx, int startColIdx) {
		logger.traceEntry();
		int rowIdx = startRowIdx;
		
		if(value instanceof JSONObject) {
			rowIdx = this.writeJsonObject(workSheet, columnsWidth, (JSONObject) value, rowIdx, startColIdx+1);
		} else if(value instanceof Map) {
			logger.warn("instanceof Map");
		} else if(value instanceof JSONArray) {
			rowIdx = this.writeJsonArray(workSheet, columnsWidth, (JSONArray) value, rowIdx, startColIdx+1);	
		} else if(value instanceof List) {
			logger.warn("instanceof List");
		} else {
			logger.debug("value " + value + " @" + rowIdx + ", " + startColIdx+1);
			org.apache.poi.xssf.streaming.SXSSFCell contentCell = contentRow.getCell(startColIdx+1);
			if(contentCell == null) {
				contentCell = contentRow.createCell(startColIdx+1);
			}
			contentCell.setCellValue(value.toString());
			contentCell.setCellStyle(this.defaultCellStyle);
			rowIdx++;
		}
		
		if(rowIdx - startRowIdx > 1) {
			workSheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(startRowIdx, rowIdx-1, startColIdx, startColIdx));
		}
		
		return logger.traceExit(rowIdx);
	}
	
}
