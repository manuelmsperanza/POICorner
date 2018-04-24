package net.dtdns.hoffunungland.poi.corner.orcxlsreport;

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
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import net.dtdns.hoffunungland.db.corner.dbconn.StatementCached;



/**
 * Manage the work-sheet data read and write.
 * @version 0.7
 * @author ***REMOVED***
 * @since 31-08-2016
 */

public class ExcelManager {

	private static final Logger logger = LogManager.getLogger(ExcelManager.class);
	//private static String ls = System.getProperty("line.separator");
	private String name;
	private org.apache.poi.ss.usermodel.Workbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook();
	private org.apache.poi.xssf.usermodel.XSSFCellStyle headerCellStyle;
	private org.apache.poi.xssf.usermodel.XSSFCellStyle defaultCellStyle;
	private org.apache.poi.xssf.usermodel.XSSFCellStyle dateCellStyle;
	private org.apache.poi.ss.usermodel.CreationHelper createHelper = wb.getCreationHelper();

	/**
	 * Constructor with input name string. Define also the styles.
	 * @param name The target excel file name prefix
	 * @author ***REMOVED***
	 * @since 31-08-2016
	 */
	
	public ExcelManager(String name){
		this.name = name;
		
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
		
		this.dateCellStyle = (XSSFCellStyle) this.wb.createCellStyle();
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

		org.apache.poi.ss.usermodel.Sheet workSheet = this.wb.createSheet(sheetName);
		this.createSheetHeader(workSheet, prepStm);
		this.createSheetContent(workSheet, prepStm);		
		workSheet.createFreezePane(0, 1);
		workSheet.setZoom(85);
		
		logger.traceExit();
	}

	/**
	 * Add the top row of the work-sheet. Get the information from the ResultSetMetaData of query's ResultSet.
	 * @param workSheet the working work-sheet
	 * @param query the executed query having a valid ResultSet.
	 * @throws SQLException
	 * @author ***REMOVED***
	 * @since 31-08-2016 
	 */
	private void createSheetHeader(org.apache.poi.ss.usermodel.Sheet workSheet, StatementCached<PreparedStatement> prepStm) throws SQLException{
		logger.traceEntry();
		org.apache.poi.ss.usermodel.Row headerRow = workSheet.createRow(0);
		ResultSet resRs = prepStm.getStm().getResultSet();

		ResultSetMetaData rsmd = resRs.getMetaData();
		for(int headerIdx = 0; headerIdx < rsmd.getColumnCount(); headerIdx++){
			org.apache.poi.ss.usermodel.Cell columnNameCell = headerRow.createCell(headerIdx);
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
	 * Add a row for each record within the query's ResultSet.
	 * It manage the following Oracle data type: VARCHAR, CHAR, CLOB, DATE, TIME, TIMESTAMP and NUMBER
	 * @param workSheet the working work-sheet
	 * @param query the executed query having a valid ResultSet.
	 * @throws SQLException
	 * @throws IOException
	 * @author ***REMOVED***
	 * @since 31-08-2016
	 */
	private void createSheetContent(org.apache.poi.ss.usermodel.Sheet workSheet, StatementCached<PreparedStatement> prepStm) throws SQLException, IOException{
		logger.traceEntry();
		int rowId = 1;
		ResultSet resRs = prepStm.getStm().getResultSet();
		ResultSetMetaData rsmd = resRs.getMetaData();
		
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
		Calendar tsCal = Calendar.getInstance();
		
		while (resRs.next()) {
			org.apache.poi.ss.usermodel.Row bodyRow = workSheet.createRow(rowId);
			for(int colIdx = 0; colIdx < rsmd.getColumnCount(); colIdx++){

				org.apache.poi.ss.usermodel.Cell contentCell = bodyRow.createCell(colIdx);
				contentCell.setCellStyle(this.defaultCellStyle);
				
				if(resRs.getObject(colIdx + 1) != null){
					
					int columnType = rsmd.getColumnType(colIdx + 1);
					String columnTypeName = rsmd.getColumnTypeName(colIdx + 1);
					
					if(columnType == Types.VARCHAR || columnType == Types.CHAR){
						contentCell.setCellValue(resRs.getString(colIdx + 1));
						
					} else if(columnType == Types.LONGVARCHAR){
						
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
						java.sql.Date dateVal = resRs.getDate(colIdx + 1);
						if(dateVal != null){
							tsCal.setTime(dateVal);
							//contentCell.setCellValue(df.format(tsCal.getTime()));
							contentCell.setCellValue(tsCal.getTime());
							contentCell.setCellStyle(this.dateCellStyle);
						}
					} else if(columnType == Types.TIME){
						
						Time timeVal = resRs.getTime(colIdx + 1);
						if(timeVal != null){
							tsCal.setTime(timeVal);
							//contentCell.setCellValue(df.format(tsCal.getTime()));
							contentCell.setCellValue(tsCal.getTime());
							contentCell.setCellStyle(this.dateCellStyle);
						}
						
					} else if(columnType == Types.TIMESTAMP || columnTypeName.equals("TIMESTAMP WITH LOCAL TIME ZONE")){
						Timestamp tsVal = resRs.getTimestamp(colIdx + 1);
						if(tsVal != null){
							tsCal.setTimeInMillis(tsVal.getTime());
							//contentCell.setCellValue(df.format(tsCal.getTime()));
							contentCell.setCellValue(tsCal.getTime());
							contentCell.setCellStyle(this.dateCellStyle);
						}
					} else if(columnType == Types.BLOB){
						
					} else {
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
	public void createSummaryPage(){
		logger.traceEntry();
		
		org.apache.poi.ss.usermodel.Sheet selectedSheet = this.wb.getSheetAt(this.wb.getActiveSheetIndex());
		org.apache.poi.ss.usermodel.Sheet summarySheet = this.wb.createSheet("Summary");
		
		this.wb.setSheetOrder("Summary", 0);
		
		summarySheet.setSelected(true);
		selectedSheet.setSelected(false);

		this.wb.setActiveSheet(0);
		
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
