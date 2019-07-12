package com.hoffnungland.poi.corner.orcxlsreport;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileFilter;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.hoffnungland.db.corner.dbconn.ConnectionManager;
import com.hoffnungland.db.corner.dbconn.StatementCached;
import com.hoffnungland.db.corner.oracleconn.OrclConnectionManager;
import com.hoffnungland.poi.corner.orcxlsreport.ExcelManager;
import com.hoffnungland.poi.corner.orcxlsreport.XlsWrkSheetException;



/**
 * Main class
 * @author ***REMOVED***
 * @since 31-08-2016
 * @version 0.1
 */
public class App 
{
	private static final Logger logger = LogManager.getLogger(App.class);

	public static void main( String[] args )
	{
		logger.traceEntry();
		
		if(args.length < 4){
			logger.error("Wrong input parameters. Params are: ConnectionName ProjectName ExcelName TargetPath");
			return;
		}
		
		String connectionName = args[0];
		String ProjectName = args[1];
		String inExcelName = args[2];
		String targetPath  = args[3];
		
		OrclConnectionManager dbManager = new OrclConnectionManager();
		
		ExcelManager xlsMng = null;
		
		try {
			logger.info("DB Manager connecting to " + connectionName);
			String connectionPropertyPath = "./etc/connections/" + connectionName + ".properties";
			dbManager.connect(connectionPropertyPath);

			logger.info("Initialize the list of queries");
			FileFilter queriesFilter = new FileFilter(){

				@Override
				public boolean accept(File pathname) {
					if(pathname.isFile()){
						if(pathname.getName().endsWith(".sql")){
							return true;
						}
					}
					return false;
				}

			};
			
			File queriesDir = new File("./" + ProjectName + "/queries");
			File[] queriesDirList = queriesDir.listFiles(queriesFilter);
			if(queriesDirList != null && queriesDirList.length > 0){
				if(xlsMng == null) {
					logger.info("Initialize the excel");
					xlsMng = new ExcelManager(inExcelName);
				}
				for (File curFile : queriesDirList){
					logger.debug("Loading " + curFile.getName());
					
					logger.info("Executing the query " + curFile.getName());
					StatementCached<PreparedStatement> prepStm = dbManager.executeQuery("./" + ProjectName + "/queries/" + curFile.getName());
	
					logger.info("Put query " + curFile.getName() + " result into the excel file");
					xlsMng.getQueryResult(prepStm);
	
				}
			}
			
			File queriesJntDir = new File("./" + ProjectName + "/queriesJnt");
			File[] queriesJntDirList = queriesJntDir.listFiles(queriesFilter);
			if(queriesJntDirList != null && queriesJntDirList.length > 0){
				if(xlsMng == null) {
					logger.info("Initialize the excel");
					xlsMng = new ExcelManager(inExcelName);
				}
				for (File curFile : queriesJntDirList){
					logger.debug("Loading " + curFile.getName());
	
					StatementCached<PreparedStatement> prepStm =  dbManager.generateAndExecuteQueryWithJunction("./" + ProjectName + "/queriesJnt/" + curFile.getName());
	
					logger.info("Put query " + curFile.getName() + " result into the excel file");
					xlsMng.getQueryResult(prepStm);
				}
			}
			File queriesJntCacheDir = new File("./" + ProjectName + "/queriesJntCached");
			File[] queriesJntCacheDirList = queriesJntCacheDir.listFiles(queriesFilter);
			if(queriesJntCacheDirList != null && queriesJntCacheDirList.length > 0){
				if(xlsMng == null) {
					logger.info("Initialize the excel");
					xlsMng = new ExcelManager(inExcelName);
				}
				for (File curFile : queriesJntCacheDirList){
					logger.debug("Loading " + curFile.getName());
	
					StatementCached<PreparedStatement> prepStm =  dbManager.executeQueryWithJunction("./" + ProjectName + "/queriesJntCached/" + curFile.getName());
	
					logger.info("Put query " + curFile.getName() + " result into the excel file");
					xlsMng.getQueryResult(prepStm);
				}
			}
			
			if(xlsMng != null) {
				xlsMng.cleanNoRecordSheets();
				xlsMng.createSummaryPage(1);
			}


		} catch (SQLException e) {
			logger.error(e.getMessage(), e);
		} catch (FileNotFoundException e) {
			logger.error(e.getMessage(), e);
		} catch (IOException e) {
			logger.error(e.getMessage(), e);
		} catch (XlsWrkSheetException e) {
			logger.error(e.getMessage(), e);
		} finally {

			if(xlsMng != null) {
				xlsMng.finalWrite(targetPath);
				xlsMng = null;
			}
			//dbManager.disconnect();

		}

		logger.info("Initialize the list of metatables files");
		FileFilter txtFilter = new FileFilter(){

			@Override
			public boolean accept(File pathname) {
				if(pathname.isFile()){
					if(pathname.getName().endsWith(".txt")){
						return true;
					}
				}
				return false;
			}

		};



		File tablesDir = new File("./" + ProjectName + "/tables");
		writeTables(tablesDir, txtFilter, dbManager, targetPath);
		File metadataDir = new File("./" + ProjectName + "/metadata");
		writeTables(metadataDir, txtFilter, dbManager, targetPath);
		
		dbManager.disconnect();
		logger.traceExit();
	}
	
	
	public static void writeTables(File tablesDir, FileFilter txtFilter, ConnectionManager dbManager, String targetPath) {
		logger.traceEntry();
		for (File curFile : tablesDir.listFiles(txtFilter)){
			
			ExcelManager xlsMng = null;
			try{
				logger.info("Working " + curFile.getName());

				BufferedReader reader = new BufferedReader( new FileReader (curFile));
				String         line = null;
				
				Pattern p = Pattern.compile("(\\w+\\.)?(\\w+)");
				
				int suffixPos = curFile.getName().lastIndexOf('.');
				String excelName = curFile.getName().substring(0, suffixPos);
				
				xlsMng  = new ExcelManager(excelName);
				
				while( ( line = reader.readLine() ) != null ) {

					Matcher msgMatcher = p.matcher(line);

					if(msgMatcher.find()){
						String queryFileName = msgMatcher.group(2);
						
						//StatementCached<PreparedStatement> prepStm = dbManager.executeFullTableQuery("./" + ProjectName + "/tables/" + queryFileName, line);
						StatementCached<PreparedStatement> prepStm = dbManager.executeFullTableQuery(queryFileName, line);
						
						logger.info("Put query " + queryFileName + " result into the excel file");
						if("metadata".equals(tablesDir.getName())) {
							xlsMng.getMetadataResult(prepStm);
						}else {
							xlsMng.getQueryResult(prepStm);
						}
					}
				}
				reader.close();

				//xlsMng.cleanNoRecordSheets();
				xlsMng.createSummaryPage(2);

			} catch (SQLException e) {
				logger.error(e.getMessage(), e);
			} catch (FileNotFoundException e) {
				logger.error(e.getMessage(), e);
			} catch (IOException e) {
				logger.error(e.getMessage(), e);
			} catch (XlsWrkSheetException e) {
				logger.error(e.getMessage(), e);
			} finally {

				xlsMng.finalWrite(targetPath);
			}
		}
		logger.traceExit();
		
	}
	
}

