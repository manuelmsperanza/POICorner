package com.hoffnungland.poi.corner.orcxlsloader;

import java.io.IOException;
import java.sql.SQLException;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.xml.sax.SAXException;

import com.hoffnungland.db.corner.oracleconn.OrclConnectionManager;
import com.hoffnungland.poi.corner.dbxlsreport.XlsWrkSheetException;


/**
 * Hello world!
 *
 */
public class App 
{
	private static final Logger logger = LogManager.getLogger(App.class);

	public static void main( String[] args )
	{
		logger.traceEntry();
        
		if(args.length < 3){
			logger.error("Wrong input parameters. Params are: ConnectionName ExcelName Source_Path");
			return;
		}
		
		String connectionName = args[0];
		String inExcelName = args[1];
		String sourcePath  = args[2];
		
		OrclConnectionManager dbManager = new OrclConnectionManager();
		
		logger.info("Initialize the excel " + inExcelName);
		ExcelLoader xlsMng = new ExcelLoader(sourcePath, inExcelName, connectionName);
		
		try {
			
			
			logger.info("DB Manager connecting to " + connectionName);
			String connectionPropertyPath = "./etc/connections/" + connectionName + ".properties";
			dbManager.connect(connectionPropertyPath);
			
			xlsMng.loadWb(dbManager);
        
		} catch (SQLException e) {
			logger.error(e.getMessage(), e);
		} catch (IOException e) {
			logger.error(e.getMessage(), e);
		} catch (SAXException e) {
			logger.error(e.getMessage(), e);
		} catch (ParserConfigurationException e) {
			logger.error(e.getMessage(), e);
		} catch (XlsWrkSheetException e) {
			logger.error(e.getMessage(), e);
		}
		
        dbManager.disconnect();
		logger.traceExit();
    }
}
