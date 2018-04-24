package net.dtdns.hoffnungland.poi.corner.orcxlsreport;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class Pega_data_tablesApp {

	private static DateFormat datefileFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
	
	public static void main(String[] args) {
		
		String ProjectName = "Pega_data_tables";
		
		String connectionName = "PEGA_DEV";
		//String connectionName = "PEGA_TEMP";
		//String connectionName = "PEGA_PROD";
		//String connectionName = "PEGA_CLONE";
		
		Date today = Calendar.getInstance().getTime();  
		String reportDate = datefileFormat.format(today);
		
		String excelName = "Pega_data_tables_" + connectionName + "_" + reportDate;
		String targetPath  = null;
		
		
		String[] inArgs = new String[4];
		
		inArgs[0]= connectionName;
		inArgs[1]= ProjectName;
		inArgs[2]= excelName;
		inArgs[3]= targetPath;
		
		App.main(inArgs);

	}

}
