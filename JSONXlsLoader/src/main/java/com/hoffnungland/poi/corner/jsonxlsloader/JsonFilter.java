package com.hoffnungland.poi.corner.jsonxlsloader;

import java.io.File;

import javax.swing.filechooser.FileFilter;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public class JsonFilter extends FileFilter {
	
	private static final Logger logger = LogManager.getLogger(JsonFilter.class);
	
	@Override
	public boolean accept(File f) {
		logger.traceEntry();
		if (f.isDirectory()) {
	        return logger.traceExit(true);
	    }
		
		String ext = null;
        String s = f.getName();
        int i = s.lastIndexOf('.');

        if (i > 0 &&  i < s.length() - 1) {
            ext = s.substring(i+1).toLowerCase();
            if(ext.equalsIgnoreCase("json")) {
            	return logger.traceExit(true);
            }
        }
		return logger.traceExit(false);
	}

	@Override
	public String getDescription() {
		return "JSON *.json";
	}

}
