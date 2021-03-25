package com.hoffnungland.poi.corner.jsonxlsloader;

import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

public class ColumnHeader {
	public String name;
	public int row;
	public int column;
	public Map<String, ColumnHeader> childrenCols = new HashMap<String, ColumnHeader>();
	
	@Override
	public String toString() {
		
		StringBuilder strBuild = new StringBuilder();
		
		strBuild.append(this.name + "\t" + this.row + "\t" + this.column + "\n");
		
		for(Entry<String, ColumnHeader> curEntry : this.childrenCols.entrySet()) {
			strBuild.append(curEntry);
		}
		return strBuild.toString();
	}
	
	
	
}
