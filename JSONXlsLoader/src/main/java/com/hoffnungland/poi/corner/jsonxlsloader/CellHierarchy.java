package com.hoffnungland.poi.corner.jsonxlsloader;

import java.util.ArrayList;
import java.util.List;
import java.util.Map.Entry;

public class CellHierarchy {
	public Object cellObject;
	public int row;
	public int column;
	public int width;
	public int depth;
	public List<CellHierarchy> listChild = new ArrayList<CellHierarchy>();
	
	@Override
	public String toString() {
		
		StringBuilder strBuild = new StringBuilder();
		
		strBuild.append(this.cellObject + "\t" + this.row + "\t" + this.column + "\n");
		
		for(CellHierarchy curCellHierarchy : this.listChild) {
			strBuild.append(curCellHierarchy);
		}
		return strBuild.toString();
	}
	
}
