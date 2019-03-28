package me.hoffnungland.poi.corner.extractshapesvsdx;

import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xdgf.usermodel.XDGFPage;
import org.apache.poi.xdgf.usermodel.XDGFShape;
import org.apache.poi.xdgf.usermodel.XmlVisioDocument;

/**
 * Print out all the shape text split by sheet
 *
 */
public class App 
{
	private static Logger logger = LogManager.getLogger(App.class);

	public static void main( String[] args ){


		try {
			logger.info("Initialize the list of visio");
			FileFilter visioFilter = new FileFilter(){

				@Override
				public boolean accept(File pathname) {
					if(pathname.isFile()){
						if(pathname.getName().endsWith(".vsdx")){
							return true;
						}
					}
					return false;
				}

			};

			File visioDir = new File(".");
			for (File curFile : visioDir.listFiles(visioFilter)){
				logger.debug("Loading " + curFile.getName());
				//File visioFile = new File("");
				FileInputStream visioFileIn = new FileInputStream(curFile);
				XmlVisioDocument visioDoc = new XmlVisioDocument(visioFileIn);

				for (XDGFPage curPage : visioDoc.getPages()){
					
					String pageName = curPage.getName().trim().replaceAll("^\\[(\\w+)\\]:?\\s", "$1\t");
					
					for(XDGFShape curShape : curPage.getContent().getShapes()){
						if(curShape.getText() != null){
							String shapeText = curShape.getText().getTextContent().trim().replace('‘', '\'')
									.replace('’', '\'')
									.replace('“', '"')
									.replace('”', '"')
									.replaceAll("\\s+", " ")
									.replaceAll("\n", "\t")
									.replaceAll("^(\\w+\\d+):?\\s", "$1\t");
							
							logger.info(pageName + "\t" + shapeText);
						}
					}
				}

				visioDoc.close();
				visioFileIn.close();

			}

		} catch (IOException e) {
			logger.error(e.getMessage(), e);
		}
	}
}
