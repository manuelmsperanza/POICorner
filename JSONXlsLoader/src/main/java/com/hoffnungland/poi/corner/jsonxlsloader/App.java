package com.hoffnungland.poi.corner.jsonxlsloader;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.SpringLayout;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.json.JSONArray;
import org.json.JSONML;
import org.json.JSONObject;

import java.awt.Component;
import javax.swing.Box;
import java.awt.Dimension;
import javax.swing.JTextField;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.awt.event.ActionEvent;
import javax.swing.JLabel;
import javax.swing.JFormattedTextField;
import javax.swing.JRadioButton;
import javax.swing.JCheckBox;

public class App implements ActionListener{
	private static final Logger logger = LogManager.getLogger(App.class);
	
	private JFrame frame;
	private JTextField jsonTextField;
	private JTextField targetDirTextField;
	private JTextField xlsxNameTextField;

	private JRadioButton flatRadioButton;

	private JRadioButton differentSheetsRadioButton;

	private JCheckBox skipArrayCheckBox;
	
	private static final String LOAD_JSON_ACTION = "Load JSON Action";
	private static final String TARGET_DIR_ACTION = "Targer DIR Action";
	private static final String CONVERT_ACTION = "Convert Action";

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		
		logger.traceEntry();
		
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
				| UnsupportedLookAndFeelException e) {
			logger.error(e);
			e.printStackTrace();
		}
		
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					App ***REMOVED***ow = new App();
					***REMOVED***ow.frame.setVisible(true);
				} catch (Exception e) {
					logger.error(e);
					e.printStackTrace();
				}
			}
		});
		
		logger.traceExit();
	}

	/**
	 * Create the application.
	 */
	public App() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		
		logger.traceEntry();
		
		frame = new JFrame();
		frame.setResizable(false);
		frame.setBounds(100, 100, 490, 180);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		SpringLayout springLayout = new SpringLayout();
		frame.getContentPane().setLayout(springLayout);
		
		JButton loadJsonButton = new JButton("Load *.json");
		loadJsonButton.setActionCommand(LOAD_JSON_ACTION);
		loadJsonButton.addActionListener(this);
		
		springLayout.putConstraint(SpringLayout.NORTH, loadJsonButton, 10, SpringLayout.NORTH, frame.getContentPane());
		springLayout.putConstraint(SpringLayout.WEST, loadJsonButton, 10, SpringLayout.WEST, frame.getContentPane());
		frame.getContentPane().add(loadJsonButton);
		
		jsonTextField = new JTextField();
		springLayout.putConstraint(SpringLayout.NORTH, jsonTextField, 1, SpringLayout.NORTH, loadJsonButton);
		springLayout.putConstraint(SpringLayout.WEST, jsonTextField, 10, SpringLayout.EAST, loadJsonButton);
		springLayout.putConstraint(SpringLayout.EAST, jsonTextField, -10, SpringLayout.EAST, frame.getContentPane());
		frame.getContentPane().add(jsonTextField);
		jsonTextField.setColumns(10);
		
		JButton targetDirButton = new JButton("Target Dir");
		targetDirButton.setActionCommand(TARGET_DIR_ACTION);
		targetDirButton.addActionListener(this);
		springLayout.putConstraint(SpringLayout.NORTH, targetDirButton, 15, SpringLayout.SOUTH, loadJsonButton);
		springLayout.putConstraint(SpringLayout.WEST, targetDirButton, 0, SpringLayout.WEST, loadJsonButton);
		springLayout.putConstraint(SpringLayout.EAST, targetDirButton, 0, SpringLayout.EAST, loadJsonButton);
		frame.getContentPane().add(targetDirButton);
		
		targetDirTextField = new JTextField();
		springLayout.putConstraint(SpringLayout.NORTH, targetDirTextField, 1, SpringLayout.NORTH, targetDirButton);
		springLayout.putConstraint(SpringLayout.WEST, targetDirTextField, 0, SpringLayout.WEST, jsonTextField);
		springLayout.putConstraint(SpringLayout.EAST, targetDirTextField, -10, SpringLayout.EAST, frame.getContentPane());
		frame.getContentPane().add(targetDirTextField);
		targetDirTextField.setColumns(10);
		
		JLabel xlsxNameLabel = new JLabel("Excel Name");
		springLayout.putConstraint(SpringLayout.NORTH, xlsxNameLabel, 15, SpringLayout.SOUTH, targetDirButton);
		springLayout.putConstraint(SpringLayout.WEST, xlsxNameLabel, 0, SpringLayout.WEST, loadJsonButton);
		frame.getContentPane().add(xlsxNameLabel);
		
		xlsxNameTextField = new JTextField();
		springLayout.putConstraint(SpringLayout.NORTH, xlsxNameTextField, -3, SpringLayout.NORTH, xlsxNameLabel);
		springLayout.putConstraint(SpringLayout.WEST, xlsxNameTextField, 10, SpringLayout.EAST, xlsxNameLabel);
		springLayout.putConstraint(SpringLayout.EAST, xlsxNameTextField, -10, SpringLayout.EAST, frame.getContentPane());
		frame.getContentPane().add(xlsxNameTextField);
		xlsxNameTextField.setColumns(10);
		
		flatRadioButton = new JRadioButton("One Sheet Flat");
		flatRadioButton.setSelected(true);
		springLayout.putConstraint(SpringLayout.NORTH, flatRadioButton, 15, SpringLayout.SOUTH, xlsxNameLabel);
		springLayout.putConstraint(SpringLayout.WEST, flatRadioButton, 10, SpringLayout.WEST, frame.getContentPane());
		frame.getContentPane().add(flatRadioButton);
		
		differentSheetsRadioButton = new JRadioButton("Separate sheets");
		springLayout.putConstraint(SpringLayout.NORTH, differentSheetsRadioButton, 0, SpringLayout.NORTH, flatRadioButton);
		springLayout.putConstraint(SpringLayout.WEST, differentSheetsRadioButton, 10, SpringLayout.EAST, flatRadioButton);
		frame.getContentPane().add(differentSheetsRadioButton);
		
		ButtonGroup radioBtnGroup = new ButtonGroup();
        radioBtnGroup.add(flatRadioButton);
        radioBtnGroup.add(differentSheetsRadioButton);
		
        skipArrayCheckBox = new JCheckBox("Skip non-object array");
        springLayout.putConstraint(SpringLayout.NORTH, skipArrayCheckBox, 0, SpringLayout.NORTH, differentSheetsRadioButton);
        springLayout.putConstraint(SpringLayout.WEST, skipArrayCheckBox, 10, SpringLayout.EAST, differentSheetsRadioButton);
		frame.getContentPane().add(skipArrayCheckBox);
        
		JButton convertButton = new JButton("Convert");
		springLayout.putConstraint(SpringLayout.NORTH, convertButton, 12, SpringLayout.SOUTH, xlsxNameTextField);
		springLayout.putConstraint(SpringLayout.WEST, convertButton, 10, SpringLayout.EAST, skipArrayCheckBox);
		convertButton.setActionCommand(CONVERT_ACTION);
		convertButton.addActionListener(this);
		frame.getContentPane().add(convertButton);
		
		
		
		logger.traceExit();
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		logger.traceEntry();
		
		switch (e.getActionCommand()) {
		case LOAD_JSON_ACTION:
			this.loadJsonFile();
			break;
		case TARGET_DIR_ACTION:
			this.chooseTargetDir();
			break;
		case CONVERT_ACTION :
			this.convertJsonToExcel();
			break;
		}
		logger.traceExit();
	}
	
	

	private void loadJsonFile() {
		
		logger.traceEntry();
		
		JFileChooser fc = new JFileChooser();
		JsonFilter fcJsonFiler = new JsonFilter();
		fc.setMultiSelectionEnabled(false);
		fc.setFileFilter(fcJsonFiler);
		fc.addChoosableFileFilter(fcJsonFiler);
		fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		
		int returnVal = fc.showOpenDialog(this.frame);
		
		if(returnVal == JFileChooser.APPROVE_OPTION) {
			String jsonFilePath = fc.getSelectedFile().getPath();
			this.jsonTextField.setText(fc.getSelectedFile().getPath());
			String jsonFolderPath = jsonFilePath.substring(0, jsonFilePath.lastIndexOf(File.separator));
			this.targetDirTextField.setText(jsonFolderPath);
			String jsonFileName = fc.getSelectedFile().getName();
			this.xlsxNameTextField.setText(jsonFileName.substring(0, jsonFileName.lastIndexOf(".")) + ".xlsx");
		}
		
		logger.traceExit();
	}
	
	private void chooseTargetDir() {
		
		logger.traceEntry();
		
		JFileChooser fc = new JFileChooser();
		fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		fc.setMultiSelectionEnabled(false);
		int returnVal = fc.showOpenDialog(this.frame);
		
		if(returnVal == JFileChooser.APPROVE_OPTION) {
			this.targetDirTextField.setText(fc.getSelectedFile().getPath());
		}
		
		logger.traceExit();
	}
	
	
	private void convertJsonToExcel() {
		
		logger.traceEntry();
		try {
			String jsonStr = Files.readString(Path.of(this.jsonTextField.getText()));
			
			Map<String, ColumnHeader> mapColumnHeader = new HashMap<String, ColumnHeader>();
			List<CellHierarchy> listCells = new ArrayList<CellHierarchy>();

			if(jsonStr.startsWith("{")) {
			
				JSONObject jsonObj = new JSONObject(jsonStr);

				Iterator<String> jsonObjKeyIter = jsonObj.keys();
				while(jsonObjKeyIter.hasNext()) {
					String jsonKey = jsonObjKeyIter.next();
					this.manageJsonEntryObject(jsonKey, jsonObj.get(jsonKey), mapColumnHeader, listCells, 0, 0);
				}
				
				/*for(Entry<String, Object> curEntry : jsonObj.toMap().entrySet()) {
					this.manageJsonEntryObject(curEntry.getKey(), curEntry.getValue(), mapColumnHeader, listCells, 0, 0);
				}*/
			} else {
				JSONArray jsonArray = new JSONArray(jsonStr);
				this.manageJsonArrayObject(null, jsonArray, mapColumnHeader, listCells, 0);
			}
			
			logger.debug(mapColumnHeader);
			logger.debug(listCells);
			
		} catch (IOException e) {
			logger.error(e);
			e.printStackTrace();
		}
		logger.traceExit();
	}
	
	private int manageJsonEntryObject(String key, Object value, Map<String, ColumnHeader> mapCols, List<CellHierarchy> listCells, int headerIdx, int arrayIdx) {
		logger.traceEntry();
		int depth = 0;
		ColumnHeader columnHeader = null;
		if(mapCols.containsKey(key)) {
			columnHeader = mapCols.get(key);
		} else {
			columnHeader = new ColumnHeader();
			columnHeader.name = key;
			columnHeader.row = headerIdx;
			columnHeader.column = mapCols.size();
			mapCols.put(key, columnHeader);
		}
		
		CellHierarchy cell = new CellHierarchy();
		cell.row = arrayIdx;
		cell.column = columnHeader.column;
		listCells.add(cell);
		
		if(value instanceof JSONObject) {
			
			JSONObject jsonObj = (JSONObject) value;

			Iterator<String> jsonObjKeyIter = jsonObj.keys();
			while(jsonObjKeyIter.hasNext()) {
				String jsonKey = jsonObjKeyIter.next();
				this.manageJsonEntryObject(jsonKey, jsonObj.get(jsonKey), columnHeader.childrenCols, cell.listChild, headerIdx + 1, 0);
			}
			
		} else if(value instanceof Map) {
			logger.warn("instanceof Map");
			Map<String, Object> objects = (Map<String, Object>) value;
			for(Entry<String, Object> curEntry : objects.entrySet()) {
				this.manageJsonEntryObject(curEntry.getKey(), curEntry.getValue(), columnHeader.childrenCols, cell.listChild, headerIdx + 1, 0);
			}
		} else if(value instanceof JSONArray) {
			JSONArray jsonArray = (JSONArray) value;
			this.manageJsonArrayObject(key, jsonArray, columnHeader.childrenCols, cell.listChild, headerIdx + 1);
			
		} else if(value instanceof List) {
			logger.warn("instanceof List");
			//List<Object> arrayObj = (List<Object>)value;
			//this.manageJsonArrayObject(key, arrayObj, columnHeader.childrenCols, cell.listChild, headerIdx + 1);
		} else {
			cell.cellObject = value;
		}
		
		
		return logger.traceExit(depth);
	}
	
	private int manageJsonArrayObject(String key, JSONArray jsonArray, Map<String, ColumnHeader> mapCols, List<CellHierarchy> listCells, int headerIdx) {
		logger.traceEntry();
		int depth = 0;
		int idx = 0;
		Iterator<Object> jsonArrayItr = jsonArray.iterator();
		while(jsonArrayItr.hasNext()) {
		
			Object curObject = jsonArrayItr.next();
		
			if(curObject instanceof JSONObject) {
				
				JSONObject jsonObj = (JSONObject) curObject;

				Iterator<String> jsonObjKeyIter = jsonObj.keys();
				while(jsonObjKeyIter.hasNext()) {
					String jsonKey = jsonObjKeyIter.next();
					this.manageJsonEntryObject(jsonKey, jsonObj.get(jsonKey), mapCols, listCells, headerIdx, idx);
				}
				
			} else if(curObject instanceof Map) {
				logger.warn("instanceof Map");
				Map<String, Object> objects = (Map<String, Object>) curObject;
				for(Entry<String, Object> curEntry : objects.entrySet()) {
					this.manageJsonEntryObject(curEntry.getKey(), curEntry.getValue(), mapCols, listCells, headerIdx, idx);
				}
			} else if(! (curObject instanceof List) && ! (curObject instanceof JSONArray) && !this.skipArrayCheckBox.isSelected()) {
								
				CellHierarchy cell = new CellHierarchy();
				cell.row = idx;
				cell.column = 0;
				cell.cellObject = curObject;
				listCells.add(cell);
			}
			idx++;
		}
		
		return logger.traceExit(depth);
	}
}
