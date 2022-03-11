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

import com.hoffnungland.poi.corner.dbxlsreport.ExcelManager;
import com.hoffnungland.poi.corner.dbxlsreport.XlsWrkSheetException;

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
import javax.swing.JOptionPane;
import javax.swing.JFormattedTextField;
import javax.swing.JRadioButton;
import javax.swing.JCheckBox;

public class App implements ActionListener{
	private static final Logger logger = LogManager.getLogger(App.class);
	
	private JFrame frame;
	private JTextField jsonTextField;
	private JTextField targetDirTextField;
	private JTextField xlsxNameTextField;
	
	private static final String LOAD_JSON_ACTION = "Load JSON Action";
	private static final String TARGET_DIR_ACTION = "Targer DIR Action";
	private static final String CONVERT_ACTION = "Convert Action";
	private File[] selectedJsonFiles;
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
					App window = new App();
					window.frame.setVisible(true);
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
		jsonTextField.setEditable(false);
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
        
		JButton convertButton = new JButton("Convert");
		springLayout.putConstraint(SpringLayout.NORTH, convertButton, 12, SpringLayout.SOUTH, xlsxNameTextField);
		springLayout.putConstraint(SpringLayout.WEST, convertButton, 371, SpringLayout.WEST, frame.getContentPane());
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
		fc.setMultiSelectionEnabled(true);
		fc.setFileFilter(fcJsonFiler);
		fc.addChoosableFileFilter(fcJsonFiler);
		fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		
		int returnVal = fc.showOpenDialog(this.frame);
		
		if(returnVal == JFileChooser.APPROVE_OPTION) {
			this.selectedJsonFiles = fc.getSelectedFiles();
				
			File selectedFile = this.selectedJsonFiles[0];
			String jsonFilePath = selectedFile.getPath();
			String jsonFolderPath = jsonFilePath.substring(0, jsonFilePath.lastIndexOf(File.separator)+1);
			this.targetDirTextField.setText(jsonFolderPath);
			
			if(this.selectedJsonFiles.length == 1) {
				String jsonFileName = selectedFile.getName();
				this.xlsxNameTextField.setText(jsonFileName.substring(0, jsonFileName.lastIndexOf(".")) + ".xlsx");				
			} else {
				this.xlsxNameTextField.setText(Path.of(jsonFolderPath).getFileName() + ".xlsx");
			}
			
			String jsonTextFieldStr = null;
			for(File curJsonFile :  this.selectedJsonFiles) {
				if(jsonTextFieldStr == null) {
					jsonTextFieldStr = curJsonFile.getPath();
				} else {
					jsonTextFieldStr += ";" + curJsonFile.getName();
				}
			}
			this.jsonTextField.setText(jsonTextFieldStr);
			
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
			String targetPath = fc.getSelectedFile().getPath();
			if(!targetPath.endsWith(File.separator)) {
				targetPath += File.separator;
			}
			this.targetDirTextField.setText(targetPath);
		}
		
		logger.traceExit();
	}
	
	
	private void convertJsonToExcel() {
		
		logger.traceEntry();
		ExcelManager xlsMng = null;
		try {
			String excelFileName = this.xlsxNameTextField.getText();
			logger.info("Initialize the excel");
			xlsMng = new ExcelManager(excelFileName.substring(0, excelFileName.lastIndexOf(".")));
			for(File curJsonFile :  this.selectedJsonFiles) {				
				String jsonStr = Files.readString(Path.of(curJsonFile.getAbsolutePath()));
				//String sheetName = (excelFileName.length() > 31) ? excelFileName.substring(0, 31) : excelFileName;			
				xlsMng.getJsonResult(curJsonFile.getName(), null, jsonStr);
			}
			
		} catch (IOException e) {
			logger.error(e);
			e.printStackTrace();
		} catch (XlsWrkSheetException e) {
			logger.error(e);
			e.printStackTrace();
		} finally{
			String targetPath = this.targetDirTextField.getText();
			if(!targetPath.endsWith(File.separator)) {
				targetPath += File.separator;
			}
			if(xlsMng != null) {
				xlsMng.finalWrite(targetPath);
				xlsMng = null;
			}
			JOptionPane.showMessageDialog(this.frame, "JSON Tabulation completed", "Conversion completed", JOptionPane.INFORMATION_MESSAGE);
		}
		logger.traceExit();
	}
}
