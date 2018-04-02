package com.ham.File;

import java.awt.Color;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileOutput {
	XSSFWorkbook workBook;
	XSSFSheet sheet;
	XSSFRow row;
	XSSFCell cell;
	
	String filePath;
	
	FileOutputStream fileOutputStream;
	
	public FileOutput(HashMap<String, HashMap> sheetMap, String filePath) {
		workBook = new XSSFWorkbook(); //workbook����
		this.filePath = filePath;
		writeFile(sheetMap);
	}
	
	private void writeFile(HashMap<String, HashMap> sheetMap){
		if(sheetMap != null && sheetMap.size() > 0){
			sheet = workBook.createSheet("��Ʈ��"); //sheet����
			
			//sheet���� �ݺ�
			for(String key : sheetMap.keySet()){
				String fileName = key;
				HashMap<String, ArrayList> marginMap = sheetMap.get(key); 

				//margin���� �ݺ�
				for(String margin : marginMap.keySet()){
					int rowIndex = 0;
					String currentMargin = margin;
					ArrayList<ArrayList<String>> etcList = marginMap.get(margin);
					
					
						//etc���� �ݺ�
						for(ArrayList<String> etcData : etcList){
							row = sheet.createRow((short)rowIndex);
							
//							XSSFCellStyle style = workBook.createCellStyle();
//							//style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//							row.setRowStyle(style);
							
							cell = row.createCell(0);
							cell.setCellValue(String.valueOf(rowIndex));

							//column���� �ݺ�
							for(int columnIndex=0; columnIndex < etcData.size(); columnIndex++){
//								if(rowIndex == 0){
//									style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
//									style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//									row.setRowStyle(style);
//								}
								//������ row�� column����
								cell = row.createCell(columnIndex+1);
								//����Ʈ�� ��� �����͸� ������ cell�� add��
								cell.setCellValue(etcData.get(columnIndex));
							}
							rowIndex++;
						}
					
																
					//���� ����
					try {
						fileOutputStream = new FileOutputStream(getFilePath() + fileName + currentMargin + ".xlsx");
					} catch (FileNotFoundException e) {}
					try {
						workBook.write(fileOutputStream);
					} catch (IOException e) {}
				}
			}
			try {
				fileOutputStream.close();
			} catch (IOException e) {}
		}
	}
	
	private String getFilePath(){
		String fileFolderPath = this.filePath.substring(0, this.filePath.lastIndexOf("\\")+1); 
		return fileFolderPath;
	}
}
