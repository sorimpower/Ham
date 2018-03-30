package com.ham.File;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileInput{
	private HashMap<String, HashMap> SheetMap = new HashMap<String, HashMap>();
	private HashMap<String, ArrayList> MarginMap;
	private ArrayList<String> EtcList;
	private ArrayList<ArrayList<String>> tmpList;
	
	private int sheets;
	private int rows;
	private int cells;
	
	XSSFWorkbook workBook;
	
	public FileInput(String filePath) throws Exception {
		FileInputStream inputStream = new FileInputStream(filePath);
		workBook = new XSSFWorkbook(inputStream); //workbook생성
		readFile();
	}
	
	private void readFile(){
		sheets = workBook.getNumberOfSheets(); //시트 수
		
		//시트 수만큼 반복
		for(int sheetIndex= 0; sheetIndex < sheets; sheetIndex++){
			MarginMap = new HashMap<String, ArrayList>();
			
			XSSFSheet sheet = workBook.getSheetAt(sheetIndex);
			String sheetName = sheet.getSheetName();
			rows = sheet.getPhysicalNumberOfRows();
					
			//행의 수만큼 반복
			for(int rowIndex = 8; rowIndex<rows; rowIndex++){
				EtcList = new ArrayList<String>();
				String currentMargin = null;
				XSSFRow row = sheet.getRow(rowIndex);
				
				if(row == null) continue;
				
				cells = row.getPhysicalNumberOfCells();
				
				//셀의 수만큼 반복
				for(int columnIndex = 0; columnIndex <= cells; columnIndex++){
					XSSFCell cell = row.getCell(columnIndex);
					String value = null;
					
					if(cell == null) continue;
					
					switch(cell.getCellType()){
						case XSSFCell.CELL_TYPE_FORMULA:
		                    value=cell.getCellFormula();
		                    break;
		                case XSSFCell.CELL_TYPE_NUMERIC:
		                	cell.setCellType(XSSFCell.CELL_TYPE_STRING);
		                    value=cell.getStringCellValue()+"";
		                    break;
		                case XSSFCell.CELL_TYPE_STRING:
		                    value=cell.getStringCellValue()+"";
		                    break;
		                
		                case XSSFCell.CELL_TYPE_ERROR:
		                    value=cell.getErrorCellValue()+"";
		                    break;
					}
					
					//시작행
					if(rowIndex >= 8){
						switch(columnIndex){
							case 0 : //시작일
							case 1 : //종료일
							case 4 : //단품코드(바코드)
							case 11 : //행사매가(행사가)
								EtcList.add(value);
								break;
							case 12 : //행사마진			
								currentMargin = value;
								break;
						}
					}
				}
				if(currentMargin == null) continue;
				
				if(MarginMap.containsKey(currentMargin)){
					tmpList = MarginMap.get(currentMargin);
					tmpList.add(EtcList);
				}else{
					tmpList = new ArrayList<ArrayList<String>>();
					tmpList.add(EtcList);
					MarginMap.put(currentMargin, tmpList);
				}
			}
			SheetMap.put(sheetName, MarginMap);
		}
	}

	public HashMap<String, HashMap> getSheetMap() {
		return SheetMap;
	}

}
