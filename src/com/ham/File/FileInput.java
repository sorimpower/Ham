package com.ham.File;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ham.Config.FileInputConfig;

public class FileInput{
	private HashMap<String, HashMap> sheetMap = new HashMap<String, HashMap>();
	private HashMap<String, ArrayList> marginMap;
	private ArrayList<String> etcList;
	private ArrayList<ArrayList<String>> tmpList;
	
	private int sheets;
	private int rows;
	private int cells;
	
	private int jumpoCode;
	
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
			marginMap = new HashMap<String, ArrayList>();
			
			XSSFSheet sheet = workBook.getSheetAt(sheetIndex);
			String sheetName = sheet.getSheetName();
			rows = sheet.getPhysicalNumberOfRows();
			
			
			switch(sheetName){
				case "광주" :
					jumpoCode = FileInputConfig.GWANGJU;
					break;
				case "마산" :
					jumpoCode = FileInputConfig.MASAN;
					break;
				case "대구" :
					jumpoCode = FileInputConfig.DAEGU;
					break;
				case "경기" :
					jumpoCode = FileInputConfig.GYEONGGI;
					break;
				case "명동" :
					jumpoCode = FileInputConfig.MYEONGDONG;
					break;
				
			}
			
					
			//행의 수만큼 반복
			for(int rowIndex = 0; rowIndex<rows; rowIndex++){
				etcList = new ArrayList<String>();
				XSSFRow row = sheet.getRow(rowIndex);
				
				String currentMargin = null;
				String startDate = null;
				String endDate = null;
				String barCode = null;
				String price = null;				
				
				
				
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
								startDate = value;
								break;
							case 1 : //종료일
								endDate = value;
								break;
							case 4 : //단품코드(바코드)
								barCode = value;
								break;
							case 11 : //행사매가(행사가)
								price = value;
								break;
							case 12 : //행사마진
								currentMargin = value;
								break;
						}
					}
				}
				etcList.add(String.valueOf(jumpoCode));
				etcList.add(barCode);
				etcList.add(price);
				etcList.add(startDate);
				etcList.add(endDate);				
				
				if(currentMargin == null) continue;
				
				if(marginMap.containsKey(currentMargin)){
					tmpList = marginMap.get(currentMargin);
					tmpList.add(etcList);
				}else{
					tmpList = new ArrayList<ArrayList<String>>();
					tmpList.add(etcList);
					marginMap.put(currentMargin, tmpList);
				}
			}
			sheetMap.put(sheetName, marginMap);
		}
	}

	public HashMap<String, HashMap> getSheetMap() {
		return sheetMap;
	}

}
