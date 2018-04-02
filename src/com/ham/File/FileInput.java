package com.ham.File;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;

import javax.swing.JOptionPane;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ham.Config.FileInputConfig;

public class FileInput{
	private HashMap<String, HashMap> sheetMap = new HashMap<String, HashMap>();
	private HashMap<String, ArrayList> marginMap;
	private ArrayList<String> barcodeList;
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
			barcodeList = new ArrayList<String>();
			
			XSSFSheet sheet = workBook.getSheetAt(sheetIndex);
			String sheetName = sheet.getSheetName();
			rows = sheet.getPhysicalNumberOfRows();
			
			switch(sheetName){
				case "광주점" :
					jumpoCode = FileInputConfig.GWANGJU;
					break;
				case "마산점" :
					jumpoCode = FileInputConfig.MASAN;
					break;
				case "대구점" :
					jumpoCode = FileInputConfig.DAEGU;
					break;
				case "경기점" :
					jumpoCode = FileInputConfig.GYEONGGI;
					break;
				case "명동" :
					jumpoCode = FileInputConfig.MYEONGDONG;
					break;
				default :
					continue;
			}
			
			//바코드 중복 체크를 위한 배열
			for(int i= 0; i<rows;i++){
				XSSFRow row = sheet.getRow(i);
				XSSFCell cell = row.getCell(4);
				barcodeList.add(cell.getStringCellValue());
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
						if(value != null){
							switch(columnIndex){
								case 0 : //시작일
									startDate = value;
									break;
								case 1 : //종료일
									endDate = value;
									break;
								case 4 : //단품코드(바코드)
									//단품코드 13자리인지 확인
									if(value.length() != 13){
										System.out.println(value);
										JOptionPane.showMessageDialog(null, "시트 : "+ sheetName + "에서 " + (rowIndex+1) + "행의 바코드를 확인해주세요!", "메시지", JOptionPane.ERROR_MESSAGE);
										System.exit(0);
									}
									//단품코드가 중복되는지 확인
									if(barcodeList.contains(value) == true){
										JOptionPane.showMessageDialog(null, "시트 : "+ sheetName + "에서 " + value + "바코드가 중복됩니다. 확인해주세요!", "메시지", JOptionPane.ERROR_MESSAGE);
										System.exit(0);
									}
									barCode = value;
									break;
								case 11 : //행사매가(행사가)
									price = value;
									break;
								case 12 : //행사마진
									Double tmpValue = Double.parseDouble(value);  
									int tmpIntValue = 0;
									if(1 <= tmpValue && tmpValue <= 15){
										if(sheetName == "대구"){
											tmpIntValue = 11;
										}else{
											tmpIntValue = 10;
										}
									}else if(15 < tmpValue && tmpValue <= 17.5){
										if(sheetName == "경기" || sheetName == "대구"){
											tmpIntValue = 16;
										}else{
											tmpIntValue = 15;
										}
									}else if(17.5 < tmpValue && tmpValue <= 19.5){
										tmpIntValue = 18;
									}else if(19.5 < tmpValue && tmpValue <= 22){
										if(sheetName == "마산"){
											tmpIntValue = 20;
										}else{
											tmpIntValue = 21;
										}
									}else if(22 < tmpValue){
										if(sheetName == "마산"){
											tmpIntValue = 20;
										}else if(sheetName == "광주"){
											tmpIntValue = 21;
										}else{
											tmpIntValue = 23;
										}
									}
									currentMargin = String.valueOf(tmpIntValue);
									break;
							}
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
