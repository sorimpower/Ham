package com.ham.File;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

public class FileOutput {
	HSSFWorkbook workBook;
	HSSFSheet sheet;
	HSSFRow row;
	HSSFCell cell;
	HSSFCellStyle style;
	HSSFFont font;
	
	String filePath;
	
	FileOutputStream fileOutputStream;
	
	public FileOutput(HashMap<String, HashMap> sheetMap, String filePath) {
		this.filePath = filePath;
		writeFile(sheetMap);
	}
	
	private void writeFile(HashMap<String, HashMap> sheetMap){
		if(sheetMap != null && sheetMap.size() > 0){
			//sheet단위 반복
			for(String key : sheetMap.keySet()){
				String fileName = key;
				HashMap<String, ArrayList> marginMap = sheetMap.get(key); 

				//margin단위 반복
				for(String margin : marginMap.keySet()){
					int rowIndex = 0;
					workBook = new HSSFWorkbook(); //workbook생성
					sheet = workBook.createSheet(key); //sheet생성
					
					sheet.setColumnWidth((short)0, (short)2666);
					sheet.setColumnWidth((short)1, (short)3481);
					sheet.setColumnWidth((short)2, (short)5296);
					sheet.setColumnWidth((short)3, (short)7555);
					sheet.setColumnWidth((short)4, (short)5814);
					sheet.setColumnWidth((short)5, (short)5407);
					
					font = workBook.createFont();
					
					String currentMargin = margin;
					ArrayList<ArrayList<String>> etcList = marginMap.get(margin);
					
					style = workBook.createCellStyle();
					row = sheet.createRow((short)rowIndex);
					
					//정렬
					style.setAlignment(HorizontalAlignment.CENTER);
					style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
					
					//테두리
					style.setBorderTop(BorderStyle.THIN);
					style.setBorderBottom(BorderStyle.THIN);
					style.setBorderRight(BorderStyle.THIN);
					style.setBorderLeft(BorderStyle.THIN);
					
					//폰트
					font.setFontHeightInPoints((short)11); //글씨크기 11
					font.setFontName("맑은 고딕");
					
					if(rowIndex==0){
						font.setBold(true);
						style.setWrapText(true); //개행
						style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
						style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
						style.setFont(font);
						
						//column단위 반복
						for(int columnIndex=0; columnIndex < 6; columnIndex++){
							//생성된 row에 column생성
							cell = row.createCell(columnIndex);
							cell.setCellStyle(style);
							
							switch(columnIndex){
								case 0 :
									cell.setCellValue("no");
									break;
								case 1 :
									cell.setCellValue("점");
									break;
								case 2 :
									cell.setCellValue("단품코드");
									break;
								case 3 :
									cell.setCellValue("행사매가\n(단위:원)");
									break;
								case 4 :
									cell.setCellValue("시작일\n(년월일)");
									break;
								case 5 :
									cell.setCellValue("종료일\n(년월일)");
									break;
							}
						}
						rowIndex++;
					}
					font = workBook.createFont();
					style = workBook.createCellStyle();
					
					//정렬
					style.setAlignment(HorizontalAlignment.CENTER);
					style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
					
					//테두리
					style.setBorderTop(BorderStyle.THIN);
					style.setBorderBottom(BorderStyle.THIN);
					style.setBorderRight(BorderStyle.THIN);
					style.setBorderLeft(BorderStyle.THIN);
					
					//폰트
					font.setFontHeightInPoints((short)11); //글씨크기 11
					font.setFontName("맑은 고딕");
					
					//etc단위 반복
					for(ArrayList<String> etcData : etcList){
						font.setBold(false);
						style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
						style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
						style.setFont(font);
						
						row = sheet.createRow((short)rowIndex);
						//column단위 반복
						for(int columnIndex=0; columnIndex <= etcData.size(); columnIndex++){
							//생성된 row에 column생성
							cell = row.createCell(columnIndex);
							cell.setCellStyle(style);

							//리스트에 담긴 데이터를 가져와 cell에 add함
							if(columnIndex == 0){
								cell.setCellValue(String.valueOf(rowIndex));
							}else if(columnIndex == 3){
			                    cell.setCellValue(Integer.parseInt(etcData.get(columnIndex-1)));
							}else {
								cell.setCellValue(etcData.get(columnIndex-1));
							}
						}
						rowIndex++;
					}
					//파일 쓰기
					try {
						fileOutputStream = new FileOutputStream(getFilePath() + fileName + currentMargin + ".xls");
						workBook.write(fileOutputStream);
						fileOutputStream.close();
					} catch (Exception e) {}
				}
			}
		}
	}
	
	private String getFilePath(){
		String fileFolderPath = this.filePath.substring(0, this.filePath.lastIndexOf("\\")+1); 
		return fileFolderPath;
	}

}
