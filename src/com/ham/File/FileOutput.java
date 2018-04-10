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
			//sheet���� �ݺ�
			for(String key : sheetMap.keySet()){
				String fileName = key;
				HashMap<String, ArrayList> marginMap = sheetMap.get(key); 

				//margin���� �ݺ�
				for(String margin : marginMap.keySet()){
					int rowIndex = 0;
					workBook = new HSSFWorkbook(); //workbook����
					sheet = workBook.createSheet(key); //sheet����
					
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
					
					//����
					style.setAlignment(HorizontalAlignment.CENTER);
					style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
					
					//�׵θ�
					style.setBorderTop(BorderStyle.THIN);
					style.setBorderBottom(BorderStyle.THIN);
					style.setBorderRight(BorderStyle.THIN);
					style.setBorderLeft(BorderStyle.THIN);
					
					//��Ʈ
					font.setFontHeightInPoints((short)11); //�۾�ũ�� 11
					font.setFontName("���� ���");
					
					if(rowIndex==0){
						font.setBold(true);
						style.setWrapText(true); //����
						style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
						style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
						style.setFont(font);
						
						//column���� �ݺ�
						for(int columnIndex=0; columnIndex < 6; columnIndex++){
							//������ row�� column����
							cell = row.createCell(columnIndex);
							cell.setCellStyle(style);
							
							switch(columnIndex){
								case 0 :
									cell.setCellValue("no");
									break;
								case 1 :
									cell.setCellValue("��");
									break;
								case 2 :
									cell.setCellValue("��ǰ�ڵ�");
									break;
								case 3 :
									cell.setCellValue("���Ű�\n(����:��)");
									break;
								case 4 :
									cell.setCellValue("������\n(�����)");
									break;
								case 5 :
									cell.setCellValue("������\n(�����)");
									break;
							}
						}
						rowIndex++;
					}
					font = workBook.createFont();
					style = workBook.createCellStyle();
					
					//����
					style.setAlignment(HorizontalAlignment.CENTER);
					style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
					
					//�׵θ�
					style.setBorderTop(BorderStyle.THIN);
					style.setBorderBottom(BorderStyle.THIN);
					style.setBorderRight(BorderStyle.THIN);
					style.setBorderLeft(BorderStyle.THIN);
					
					//��Ʈ
					font.setFontHeightInPoints((short)11); //�۾�ũ�� 11
					font.setFontName("���� ���");
					
					//etc���� �ݺ�
					for(ArrayList<String> etcData : etcList){
						font.setBold(false);
						style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
						style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
						style.setFont(font);
						
						row = sheet.createRow((short)rowIndex);
						//column���� �ݺ�
						for(int columnIndex=0; columnIndex <= etcData.size(); columnIndex++){
							//������ row�� column����
							cell = row.createCell(columnIndex);
							cell.setCellStyle(style);

							//����Ʈ�� ��� �����͸� ������ cell�� add��
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
					//���� ����
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
