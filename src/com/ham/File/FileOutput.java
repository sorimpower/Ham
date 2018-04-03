package com.ham.File;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import javax.jws.soap.SOAPBinding.Style;
import javax.swing.GroupLayout.Alignment;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddressBase.CellPosition;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xssf.usermodel.XSSFCell;

public class FileOutput {
	HSSFWorkbook workBook;
	HSSFSheet sheet;
	HSSFRow row;
	HSSFCell cell;
	
	String filePath;
	
	FileOutputStream fileOutputStream;
	
	public FileOutput(HashMap<String, HashMap> sheetMap, String filePath) {
		workBook = new HSSFWorkbook(); //workbook����
		this.filePath = filePath;
		writeFile(sheetMap);
	}
	
	private void writeFile(HashMap<String, HashMap> sheetMap){
		if(sheetMap != null && sheetMap.size() > 0){
			sheet = workBook.createSheet("��Ʈ��"); //sheet����
			
			sheet.setColumnWidth((short)0, (short)2666);
			sheet.setColumnWidth((short)1, (short)3481);
			sheet.setColumnWidth((short)2, (short)5296);
			sheet.setColumnWidth((short)3, (short)7555);
			sheet.setColumnWidth((short)4, (short)5814);
			sheet.setColumnWidth((short)5, (short)5407);
			
			
			
			//sheet���� �ݺ�
			for(String key : sheetMap.keySet()){
				
				String fileName = key;
				HashMap<String, ArrayList> marginMap = sheetMap.get(key); 

				//margin���� �ݺ�
				for(String margin : marginMap.keySet()){
					int rowIndex = 0;
					String currentMargin = margin;
					ArrayList<ArrayList<String>> etcList = marginMap.get(margin);

					HSSFCellStyle style = workBook.createCellStyle();
					style.setAlignment(HorizontalAlignment.CENTER);
					style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
					
					HSSFFont font = workBook.createFont();
					font.setFontHeightInPoints((short)11); //�۾�ũ�� 11
					
					//ù°�� ��Ÿ�� ����
					if(rowIndex==0){
						font.setBold(true);
						style.setWrapText(true); //����
						style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
						style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					}else{
						style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
						style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

//						HSSFDataFormat dataFormat = workBook.createDataFormat();
//						style.setDataFormat((short) dataFormat.getNumberOfBuiltinBuiltinFormats());
					}
					
					style.setFont(font);
					font=null;
						//etc���� �ݺ�
						for(ArrayList<String> etcData : etcList){
							
							row = sheet.createRow((short)rowIndex);

							//column���� �ݺ�
							for(int columnIndex=0; columnIndex <= etcData.size(); columnIndex++){
								//������ row�� column����
								cell = row.createCell(columnIndex);
								cell.setCellStyle(style);
								style = null;
								if(rowIndex == 0){
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
									
									
								}else{
									//����Ʈ�� ��� �����͸� ������ cell�� add��
									if(columnIndex == 0){
										cell.setCellStyle(style);
										cell.setCellValue(rowIndex);
									}else if(columnIndex == 3 || columnIndex == 5){
										//style.setDataFormat(HSSFDataFormat.getBuiltinFormat(etcData.get(columnIndex-1)));
//										cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
//					                    value=cell.getStringCellValue()+"";
					                    cell.setCellValue(etcData.get(columnIndex-1));
									}else {
										cell.setCellValue(etcData.get(columnIndex-1));
									}
								}
								
							}
							rowIndex++;
						}
																
					//���� ����
					try {
						fileOutputStream = new FileOutputStream(getFilePath() + fileName + currentMargin + ".xls");
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
