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
		workBook = new XSSFWorkbook(inputStream); //workbook����
		readFile();
	}
	
	private void readFile(){
		sheets = workBook.getNumberOfSheets(); //��Ʈ ��
		
		//��Ʈ ����ŭ �ݺ�
		for(int sheetIndex= 0; sheetIndex < sheets; sheetIndex++){
			marginMap = new HashMap<String, ArrayList>();
			barcodeList = new ArrayList<String>();
			
			XSSFSheet sheet = workBook.getSheetAt(sheetIndex);
			String sheetName = sheet.getSheetName();
			rows = sheet.getPhysicalNumberOfRows();
			
			switch(sheetName){
				case "������" :
					jumpoCode = FileInputConfig.GWANGJU;
					break;
				case "������" :
					jumpoCode = FileInputConfig.MASAN;
					break;
				case "�뱸��" :
					jumpoCode = FileInputConfig.DAEGU;
					break;
				case "�����" :
					jumpoCode = FileInputConfig.GYEONGGI;
					break;
				case "��" :
					jumpoCode = FileInputConfig.MYEONGDONG;
					break;
				default :
					continue;
			}
			
			//���ڵ� �ߺ� üũ�� ���� �迭
			for(int i= 0; i<rows;i++){
				XSSFRow row = sheet.getRow(i);
				XSSFCell cell = row.getCell(4);
				barcodeList.add(cell.getStringCellValue());
			}
					
			//���� ����ŭ �ݺ�
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
				
				//���� ����ŭ �ݺ�
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
					
					//������
					if(rowIndex >= 8){		
						if(value != null){
							switch(columnIndex){
								case 0 : //������
									startDate = value;
									break;
								case 1 : //������
									endDate = value;
									break;
								case 4 : //��ǰ�ڵ�(���ڵ�)
									//��ǰ�ڵ� 13�ڸ����� Ȯ��
									if(value.length() != 13){
										System.out.println(value);
										JOptionPane.showMessageDialog(null, "��Ʈ : "+ sheetName + "���� " + (rowIndex+1) + "���� ���ڵ带 Ȯ�����ּ���!", "�޽���", JOptionPane.ERROR_MESSAGE);
										System.exit(0);
									}
									//��ǰ�ڵ尡 �ߺ��Ǵ��� Ȯ��
									if(barcodeList.contains(value) == true){
										JOptionPane.showMessageDialog(null, "��Ʈ : "+ sheetName + "���� " + value + "���ڵ尡 �ߺ��˴ϴ�. Ȯ�����ּ���!", "�޽���", JOptionPane.ERROR_MESSAGE);
										System.exit(0);
									}
									barCode = value;
									break;
								case 11 : //���Ű�(��簡)
									price = value;
									break;
								case 12 : //��縶��
									Double tmpValue = Double.parseDouble(value);  
									int tmpIntValue = 0;
									if(1 <= tmpValue && tmpValue <= 15){
										if(sheetName == "�뱸"){
											tmpIntValue = 11;
										}else{
											tmpIntValue = 10;
										}
									}else if(15 < tmpValue && tmpValue <= 17.5){
										if(sheetName == "���" || sheetName == "�뱸"){
											tmpIntValue = 16;
										}else{
											tmpIntValue = 15;
										}
									}else if(17.5 < tmpValue && tmpValue <= 19.5){
										tmpIntValue = 18;
									}else if(19.5 < tmpValue && tmpValue <= 22){
										if(sheetName == "����"){
											tmpIntValue = 20;
										}else{
											tmpIntValue = 21;
										}
									}else if(22 < tmpValue){
										if(sheetName == "����"){
											tmpIntValue = 20;
										}else if(sheetName == "����"){
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
