package com.ham.File;

import java.util.HashMap;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileOutput {
	XSSFWorkbook workBook;
	XSSFSheet sheet;
	XSSFRow row;
	XSSFCell cell;
	
	public FileOutput(HashMap<String, HashMap> getSheetMap) {
		workBook = new XSSFWorkbook(); //workbook생성
		sheet = workBook.createSheet("시트명"); //sheet생성
		writeFile();
	}
	
	private void writeFile(){
		
	}
}
