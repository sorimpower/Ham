package com.ham.File;

import java.io.File;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;

import com.ham.Config.FileChooserConfig;

public class FileChooser{
	private static JFileChooser FileChooser;
	private static String folderPath;
	private int fileOpenReturnVal; 
	
	public FileChooser() {
		FileChooser = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
		
		setDefault(); //파일 탐색기 설정
		chooseFile(); //파일 탐색
	}
	
	private void setDefault(){
		FileChooser.setCurrentDirectory(new File(FileChooserConfig.ROUTE_PATH)); //현재 사용 디렉토리를 지정
		FileChooser.setAcceptAllFileFilterUsed(true); //Filter 모든 파일 적용
		FileChooser.setDialogTitle(FileChooserConfig.TITLE); //창의 제목
		FileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES); //파일 선택 모드
		
		FileNameExtensionFilter filter = new FileNameExtensionFilter("excel file", "xls", "xlsx");
		FileChooser.setFileFilter(filter); //파일 필터 추가
	}
	
	private void chooseFile(){
		fileOpenReturnVal = FileChooser.showOpenDialog(null); //열기용 창 오픈
		
		//열기를 클릭
		switch(fileOpenReturnVal){
			case JFileChooser.APPROVE_OPTION :
				folderPath = FileChooser.getSelectedFile().toString();
				
				if(!folderPath.contains("xlsx")){
					JOptionPane.showMessageDialog(null, ".xlsx 확장자를 선택해 주세요", "메시지", JOptionPane.ERROR_MESSAGE);
					System.exit(0);
				}
				break;
			case JFileChooser.CANCEL_OPTION :
			default:
				JOptionPane.showMessageDialog(null, "파일을 선택해 주세요", "메시지", JOptionPane.INFORMATION_MESSAGE);
				System.exit(0);
				break;
		}
	}
	
	public String getFilePath(){
		return folderPath;
	}

}
