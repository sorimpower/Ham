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
		
		setDefault(); //���� Ž���� ����
		chooseFile(); //���� Ž��
	}
	
	private void setDefault(){
		FileChooser.setCurrentDirectory(new File(FileChooserConfig.ROUTE_PATH)); //���� ��� ���丮�� ����
		FileChooser.setAcceptAllFileFilterUsed(true); //Filter ��� ���� ����
		FileChooser.setDialogTitle(FileChooserConfig.TITLE); //â�� ����
		FileChooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES); //���� ���� ���
		
		FileNameExtensionFilter filter = new FileNameExtensionFilter("excel file", "xls", "xlsx");
		FileChooser.setFileFilter(filter); //���� ���� �߰�
	}
	
	private void chooseFile(){
		fileOpenReturnVal = FileChooser.showOpenDialog(null); //����� â ����
		
		//���⸦ Ŭ��
		switch(fileOpenReturnVal){
			case JFileChooser.APPROVE_OPTION :
				folderPath = FileChooser.getSelectedFile().toString();
				
				if(!folderPath.contains("xlsx")){
					JOptionPane.showMessageDialog(null, ".xlsx Ȯ���ڸ� ������ �ּ���", "�޽���", JOptionPane.ERROR_MESSAGE);
					System.exit(0);
				}
				break;
			case JFileChooser.CANCEL_OPTION :
			default:
				JOptionPane.showMessageDialog(null, "������ ������ �ּ���", "�޽���", JOptionPane.INFORMATION_MESSAGE);
				System.exit(0);
				break;
		}
	}
	
	public String getFilePath(){
		return folderPath;
	}

}
