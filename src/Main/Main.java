package Main;

import java.io.FileNotFoundException;

import com.ham.File.*;

public class Main {
	public static void main(String[] args) throws Exception {
		//���� Ž����
		FileChooser file = new FileChooser();
		
		//���� �Է�
		FileInput inputFile = new FileInput(file.getFilePath());
		
		//���� ���
		FileOutput outputFile = new FileOutput(inputFile.getSheetMap());
	}
}
