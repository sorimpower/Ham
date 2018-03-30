package Main;

import java.io.FileNotFoundException;

import com.ham.File.*;

public class Main {
	public static void main(String[] args) throws Exception {
		//파일 탐색기
		FileChooser file = new FileChooser();
		
		//파일 입력
		FileInput inputFile = new FileInput(file.getFilePath());
		
		//파일 출력
		FileOutput outputFile = new FileOutput(inputFile.getSheetMap());
	}
}
