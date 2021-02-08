package POI;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POISample {
	public static void main(String[] args) throws IOException {
		Workbook workbook = null;
		OutputStream os = null;
		String outputFilePath = "testfile.xlsx";
		try {
			os = new FileOutputStream(outputFilePath);
			workbook = new XSSFWorkbook();
			
			// �V�[�g�쐬
			Sheet sheet = workbook.createSheet("�V�[�gA");
			
			Row row = sheet.createRow(0);
			Cell cell = row.createCell(0);
			// 1�s�ڂ�A��ɒl��ݒ�
			cell.setCellValue("A-1");
			cell = row.createCell(1);
			// 1�s�ڂ�B��ɒl��ݒ�
			cell.setCellValue("B-1");
			
			row = sheet.createRow(1);
			cell = row.createCell(0);
			// 2�s�ڂ�A��ɒl��ݒ�
			cell.setCellValue("A-2");
			
			cell = row.createCell(1);
			// �G�N�Z���t�@�C�����o��
			workbook.write(os);
			
		} finally {
			if(os != null) {
				os.close();
			}
			if(workbook != null) {
				workbook.close();
			}
		}
	}
}
