package POI;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellStyleTest {
	public static void main(String[] args)throws IOException{
		Workbook workbook = null;
		OutputStream os = null;
		String outputFilePath = "testfile.xlsx";
		try {
			os = new FileOutputStream(outputFilePath);
			workbook = new XSSFWorkbook();
			Sheet sheet = workbook.createSheet("�V�[�g�P");
			Row row = sheet.createRow(1);
			XSSFCell cell = (XSSFCell)row.createCell(1);
			cell.setCellValue("����");
			CellStyle cellStyle = workbook.createCellStyle();
			// �Z���̒l�̐��������̈ʒu��ݒ�
			cellStyle.setAlignment(HorizontalAlignment.CENTER);
			cell.setCellStyle(cellStyle);
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
