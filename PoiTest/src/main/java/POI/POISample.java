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
			
			// シート作成
			Sheet sheet = workbook.createSheet("シートA");
			
			Row row = sheet.createRow(0);
			Cell cell = row.createCell(0);
			// 1行目のA列に値を設定
			cell.setCellValue("A-1");
			cell = row.createCell(1);
			// 1行目のB列に値を設定
			cell.setCellValue("B-1");
			
			row = sheet.createRow(1);
			cell = row.createCell(0);
			// 2行目のA列に値を設定
			cell.setCellValue("A-2");
			
			cell = row.createCell(1);
			// エクセルファイルを出力
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
