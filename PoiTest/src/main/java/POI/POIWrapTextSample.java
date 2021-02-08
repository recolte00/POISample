package POI;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIWrapTextSample {
	
	public static void main(String[] args) throws IOException {
		Workbook workbook = null;
		OutputStream os = null;
		String outputFilePath = "testfile.xlsx";
		try {
			os = new FileOutputStream(outputFilePath);
			workbook = new XSSFWorkbook();
			
			// シート作成
			Sheet sheet = workbook.createSheet("シートA");
			
			Row row = sheet.createRow(1);
			XSSFCell cell = (XSSFCell)row.createCell(1);
			cell.setCellValue("このセルは折返し設定にする");
			CellStyle cellStyle = workbook.createCellStyle();
			
			// 折返しの設定
			cellStyle.setWrapText(true);
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
