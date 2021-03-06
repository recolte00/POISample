package POI;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class EditExcelSheet {
	// エクセルファイルを置いているフォルダ
	static final String INPUT_DIR = "./";
	public static void main(String[] args)throws IOException {
		// 変更するエクセルファイルを指定
		FileInputStream in = new FileInputStream(INPUT_DIR + "testfile.xlsx");
		Workbook wb = null;
		
		try {
			// 既存のエクセルファイルを編集する際は、WorkbookFactoryを使用
			wb = WorkbookFactory.create(in);
		} catch (Exception e) {
			e.printStackTrace();
		}
		Sheet sheet = wb.getSheet("シートA");
		Row row = sheet.getRow(0);
		Cell cell = row.getCell(0);
		cell.setCellValue("出力");
		FileOutputStream out = null;
		
		try {
			// 変更するエクセルファイルを指定
			out = new FileOutputStream(INPUT_DIR + "newTestfile.xlsx");
			// 書き込み
			wb.write(out);
		} catch(Exception e) {
			e.printStackTrace();
		} finally {
			out.close();
			wb.close();
		}
	}
}
