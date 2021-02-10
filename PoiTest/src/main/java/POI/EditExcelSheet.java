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
	// �G�N�Z���t�@�C����u���Ă���t�H���_
	static final String INPUT_DIR = "./";
	public static void main(String[] args)throws IOException {
		// �ύX����G�N�Z���t�@�C�����w��
		FileInputStream in = new FileInputStream(INPUT_DIR + "testfile.xlsx");
		Workbook wb = null;
		
		try {
			// �����̃G�N�Z���t�@�C����ҏW����ۂ́AWorkbookFactory���g�p
			wb = WorkbookFactory.create(in);
		} catch (Exception e) {
			e.printStackTrace();
		}
		Sheet sheet = wb.getSheet("�V�[�gA");
		Row row = sheet.getRow(0);
		Cell cell = row.getCell(0);
		cell.setCellValue("�o��");
		FileOutputStream out = null;
		
		try {
			// �ύX����G�N�Z���t�@�C�����w��
			out = new FileOutputStream(INPUT_DIR + "newTestfile.xlsx");
			// ��������
			wb.write(out);
		} catch(Exception e) {
			e.printStackTrace();
		} finally {
			out.close();
			wb.close();
		}
	}
}