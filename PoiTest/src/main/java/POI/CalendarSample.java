package POI;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CalendarSample {
	public static void main(String[] args) {
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet();
		
		sheet.setColumnWidth(0, 4096);
		sheet.setColumnWidth(1, 4096);
		
		Row row1 = sheet.createRow(1);
		Row row2 = sheet.createRow(2);
		
		Cell cell1_0 = row1.createCell(0);
		Cell cell1_1 = row1.createCell(1);
		Cell cell2_0 = row2.createCell(0);
		Cell cell2_1 = row2.createCell(1);
		
		cell1_0.setCellValue(new Date());
	    cell1_1.setCellValue(new Date());
	    cell2_0.setCellValue(Calendar.getInstance());
	    cell2_1.setCellValue(Calendar.getInstance());
		
		CreationHelper createHelper = wb.getCreationHelper();
		CellStyle cellStyle = wb.createCellStyle();
		short style = createHelper.createDataFormat().getFormat("yyyy/mm/dd h:mm");
		cellStyle.setDataFormat(style);
		
		cell1_1.setCellStyle(cellStyle);
		cell2_1.setCellStyle(cellStyle);
		
		FileOutputStream out = null;
		try {
			out = new FileOutputStream("sampleCalendar.xlsx");
			wb.write(out);
		}catch(IOException e) {
			System.out.println(e.toString());
		}finally{
			try {
				out.close();
			}catch(IOException e){
				System.out.println(e.toString());
			}
		}
	}
}
