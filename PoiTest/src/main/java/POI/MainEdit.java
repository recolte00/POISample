package POI;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import Dto.DataDto;
import Dto.WorkbookWrapper;

public class MainEdit {
	public static void main(String[] args) {
		String sheetName = "シートA";
		try {
			// 使用例１：物理ファイルをtemplateにする場合
			File template = new File("calendarSample.xlsx");
			WorkbookWrapper wr = new WorkbookWrapper(template);
			
			// 使用例２：メモリ上のWorkbookをtemplateにする場合
//			XSSFWorkbook wb = new XSSFWorkbook();
//			wb.createSheet(sheetName);
//			WorkbookWrapper wr = new WorkbookWrapper(wb);

			// Write xmls
			List<DataDto> dataDtos = prepareData();
			System.out.println("start");
			long startTime = (new Date()).getTime();
			wr.writeSheet(sheetName, dataDtos);
			long endTime = (new Date()).getTime();
			System.out.println("end:" + ((long)(endTime - startTime)/1000) + "s");

			// Generate zip
			FileOutputStream output = new FileOutputStream(new File(generateFileName() + ".xlsx"));
			wr.write(output);
			output.close();
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
	/**
	 * テストデータ作成
	 * @return
	 */
	private static List<DataDto> prepareData() {
		String startDate = "2021/02"//+"/01 00:00:00"
		;
		Calendar cal = Calendar.getInstance();
		Date date1 = toDate(startDate,"yyyy/MM");
		cal.setTime(date1);
		cal.add(Calendar.MONTH, 1);
		Date date2 = cal.getTime();
		List<DataDto> dataDtos = new ArrayList<DataDto>();
		for (int i=0; i<50000; i++) {
			dataDtos.add(new DataDto("こんにちは", (double)i/100, new BigDecimal("-0.5"), String.valueOf(date1), String.valueOf(date2), String.format("%1$02d", i%47+1)));
			dataDtos.add(new DataDto("", (double)i/100, new BigDecimal("0.5"), null, String.valueOf(date1), String.format("%1$02d", i%47+1)));
		}
		return dataDtos;
	}
	
	public static Date toDate(String str, String format){
		SimpleDateFormat df = new SimpleDateFormat(format);
		df.setLenient(false);
		Date date= new Date();
		try {
			date = df.parse(str);
		}catch(Exception e) {
			e.printStackTrace();
		}
		return date;
	}
	/**
	 * テスト用ファイル名生成
	 * @return
	 */
	private static String generateFileName() {
		SimpleDateFormat sdf1 = new SimpleDateFormat("yyyyMMdd_HHmmss");
		return sdf1.format(new Date());
	}
}
