package POI;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import Dto.DataDto;
import Dto.WorkbookWrapper;

public class MainEdit {
	public static void main(String[] args) {
		String sheetName = "�V�[�gA";
		try {
			// �g�p��P�F�����t�@�C����template�ɂ���ꍇ
			File template = new File("testfile.xlsx");
			WorkbookWrapper wr = new WorkbookWrapper(template);
			
			// �g�p��Q�F���������Workbook��template�ɂ���ꍇ
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
	 * �e�X�g�f�[�^�쐬
	 * @return
	 */
	private static List<DataDto> prepareData() {
		SimpleDateFormat df = new SimpleDateFormat("yyyy/MM");
		Calendar cal = Calendar.getInstance();
		List<DataDto> dataDtos = new ArrayList<DataDto>();
		for (int i=0; i<50000; i++) {
			dataDtos.add(new DataDto("����ɂ���", (double)i/100, new BigDecimal("-0.5"), df.format(new Date()), df.format(cal.getTime()), String.format("%1$02d", i%47+1)));
			dataDtos.add(new DataDto("", (double)i/100, new BigDecimal("0.5"), null, df.format(cal.getTime()), String.format("%1$02d", i%47+1)));
		}
		return dataDtos;
	}

	/**
	 * �e�X�g�p�t�@�C��������
	 * @return
	 */
	private static String generateFileName() {
		SimpleDateFormat sdf1 = new SimpleDateFormat("yyyyMMdd_HHmmss");
		return sdf1.format(new Date());
	}
}