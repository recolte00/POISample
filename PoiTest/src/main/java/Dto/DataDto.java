package Dto;

import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class DataDto implements XlsxWritable {
	public DataDto(String str, double d, BigDecimal bigDecimal, String date, String cal, String prefectureCode) {
		super();
		this.str = str;
		this.str = str;
		this.d = d;
		this.bigDecimal = bigDecimal;
		this.date = date;
		this.cal = cal;
		this.prefectureCode = prefectureCode;
	}
	private String str;
	private double d;
	private BigDecimal bigDecimal;
	private String date;
	private String cal;
	private String prefectureCode;

	public Map<Integer, Object> getMap() {
		Map<Integer, Object> map = new HashMap<Integer, Object>();
		map.put(0, str);
		map.put(1, d);
		map.put(2, bigDecimal);
		map.put(4, date);
		map.put(5, cal);
		map.put(6, prefectureCode);
		return map;
	}
}