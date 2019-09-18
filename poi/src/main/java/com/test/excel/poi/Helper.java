package com.test.excel.poi;

import java.text.NumberFormat;

public class Helper {

	public static String doubleToString(double numericCellValue) {
		Double dou_obj = new Double(numericCellValue);
		NumberFormat nf = NumberFormat.getInstance();
		nf.setGroupingUsed(false);
		return nf.format(dou_obj);
	}

}
