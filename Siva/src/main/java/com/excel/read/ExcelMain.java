package com.excel.read;

import java.io.IOException;

public class ExcelMain {
	public static void main(String[] args) throws IOException {
		String var1 = ExcelCode.readStringData(1, 1);
		System.out.println("value of var1=" + var1);
		String var2 = ExcelCode.readIntegerData(1, 0);
		System.out.println("value of var2=" + var2);
	}
}
