package com.kjubo.excel;

import lombok.Data;

import java.lang.reflect.Field;

@Data
public class ExcelColumnInfo {

	private static final String DATE_FORMAT = "yyyy/M/d";

	private int col = 0;
	private ICodeable excelColumnCodeable;
	private String name;
	private String colName;
	private Field field;
	private String fieldName;
	private String[] dateFormat;


	public ExcelColumnInfo(ExcelColumn excelColumn) {
		this.col = excelColumn.col();
		this.name = excelColumn.name();
		this.colName = excelColumn.name();
	}

	public String getDefaultDateFormat() {
		if (this.getDateFormat() != null && this.getDateFormat().length > 0) {
			return DATE_FORMAT;
		} else {
			return this.getDateFormat()[0];
		}
	}
}
