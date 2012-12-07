package se.technipelago

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFCell

import se.technipelago.ExcelBuilder

class ExcelBuilderX extends ExcelBuilder{
	static Class getRowClass(){
		XSSFRow;
	};
	static Class getCellClass(){
		XSSFCell;
	};

	static Class getWorkbookClass(){
		XSSFWorkbook;
	};

	ExcelBuilderX(String fileName) {
		super(fileName)
	};
}