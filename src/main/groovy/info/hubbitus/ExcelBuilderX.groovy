package info.hubbitus

import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFCell

import info.hubbitus.ExcelBuilder

class ExcelBuilderX extends ExcelBuilder{
	static Class getRowClass(){
		XSSFRow;
	};

	static Class getWorkbookClass(){
		XSSFWorkbook;
	};

	ExcelBuilderX(String fileName) {
		super(fileName)
	};
}