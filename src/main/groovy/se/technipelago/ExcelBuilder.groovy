package se.technipelago

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFDateUtil
import org.apache.poi.ss.usermodel.Workbook

import se.technipelago.ExcelBuilderX

/**
 * Groovy Builder that extracts data from
 * Microsoft Excel spreadsheets.
 * @author Goran Ehrsson
 * @url http://www.technipelago.se/content/technipelago/blog/44
 */
class ExcelBuilder {

	Workbook workbook
	def labels
	def row

	static Class getRowClass(){
		HSSFRow;
	};
	static Class getCellClass(){
		HSSFCell;
	};

	static Class getWorkbookClass(){
		HSSFWorkbook;
	};

	/**
	 * Use factory method to create object.
	 *
	 * @param fileName
	 */
	ExcelBuilder(String fileName) {
		rowClass.metaClass.getAt = { int idx ->
			def cell = delegate.getCell(idx)
			if(!cell) {
				return null
			}
			def value
			switch(cell.cellType) {
				case cellClass.CELL_TYPE_NUMERIC:
					if(HSSFDateUtil.isCellDateFormatted(cell)) {
						value = cell.dateCellValue
					} else {
						value = cell.numericCellValue
					}
					break
				case cellClass.CELL_TYPE_BOOLEAN:
					value = cell.booleanCellValue
					break
				default:
					value = cell.stringCellValue
					break
			}
			return value
		}

		new File(fileName).withInputStream { is ->
			workbook = workbookClass.newInstance(is);
		}
	}

	def getSheet(idx) {
		def sheet
		if(!idx) idx = 0
		if(idx instanceof Number) {
			sheet = workbook.getSheetAt(idx)
		} else if(idx ==~ /^\d+$/) {
			sheet = workbook.getSheetAt(Integer.valueOf(idx))
		} else {
			sheet = workbook.getSheet(idx)
		}
		return sheet
	}

	def cell(idx) {
		if(labels && (idx instanceof String)) {
			idx = labels.indexOf(idx.toLowerCase())
		}
		return row[idx]
	}

	def propertyMissing(String name) {
		cell(name)
	}

	def eachLine(Map params = [:], Closure closure) {
		def offset = params.offset ?: 0
		def max = params.max ?: 9999999
		def sheet = getSheet(params.sheet)
		def rowIterator = sheet.rowIterator()
		def linesRead = 0

		if(params.labels) {
			labels = rowIterator.next().collect { it.toString().toLowerCase() }
		}
		offset.times { rowIterator.next() }

		closure.setDelegate(this)

		while(rowIterator.hasNext() && linesRead++ < max) {
			row = rowIterator.next()
			closure.call(row)
		}
	}

	public static ExcelBuilder factory(String fileName){
		if ( (fileName =~ /(?i)\.xlsx$/ ) ){
			return new ExcelBuilderX(fileName);
		}
		else{
			return new ExcelBuilder(fileName);
		}
	}
}
