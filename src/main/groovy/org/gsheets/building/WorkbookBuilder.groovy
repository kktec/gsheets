package org.gsheets.building

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.DataFormat
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.gsheets.NonXmlWorkbookSupport
import org.gsheets.WorkbookSupport
import org.gsheets.XmlWorkbookSupport

/**
 * Provides basic support for building xml or non-xml spreadsheets.
 * 
 * @author Ken Krebs 
 */
class WorkbookBuilder {
	
	private final WorkbookSupport support

	Workbook wb
	Sheet currentSheet
	int nextRowNum
	Row currentRow
	String defaultWorkbookDateFormat = 'yyyy-mm-dd hh:mm'

	WorkbookBuilder(boolean xml) {
		support = xml ? new XmlWorkbookSupport() : new NonXmlWorkbookSupport()
		wb = support.workbookType().newInstance()
	}
	
	static {
		WorkbookBuilder.metaClass.methodMissing = { String name, args ->
			if (name == 'cell' && !currentRow) { throw new IllegalStateException('can NOT build a cell outside a row') }
			else { throw new MissingMethodException(name, WorkbookBuilder, args) }
		}
	}
	
	/**
	 * Provides the root of a Workbook DSL.
	 *
	 * @param closure to support nested method calls
	 * 
	 * @return the created Workbook
	 */
	Workbook workbook(Closure closure) {
		assert closure

		closure.delegate = this
		closure.call()
		wb
	}

	/**
	 * Builds a new Sheet.
	 *
	 * @param closure to support nested method calls
	 * 
	 * @return the created Sheet
	 */
	Sheet sheet(String name, Closure closure) {
		assert name
		assert closure

		currentSheet = wb.createSheet(name)
		nextRowNum = 0
		closure.delegate = currentSheet
		closure.call()
		currentSheet
	}

	Row row(... values) { row(values as List) }
	
	Row row(Iterable values) {
		if (!currentSheet) { throw new IllegalStateException('can NOT build a row outside a sheet') }
		
		currentRow = currentSheet.createRow(nextRowNum++)		
		if (values) {
			values.eachWithIndex { value, column -> cell value, column }
		}
		currentRow
	}
	
	void autoColumnWidth(int columns) {
		for (i in 0..<columns) {
			currentSheet.autoSizeColumn(i)
		}
	}
	
	Cell cell(String value, int column) { createCell value, column, Cell.CELL_TYPE_STRING }
	
	Cell cell(Boolean value, int column) { createCell value, column, Cell.CELL_TYPE_BOOLEAN }
	
	Cell cell(Number value, int column) { createCell value, column, Cell.CELL_TYPE_NUMERIC }
	
	Cell cell(Date date, int column) { 
		Cell cell = createCell date, column, Cell.CELL_TYPE_NUMERIC
		CellStyle cellStyle = wb.createCellStyle()
		cellStyle.dataFormat = wb.creationHelper.createDataFormat().getFormat(defaultWorkbookDateFormat)
		cell.setCellStyle cellStyle
	}
	
	Cell cell(Formula formula, int column) { createCell formula.text, column, Cell.CELL_TYPE_FORMULA }
	
	Cell cell(value, int column) { createCell value.toString(), column, Cell.CELL_TYPE_STRING }
	
	private Cell createCell(value, int column, int cellType) {
		Cell cell = currentRow.createCell(column)
		cell.cellType = cellType
		cell.setCellValue(value)
		cell
	}
}

