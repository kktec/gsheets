package org.gsheets.building

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
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
	
	final static String DEFAULT_DATE_FORMAT = 'yyyy-mm-dd hh:mm'
	
	private final WorkbookSupport support
	
	private final boolean xml

	final Workbook wb
	final Map cellStyles = [:]
	
	Sheet currentSheet
	int nextRowNum
	Row currentRow
	CellStyle currentStyle
	CellStyle previousStyle
	String currentDateFormat = DEFAULT_DATE_FORMAT
	
	WorkbookBuilder(boolean xml) {
		this.xml = xml
		support = xml ? new XmlWorkbookSupport() : new NonXmlWorkbookSupport()
		wb = support.workbookType().newInstance()
		cellStyles['default'] = currentStyle = previousStyle = wb.getCellStyleAt(0 as short)
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
	
	CellStyle useStyle(String name) {
		assert name
		CellStyle style = cellStyles[name] 
		assert style
		previousStyle = currentStyle
		currentStyle = style
	}
	
	CellStyle restoreStyle() {
		currentStyle = previousStyle
	}
	
	CellStyle cellStyle(String name, Map fontConfig, Map stylingConfig) {
		Font font = wb.createFont()
		fontConfig.each { property, value ->
			font[property] = value
		}
		CellStyle style = wb.createCellStyle()
		stylingConfig.each { property, value ->
			style[property] = value
		}
		style.font = font
		cellStyles[name] = style
		style
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

	Row row(... values) { 
		row(values as List)
	}
	
	Row row(Iterable values) {
		if (!currentSheet) { throw new IllegalStateException('can NOT build a row outside a sheet') }
		
		currentRow = currentSheet.createRow(nextRowNum++)
		if (values) {
			values.eachWithIndex { value, column ->
				cell(value, column)
			}
		}
		currentRow
	}
	
	void autoColumnWidth(int columns) {
		for (i in 0..<columns) {
			currentSheet.autoSizeColumn(i)
		}
	}
	
	Cell cell(String value, int column) { 
		createCell value, column, Cell.CELL_TYPE_STRING
	}
	
	Cell cell(Boolean value, int column) {
		createCell value, column, Cell.CELL_TYPE_BOOLEAN
	}
	
	Cell cell(Number value, int column) {
		createCell value, column, Cell.CELL_TYPE_NUMERIC
	}
	
	Cell cell(Date date, int column) { 
		CellStyle style = wb.createCellStyle()
		style.cloneStyleFrom(currentStyle)
		style.dataFormat = wb.creationHelper.createDataFormat().getFormat(currentDateFormat)
		createCell date, column, Cell.CELL_TYPE_NUMERIC, style
	}
	
	Cell cell(Formula formula, int column) { 
		createCell formula.text, column, Cell.CELL_TYPE_FORMULA
	}
	
	Cell cell(value, int column) {
		createCell value.toString(), column, Cell.CELL_TYPE_STRING
	}
	
	private Cell createCell(value, int column, int cellType, CellStyle style = currentStyle) {
		Cell cell = currentRow.createCell(column)
		cell.cellType = cellType
		cell.setCellValue(value)
		cell.cellStyle = style
		cell
	}
}

