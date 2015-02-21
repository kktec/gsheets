package org.gsheets.building

import static org.hamcrest.Matchers.*
import static spock.util.matcher.HamcrestSupport.*

import java.text.SimpleDateFormat

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

import spock.lang.Specification

abstract class WorkbookBuilderSpec extends Specification {
	
	WorkbookBuilder builder
	
	def headers = ['a', 'b', 'c']
	
	abstract protected newBuilder()
	
	def setup() {
		builder = newBuilder()
	}

	def cleanup() {
		assert builder.wb.class == builder.support.workbookType()
	}
	
	def 'can build an empty workbook'() {
		when:
		builder.workbook {
		}

		then:
		fetchWb().numberOfSheets == 0
	}
	
	def 'can build a workbook with multiple sheets'() {
		when:
		builder.workbook {
			sheet('sheet1') {
			}
			sheet('sheet2') {
			}
		}
		
		then:
		fetchWb().getSheetAt(0).sheetName == 'sheet1'
		fetchWb().getSheetAt(1).sheetName == 'sheet2'
	}
	
	def 'can build a single sheet directly into the workbook'() {
		when:
		builder.sheet('sheet1') {
		}
		
		then:
		fetchWb().getSheetAt(0).sheetName == 'sheet1'
	}
	
	def 'can NOT build a row outside a sheet'() {
		when:
		builder.workbook {
				row()
		}
		
		then:
		Exception x = thrown()
		x.message == 'can NOT build a row outside a sheet'
	}
	
	def 'can NOT build a cell outside a row'() {
		when:
		builder.workbook {
			sheet('a') {
				cell(true)
			}
		}
		
		then:
		IllegalStateException x = thrown()
		x.message == 'can NOT build a cell outside a row'
	}
	
	def 'can build 1 empty row'() {
		when:
		builder.workbook {
			sheet('a') {
				row()
			}
		}
		
		then:
		fetchCell0(0) == null
		fetchSheet().lastRowNum == 0
	}
	
	def 'can build multiple rows'() {
		when:
		builder.workbook {
			sheet('b') {
				row()
				row()
			}
		}
		
		then:
		fetchCell0(0) == null
		fetchCell0(1) == null
		fetchSheet().lastRowNum == 1
	}
	
	def 'can build a String cell'() {
		when:
		builder.workbook {
			sheet('c') {
				row('s')
			}
		}
		
		then:
		fetchCell0(0).stringCellValue == 's'
		fetchSheet().getColumnWidth(0) == 2048
	}
	
	def 'can build a Boolean cell'() {
		when:
		builder.workbook {
			sheet('d') {
				row(true)
				row(false)
			}
		}
		
		then:
		fetchCell0(0).booleanCellValue
		!fetchCell0(1).booleanCellValue
	}
	
	def 'can build a Double cell'() {
		when:
		builder.workbook {
			sheet('e') {
				row(12.345D)
				row(13)
				row(14.18)
				row(23.45F)
				row(123456789L)
				row(123 as Short)
			}
		}
		
		then:
		fetchCell0(0).numericCellValue == 12.345D
		fetchCell0(1).numericCellValue == 13D
		fetchCell0(2).numericCellValue == 14.18D
		that fetchCell0(3).numericCellValue, closeTo(23.45D, 0.000001D)
		fetchCell0(4).numericCellValue == 123456789
		fetchCell0(5).numericCellValue == 123
	}
	
	def 'can build a Date cell'() {
		given:
		Date date = new Date()
		
		when:
		builder.workbook {
			sheet('f') {
				row(date)
			}
		}
		
		then:
		fetchCell0(0).dateCellValue == date
		fetchCell0(0).cellStyle.dataFormatString == 'yyyy-mm-dd hh:mm'
	}
	
	def 'can build a Date cell with a particular format string'() {
		given:
		SimpleDateFormat sdf = new SimpleDateFormat('yyyy-MM-dd', Locale.default)
		Date date = sdf.parse('1951-08-06')
		
		when:
		builder.workbook {
			sheet('f') {
				row('x', date)
			}
		}
		
		then:
		fetchCell(0, 1).dateCellValue == date
		fetchCell(0, 1).cellStyle.dataFormatString == 'yyyy-mm-dd hh:mm'
	}
	
	def 'can build an object cell as a String'() {
		given:
		SomeClass someObject = new SomeClass()
		
		when:
		builder.workbook {
			sheet('g') {
				row(someObject)
			}
		}
		
		then:
		fetchCell0(0).stringCellValue == someObject.toString()

	}
	
	def 'can build a Formula cell'() {
		given:
		Formula formula = new Formula('A1*B1')
		
		when:
		builder.workbook {
			sheet('h') {
				row(13, 3, formula)
			}
		}
		
		then:
		fetchCell(0, 0).numericCellValue == 13
		fetchCell(0, 1).numericCellValue == 3
		fetchCell(0, 2).stringCellValue == 'A1*B1'
		fetchCell(0, 2).cellType == Cell.CELL_TYPE_FORMULA
	}
	
	def 'can build multiple cells in a row'() {
		when:
		builder.workbook {
			sheet('i') {
				row(13, 'x', true)
			}
		}
		
		then:
		fetchCell(0, 0).numericCellValue == 13
		fetchCell(0, 1).stringCellValue == 'x'
		fetchCell(0, 2).booleanCellValue
	}
	
	def 'cannot build a row outside a sheet'() {
		when:
		builder.workbook {
			row()
		}
		
		then:
		IllegalStateException x = thrown()
		x.message == 'can NOT build a row outside a sheet'
	}
	
	def 'can auto size columns'() {
		when:
		builder.workbook {
			sheet('c') {
				row('s', 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa')
				autoColumnWidth(2)
			}
		}
		
		then:
		fetchCell0(0).stringCellValue == 's'
		fetchSheet().getColumnWidth(0) < 2048
		fetchCell(0, 1).stringCellValue == 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
		fetchSheet().getColumnWidth(1) > 2048
	}
	
	def 'adding an additional sheet resets the the rowIndex counter'() {
		when:
		builder.workbook {
			sheet('s1') {
				row('a', 'x')
			}
			sheet('s2') {
			}
		}

        then:
		builder.nextRowNum == 0
	}
	
	static class SomeClass {
		String toString() {
			'someObject'
		}
	}
	
	protected Cell fetchCell(rowNum, cellNum) { fetchCell(builder.currentSheet.getRow(rowNum), cellNum) }
	
	protected Cell fetchCell(Row row, cellNum) { row.getCell cellNum }

	protected Cell fetchCell0(rowNum) { fetchCell rowNum, 0 }

	protected Cell fetchCell(cellNum) { fetchCell builder.currentRow, cellNum }

	protected Workbook fetchWb() { builder.wb }
	
	protected Sheet fetchSheet() { builder.currentSheet }
	
	static demospec = {
		workbook {
			def fmt = new SimpleDateFormat('yyyy-MM-dd HHmm', Locale.default)
			sheet('sheet 1') {
				row('Name', 'Date', 'Count', 'Value', 'Active')
				row('a', fmt.parse('2012-09-12 1012'), 69, 12.34, true)
				row('b', fmt.parse('2012-09-13 2213'), 666, 43.21, false)
				autoColumnWidth(5)
			}
		}
	}
	
	static file(name, builder) {
		Workbook workbook = builder.workbook(demospec)
		
		File file = new File(name)
		if (!file.exists()) {
			file.createNewFile()
		}
		OutputStream out = new FileOutputStream(file)
		workbook.write out
		out.close()
	}
}
