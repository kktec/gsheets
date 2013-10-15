package org.gsheets

import static org.junit.Assert.*

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.junit.Test


/**
 * see README for copyright/licensing info 

 * @author Ken Krebs - kktec1@gmail.com @kktec
 */
class ExcelFileTests {

	Workbook wb = workbook { }
	Sheet sheet

	@Test
	void empty_workbook() {
		assert wb instanceof Workbook
		assert wb.numberOfSheets == 0
	}

	@Test
	void workbook_with_sheets() {
		workbook {
			data {
				sheet('sheet1') { }
				sheet('sheet2') { }
			}
		}

		wb.with {
			assert getSheetAt(0).sheetName == 'sheet1'
			assert getSheetIndex('sheet2') == 1
			assert numberOfSheets == 2
		}
	}

	@Test
	void workbook_with_data_grid() {
		def headers = ['String', 'Boolean', 'Date', 'Object', 'Integer', 'Double', 'BigDecimal']
		Date date = new Date()
		def o = new Object()
		workbook {
			data {
				sheet('sheet') { 
					header(headers) 
					row(['ken', true, date, o, 13, 1 / 3, 13.267])
					row([null])
				}
			}
		}

		sheet = wb.getSheetAt(0)
		assertHeaderRow headers
		
		Row r1 = sheet.getRow(1)
		assert_stringCellValue r1, 0, 'ken'
		assert_stringCellValue r1, 1, 'true'
		assert_dateCellValue r1, 2, date		
		assert_stringCellValue r1, 3, o
		assert_numericCellValue r1, 4, 13
		assert_numericCellValue r1, 5, 0.3333333333
		assert_numericCellValue r1, 6, 13.267
		
		Row r2 = sheet.getRow(2)
		assert_stringCellValue r2, 0, ''
	}
	
	private assert_stringCellValue(row, cellNum, value) {
		Cell cell = row.getCell(cellNum)
		assert cell.cellType == Cell.CELL_TYPE_STRING
		assert cell.stringCellValue == value.toString()
	}

	private assert_dateCellValue(row, cellNum, Date value) {
		Cell cell = row.getCell(cellNum)
		assert cell.cellType == Cell.CELL_TYPE_NUMERIC
		assert cell.dateCellValue == value
	}
	
	private assert_numericCellValue(row, cellNum, value) {
		Cell cell = row.getCell(cellNum)
		assert cell.cellType == Cell.CELL_TYPE_NUMERIC
		assert cell.numericCellValue == value
	}
	
	private assertHeaderRow(List headers) {
		Row header = sheet.getRow(0)
		headers.eachWithIndex { h, i ->
			Cell cell = header.getCell(i)
			assert cell.cellType == Cell.CELL_TYPE_STRING
			assert cell.stringCellValue == h
		}
	}

	private workbook(Closure c) {
		wb = new ExcelFile().workbook(c)
	}
}
