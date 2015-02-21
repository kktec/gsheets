package org.gsheets.parsing

import groovy.util.logging.Log

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.gsheets.building.WorkbookBuilder

@Log
class XmlWorkbookParserSpec extends WorkbookParserSpec {

	protected WorkbookParser newParser(Workbook workbook) { new WorkbookParser(workbook) }
	
	protected WorkbookBuilder newBuilder() { new WorkbookBuilder(true) }
	
	// Note: run the corresponding builder spec as an app first to create the file 
	static void main(String[] args) {
		FileInputStream ins = new FileInputStream('demo_test.xlsx')
		Workbook workbook = new XSSFWorkbook(ins)
		WorkbookParser parser = new WorkbookParser(workbook)
		log.info parser.grid {
			startRowIndex = 1
			columns name: 'int', date: 'date', count: 'int', value: 'decimal', active: 'boolean'
		}
		ins.close()
	}
	
}
