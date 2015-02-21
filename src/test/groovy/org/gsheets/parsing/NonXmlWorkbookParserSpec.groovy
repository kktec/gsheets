package org.gsheets.parsing

import groovy.util.logging.Log

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook
import org.gsheets.building.WorkbookBuilder

@Log
class NonXmlWorkbookParserSpec extends WorkbookParserSpec {

	protected WorkbookParser newParser(Workbook workbook) { new WorkbookParser(workbook) }
	
	protected WorkbookBuilder newBuilder() { new WorkbookBuilder(false) }
	
	// Note: run the corresponding builder spec as an app first to create the file 
	static void main(String[] args) {
		FileInputStream ins = new FileInputStream('demo.xls')
		Workbook workbook = new HSSFWorkbook(ins)
		WorkbookParser parser = new WorkbookParser(workbook)
		log.info parser.grid {
			startRowIndex = 1
			columns name: 'int', date: 'date', count: 'int', value: 'decimal', active: 'boolean'
		}
		ins.close()
	}
}
