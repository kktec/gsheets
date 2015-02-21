package org.gsheets

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Workbook

/** Support for .xls spreadsheets */
class NonXmlWorkbookSupport implements WorkbookSupport {
	
	Class<? extends Workbook> workbookType() { HSSFWorkbook }
	
}
