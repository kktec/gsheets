package org.gsheets

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/** Support for .xlsx spreadsheets */
class XmlWorkbookSupport implements WorkbookSupport {

	Class<? extends Workbook> workbookType() { XSSFWorkbook }
	
}
