/**
	see README for copyright/licensing info 
 */
package org.gsheets

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Workbook

/**
 * Simple smoke tests for ExcelFile
 * 
 * @author me@andresteingress.com
 */
class ExcelFileSmokeTests extends GroovyTestCase {
	
    private File excel

    void setUp() {
		super.setUp()
        excel = new File('test.xls')
        if (!excel.exists()) {
			 excel.createNewFile()
        }
    }

    void testCreateSimpleWorkbook()  {
        Workbook workbook = new ExcelFile().workbook {

            styles {
                font('bold')  { Font font ->
                    font.setBoldweight(Font.BOLDWEIGHT_BOLD)
                }

                cellStyle ('header')  { CellStyle cellStyle ->
                    cellStyle.setAlignment(CellStyle.ALIGN_CENTER)

                }
            }

            data {
                sheet ('Export')  {
                    header(['Column1', 'Column2', 'Column3'])

                    row([10, 20, '=A2*B2'])
                    row(['', '', '=sum(C2)'])
                }
            }

            commands {
                applyCellStyle(cellStyle: 'header', font: 'bold', rows: 1, columns: 1..3)
                applyColumnWidth(columns: 1..2, width: 200)
                // mergeCells(rows: 1, columns: 1..3)
            }
        }
		
		assert workbook.numberOfSheets == 1

        def excelOut = new FileOutputStream(excel)
        workbook.write(excelOut)
        excelOut.close()
    }
}
