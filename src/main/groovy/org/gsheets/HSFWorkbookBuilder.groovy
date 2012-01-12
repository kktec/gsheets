package org.gsheets

import org.apache.poi.hssf.usermodel.HSSFRichTextString
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.util.CellRangeAddress

/**
 * @author me@andresteingress.com
 */
class HSSFWorkbookBuilder {

    private Workbook workbook = new HSSFWorkbook()
    private Sheet sheet
    private int rowsCounter

    private Map<String, CellStyle> cellStyles = [:]
    private Map<String, Font> fonts = [:]

    Workbook workbook(Closure closure) {
        assert closure

        closure.delegate = this
        closure.call()
        workbook
    }

    void sheet(String name, Closure closure) {
        assert name
        assert closure

        sheet = workbook.createSheet(name)
        rowsCounter = 0
        closure.delegate = this
        closure.call()
    }

    void cellStyle(String cellStyleId, Closure closure)  {
        assert cellStyleId
        assert !cellStyles.containsKey(cellStyleId)
        assert closure

        CellStyle cellStyle = workbook.createCellStyle()
        cellStyles.put(cellStyleId, cellStyle)

        closure.call(cellStyle)
    }

    void font(String fontId, Closure closure)  {
        assert fontId
        assert !fonts.containsKey(fontId)
        assert closure

        Font font = workbook.createFont()
        fonts.put(fontId, font)

        closure.call(font)
    }

    void applyCellStyle(Map<String, Object> args)  {
        String cellStyleId = args.cellStyle
        def rows = args.rows
        def cells = args.columns
        def fontId = args.font

        assert cellStyleId && cellStyles.containsKey(cellStyleId)
        assert rows && (rows instanceof Number || rows instanceof Range<Number>)
        assert cells && (cells instanceof Number || cells instanceof Range<Number>)
        
        if (fontId && !fonts.containsKey(fontId)) fontId = null
        if (rows instanceof  Number) rows = [rows]
        if (cells instanceof  Number) cells = [cells]

        rows.each { Number rowIndex ->
            assert rowIndex

            cells.each { Number cellIndex ->
                assert cellIndex

                Row row = sheet.getRow(rowIndex.intValue() - 1)
                if (!row) return

                Cell cell = row.getCell(cellIndex.intValue() - 1)
                if (!cell) return

                cell.setCellStyle(cellStyles.get(cellStyleId))
                if (fontId) cell.getCellStyle().setFont(fonts.get(fontId))
            }
        }
    }

    void mergeCells(Map<String, Object> args)  {
        def rows = args.rows
        def cols = args.columns
        
        assert rows && (rows instanceof Number || rows instanceof Range<Number>)
        assert cols && (cols instanceof Number || cols instanceof Range<Number>)

        if (rows instanceof Number) rows = [rows]
        if (cols instanceof Number) cols = [cols]

        sheet.addMergedRegion(new CellRangeAddress(rows.first() - 1, rows.last() - 1, cols.first() - 1, cols.last() - 1))
    }

    void header(List<String> names)  {
        assert names

        Row row = sheet.createRow(rowsCounter++ as int)
        names.eachWithIndex { String value, col ->
            Cell cell = row.createCell(col)
            cell.setCellValue(value)
        }
    }

    void row(values) {
        assert values

        Row row = sheet.createRow(rowsCounter++ as int)
        values.eachWithIndex {value, col ->
            Cell cell = row.createCell(col)
            switch (value) {
                case Date: cell.setCellValue((Date) value); break
                case Double: cell.setCellValue((Double) value); break
                case BigDecimal: cell.setCellValue(((BigDecimal) value).doubleValue()); break
                default: cell.setCellValue(new HSSFRichTextString("" + value)); break
            }
        }
    }

}