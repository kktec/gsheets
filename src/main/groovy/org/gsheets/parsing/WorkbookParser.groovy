package org.gsheets.parsing

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

/**
 * Provides basic support for parsing a grid of data from xml or non-xml spreadsheets.
 * 
 * @author Ken Krebs 
 */
class WorkbookParser {
	
	private final Workbook workbook
	
	private String sheetName
	
	private int sheetIndex
	
	int startRowIndex
	
	int startColumnIndex
	
	private Map columnMap = [:]
	
	private final List errors = []
	
	private final Map extractors = [
		string: this.&cellAsString,
		decimal: this.&cellAsBigDecimal,
		'int': this.&cellAsInteger,
		'boolean': this.&cellAsBoolean,
		'double': this.&cellAsDouble,
		'float': this.&cellAsFloat,
		date: this.&cellAsDate,
		'long': this.&cellAsLong,
	]
	
	/**
	 * Constructs a WorkbookParser.
	 * 
	 * @param workbook to be parsed
	 */
	WorkbookParser(Workbook workbook) {
		assert workbook
		
		this.workbook = workbook
	}
	
	/**
	 * Parses a grid of data from the provided Workbook.
	 * Cell data extraction are collected in a List named errors which will be empty if there none.
	 * Application code may check this list to determine how to deal with it.
	 * 
	 * @param closure declares a parsing strategy
	 * 
	 * @return a List of data Maps
	 * 
	 * @throws IllegalArgumentException for unsupported conversions
	 */
	List<Map> grid(Closure closure) {
		assert closure
		
		closure.delegate = this
		closure.call()
		
		data(workbook)
	}
	
	/**
	 * Sets the name of the worksheet of interest.
	 * 
	 * @param name worksheet name
	 */
	void sheet(String name) { sheetName = name }
	
	/**
	 * Sets the index of the worksheet of interest.
	 * 
	 * @param name worksheet index
	 */
	void sheet(int index) { sheetIndex = index }
	
	/**
	 * Sets the column data extractor strategy. Columns are processed in order.
	 * 'skip' is a special strategy that skips over a column.
	 * 
	 * @param columns LinkedHashMap of column names to extractor names
	 */
	void columns(Map columns) { columnMap = columns }
	
	/**
	 * Adds or replaces an extractor.
	 * 
	 * @param name of the extractor
	 * @param extractor Closure that extracts data from a Cell
	 */
	void extractor(String name, Closure extractor) { extractors[name] = extractor }
	
	private List<Map> data(Workbook workbook) {
		List data = []
		Sheet sheet 
		if (sheetName) { sheet = workbook.getSheet(sheetName) }
		sheet = sheet ?: workbook.getSheetAt(sheetIndex)

		int rows = sheet.physicalNumberOfRows - startRowIndex
		rows.times {
			Row row = sheet.getRow(it + startRowIndex)
			if (row) { data << rowData(row) }
		}
		data
	}
	
	private Map rowData(Row row) {
		Map data = [:]
		columnMap.eachWithIndex { column, extractorName, index ->
			Cell cell = row.getCell(index + startColumnIndex)
			if (extractorName.toLowerCase() != 'skip') {
				Closure extractor = extractors[extractorName]
				if (extractor) {
					try { 
						data[column] = extractor cell
					} catch (Exception e) {
						data[column] = null
						errors << [rowIndex: row.rowNum, columnIndex: index, column: column, error: e.toString(), value: cell.toString()]
					}
				}
				else { 
					throw new IllegalArgumentException("$extractorName is not a supported extractor for column $column")
				}
			}
		}
		data
	}
	
	private String cellAsString(Cell cell) { cell.toString() }
	
	private Boolean cellAsBoolean(Cell cell) { cell.booleanCellValue }
	
	private BigDecimal cellAsBigDecimal(Cell cell) { new BigDecimal(cell.toString()) }

	private Double cellAsDouble(Cell cell) { cell.numericCellValue }
	
	private Float cellAsFloat(Cell cell) { cellAsDouble(cell).floatValue() }
	
	private Long cellAsLong(Cell cell) { cellAsBigDecimal(cell).longValue() }
	
	private Integer cellAsInteger(Cell cell) { cellAsBigDecimal(cell).intValue() }
	
	private Date cellAsDate(Cell cell) { cell.dateCellValue }
	
}
