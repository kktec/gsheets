package org.gsheets.parsing

import org.apache.poi.ss.usermodel.Workbook
import org.gsheets.building.WorkbookBuilder

import spock.lang.Shared
import spock.lang.Specification
import spock.lang.Unroll

abstract class WorkbookParserSpec extends Specification {

	@Shared WorkbookParser parser
	
	WorkbookBuilder builder
	
	abstract protected WorkbookParser newParser(Workbook workbook)
	
	abstract protected WorkbookBuilder newBuilder()
	
	Date date1 = new Date()
	Date date2 = new Date(date1.time + 1000 * 60 * 60 * 24)
	List rowA = ['a', 'rubbish', 12.34, 13I, true, 3.14159D, 2.345F, date1, 123456789L]
	List rowB = ['b', 'garbage', -98.62, -69I, false, 2.71821828D, 13.01F, date2, 999999999999999L]
	Map mapA = [
		abbreviation: 'a', cost: 12.34, num: 13, status: true,
		irr: 3.14159D, flt: 2.345F, date: date1, lng: 123456789L
	]
	Map mapB = [
		abbreviation: 'b', cost: -98.62, num: -69, status: false,
		irr: 2.71821828D, flt: 13.01F, date: date2, lng: 999999999999999L
	]
	
	Closure buildIt = {
		sheet('a sheet') {
			row rowA
			row rowB
		}
	}
	
	Closure gridIt = {
		columns(
			abbreviation: 'string', disregard: 'skip', cost: 'decimal',
			num: 'int', status: 'boolean', irr: 'double', flt: 'float', date: 'date', lng: 'long'
		)
	}
	
	def setup() { builder = newBuilder() }
	
	def 'can parse a grid without header rows of simple types from the first worksheet by default'() {
		given:
		builder.workbook buildIt
		parser = newParser(builder.wb)
		
		when:
		List data = parser.grid gridIt
		
		then:
		data == [mapA, mapB]
		!parser.errors
	}

	@Unroll
	def 'from parsed grid: #a plus #b equals #c'() {
		expect:
		a + b == c
		!parser.errors
		
		where:
		[a, b, c] << abcData().collect { it*.value }
		
		// Alternatively
//		row << abcData()
//		a = row.a
//		b = row.b
//		c = row.c
	}
	
	private List abcData() {
		setup()
		builder.workbook {
			sheet('a sheet') {
				row 2, 4, 6
				row 23, 46, 69
			}
		}
		parser = newParser(builder.wb)
		parser.grid { columns a: 'int', b: 'int', c: 'int' }
	}
	
	def 'can parse a grid with ignored header rows of simple types'() {
		given:
		builder.workbook {
			sheet('a sheet') {
				row('Abbreviation', 'Cost', '')
				row('ab', 'x', 'y')
				row rowA
				row rowB
			}
		}
		parser = newParser(builder.wb)
		
		when:
		List data = parser.grid {
			startRowIndex = 2
			columns(
				abbreviation: 'string', disregard: 'SKIP', cost: 'decimal',
				num: 'int', status: 'boolean', irr: 'double', flt: 'float', date: 'date', lng: 'long'
			)
		}
		parser = newParser(builder.wb)
		
		then:
		data == [mapA, mapB]
		!parser.errors
	}
	
	def 'parsing a grid fails fast with an unsupported extractor'() {
		given:
		builder.workbook { 
			sheet('a') {
				row 'whatever'
			}
		}
		parser = newParser(builder.wb)
		
		when: 
		parser.grid { columns y: 'map' }
		
		then:
		IllegalArgumentException x = thrown()
		x.message == 'map is not a supported extractor for column y'
	}
	
	def 'can parse a grid from a worksheet by name'() {
		given:
		builder.workbook {
			sheet('one') { }
			sheet('special') {
				row 'something'
			}
		}
		parser = newParser(builder.wb)	
		
		when:
		List data = parser.grid {
			sheet('special')
			columns x: 'string'
		}
		
		then:
		data[0] == [x: 'something']
	}
	
	def 'can parse a grid from a worksheet by index'() {
		given:
		builder.workbook {
			sheet('one') { }
			sheet('special') {
				row 'something'
			}
		}
		parser = newParser(builder.wb)		
		
		when:
		List data = parser.grid {
			sheet(1)
			columns x: 'string'
		}
		
		then:
		data[0] == [x: 'something']
	}
	
	def 'can parse a grid skipping columns'() {
		given:
		builder.workbook {
			sheet('special') {
				row 'a', 'b', 'something'
			}
		}
		parser = newParser(builder.wb)
		
		when:
		List data = parser.grid {
			startColumnIndex = 2
			columns x: 'string'
		}
		
		then:
		data == [[x: 'something']]
	}
	
	def 'can parse a grid using a custom cell extractor'() {
		given:
		builder.workbook {
			sheet('special') {
				row 'something'
				row 'else'
			}
		}
		parser = newParser(builder.wb)
		
		when:
		List data = parser.grid {
			extractor('toUpper') { it.toString().toUpperCase() }
			columns x: 'toUpper'
		}
		
		then:
		data == [[x: 'SOMETHING'], [x: 'ELSE']]
		!parser.errors
	}
	
	def 'can report parsing errors on individual cells'() {
		given:
		builder.workbook {
			sheet('special') {
				row 'Thingy', 'Integer', 'Status'
				row 'thing', 3, 'ok'
				row 'something', 'notAnInt', 'bad'
				row 'athing', 13, 'a-ok'
			}
		}
		parser = newParser(builder.wb)
		
		when:
		List data = parser.grid {
			startRowIndex = 1
			columns x: 'string', y: 'int', z: 'string'
		}
		
		then:
		data == [
			[x: 'thing', y: 3, z: 'ok'],
			[x: 'something', y: null, z: 'bad'],
			[x: 'athing', y: 13, z: 'a-ok'],
		]
		parser.errors.size() == 1
		parser.errors[0] == [
			rowIndex: 2, columnIndex: 1, column: 'y',
			error: 'java.lang.NumberFormatException', value: 'notAnInt'
		]
	}

}
