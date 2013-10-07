GSheets
========

kktec/gsheets is a [Groovy](http://groovy.codehaus.org) [DSL](http://martinfowler.com/books/dsl.html) wrapper over [Apache POI](http://poi.apache.org) based on code forked from [andresteingress/gsheets](https://github.com/andresteingress/gsheets).

Overview
--------

It can be used to declaratively build or parse spreadsheets.

The original code, ExcelFile, does not support xml spreadsheets and is provided as a convenience and to provide building functionality not yet provided.

Changes
-------

0.3 adds grid parsing functionality for declaratively reading spreadsheets.

0.3.1 adds support for a default Workbook Date format of 'yyyy-mm-dd hh:mm', showing military style time hours (0-23)
0.3.2 adds the ability to autosize a specific no. of columns - call this after the sheet has been populated 

Plans as of 2013-08 include additional spreadsheet building features, better documentation.

Usage
-----

It assumes a simple grid on the specified worksheet, by name or index, originating at a specified startRowIndex (default is 0) and columnIndex (default is 0).
If no worksheet is specified, the first will be used. 

Examples
--------

Check the tests for more complete examples of usage. There are main methods on the tests that can be used to demonstrate building and parsing spreadsheets.

There are simple loading/parsing examples in the integration tests.


A simple example of building a Workbook with a Sheet with 1 header row and 3 data rows of 5 columns:

    Workbook workbook = builder.workbook {
        workbook {
            def fmt = new SimpleDateFormat('yyyy-MM-dd', Locale.default)
            sheet('sheet 1') {
                row('Name', 'Date', 'Count', 'Value', 'Active')
                row('a', fmt.parse('2012-09-12'), 69, 12.34, true)
                row('b', fmt.parse('2012-09-13'), 666, 43.21, false)
                autoColumnWidth(5)
            }
        }	
    }

    File file = new File(name)
    if (!file.exists()) {
        file.createNewFile()
    }
    OutputStream out = new FileOutputStream(file)
    workbook.write out
    out.close()




A simple example of a parsing a Workbook that pulls in all physically existing rows, disregarding a header row, with columns of various simple types as a List of Maps:

    FileInputStream ins = new FileInputStream('demo.xlsx')
    Workbook workbook = new XSSFWorkbook(ins)
    WorkbookParser parser = new WorkbookParser(workbook)
    List data = parser.grid {
        startRowIndex = 1
        columns name: 'string', date: 'date', count: 'int', value: 'decimal', active: 'boolean'
    }
    ins.close()

If you don't like a provided extractor or need a new one, you can replace an existing one or provide a new one by adding a line to your parsing Closure, i.e.:

    extractor('toUpper') { Cell cell -> cell.toString().toUpperCase() }

As to error handling, parsing will collect a List of individual cell data extraction errors. It will also fail fast on an unsupported extractor.

Note: the 'long' extractor is limited to 15 decimal digit Longs.
 
 
