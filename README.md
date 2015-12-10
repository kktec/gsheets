GSheets
========

kktec/gsheets is a [Groovy](http://groovy.codehaus.org) DSL wrapper over [Apache POI](http://poi.apache.org) based on code forked from [andresteingress/gsheets](https://github.com/andresteingress/gsheets).

Overview
--------

It can be used to declaratively parse or build spreadsheets.

The original code, ExcelFile, does not support xml spreadsheets and is provided as a convenience and to provide building functionality not yet provided.

Plans as of 2015-08 include additional spreadsheet building/parsing features, additional samples, and better documentation.
See the GitHub issue tracker for more or to make feature requests or report bugs.



Release Notes & Versions
--------

__0.4.2 THE_CURRENT_VERSION__

__Fixed:__ Issue #19 - startRowIndex and maxRows ignored in closure for WorkbookParser.grid()

Technology upgrades:

* Update Groovy to 2.3.11
* Update Spock to 1.0
* Upgrade Gradle to 2.9
* Updated Cobertura Gradle plugin to 2.2.28
* Updated CodeNarc to 0.24


__0.4.1

* #17 Add building support for styling a workbook, sheet, column, row, or individual cell
* #10 Specify a maximum number of rows to parse
* Add support for using a custom Date format 
 

__0.4.0__ 

This is a technology upgrade release, primarily Groovy2.

* Updated Groovy compiler to 2.0
* Updated Spock for Groovy2
* Updated Java to 7
* Updated Gradle to 2.3
* Updated Codenarc to 0.22
* Updated Cobertura Gradle plugin
* Fixed issue #13, builder rowIndex counter is now reset on building a new sheet

__0.3.2a THE_FIRST_VERSION__

* now using Gradle 1.10 
* built on Groovy 1.8 and is therefore usable in any Grails 2 app
* first published version with binaries deployed to jcenter at binTray
* allows the special extractor 'skip' to be specified in a case insensitive way as all upper case reads better
 
__0.3.2__

* adds the ability to autosize the width of a specific no. of columns - call this after the sheet has been populated

__0.3.1__

* adds support for building a Workbook with a specified Date format, default format is 'yyyy-mm-dd hh:mm' showing military style time hours (0-23)

__0.3__

* adds grid parsing functionality for declaratively reading spreadsheets



Notes
-----

Check the tests for more detailed examples of usage !!!

There are main methods on the tests that can be used to demonstrate building and parsing spreadsheets.

There are other simple building/parsing examples in the integration tests.

Thanks to Sean Gilligan for helping out. - Ken Krebs

To use it in Grails2 or a gradle build, simply add a compile dependency to *THE_CURRENT_VERSION* against jcenter or mavenRepo 'http://dl.bintray.com/kktec/maven'




Parsing
-------

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
    
To skip over a column, give it any name and specify an extractor of 'skip' (any case as of v0.3.2a).

As to error handling, parsing will collect a List of individual cell data extraction errors. It will also fail fast on an unsupported extractor.

NOTE: the 'long' extractor is limited to 15 decimal digit Longs.
 
 


Building
--------

NOTE:
The building feature is provided to allow simple data dumps and some simple styling. It does NOT currently support cell/row merging.

The older feature, ExcelFile, can be used to provide some of this but there is no intent to further build on this. It will be deprecated once all of its features are provided. The tests clearly show how to use the new styling support.

It assumes a simple grid on the specified worksheet, by name or index, originating at a specified startRowIndex (default is 0) and columnIndex (default is 0).
If no worksheet is specified, the first will be used. 

A simple example of building a Workbook with a Sheet with 1 header row and 3 data rows of 5 columns:

	WorkbookBuilder builder = new WorkbookBuilder(true)     // true for .xlsx, false for .xls
    Workbook workbook = builder.workbook {
        def fmt = new SimpleDateFormat('yyyy-MM-dd', Locale.default)
        sheet('sheet 1') {
            row('Name', 'Date', 'Count', 'Value', 'Active')
            row('a', fmt.parse('2012-09-12'), 69, 12.34, true)
            row('b', fmt.parse('2012-09-13'), 666, 43.21, false)
            autoColumnWidth(5)
        }
    }

    File file = new File(name)
    if (!file.exists()) {
        file.createNewFile()
    }
    OutputStream out = new FileOutputStream(file)
    workbook.write out
    out.close()


