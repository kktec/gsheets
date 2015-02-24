package org.gsheets.building

import org.apache.poi.ss.usermodel.Font

class XmlWorkbookBuilderSpec extends WorkbookBuilderSpec {

	protected newBuilder() { new WorkbookBuilder(true) }
	
	protected int expectedNumberOfDefaultFonts() { 1 }
	
	protected boolean assertDifferingFontCharacteristics(Font font) {
		with(font) {
			fontName == 'Calibri'
			fontHeightInPoints == 11
			color == 8
		}
		true
	}
	
	static void main(String[] args) {
		String filename = args.length ? args[0] : 'demo.xlsx'
		file  filename, new XmlWorkbookBuilderSpec().newBuilder()
	}
	
}
