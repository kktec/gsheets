package org.gsheets.building

import org.apache.poi.ss.usermodel.Font

class NonXmlWorkbookBuilderSpec extends WorkbookBuilderSpec {

	protected newBuilder() { new WorkbookBuilder(false) }
	
	protected int expectedNumberOfDefaultFonts() { 4 }
	
	protected boolean assertDifferingFontCharacteristics(Font font) {
		with(font) {
			fontName == 'Arial'
			fontHeightInPoints == 10
			color == Font.COLOR_NORMAL
		}
		true
	}
	
	static void main(String[] args) {
		String filename = args.length ? args[0] : 'demo.xls'
		file filename, new NonXmlWorkbookBuilderSpec().newBuilder()
	}
	
}
