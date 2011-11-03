/********************************************************************************
* (C) Copyright 2011, by James Talbut.
*
*   This program is free software: you can redistribute it and/or modify
*   it under the terms of the GNU General Public License as published by
*   the Free Software Foundation, either version 3 of the License, or
*   (at your option) any later version.
*
*   This program is distributed in the hope that it will be useful,
*   but WITHOUT ANY WARRANTY; without even the implied warranty of
*   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
*   GNU General Public License for more details.
*
*   You should have received a copy of the GNU General Public License
*   along with this program.  If not, see <http://www.gnu.org/licenses/>.
*
*   [Java is a trademark or registered trademark of Sun Microsystems, Inc.
*   in the United States and other countries.]
********************************************************************************/

package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class FontsReportTest extends ReportRunner {

	private void assertFontCell( Sheet sheet, int row, int col, String contents, String fontName, int fontHeight ) {
		
		Cell cell = sheet.getRow(row).getCell(col);
		CellStyle style = cell.getCellStyle();
		
		assertEquals( contents, cell.getStringCellValue());
		assertEquals( fontName, sheet.getWorkbook().getFontAt( style.getFontIndex() ).getFontName().replace("\"", "") );
		assertEquals( fontHeight, sheet.getWorkbook().getFontAt( style.getFontIndex() ).getFontHeightInPoints() );
	}
	
	@Test
	public void testRunReport() throws BirtException, IOException {

		InputStream inputStream = runAndRenderReport("Fonts.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Fonts Test Report", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertNotNull( sheet.getRow(0) );
			assertNotNull( sheet.getRow(1) );
			assertNotNull( sheet.getRow(2) );
			assertNotNull( sheet.getRow(3) );
			assertNotNull( sheet.getRow(4) );
			assertNotNull( sheet.getRow(5) );
			assertNull( sheet.getRow(6) );
			
			assertFontCell( sheet, 0, 0, "Sans Serif 10pt"       , "Arial"          , 10);
			assertFontCell( sheet, 1, 0, "Sans Serif 10pt Underlined"       , "Arial"          , 10);
			assertFontCell( sheet, 2, 0, "Serif 10pt"            , "Times New Roman", 10);
			assertFontCell( sheet, 3, 0, "Tahoma Medium"         , "Tahoma"         , 12);
			assertFontCell( sheet, 4, 0, "Comic Sans MS X Small" , "Comic Sans MS"  , 8);
			assertFontCell( sheet, 5, 0, "Verdana XX Large"      , "Verdana"        , 20);
			
		} finally {
			inputStream.close();
		}
	}

}
