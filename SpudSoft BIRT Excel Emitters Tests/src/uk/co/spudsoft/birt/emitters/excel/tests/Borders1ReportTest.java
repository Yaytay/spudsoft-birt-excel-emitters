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

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class Borders1ReportTest extends ReportRunner {

	private void assertBorder( Sheet sheet, int row, int col, short bottom, short left, short right, short top ) {
		
		Cell cell = sheet.getRow(row).getCell(col);
		CellStyle style = cell.getCellStyle();
		
		assertEquals( bottom, style.getBorderBottom() );
		assertEquals( left,   style.getBorderLeft() );
		assertEquals( right,  style.getBorderRight() );
		assertEquals( top,    style.getBorderTop() );
	}
	
	@Test
	public void testRunReport() throws BirtException, IOException {

		InputStream inputStream = runAndRenderReport("Borders1.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Borders Test Report 1", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 12, firstNullRow(sheet));
			
			int i = 0;
			assertBorder( sheet, i++, 0, CellStyle.BORDER_THIN, CellStyle.BORDER_THIN, CellStyle.BORDER_THIN, CellStyle.BORDER_THIN );
			assertBorder( sheet, i++, 0, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM );
			assertBorder( sheet, i++, 0, CellStyle.BORDER_THICK, CellStyle.BORDER_THICK, CellStyle.BORDER_THICK, CellStyle.BORDER_THICK );
			
			assertBorder( sheet, i++, 0, CellStyle.BORDER_DOTTED, CellStyle.BORDER_DOTTED, CellStyle.BORDER_DOTTED, CellStyle.BORDER_DOTTED );
			assertBorder( sheet, i++, 0, CellStyle.BORDER_DOTTED, CellStyle.BORDER_DOTTED, CellStyle.BORDER_DOTTED, CellStyle.BORDER_DOTTED );
			assertBorder( sheet, i++, 0, CellStyle.BORDER_DOTTED, CellStyle.BORDER_DOTTED, CellStyle.BORDER_DOTTED, CellStyle.BORDER_DOTTED );
			
			assertBorder( sheet, i++, 0, CellStyle.BORDER_DASHED, CellStyle.BORDER_DASHED, CellStyle.BORDER_DASHED, CellStyle.BORDER_DASHED );
			assertBorder( sheet, i++, 0, CellStyle.BORDER_MEDIUM_DASHED, CellStyle.BORDER_MEDIUM_DASHED, CellStyle.BORDER_MEDIUM_DASHED, CellStyle.BORDER_MEDIUM_DASHED );
			assertBorder( sheet, i++, 0, CellStyle.BORDER_MEDIUM_DASHED, CellStyle.BORDER_MEDIUM_DASHED, CellStyle.BORDER_MEDIUM_DASHED, CellStyle.BORDER_MEDIUM_DASHED );
			
			assertBorder( sheet, i++, 0, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE );
			assertBorder( sheet, i++, 0, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE );
			assertBorder( sheet, i++, 0, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE, CellStyle.BORDER_DOUBLE );
			
		} finally {
			inputStream.close();
		}
	}

}
