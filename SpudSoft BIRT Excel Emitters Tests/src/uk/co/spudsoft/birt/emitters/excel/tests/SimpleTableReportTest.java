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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class SimpleTableReportTest extends ReportRunner {

	@Test
	public void testRunReport() throws BirtException, IOException {

		debug = true;
		InputStream inputStream = runAndRenderReport("SimpleTable.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Simple Table Report", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals(2, firstNullRow(sheet));
			
			assertEquals( "1", sheet.getRow(0).getCell(0).getStringCellValue() );
			assertEquals( "2", sheet.getRow(1).getCell(0).getStringCellValue() );
			assertEquals( 3.0, sheet.getRow(0).getCell(1).getNumericCellValue(), 0.001 );
			assertEquals( Cell.CELL_TYPE_BLANK, sheet.getRow(1).getCell(1).getCellType() );
			
			assertEquals( "Title\nSubtitle", 	sheet.getHeader().getLeft() );
			assertEquals( "The Writer", 		sheet.getFooter().getLeft() );
			assertEquals( "1", 					sheet.getFooter().getCenter() );
		} finally {
			inputStream.close();
		}
	}
}
