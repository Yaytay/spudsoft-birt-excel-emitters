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

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class NestedTablesReportTest extends ReportRunner {

	@Test
	public void testRunReport() throws BirtException, IOException {

		InputStream inputStream = runAndRenderReport("NestedTables.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Nested Tables Test Report", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertNotNull( sheet.getRow(0) );
			assertNotNull( sheet.getRow(1) );
			assertNotNull( sheet.getRow(2) );
			assertNotNull( sheet.getRow(3) );
			assertNull( sheet.getRow(4) );
			
			assertEquals( 1.0, sheet.getRow(1).getCell(0).getNumericCellValue(), 0.001);
			assertEquals( 2.0, sheet.getRow(1).getCell(1).getNumericCellValue(), 0.001);
			assertEquals( 3.0, sheet.getRow(1).getCell(2).getNumericCellValue(), 0.001);
			assertEquals( 2.0, sheet.getRow(2).getCell(0).getNumericCellValue(), 0.001);
			assertEquals( 4.0, sheet.getRow(2).getCell(1).getNumericCellValue(), 0.001);
			assertEquals( 6.0, sheet.getRow(2).getCell(2).getNumericCellValue(), 0.001);
			assertEquals( 3.0, sheet.getRow(3).getCell(0).getNumericCellValue(), 0.001);
			assertEquals( 6.0, sheet.getRow(3).getCell(1).getNumericCellValue(), 0.001);
			assertEquals( 9.0, sheet.getRow(3).getCell(2).getNumericCellValue(), 0.001);
			
			assertEquals( 3510, sheet.getColumnWidth(0) );
			assertEquals( 3510, sheet.getColumnWidth(1) );
			assertEquals( 3510, sheet.getColumnWidth(2) );			
		} finally {
			inputStream.close();
		}
	}
}
