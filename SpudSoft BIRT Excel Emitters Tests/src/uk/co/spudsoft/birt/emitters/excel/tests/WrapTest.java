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

import static org.junit.Assert.*;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class WrapTest extends ReportRunner {

	@Test
	public void testRunReport() throws BirtException, IOException {

		InputStream inputStream = runAndRenderReport("Wrap.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 4, workbook.getNumberOfSheets() );
			assertEquals( "Auto", workbook.getSheetAt(0).getSheetName());
			assertEquals( "NoWrap", workbook.getSheetAt(1).getSheetName());
			assertEquals( "Normal", workbook.getSheetAt(2).getSheetName());
			assertEquals( "Preformatted", workbook.getSheetAt(3).getSheetName());
			
			assertTrue( ! workbook.getSheetAt( 0 ).getRow( 1 ).getCell( 1 ).getCellStyle().getWrapText() );
			assertTrue( workbook.getSheetAt( 0 ).getRow( 1 ).getCell( 2 ).getCellStyle().getWrapText() );
			
			assertTrue( ! workbook.getSheetAt( 1 ).getRow( 1 ).getCell( 1 ).getCellStyle().getWrapText() );
			assertTrue( ! workbook.getSheetAt( 1 ).getRow( 1 ).getCell( 2 ).getCellStyle().getWrapText() );
			
			assertTrue( ! workbook.getSheetAt( 2 ).getRow( 1 ).getCell( 1 ).getCellStyle().getWrapText() );
			assertTrue( workbook.getSheetAt( 2 ).getRow( 1 ).getCell( 2 ).getCellStyle().getWrapText() );
			
			assertTrue( workbook.getSheetAt( 3 ).getRow( 1 ).getCell( 1 ).getCellStyle().getWrapText() );
			assertTrue( workbook.getSheetAt( 3 ).getRow( 1 ).getCell( 2 ).getCellStyle().getWrapText() );
			
		} finally {
			inputStream.close();
		}
	}

}
