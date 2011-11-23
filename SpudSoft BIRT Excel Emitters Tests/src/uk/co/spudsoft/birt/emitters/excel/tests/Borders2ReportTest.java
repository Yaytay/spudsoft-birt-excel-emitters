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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class Borders2ReportTest extends ReportRunner {

	/**
	 * Check that the borders for a given cell match the expected values.
	 * This is complicated by the fact that POI will not always give a particular cell the borders that are seen in Excel
	 * - neighbouring cells may override the values for the chosen cell.
	 * I don't know how to tell which takes precedence, but the following works for the tests I've carried out.
	 */
	private void assertBorder( Sheet sheet, int row, int col, short bottom, short left, short right, short top ) {
		
		Row curRow = sheet.getRow( row );
		Row prevRow = ( row > 0 ) ? sheet.getRow( row - 1 ) : null;
		Row nextRow = sheet.getRow( row + 1 );
		Cell cell = curRow.getCell(col);
		CellStyle style = cell.getCellStyle();
		
		Cell cellUp = ( prevRow == null ) ? null : prevRow.getCell( col );
		Cell cellDown = ( nextRow == null ) ? null : nextRow.getCell( col );
		Cell cellLeft = ( col == 0 ) ? null : curRow.getCell( col - 1 ); 
		Cell cellRight = curRow.getCell( col + 1 ); 
		
		CellStyle styleUp = ( cellUp == null ) ? null : cellUp.getCellStyle();
		CellStyle styleDown = ( cellDown == null ) ? null : cellDown.getCellStyle();
		CellStyle styleLeft = ( cellLeft == null ) ? null : cellLeft.getCellStyle();
		CellStyle styleRight = ( cellRight == null ) ? null : cellRight.getCellStyle();
		
		if( ( top != style.getBorderTop() ) && 
				( styleUp == null ) || ( top != styleUp.getBorderBottom() ) ) {
			assertEquals( top,    style.getBorderTop() );
		}
		if( ( bottom != style.getBorderBottom() ) && 
				( styleDown == null ) || ( top != styleDown.getBorderTop() ) ) {
			assertEquals( bottom, style.getBorderBottom() );
		}
		if( ( left != style.getBorderLeft() ) && 
				( styleLeft == null ) || ( top != styleLeft.getBorderRight() ) ) {
			assertEquals( left,   style.getBorderLeft() );
		}
		if( ( right != style.getBorderRight() ) && 
				( styleRight == null ) || ( right != styleRight.getBorderLeft() ) ) {
			assertEquals( right,  style.getBorderRight() );
		}
	}
	
	@Test
	public void testRunReport() throws BirtException, IOException {

		debug = true;
		removeEmptyRows = false;
		InputStream inputStream = runAndRenderReport("Borders2.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Borders Test Report 2", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 4, firstNullRow(sheet));
			
			assertBorder( sheet, 1, 2, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM );

			assertBorder( sheet, 1, 4, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM, CellStyle.BORDER_MEDIUM );
			
		} finally {
			inputStream.close();
		}
	}

}
