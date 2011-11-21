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

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class NestedTables2ReportTest extends ReportRunner {
	
	@Test
	public void testRunReport() throws BirtException, IOException {

		InputStream inputStream = runAndRenderReport("NestedTables2.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Nested Tables Test Report", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals(6, firstNullRow(sheet));
			
			assertEquals( "1\n2\n3\n1\n2\n3", sheet.getRow(0).getCell(0).getStringCellValue());
			
			XSSFColor bgColour = ((XSSFCell)sheet.getRow(0).getCell(0)).getCellStyle().getFillForegroundColorColor();
			assertEquals( "FFFFFFFF", bgColour.getARGBHex() );
			// assertEquals( null, bgColour );
			XSSFColor baseColour = ((XSSFCell)sheet.getRow(0).getCell(0)).getCellStyle().getFont().getXSSFColor();
			assertEquals( "FFFFFFFF", baseColour.getARGBHex() );
			// assertTrue( !bgColour.equals( baseColour ) );
			XSSFRichTextString rich = (XSSFRichTextString)sheet.getRow(0).getCell(0).getRichStringCellValue();
			assertEquals( 2, rich.numFormattingRuns() );
			assertEquals( 11, rich.getString().length() );
		} finally {
			inputStream.close();
		}
	}

	@Test
	public void testRunReportXls() throws BirtException, IOException {

		InputStream inputStream = runAndRenderReport("NestedTables2.rptdesign", "xls");
		assertNotNull(inputStream);
		try {
			
			HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Nested Tables Test Report", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals(6, firstNullRow(sheet));
			
			assertEquals( "1\n2\n3\n1\n2\n3", sheet.getRow(0).getCell(0).getStringCellValue());
			
			short bgColour = ((HSSFCell)sheet.getRow(0).getCell(0)).getCellStyle().getFillBackgroundColor();
			assertEquals( "0:0:0", workbook.getCustomPalette().getColor(bgColour).getHexString() );
			short baseColour = workbook.getFontAt(((HSSFCell)sheet.getRow(0).getCell(0)).getCellStyle().getFontIndex()).getColor();
			assertEquals( "0:0:0", workbook.getCustomPalette().getColor(baseColour).getHexString() );
			// Someone else can explain how it makes sense for these two to need to be the same, given that the result is them being different!
			assertEquals( workbook.getCustomPalette().getColor(bgColour).getHexString(), workbook.getCustomPalette().getColor(baseColour).getHexString() );
			HSSFRichTextString rich = (HSSFRichTextString)sheet.getRow(0).getCell(0).getRichStringCellValue();
			assertEquals( 3, rich.numFormattingRuns() );
			assertEquals( 11, rich.getString().length() );
		} finally {
			inputStream.close();
		}
	}
}
