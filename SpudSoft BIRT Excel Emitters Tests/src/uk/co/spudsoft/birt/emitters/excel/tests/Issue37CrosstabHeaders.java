package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class Issue37CrosstabHeaders extends ReportRunner {

	@Test
	public void testCrosstabHeaders() throws BirtException, IOException {

		debug = false;
		InputStream inputStream = runAndRenderReport("issue_35_36_37_38.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
	
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 256, this.firstNullRow(sheet));

			assertEquals( "Country",          sheet.getRow(2).getCell(0).getStringCellValue() );
			assertEquals( "Customer",         sheet.getRow(2).getCell(1).getStringCellValue() );
			assertEquals( "Transaction Date", sheet.getRow(2).getCell(2).getStringCellValue() );
			assertEquals( "Title",            sheet.getRow(2).getCell(3).getStringCellValue() );
			assertEquals( "Version",          sheet.getRow(2).getCell(4).getStringCellValue() );
			assertEquals( "Turns",            sheet.getRow(2).getCell(5).getStringCellValue() );
			
			assertEquals( "Canada",           sheet.getRow(3).getCell(0).getStringCellValue() );
			assertEquals( "Customer #1",      sheet.getRow(3).getCell(1).getStringCellValue() );
			assertEquals( 1326153600000L,     sheet.getRow(3).getCell(2).getDateCellValue().getTime() );
			assertEquals( "TITLE #0",         sheet.getRow(3).getCell(3).getStringCellValue() );
			assertEquals( "VERSION",          sheet.getRow(3).getCell(4).getStringCellValue() );
			assertEquals( 1000.0,             sheet.getRow(3).getCell(5).getNumericCellValue(), 0.1 );

			
			
		} finally {
			inputStream.close();
		}
	}
	
	
}
