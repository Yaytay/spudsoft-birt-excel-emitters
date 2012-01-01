package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class Issue24 extends ReportRunner {
	
	@Test
	public void testExternalCss() throws BirtException, IOException {

		debug = false;
		InputStream inputStream = runAndRenderReport("Issue24.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
	
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 3, this.firstNullRow(sheet));

			assertEquals( "FF206090",              ((XSSFColor)sheet.getRow(1).getCell(0).getCellStyle().getFillForegroundColorColor()).getARGBHex());
			assertEquals( "FF206090",              ((XSSFColor)sheet.getRow(1).getCell(1).getCellStyle().getFillForegroundColorColor()).getARGBHex());
			assertEquals( "FF206090",              ((XSSFColor)sheet.getRow(1).getCell(2).getCellStyle().getFillForegroundColorColor()).getARGBHex());
			assertEquals( "FF206090",              ((XSSFColor)sheet.getRow(1).getCell(3).getCellStyle().getFillForegroundColorColor()).getARGBHex());
			
			assertEquals( "FF6495ED",              ((XSSFColor)sheet.getRow(2).getCell(0).getCellStyle().getFillForegroundColorColor()).getARGBHex());
			assertEquals( "FF6495ED",              ((XSSFColor)sheet.getRow(2).getCell(1).getCellStyle().getFillForegroundColorColor()).getARGBHex());
			assertEquals( "FF6495ED",              ((XSSFColor)sheet.getRow(2).getCell(2).getCellStyle().getFillForegroundColorColor()).getARGBHex());
			assertEquals( "FF6495ED",              ((XSSFColor)sheet.getRow(2).getCell(3).getCellStyle().getFillForegroundColorColor()).getARGBHex());
		} finally {
			inputStream.close();
		}
	}
	

}
