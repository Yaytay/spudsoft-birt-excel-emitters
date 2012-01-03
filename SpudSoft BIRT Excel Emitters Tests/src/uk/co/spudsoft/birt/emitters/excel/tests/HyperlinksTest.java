package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class HyperlinksTest extends ReportRunner {

	@Test
	public void testHyperlinksXlsx() throws BirtException, IOException {

		debug = false;
		InputStream inputStream = runAndRenderReport("Hyperlinks.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
	
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 2002, this.firstNullRow(sheet));

			for(int i = 1; i < 2000; ++i ) {
				assertEquals( "http://www.spudsoft.co.uk/?p=" + i,              sheet.getRow(i).getCell(0).getHyperlink().getAddress());
				assertEquals( 1,             workbook.getFontAt( sheet.getRow(i).getCell(0).getCellStyle().getFontIndex() ).getUnderline() );
				assertEquals( "FF0000FF",    ((XSSFFont)workbook.getFontAt( sheet.getRow(i).getCell(0).getCellStyle().getFontIndex() ) ).getXSSFColor().getARGBHex() );
			}
		
		} finally {
			inputStream.close();
		}
	}

	@Test
	public void testHyperlinksXls() throws BirtException, IOException {

		debug = false;
		InputStream inputStream = runAndRenderReport("Hyperlinks.rptdesign", "xls");
		assertNotNull(inputStream);
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
	
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 2002, this.firstNullRow(sheet));

			for(int i = 1; i < 2000; ++i ) {
				assertEquals( "http://www.spudsoft.co.uk/?p=" + i,              sheet.getRow(i).getCell(0).getHyperlink().getAddress());
				assertEquals( 1,           workbook.getFontAt( sheet.getRow(i).getCell(0).getCellStyle().getFontIndex() ).getUnderline() );
				assertEquals( "0:0:FFFF",  workbook.getCustomPalette().getColor( workbook.getFontAt( sheet.getRow(i).getCell(0).getCellStyle().getFontIndex() ).getColor() ).getHexString() );
			}
		
		} finally {
			inputStream.close();
		}
	}
	

}
