package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class Lists extends ReportRunner {
	
	@Test
	public void testLists() throws BirtException, IOException {

		debug = true;
		InputStream inputStream = runAndRenderReport("Lists.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 28, workbook.getNumberOfSheets() );
	
			assertEquals( 17, this.firstNullRow(workbook.getSheetAt(0)));
			assertEquals( 11, this.firstNullRow(workbook.getSheetAt(1)));
			assertEquals( 11, this.firstNullRow(workbook.getSheetAt(2)));
			assertEquals( 13, this.firstNullRow(workbook.getSheetAt(3)));
			assertEquals( 11, this.firstNullRow(workbook.getSheetAt(4)));
			assertEquals( 13, this.firstNullRow(workbook.getSheetAt(5)));
			assertEquals( 31, this.firstNullRow(workbook.getSheetAt(6)));
			assertEquals( 33, this.firstNullRow(workbook.getSheetAt(7)));
			assertEquals( 9, this.firstNullRow(workbook.getSheetAt(8)));
			assertEquals( 11, this.firstNullRow(workbook.getSheetAt(9)));

			assertEquals( "Australia", workbook.getSheetAt(0).getSheetName() );
			assertEquals( "Austria", workbook.getSheetAt(1).getSheetName() );
			
		} finally {
			inputStream.close();
		}
	}
	

}
