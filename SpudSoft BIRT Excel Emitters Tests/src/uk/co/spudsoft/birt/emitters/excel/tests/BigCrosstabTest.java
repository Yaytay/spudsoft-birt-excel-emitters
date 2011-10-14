package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class BigCrosstabTest extends ReportRunner {

	@Test
	public void test1() throws Exception {

		InputStream inputStream = runAndRenderReport("BigCrosstab.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Big Crosstab Report 1", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 236, firstNullRow(sheet));
			
			assertEquals(28, greatestNumColumns(sheet));
			
		} finally {
			inputStream.close();
		}
	}
}
