package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class Groupings extends ReportRunner {

	@Test
	public void testGroupings() throws BirtException, IOException {

		InputStream inputStream = runAndRenderReport("Grouping.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 3, workbook.getNumberOfSheets() );

			Sheet sheet1 = workbook.getSheetAt(0);
			assertEquals( "HeaderAndFooter", sheet1.getSheetName());
			
			// No way to access groups in POI :(

		} finally {
			inputStream.close();
		}
	}
}
