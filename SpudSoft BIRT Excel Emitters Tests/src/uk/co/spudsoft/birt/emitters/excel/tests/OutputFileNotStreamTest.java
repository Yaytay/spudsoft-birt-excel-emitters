package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class OutputFileNotStreamTest extends ReportRunner {

	@Test
	public void testRunReport() throws BirtException, IOException {

		InputStream inputStream = runAndRenderReportFileNotStream("NumberFormats.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Number Formats Test Report", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals(18, this.firstNullRow(sheet));
		} finally {
			inputStream.close();
		}
		
	}
}
