package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class AutoRowHeightsTest extends ReportRunner {

	@Test
	public void testRunReportXlsx() throws BirtException, IOException {

		debug = true;
		InputStream inputStream = runAndRenderReport("AutoRowHeight.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Auto RowHeight Report", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 6, this.firstNullRow(sheet));
			
			assertEquals( 309, sheet.getRow(0).getHeight() );
			assertEquals( 907, sheet.getRow(1).getHeight() );
			assertEquals( 769, sheet.getRow(2).getHeight() );
			assertEquals( 1394, sheet.getRow(3).getHeight() );
			assertEquals( 3069, sheet.getRow(4).getHeight() );
			assertEquals( 4632, sheet.getRow(5).getHeight() );
			
		} finally {
			inputStream.close();
		}
	}

	@Test
	public void testRunReportXls() throws BirtException, IOException {

		InputStream inputStream = runAndRenderReport("AutoRowHeight.rptdesign", "xls");
		assertNotNull(inputStream);
		try {
			
			HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Auto RowHeight Report", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 6, this.firstNullRow(sheet));
			
			assertEquals( 309, sheet.getRow(0).getHeight() );
			assertEquals( 907, sheet.getRow(1).getHeight() );
			assertEquals( 769, sheet.getRow(2).getHeight() );
			assertEquals( 1394, sheet.getRow(3).getHeight() );
			assertEquals( 3069, sheet.getRow(4).getHeight() );
			assertEquals( 4632, sheet.getRow(5).getHeight() );

		} finally {
			inputStream.close();
		}
	}
}
