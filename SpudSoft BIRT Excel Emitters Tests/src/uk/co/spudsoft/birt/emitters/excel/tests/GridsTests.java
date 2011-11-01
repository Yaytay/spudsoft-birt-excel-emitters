package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class GridsTests extends ReportRunner {

	@Test
	public void testRunReportXlsx() throws BirtException, IOException {

		InputStream inputStream = runAndRenderReport("CombinedGrid.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Combined Grid Report", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 1, this.firstNullRow(sheet));
			
			DataFormatter formatter = new DataFormatter();
			
			assertEquals( "1",                     formatter.formatCellValue(sheet.getRow(1).getCell(1)));
			
		} finally {
			inputStream.close();
		}
	}

	@Test
	public void testRunReportXls() throws BirtException, IOException {

		InputStream inputStream = runAndRenderReport("CombinedGrid.rptdesign", "xls");
		assertNotNull(inputStream);
		try {
			
			HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Combined Grid Report", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 1, this.firstNullRow(sheet));
			
			DataFormatter formatter = new DataFormatter();
			
			assertEquals( "1",                     formatter.formatCellValue(sheet.getRow(1).getCell(1)));
			
		} finally {
			inputStream.close();
		}
	}
}
