package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class BigCrosstabTest extends ReportRunner {

	@Test
	public void testXlsx() throws Exception {

		InputStream inputStream = runAndRenderReport("BigCrosstab.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( 12, workbook.getNumCellStyles() );
			assertEquals( "Big Crosstab Report 1", workbook.getSheetAt(0).getSheetName());
			
			assertEquals( 60, workbook.getSheetAt(0).getRow(1).getCell(2).getCellStyle().getRotation());
			assertEquals( 60, workbook.getSheetAt(0).getRow(2).getCell(2).getCellStyle().getRotation());
			assertEquals( 60, workbook.getSheetAt(0).getRow(2).getCell(3).getCellStyle().getRotation());
			assertEquals(  0, workbook.getSheetAt(0).getRow(3).getCell(2).getCellStyle().getRotation());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 236, firstNullRow(sheet));
			
			assertEquals(28, greatestNumColumns(sheet));
			
		} finally {
			inputStream.close();
		}
	}

	@Test
	public void testXls() throws Exception {

		InputStream inputStream = runAndRenderReport("BigCrosstab.rptdesign", "xls");
		assertNotNull(inputStream);
		try {
			
			HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( 32, workbook.getNumCellStyles() );
			assertEquals( "Big Crosstab Report 1", workbook.getSheetAt(0).getSheetName());
			
			assertEquals( 60, workbook.getSheetAt(0).getRow(1).getCell(2).getCellStyle().getRotation());
			assertEquals( 60, workbook.getSheetAt(0).getRow(2).getCell(2).getCellStyle().getRotation());
			assertEquals( 60, workbook.getSheetAt(0).getRow(2).getCell(3).getCellStyle().getRotation());
			assertEquals(  0, workbook.getSheetAt(0).getRow(3).getCell(2).getCellStyle().getRotation());

			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 236, firstNullRow(sheet));
			
			assertEquals(28, greatestNumColumns(sheet));
			
		} finally {
			inputStream.close();
		}
	}
}
