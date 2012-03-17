package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class Issue43StructuredHeader extends ReportRunner {

	@Test
	public void testWithoutOption() throws Exception {
		
		debug = false;
		InputStream inputStream = runAndRenderReport("StructuredHeader.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 3, workbook.getNumberOfSheets() );
	
			assertEquals( 4, this.firstNullRow(workbook.getSheetAt(0)));
			assertEquals( 4, this.firstNullRow(workbook.getSheetAt(1)));
			assertEquals( 4, this.firstNullRow(workbook.getSheetAt(2)));
			
		} finally {
			inputStream.close();
		}
	}

	@Test
	public void testWithOption() throws Exception {
		
		debug = false;
		structuredHeader = true;
		InputStream inputStream = runAndRenderReport("StructuredHeader.rptdesign", "xlsx");
		structuredHeader = false;
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 3, workbook.getNumberOfSheets() );
	
			assertEquals( 8, this.firstNullRow(workbook.getSheetAt(0)));
			assertEquals( 8, this.firstNullRow(workbook.getSheetAt(1)));
			assertEquals( 8, this.firstNullRow(workbook.getSheetAt(2)));
			
		} finally {
			inputStream.close();
		}
	}

	@Test
	public void testWithUserProperty() throws Exception {
		
		debug = false;
		structuredHeader = false;
		InputStream inputStream = runAndRenderReport("StructuredHeaderWithUserProperty.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 3, workbook.getNumberOfSheets() );
	
			assertEquals( 8, this.firstNullRow(workbook.getSheetAt(0)));
			assertEquals( 8, this.firstNullRow(workbook.getSheetAt(1)));
			assertEquals( 8, this.firstNullRow(workbook.getSheetAt(2)));
			
		} finally {
			inputStream.close();
		}
	}
}
