package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class Issue29 extends ReportRunner {
	
	@Test
	public void testMultiRowEmptinessXlsx() throws BirtException, IOException {

		debug = false;
		InputStream inputStream = runAndRenderReport("Issue29.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
	
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 4, this.firstNullRow(sheet));

			for( int i = 0; i < 4; ++i ) {
				for( Cell cell : sheet.getRow(i) ) {
					assertEquals( 0, cell.getCellStyle().getBorderTop() );
					assertEquals( 0, cell.getCellStyle().getBorderLeft() );
					assertEquals( 0, cell.getCellStyle().getBorderRight() );
					assertEquals( 0, cell.getCellStyle().getBorderBottom() );
				}
			}
		
		} finally {
			inputStream.close();
		}
	}
	
	@Test
	public void testMultiRowEmptinessXls() throws BirtException, IOException {

		debug = false;
		InputStream inputStream = runAndRenderReport("Issue29.rptdesign", "xls");
		assertNotNull(inputStream);
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
	
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 4, this.firstNullRow(sheet));
			
			for( int i = 0; i < 4; ++i ) {
				for( Cell cell : sheet.getRow(i) ) {
					assertEquals( 0, cell.getCellStyle().getBorderTop() );
					assertEquals( 0, cell.getCellStyle().getBorderLeft() );
					assertEquals( 0, cell.getCellStyle().getBorderRight() );
					assertEquals( 0, cell.getCellStyle().getBorderBottom() );
				}
			}
		
		} finally {
			inputStream.close();
		}
	}
	

}
