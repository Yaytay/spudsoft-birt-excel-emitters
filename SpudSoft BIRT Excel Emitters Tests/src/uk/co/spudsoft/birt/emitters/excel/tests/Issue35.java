package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class Issue35 extends ReportRunner {
	
	@Test
	public void testIssue35() throws BirtException, IOException {

		debug = false;
		InputStream inputStream = runAndRenderReport("issue_35_36_37_38.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
	
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 25, this.firstNullRow(sheet));

			assertEquals( CellStyle.VERTICAL_TOP, sheet.getRow(3).getCell(0).getCellStyle().getVerticalAlignment() );
			assertEquals( CellStyle.VERTICAL_TOP, sheet.getRow(3).getCell(1).getCellStyle().getVerticalAlignment() );
			assertEquals( CellStyle.VERTICAL_TOP, sheet.getRow(3).getCell(2).getCellStyle().getVerticalAlignment() );
		
		} finally {
			inputStream.close();
		}
	}
	

}
