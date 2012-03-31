package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.*;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class Issue50MultiRowCrosstabHeaderGrids extends ReportRunner {

	@Test
	public void testHeader() throws Exception {
		
		debug = false;
		InputStream inputStream = runAndRenderReport("Issue50MultiRowCrosstabHeaderGrids.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( "Atelier graphique", sheet.getRow(2).getCell(1).getStringCellValue() );
			assertTrue( mergedRegion( sheet, 0, 0, 1, 0 ) );
			assertTrue( mergedRegion( sheet, 0, 1, 1, 1 ) );
			assertEquals( 34, sheet.getNumMergedRegions() );
			
			assertEquals( 100, this.firstNullRow(workbook.getSheetAt(0)));			
		} finally {
			inputStream.close();
		}
		
	}
}
