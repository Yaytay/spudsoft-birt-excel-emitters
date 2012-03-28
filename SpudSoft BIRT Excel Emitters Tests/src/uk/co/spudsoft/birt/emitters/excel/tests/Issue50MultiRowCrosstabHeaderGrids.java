package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.*;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class Issue50MultiRowCrosstabHeaderGrids extends ReportRunner {

	public boolean mergedRegion( Sheet sheet, int top, int left, int bottom, int right ) {
		for( int i = 0; i < sheet.getNumMergedRegions(); ++i ) {
			CellRangeAddress curRegion = sheet.getMergedRegion(i);
			if( ( curRegion.getFirstRow() == top )
					&& ( curRegion.getFirstColumn() == left )
					&& ( curRegion.getLastRow() == bottom )
					&& ( curRegion.getLastColumn() == right ) ) {
				return true;
			}
		}
		return false;
	}
	
	@Test
	public void testHeader() throws Exception {
		
		debug = true;
		InputStream inputStream = runAndRenderReport("Issue50MultiRowCrosstabHeaderGrids.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( "Atelier graphique", sheet.getRow(2).getCell(1).getStringCellValue() );
			assertTrue( mergedRegion( sheet, 0, 0, 1, 0 ) );
			assertEquals( 33, sheet.getNumMergedRegions() );
	
			assertEquals( 100, this.firstNullRow(workbook.getSheetAt(0)));			
		} finally {
			inputStream.close();
		}
		
	}
}
