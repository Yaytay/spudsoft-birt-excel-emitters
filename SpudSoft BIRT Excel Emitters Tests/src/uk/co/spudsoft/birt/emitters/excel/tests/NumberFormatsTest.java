package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class NumberFormatsTest extends ReportRunner {
	
	@Test
	public void testRunReport() throws BirtException, IOException {

		InputStream inputStream = runAndRenderReport("NumberFormats.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			assertEquals( "Number Formats Test Report", workbook.getSheetAt(0).getSheetName());
			
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals(18, this.firstNullRow(sheet));
			
			assertEquals( 3035,                    sheet.getColumnWidth( 0 ) );
			assertEquals( 3913,                    sheet.getColumnWidth( 1 ) );
			assertEquals( 7021,                    sheet.getColumnWidth( 2 ) );
			assertEquals( 4205,                    sheet.getColumnWidth( 3 ) );
			assertEquals( 3474,                    sheet.getColumnWidth( 4 ) );
			assertEquals( 2852,                    sheet.getColumnWidth( 5 ) );
			assertEquals( 3510,                    sheet.getColumnWidth( 6 ) );
			assertEquals( 2889,                    sheet.getColumnWidth( 7 ) );
			assertEquals( 2048,                    sheet.getColumnWidth( 8 ) );
						
			DataFormatter formatter = new DataFormatter();
			
			assertEquals( "1",                     formatter.formatCellValue(sheet.getRow(1).getCell(1)));
			assertEquals( "2019-10-11 13:18:46",   formatter.formatCellValue(sheet.getRow(1).getCell(2)));
			assertEquals( "3.1415926536",          formatter.formatCellValue(sheet.getRow(1).getCell(3)));
			assertEquals( "3.1415926536",          formatter.formatCellValue(sheet.getRow(1).getCell(4)));
			assertEquals( "false",                 formatter.formatCellValue(sheet.getRow(1).getCell(5)));
			assertEquals( "Fri, 11 Oct 2019",      formatter.formatCellValue(sheet.getRow(1).getCell(6)));
			assertEquals( "13:18",                 formatter.formatCellValue(sheet.getRow(1).getCell(7)));

			assertEquals( "2",                     formatter.formatCellValue(sheet.getRow(2).getCell(1)));
			assertEquals( "2019-10-11 13:18:46",   formatter.formatCellValue(sheet.getRow(2).getCell(2)));
			assertEquals( "6.2831853072",          formatter.formatCellValue(sheet.getRow(2).getCell(3)));
			assertEquals( "6.2831853072",          formatter.formatCellValue(sheet.getRow(2).getCell(4)));
			assertEquals( "true",                  formatter.formatCellValue(sheet.getRow(2).getCell(5)));
			assertEquals( "Fri, 11 Oct 2019",      formatter.formatCellValue(sheet.getRow(2).getCell(6)));
			assertEquals( "13:18",                 formatter.formatCellValue(sheet.getRow(2).getCell(7)));

			assertEquals( "3.1415926536",          formatter.formatCellValue(sheet.getRow(5).getCell(1)));
			assertEquals( "3.1415926536",          formatter.formatCellValue(sheet.getRow(5).getCell(2)));
			assertEquals( "£3.14",                 formatter.formatCellValue(sheet.getRow(5).getCell(3)));
			assertEquals( "3.14",                  formatter.formatCellValue(sheet.getRow(5).getCell(4)));
			assertEquals( "314.16%",               formatter.formatCellValue(sheet.getRow(5).getCell(5)));
			assertEquals( "3.14E00",               formatter.formatCellValue(sheet.getRow(5).getCell(6)));
			assertEquals( "3.14E00",               formatter.formatCellValue(sheet.getRow(5).getCell(7)));
			
			assertEquals( "6.2831853072",          formatter.formatCellValue(sheet.getRow(6).getCell(1)));
			assertEquals( "6.2831853072",          formatter.formatCellValue(sheet.getRow(6).getCell(2)));
			assertEquals( "£6.28",                 formatter.formatCellValue(sheet.getRow(6).getCell(3)));
			assertEquals( "6.28",                  formatter.formatCellValue(sheet.getRow(6).getCell(4)));
			assertEquals( "628.32%",               formatter.formatCellValue(sheet.getRow(6).getCell(5)));
			assertEquals( "6.28E00",               formatter.formatCellValue(sheet.getRow(6).getCell(6)));
			assertEquals( "6.28E00",               formatter.formatCellValue(sheet.getRow(6).getCell(7)));
			
			assertEquals( "1",                     formatter.formatCellValue(sheet.getRow(9).getCell(1)));
			assertEquals( "11/10/2019 13:18",      formatter.formatCellValue(sheet.getRow(9).getCell(2)));
			assertEquals( "3.1415926536",          formatter.formatCellValue(sheet.getRow(9).getCell(3)));
			assertEquals( "3.1415926536",          formatter.formatCellValue(sheet.getRow(9).getCell(4)));
			assertEquals( "false",                 formatter.formatCellValue(sheet.getRow(9).getCell(5)));
			assertEquals( "2019-10-11",            formatter.formatCellValue(sheet.getRow(9).getCell(6)));
			assertEquals( "13:18",                 formatter.formatCellValue(sheet.getRow(9).getCell(7)));
			
			assertEquals( "2",                     formatter.formatCellValue(sheet.getRow(10).getCell(1)));
			assertEquals( "11/10/2019 13:18",      formatter.formatCellValue(sheet.getRow(10).getCell(2)));
			assertEquals( "6.2831853072",          formatter.formatCellValue(sheet.getRow(10).getCell(3)));
			assertEquals( "6.2831853072",          formatter.formatCellValue(sheet.getRow(10).getCell(4)));
			assertEquals( "true",                  formatter.formatCellValue(sheet.getRow(10).getCell(5)));
			assertEquals( "2019-10-11",            formatter.formatCellValue(sheet.getRow(10).getCell(6)));
			assertEquals( "13:18",                 formatter.formatCellValue(sheet.getRow(10).getCell(7)));
			
			assertEquals( "MSRP $3.14",            formatter.formatCellValue(sheet.getRow(15).getCell(1)));
			
		} finally {
			inputStream.close();
		}
	}
	
}
