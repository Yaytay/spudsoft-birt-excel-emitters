package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class Issue36XlsColours extends ReportRunner {
	
	private void dumpPalette( HSSFWorkbook workbook, String sheetName ) {
		Sheet sheet = workbook.createSheet( sheetName );

	    HSSFPalette palette = workbook.getCustomPalette();
	    System.out.println( "Palette: " + palette.toString() );
	    
	    for( short i = 0; i < 100; ++i) {
		    Row row = sheet.createRow(i);
		    CellStyle style = workbook.createCellStyle();
		    style.setFillForegroundColor( i );
		    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		    
		    Cell cell = row.createCell(0);
		    cell.setCellValue( i );
		    cell.setCellStyle(style);
		    
		    cell = row.createCell(1);
		    cell.setCellValue( i );

		    cell = row.createCell(2);
		    cell.setCellValue( palette.getColor(i) == null ? "null" : palette.getColor(i).getHexString() );
	    }
	}
	
	@Test
	public void generateSpectrum() throws Exception {
		
		HSSFWorkbook workbook = new HSSFWorkbook();
		
		dumpPalette(workbook, "default");
		
	    HSSFPalette palette = workbook.getCustomPalette();
	    try {
	    	palette.addColor( (byte)111, (byte)123, (byte)146);
	    } catch( RuntimeException ex ) {
	    	System.err.println( ex.toString() );
	    }
		
		
		File outputFile = createTempFile( "Issue36", ".XLS" );
		System.err.println( outputFile );
		FileOutputStream stream = new FileOutputStream(outputFile);
		try {
			workbook.write(stream);
		} finally {
			stream.close();
		}
	}
	
	@Test
	public void testColourMatching() throws BirtException, IOException {

		debug = false;
		InputStream inputStreamX = runAndRenderReport("issue_35_36_37_38.rptdesign", "xlsx");
		assertNotNull(inputStreamX);
		try {
			XSSFWorkbook workbookX = new XSSFWorkbook(inputStreamX);
			assertNotNull(workbookX);
			
			assertEquals( 1, workbookX.getNumberOfSheets() );
			XSSFSheet sheetX = workbookX.getSheetAt(0); 
	
			InputStream inputStreamS = runAndRenderReport("issue_35_36_37_38.rptdesign", "xls");
			assertNotNull(inputStreamS);
			try {
				HSSFWorkbook workbookS = new HSSFWorkbook(inputStreamS);
				assertNotNull(workbookS);
				
				assertEquals( workbookX.getNumberOfSheets(), workbookS.getNumberOfSheets() );
				HSSFSheet sheetS = workbookS.getSheetAt(0); 

				assertEquals( firstNullRow(sheetX), firstNullRow( sheetS ) );
				
				for( int rownum = 0; rownum < firstNullRow(sheetX); ++rownum ) {
					XSSFRow rowX = sheetX.getRow(rownum);
					HSSFRow rowS = sheetS.getRow(rownum);
					
					assertEquals( rowX.getLastCellNum(), rowS.getLastCellNum() );
					
					for( int colnum = 0; colnum < rowX.getLastCellNum(); ++colnum ) {
						XSSFCell cellX = rowX.getCell(colnum);
						HSSFCell cellS = rowS.getCell(colnum);
						
						if( cellX != null ) {
							assertNotNull( cellS );
							
							XSSFCellStyle styleX = cellX.getCellStyle();
							HSSFCellStyle styleS = cellS.getCellStyle();
							
							if( styleX != null ) {
								assertNotNull( styleX );
								
								XSSFColor colorX = styleX.getFillForegroundColorColor();
								HSSFColor colorS = styleS.getFillForegroundColorColor();
								
								if( colorX != null ) {
									assertNotNull( colorS );
									byte[] rgbX = colorX.getARgb();
									short[] rgbS = colorS.getTriplet();
									
									System.out.println( "Comparing " + colorX.getARGBHex() + " with " + colorS.getHexString() );
									
									assertEquals( rgbX[ 1 ], (byte)rgbS[ 0 ] );
									assertEquals( rgbX[ 2 ], (byte)rgbS[ 1 ] );
									assertEquals( rgbX[ 3 ], (byte)rgbS[ 2 ] );
								} else {
									if( colorS != null ) {
										if( ! "0:0:0".equals( colorS.getHexString() ) ) {
											assertNull( colorS.getHexString(), colorS );
										}
									}
								}
								
							} else {
								assertNull( styleS );
							}
							
						} else {
							assertNull( cellS );
						}
					}
					
				}
				
			} finally {
				inputStreamS.close();
			}
		} finally {
			inputStreamX.close();
		}
	}
	

}
