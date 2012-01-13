package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertNotNull;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.api.RenderOption;
import org.junit.Test;
import static org.junit.Assert.*;

public class Issue27 extends ReportRunner {
	
	private Pattern pattern = Pattern.compile( "R(\\d+)C(\\d+)-R(\\d+)C(\\d+): .*" ); 
	
	private void validateCellRange( Matcher matcher, Cell cell ) {
		int desiredR1 = Integer.parseInt( matcher.group(1) );
		int desiredC1 = Integer.parseInt( matcher.group(2) );
		int desiredR2 = Integer.parseInt( matcher.group(3) );
		int desiredC2 = Integer.parseInt( matcher.group(4) );
		
		int actualR1 = cell.getRowIndex() + 1;
		int actualC1 = cell.getColumnIndex() + 1;
		int actualR2 = 0;
		int actualC2 = 0;
		
		for( int i = 0; i < cell.getSheet().getNumMergedRegions(); ++i) {
			CellRangeAddress cra = cell.getSheet().getMergedRegion(i);
			if( ( cra.getFirstRow() == cell.getRowIndex() ) && ( cra.getFirstColumn() == cell.getColumnIndex() ) ) {
				assertEquals( 0, actualR2 );
				assertEquals( 0, actualC2 );
				actualR2 = cra.getLastRow() + 1;
				actualC2 = cra.getLastColumn() + 1;
			}
		}
		assertEquals( desiredR1, actualR1 );
		assertEquals( desiredC1, actualC1 );
		assertEquals( desiredR2, actualR2 );
		assertEquals( desiredC2, actualC2 );
	}
	
	@Test
	public void testRowSpanXls() throws BirtException, IOException {

		debug = false;
		InputStream inputStream = runAndRenderReport("Issue27.rptdesign", "xls");
        assertNotNull(inputStream);
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            assertNotNull(workbook);
    
            Sheet sheet = workbook.getSheetAt(0);
            int rangesValidated = 0;
            
            for( Row row : sheet ) {
            	for( Cell cell : row ) {
            		if(cell.getCellType() == Cell.CELL_TYPE_STRING) {
            			String cellValue = cell.getStringCellValue();
            			Matcher matcher = pattern.matcher(cellValue);
            			if( matcher.matches() ) {
            				validateCellRange( matcher, cell );
            				++rangesValidated;
            			}
            		}
            	}
            }
            assertEquals( 12, rangesValidated );
        
        } finally {
            inputStream.close();
        }
	}
	    
	@Test
	public void testRowSpanXlsx() throws BirtException, IOException {

		debug = false;
		InputStream inputStream = runAndRenderReport("Issue27.rptdesign", "xlsx");
        assertNotNull(inputStream);
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            assertNotNull(workbook);
    
            Sheet sheet = workbook.getSheetAt(0);
            int rangesValidated = 0;
            
            for( Row row : sheet ) {
            	for( Cell cell : row ) {
            		if(cell.getCellType() == Cell.CELL_TYPE_STRING) {
            			String cellValue = cell.getStringCellValue();
            			
            			Matcher matcher = pattern.matcher(cellValue);
            			if( matcher.matches() ) {
            				validateCellRange( matcher, cell );
            				++rangesValidated;
            			}
            		}
            	}
            }
            assertEquals( 12, rangesValidated );
        
        } finally {
            inputStream.close();
        }
	}
	    
    protected RenderOption prepareRenderOptions(String outputFormat, FileOutputStream outputStream) {
        RenderOption option = super.prepareRenderOptions(outputFormat, outputStream);
        option.setOption("ExcelEmitter.RemoveBlankRows", Boolean.FALSE);
        return option;
    }

}
