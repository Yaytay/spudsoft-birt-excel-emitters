package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Date;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

import uk.co.spudsoft.birt.emitters.excel.DateFormatConverter;

public class Issue38DateFormat extends ReportRunner {

	private class LocaleCompare implements Comparator<Locale> {

		@Override
		public int compare(Locale o1, Locale o2) {
			return o1.toString().compareTo(o2.toString());
		}		
	}
	
	private String pad(String arg, int length) {
		while( arg.length() < length ) {
			arg = arg + " ";
		}
		return arg;
	}
	
	@Test
	public void testLocaleFormats() {
		
		Date now = new Date();
		
		Locale[] locales = DateFormat.getAvailableLocales();
		Arrays.sort(locales, new LocaleCompare());
		for( Locale locale : locales ) {
			DateFormat defaultDate = DateFormat.getDateInstance(DateFormat.DEFAULT, locale); 
			DateFormat shortDate = DateFormat.getDateInstance(DateFormat.SHORT, locale); 
			DateFormat mediumDate = DateFormat.getDateInstance(DateFormat.MEDIUM, locale); 
			DateFormat longDate = DateFormat.getDateInstance(DateFormat.LONG, locale);
			
			System.out.println( pad( locale.toString(), 16 ) 
					+ pad( ((SimpleDateFormat)defaultDate).toLocalizedPattern(), 16 ) 
						+ pad( defaultDate.format(now), 16 )
					+ pad( ((SimpleDateFormat)shortDate).toLocalizedPattern(), 16 ) 
						+ pad( shortDate.format(now), 16 )
					+ pad( ((SimpleDateFormat)mediumDate).toLocalizedPattern(), 24 ) 
						+ pad( mediumDate.format(now), 24 )
					+ pad( ((SimpleDateFormat)longDate).toLocalizedPattern(), 32 ) 
						+ pad( longDate.format(now), 32 ) 
					);
			
		}
	}	
	
	private void outputLocaleDataFormats( Date date, boolean dates, boolean times, int style, String styleName ) throws Exception {

		XSSFWorkbook workbook = new XSSFWorkbook();
		String sheetName;
		if( dates ) {
			if( times ) {
				sheetName = "DateTimes";
			} else {
				sheetName = "Dates";
			}
		} else {
			sheetName = "Times";
		}
		Sheet sheet = workbook.createSheet(sheetName);
		Row header = sheet.createRow(0);
		header.createCell(0).setCellValue("locale");
		header.createCell(0).setCellValue("DisplayName");
		header.createCell(1).setCellValue(styleName);
		header.createCell(3).setCellValue("Java");
		header.createCell(5).setCellValue("Excel");

		int rowNum = 1;
		for( Locale locale : DateFormat.getAvailableLocales() ) {
			Row row = sheet.createRow(rowNum++);
			
			row.createCell(0).setCellValue(locale.toString());
			row.createCell(1).setCellValue(locale.getDisplayName());
						
			DateFormat dateFormat;
			if( dates ) {
				if( times ) {
					dateFormat = DateFormat.getDateTimeInstance(style, style, locale); 
				} else {
					dateFormat = DateFormat.getDateInstance(style, locale); 
				}
			} else {
				dateFormat = DateFormat.getTimeInstance(style, locale); 
			}
			
			Cell cell = row.createCell(2);
			cell.setCellValue(date);
			CellStyle cellStyle = row.getSheet().getWorkbook().createCellStyle();
			
			String javaDateFormatPattern = ((SimpleDateFormat)dateFormat).toPattern();
			String excelFormatPattern = DateFormatConverter.convert(locale, javaDateFormatPattern);
			
			DataFormat poiFormat = row.getSheet().getWorkbook().createDataFormat();
			cellStyle.setDataFormat(poiFormat.getFormat(excelFormatPattern));
			
			cell.setCellStyle(cellStyle);
			
			row.createCell(3).setCellValue(javaDateFormatPattern);
			row.createCell(4).setCellValue(excelFormatPattern);
		}
		
		File outputFile = createTempFile( "Locale" + sheetName + styleName , ".XLSX");
		FileOutputStream outputStream = new FileOutputStream(outputFile);
		try {
			workbook.write(outputStream);
		} finally {
			outputStream.close();
		}
	}
	
	@Test
	public void testJavaDateFormatsInExcel() throws Exception {

		Date date = new Date();
		
		outputLocaleDataFormats(date, true, false, DateFormat.DEFAULT, "Default" );
		outputLocaleDataFormats(date, true, false, DateFormat.SHORT, "Short" );
		outputLocaleDataFormats(date, true, false, DateFormat.MEDIUM, "Medium" );
		outputLocaleDataFormats(date, true, false, DateFormat.LONG, "Long" );
		outputLocaleDataFormats(date, true, false, DateFormat.FULL, "Full" );
		
		outputLocaleDataFormats(date, true, true, DateFormat.DEFAULT, "Default" );
		outputLocaleDataFormats(date, true, true, DateFormat.SHORT, "Short" );
		outputLocaleDataFormats(date, true, true, DateFormat.MEDIUM, "Medium" );
		outputLocaleDataFormats(date, true, true, DateFormat.LONG, "Long" );
		outputLocaleDataFormats(date, true, true, DateFormat.FULL, "Full" );
		
		outputLocaleDataFormats(date, false, true, DateFormat.DEFAULT, "Default" );
		outputLocaleDataFormats(date, false, true, DateFormat.SHORT, "Short" );
		outputLocaleDataFormats(date, false, true, DateFormat.MEDIUM, "Medium" );
		outputLocaleDataFormats(date, false, true, DateFormat.LONG, "Long" );
		outputLocaleDataFormats(date, false, true, DateFormat.FULL, "Full" );
				
	}

	
	@Test
	public void outputDateFormatsReport() throws Exception {
		debug = false;
		InputStream inputStreamX = runAndRenderReport("DateFormats.rptdesign", "xlsx");
		assertNotNull(inputStreamX);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStreamX);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
			Sheet sheet = workbook.getSheetAt(0);
			
			assertEquals( "[$-1010409]MMMM d, yyyy h:mm:ss am/pm", sheet.getRow(2).getCell(1).getCellStyle().getDataFormatString() );
			assertEquals( "[$-1010409]MMMM d, yyyy h:mm:ss am/pm", sheet.getRow(2).getCell(2).getCellStyle().getDataFormatString() );
			assertEquals( "[$-1010409]MMMM d, yyyy",               sheet.getRow(2).getCell(3).getCellStyle().getDataFormatString() );
			assertEquals( "[$-1010409]MMM d, yyyy",                sheet.getRow(2).getCell(4).getCellStyle().getDataFormatString() );
			assertEquals( "[$-1010409]M/d/yy",                     sheet.getRow(2).getCell(5).getCellStyle().getDataFormatString() );
			assertEquals( "[$-1010409]h:mm:ss am/pm",              sheet.getRow(2).getCell(6).getCellStyle().getDataFormatString() );
			assertEquals( "[$-1010409]h:mm:ss am/pm",              sheet.getRow(2).getCell(7).getCellStyle().getDataFormatString() );
			assertEquals( "[$-1010409]hh:mm",                      sheet.getRow(2).getCell(8).getCellStyle().getDataFormatString() );
	
		} finally {
			inputStreamX.close();
		}
		
		System.out.println( "The default locale is " + Locale.getDefault().toString() );
	}
	
	@Test
	public void testWithBugReport() throws BirtException, IOException {

		debug = false;
		InputStream inputStream = runAndRenderReport("issue_35_36_37_38.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			assertNotNull(workbook);
			
			assertEquals( 1, workbook.getNumberOfSheets() );
	
			Sheet sheet = workbook.getSheetAt(0);
			assertEquals( 256, this.firstNullRow(sheet));

			assertEquals( 1326153600000L,     sheet.getRow(3).getCell(2).getDateCellValue().getTime() );
			assertEquals( "[$-1010409]MMM d, yyyy",      sheet.getRow(3).getCell(2).getCellStyle().getDataFormatString() );
			
		} finally {
			inputStream.close();
		}
	}
	
}
