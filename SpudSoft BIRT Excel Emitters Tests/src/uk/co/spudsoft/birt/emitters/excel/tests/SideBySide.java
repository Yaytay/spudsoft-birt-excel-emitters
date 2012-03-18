package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertNotNull;

import java.io.InputStream;

import org.junit.Test;

public class SideBySide extends ReportRunner {

	@Test
	public void singleCells() throws Exception {
		debug = false;
		InputStream inputStream = runAndRenderReport("SideBySideOneCellEach.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			
		} finally {
			inputStream.close();
		}
		
	}

	@Test
	public void multiColumns() throws Exception {
		debug = false;
		InputStream inputStream = runAndRenderReport("SideBySideMultiColumns.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			
		} finally {
			inputStream.close();
		}
		
	}
}
