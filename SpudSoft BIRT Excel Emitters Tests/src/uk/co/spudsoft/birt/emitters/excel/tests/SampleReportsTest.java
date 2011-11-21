package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertNotNull;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class SampleReportsTest extends ReportRunner {
	
	/*
	 * I really want these unit tests to cover the sample reports provided with BIRT,
	 * but I cannot package those reports up with my source due to licensing conflicts.
	 * If the BIRT source is available change this path to point to the root of the sample reports,
	 * otherwise set it to null and all these tests will be ignored.
	 */
	private static String basePath = "C:\\Temp\\birt-source-3_7_1\\plugins\\org.eclipse.birt.report.designer.samplereports\\";
	
	private InputStream runAndRenderSampleReport( String filename, String extension ) throws IOException, BirtException {
		if( basePath != null ) {
			File file = new File( basePath + filename );
			if( file.exists() ) {
				return runAndRenderReport( file.getAbsolutePath(), extension );
			}
		}
		return null;
	}
	
	@Test
	public void productCatalogTest() throws BirtException, IOException {
		
		InputStream inputStream = null;
		if( ( inputStream = runAndRenderSampleReport( "samplereports/Solution Reports/Listing/ProductCatalog.rptdesign", "xlsx") ) != null ) {
			try {
				assertNotNull(inputStream);
			} finally {
				inputStream.close();
			}
		}
	}
	
	@Test
	public void cascadeTest() throws BirtException, IOException {
		
		InputStream inputStream = null;
		parameters.put( "customer", 103);
		parameters.put( "order", 10123);
		if( ( inputStream = runAndRenderSampleReport( "samplereports/Reporting Feature Examples/Cascaded Parameter Report/cascade.rptdesign", "xlsx") ) != null ) {
			try {
				assertNotNull(inputStream);
			} finally {
				inputStream.close();
			}
		}
	}

/*	@Test
	public void customerOrdersFinal() throws BirtException, IOException {
		
		InputStream inputStream = null;
		parameters.put( "customer", 103);
		parameters.put( "order", 10123);
		if( ( inputStream = runAndRenderSampleReport( "samplereports/Reporting Feature Examples/Combination Chart/CustomerOrdersFinal.rptdesign", "xlsx") ) != null ) {
			try {
				assertNotNull(inputStream);
			} finally {
				inputStream.close();
			}
		}
	}
*/
	@Test
	public void crosstabSampleRevenue() throws BirtException, IOException {
		
		InputStream inputStream = null;
		parameters.put( "customer", 103);
		parameters.put( "order", 10123);
		if( ( inputStream = runAndRenderSampleReport( "samplereports/Reporting Feature Examples/Cross tab/CrosstabSampleRevenue.rptdesign", "xlsx") ) != null ) {
			try {
				assertNotNull(inputStream);
			} finally {
				inputStream.close();
			}
		}
	}
}
