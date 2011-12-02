package uk.co.spudsoft.birt.emitters.excel.tests;

import java.io.IOException;

import org.eclipse.birt.core.exception.BirtException;
import org.junit.Test;

public class MegaSizeTest extends ReportRunner {
	
	@Test
	public void testMegaXlsx() throws BirtException, IOException {

		debug = true;
/*		InputStream inputStream = runAndRenderReport("MegaSize.rptdesign", "xlsx");
		assertNotNull(inputStream);
		try {
			
			
		} finally {
			inputStream.close();
		}
	}
	
	@Test
	public void testMegaXls() throws BirtException, IOException {

		try {
			runAndRenderReport("MegaSize.rptdesign", "xls");
			fail( "Should have failed!" );
		} catch( IllegalArgumentException ex ) {
			assertEquals( "Invalid column index (256)", ex.getMessage().substring(0, 26));
		}
*/	}
	
}
