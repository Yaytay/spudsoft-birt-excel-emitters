/********************************************************************************
* (C) Copyright 2011, by James Talbut.
*
*   This program is free software: you can redistribute it and/or modify
*   it under the terms of the GNU General Public License as published by
*   the Free Software Foundation, either version 3 of the License, or
*   (at your option) any later version.
*
*   This program is distributed in the hope that it will be useful,
*   but WITHOUT ANY WARRANTY; without even the implied warranty of
*   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
*   GNU General Public License for more details.
*
*   You should have received a copy of the GNU General Public License
*   along with this program.  If not, see <http://www.gnu.org/licenses/>.
*
*   [Java is a trademark or registered trademark of Sun Microsystems, Inc.
*   in the United States and other countries.]
********************************************************************************/

package uk.co.spudsoft.birt.emitters.excel.tests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Sheet;
import org.eclipse.birt.core.archive.FileArchiveWriter;
import org.eclipse.birt.core.archive.IDocArchiveWriter;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.core.framework.Platform;
import org.eclipse.birt.report.engine.api.EngineConfig;
import org.eclipse.birt.report.engine.api.IRenderTask;
import org.eclipse.birt.report.engine.api.IReportDocument;
import org.eclipse.birt.report.engine.api.IReportEngine;
import org.eclipse.birt.report.engine.api.IReportEngineFactory;
import org.eclipse.birt.report.engine.api.IReportRunnable;
import org.eclipse.birt.report.engine.api.IRunTask;
import org.eclipse.birt.report.engine.api.RenderOption;

public class ReportRunner {
	
	private static byte[] getBytesFromFile(File file) throws IOException {
	    InputStream is = new FileInputStream(file);
	    try {
		    byte[] data = new byte[(int)file.length()];
		    int offset = 0;
		    int read = 0;
		    while (offset < data.length && (read=is.read(data, offset, data.length-offset)) >= 0) {
		        offset += read;
		    }
		    return data;
	    } finally {
	    	is.close();
	    }
	}
	
	protected int firstNullRow(Sheet sheet) {
		int i = 0;
		while( sheet.getRow(i) != null ) {
			++i;
		}
		return i;
	}

	protected InputStream runAndRenderReport( String filename, String outputFormat ) throws BirtException, IOException {

        EngineConfig config = new EngineConfig();

        IReportEngineFactory engineFactory = (IReportEngineFactory)Platform.createFactoryObject( IReportEngineFactory.EXTENSION_REPORT_ENGINE_FACTORY );
		assertNotNull(engineFactory);
		
		IReportEngine reportEngine = engineFactory.createReportEngine( config );
		assertNotNull(reportEngine);
		
		InputStream resourceStream = this.getClass().getResourceAsStream( filename );
		assertNotNull( resourceStream );
		try {
			IReportRunnable reportRunnable = reportEngine.openReportDesign( resourceStream );
			assertNotNull(reportRunnable);
			
			File tempDoc = File.createTempFile("runAndRenderReport", ".rptdocument");
			assertNotNull(tempDoc);
			
			try {
				IRunTask reportRunTask = reportEngine.createRunTask( reportRunnable );
				assertNotNull(reportRunTask);
				try {
					IDocArchiveWriter archiveWriter = new FileArchiveWriter( tempDoc.getCanonicalPath() );
					assertNotNull(archiveWriter);
	
					reportRunTask.run( archiveWriter );
			        assertEquals( 0, reportRunTask.getErrors().size() );
			        
			        reportRunTask.close();
			        
			        IReportDocument reportDocument = reportEngine.openReportDocument( tempDoc.getCanonicalPath() );
			        assertNotNull(reportDocument);
			        
			        IRenderTask renderTask = reportEngine.createRenderTask( reportDocument );
			        assertNotNull(renderTask);
			        try {
	
			        	File tempOutput = File.createTempFile("Render", "." + outputFormat);
			        	System.err.println( tempOutput );
			        	FileOutputStream outputStream = new FileOutputStream( tempOutput ); 
			        	
				        assertNotNull(outputStream);
				        try {
					        
					        RenderOption renderOptions = new RenderOption();
					        renderOptions.setOutputFormat(outputFormat);
					        renderOptions.setOutputStream( outputStream );
					        renderTask.setRenderOption(renderOptions);
					        
					        renderTask.render();
					        assertEquals(0, renderTask.getErrors().size());					        
				        } finally {
				        	outputStream.close();
				        }

				        return new ByteArrayInputStream(getBytesFromFile(tempOutput));
			        } finally {
				        renderTask.close();
			        }
				} finally {
					reportRunTask.close();
				}
			} finally {
				tempDoc.delete();
			}
		} finally {
			resourceStream.close();
		}
	}
}
