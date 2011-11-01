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

import static org.junit.Assert.*;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.eclipse.birt.core.archive.FileArchiveWriter;
import org.eclipse.birt.core.archive.IDocArchiveWriter;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.core.framework.Platform;
import org.eclipse.birt.report.engine.api.EngineConfig;
import org.eclipse.birt.report.engine.api.EngineException;
import org.eclipse.birt.report.engine.api.IEngineTask;
import org.eclipse.birt.report.engine.api.IRenderTask;
import org.eclipse.birt.report.engine.api.IReportDocument;
import org.eclipse.birt.report.engine.api.IReportEngine;
import org.eclipse.birt.report.engine.api.IReportEngineFactory;
import org.eclipse.birt.report.engine.api.IReportRunnable;
import org.eclipse.birt.report.engine.api.IRunAndRenderTask;
import org.eclipse.birt.report.engine.api.IRunTask;
import org.eclipse.birt.report.engine.api.RenderOption;
import org.eclipse.birt.report.engine.api.impl.ReportEngine;

import uk.co.spudsoft.birt.emitters.bugfix.FixedRenderTask;
import uk.co.spudsoft.birt.emitters.excel.tests.framework.Activator;

public class ReportRunner {
	
	protected boolean debug;
	
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
	
	protected int greatestNumColumns(Sheet sheet) {
		int result = 0;
		for(Row row : sheet) {
			if(row.getLastCellNum() > result) {
				result = row.getLastCellNum();
			}
		}
		return result;
	}

	protected InputStream runAndRenderReport( String filename, String outputFormat ) throws BirtException, IOException {
		return runAndRenderReportCustomTask(filename, outputFormat);
	}

	protected InputStream runAndRenderReportDefaultTask( String filename, String outputFormat ) throws BirtException, IOException {

        IReportEngine reportEngine = createReportEngine();
		
		String filepath = deriveFilepath(filename);

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
					addFilepathToAppContext(filepath, reportRunTask);
					
					IReportDocument reportDocument = runReport(reportEngine,
							reportRunTask, tempDoc);
			        
			        IRenderTask renderTask = reportEngine.createRenderTask( reportDocument );
			        assertNotNull(renderTask);
			        try {
			        	File tempOutput = File.createTempFile("Render", "." + outputFormat);
			        	System.err.println( tempOutput );
			        	FileOutputStream outputStream = new FileOutputStream( tempOutput ); 
			        	
				        assertNotNull(outputStream);
				        try {
					        renderTask.setRenderOption(prepareRenderOptions( outputFormat, outputStream ));
					        
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

	protected InputStream runAndRenderReportFileNotStream( String filename, String outputFormat ) throws BirtException, IOException {

        IReportEngine reportEngine = createReportEngine();
		
		String filepath = deriveFilepath(filename);

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
					addFilepathToAppContext(filepath, reportRunTask);
					
					IReportDocument reportDocument = runReport(reportEngine,
							reportRunTask, tempDoc);
			        
			        IRenderTask renderTask = reportEngine.createRenderTask( reportDocument );
			        assertNotNull(renderTask);
			        try {
			        	File outputFile = File.createTempFile("Render", "." + outputFormat);
			        	System.err.println( outputFile );
			        	
				        assertNotNull( outputFile );
				        renderTask.setRenderOption(prepareRenderOptions( outputFormat, null ));
				        renderTask.getRenderOption().setOutputFileName( outputFile.getCanonicalPath() );
				        
				        renderTask.render();
				        assertEquals(0, renderTask.getErrors().size());					        

				        InputStream result = new ByteArrayInputStream(getBytesFromFile(outputFile));
				        
				        boolean deleted = outputFile.delete();
				        assertTrue( deleted );
				        
				        return result;
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

	protected InputStream runAndRenderReportAsOne( String filename, String outputFormat ) throws BirtException, IOException {

        IReportEngine reportEngine = createReportEngine();
		
		String filepath = deriveFilepath(filename);

		InputStream resourceStream = this.getClass().getResourceAsStream( filename );
		
		assertNotNull( resourceStream );
		try {
			IReportRunnable reportRunnable = reportEngine.openReportDesign( resourceStream );
			assertNotNull(reportRunnable);
			
			File tempDoc = File.createTempFile("runAndRenderReport", ".rptdocument");
			assertNotNull(tempDoc);
			
			try {
				IRunAndRenderTask reportRunRenderTask = reportEngine.createRunAndRenderTask( reportRunnable );
				assertNotNull(reportRunRenderTask);
				try {
					addFilepathToAppContext(filepath, reportRunRenderTask);
					
		        	File tempOutput = File.createTempFile("Render", "." + outputFormat);
		        	System.err.println( tempOutput );
		        	FileOutputStream outputStream = new FileOutputStream( tempOutput ); 
		        	
			        assertNotNull(outputStream);
			        try {
				        
				        reportRunRenderTask.setRenderOption(prepareRenderOptions( outputFormat, outputStream ));
				        
				        reportRunRenderTask.run();
				        assertEquals(0, reportRunRenderTask.getErrors().size());					        
			        } finally {
			        	outputStream.close();
			        }

			        return new ByteArrayInputStream(getBytesFromFile(tempOutput));
				} finally {
					reportRunRenderTask.close();
				}
			} finally {
				tempDoc.delete();
			}
		} finally {
			resourceStream.close();
		}
	}

	protected InputStream runAndRenderReportCustomTask( String filename, String outputFormat ) throws BirtException, IOException {

        IReportEngine reportEngine = createReportEngine();
		
		String filepath = deriveFilepath(filename);

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
					addFilepathToAppContext(filepath, reportRunTask);
					
					IReportDocument reportDocument = runReport(reportEngine,
							reportRunTask, tempDoc);
			        
			        // IRenderTask renderTask = reportEngine.createRenderTask( reportDocument );
			        IRenderTask renderTask = new FixedRenderTask( (ReportEngine)reportEngine, reportRunnable, reportDocument );
			        assertNotNull(renderTask);
			        try {
			        	File tempOutput = File.createTempFile("Render", "." + outputFormat);
			        	System.err.println( tempOutput );
			        	FileOutputStream outputStream = new FileOutputStream( tempOutput ); 
			        	
				        assertNotNull(outputStream);
				        try {
					        
					        renderTask.setRenderOption(prepareRenderOptions( outputFormat, outputStream ));
					        
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

	protected IReportEngine createReportEngine() {
		EngineConfig config = new EngineConfig();

        IReportEngineFactory engineFactory = (IReportEngineFactory)Platform.createFactoryObject( IReportEngineFactory.EXTENSION_REPORT_ENGINE_FACTORY );
		assertNotNull(engineFactory);
		
		IReportEngine reportEngine = engineFactory.createReportEngine( config );
		assertNotNull(reportEngine);
		return reportEngine;
	}

	protected String deriveFilepath(String filename) throws MalformedURLException {
		String filepath = null;
		if( Activator.getContext() != null ) {
			URL bundleLocation = new URL(Activator.getContext().getBundle().getLocation()); 
			// System.err.println( "Activator.getContext().getBundle().getLocation() = " + bundleLocation );
			String bundleLocationFile = bundleLocation.getFile();
			if(bundleLocationFile.startsWith("file:/")) {
				bundleLocationFile = bundleLocationFile.substring(6);
			}
			// System.err.println( "bundleLocationFile = " + bundleLocationFile );

			URL resourceLocation = this.getClass().getResource( filename );
			String resourceLocationFile = resourceLocation.getFile();
			// System.err.println( "resourceLocationFile = " + resourceLocationFile );
			
			
			filepath = bundleLocationFile + "bin" + resourceLocationFile;
			// System.err.println( "filepath = " + filepath );
		}
		return filepath;
	}

	protected void addFilepathToAppContext(String filepath, IEngineTask task) {
		if( filepath != null ) {
			@SuppressWarnings("unchecked")
			Map<String,Object> appContext = (Map<String,Object>)task.getAppContext();
			if( appContext == null ) {
				appContext = new HashMap<String,Object>();
				task.setAppContext(appContext);
			}
			appContext.put("__report", filepath);					
		}
	}

	protected IReportDocument runReport(IReportEngine reportEngine,
			IRunTask reportRunTask, File tempDoc) throws IOException,
			EngineException {
		IDocArchiveWriter archiveWriter = new FileArchiveWriter( tempDoc.getCanonicalPath() );
		assertNotNull(archiveWriter);

		reportRunTask.run( archiveWriter );
		assertEquals( 0, reportRunTask.getErrors().size() );
		
		reportRunTask.close();
		
		IReportDocument reportDocument = reportEngine.openReportDocument( tempDoc.getCanonicalPath() );
		assertNotNull(reportDocument);
		return reportDocument;
	}

	protected RenderOption prepareRenderOptions(String outputFormat, FileOutputStream outputStream) {
		RenderOption renderOptions = new RenderOption();
		renderOptions.setOutputFormat( outputFormat );
		if( outputStream != null ) {
			renderOptions.setOutputStream( outputStream );
		}
		if( debug ) {
			renderOptions.setOption( "ExcelEmitter.DEBUG", Boolean.TRUE);
		}
		return renderOptions;
	}

}
