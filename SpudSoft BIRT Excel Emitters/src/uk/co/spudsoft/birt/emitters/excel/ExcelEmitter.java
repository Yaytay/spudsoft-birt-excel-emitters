package uk.co.spudsoft.birt.emitters.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.api.IRenderOption;
import org.eclipse.birt.report.engine.content.IAutoTextContent;
import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.content.IContainerContent;
import org.eclipse.birt.report.engine.content.IContent;
import org.eclipse.birt.report.engine.content.IDataContent;
import org.eclipse.birt.report.engine.content.IForeignContent;
import org.eclipse.birt.report.engine.content.IGroupContent;
import org.eclipse.birt.report.engine.content.IImageContent;
import org.eclipse.birt.report.engine.content.ILabelContent;
import org.eclipse.birt.report.engine.content.IListBandContent;
import org.eclipse.birt.report.engine.content.IListContent;
import org.eclipse.birt.report.engine.content.IListGroupContent;
import org.eclipse.birt.report.engine.content.IPageContent;
import org.eclipse.birt.report.engine.content.IReportContent;
import org.eclipse.birt.report.engine.content.IRowContent;
import org.eclipse.birt.report.engine.content.ITableBandContent;
import org.eclipse.birt.report.engine.content.ITableContent;
import org.eclipse.birt.report.engine.content.ITableGroupContent;
import org.eclipse.birt.report.engine.content.ITextContent;
import org.eclipse.birt.report.engine.css.engine.CSSEngine;
import org.eclipse.birt.report.engine.emitter.IContentEmitter;
import org.eclipse.birt.report.engine.emitter.IEmitterServices;

import uk.co.spudsoft.birt.emitters.excel.framework.ExcelEmitterPlugin;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;
import uk.co.spudsoft.birt.emitters.excel.handlers.PageHandler;

public abstract class ExcelEmitter implements IContentEmitter {

	public static final String DEBUG = "ExcelEmitter.DEBUG";
	public static final String REMOVE_BLANK_ROWS = "ExcelEmitter.RemoveBlankRows";
	public static final String ROTATION_PROP = "ExcelEmitter.Rotation";
	public static final String FORCEAUTOCOLWIDTHS_PROP = "ExcelEmitter.ForceAutoColWidths";
	
	/**
	 * Logger.
	 */
	protected Logger log;
	/**
	 * <p>
	 * Output stream that the report is to be written to.
	 * </p><p>
	 * This is set in initialize() and reset in end() and must not be set anywhere else.
	 * </p>
	 */
	protected OutputStream reportOutputStream;
	/**
	 * <p>
	 * Record of whether the emitter opened the report output stream itself, and it thus responsible for closing it.
	 * </p>
	 */
	protected boolean outputStreamOpened;
	/**
	 * <p>
	 * Name of the file that the report is to be written to (for tracking only).
	 * </p><p>
	 * This is set in initialize() and reset in end() and must not be set anywhere else.
	 * </p>
	 */
	protected String reportOutputFilename;
	/**
	 * The state date passed around the handlers.
	 */
	private HandlerState handlerState;
	/**
	 * <p>
	 * Set of functions for carrying out conversions between BIRT and POI. 
	 * </p>
	 */
	private StyleManagerUtils smu;
	private IRenderOption renderOptions;

	
	
	protected ExcelEmitter(StyleManagerUtils.Factory utilsFactory) {
		try {
			if( ExcelEmitterPlugin.getDefault() != null ) {
				log = ExcelEmitterPlugin.getDefault().getLogger();
			} else {
				log = new Logger( this.getClass().getPackage().getName() );
			}
			log.debug("ExcelEmitter");
			smu = utilsFactory.create(log);
		} catch( Exception ex ) {
			Throwable t = ex;
			while( t != null ) {
				log.debug( t.getMessage() );
				t.printStackTrace();
				t = t.getCause();
			}
		}
	}

	/**
	 * Constructs a new workbook to be processed by the emitter.
	 * @return
	 * The new workbook.
	 */
	protected abstract Workbook createWorkbook();
	
	
	public void initialize( IEmitterServices service ) throws BirtException {
		log.debug("inintialize");
		reportOutputStream = service.getRenderOption().getOutputStream();
		reportOutputFilename = service.getRenderOption().getOutputFileName();
		if( ( reportOutputStream == null )
				&& ( ( reportOutputFilename == null ) || reportOutputFilename.isEmpty() ) ) {
			throw new BirtException( EmitterServices.getPluginName()
					, "Neither output stream nor output filename have been specified"
					, null
					);			
		}
				
		renderOptions = service.getRenderOption();
		boolean debug = EmitterServices.booleanOption( renderOptions, DEBUG, false );
		if( debug )  {
			this.log.setDebug(debug);
		}		
	}

	public void start( IReportContent report ) throws BirtException {
		log.addPrefix('>');
		log.info( 0, "start:" + report.toString(), null);
		
	    Workbook wb = createWorkbook();
	    CSSEngine cssEngine = report.getRoot().getCSSEngine();
	    StyleManager sm = new StyleManager( wb, log, smu, cssEngine );
	    
		handlerState = new HandlerState(this, log, smu, wb, sm, renderOptions);
		handlerState.setHandler( new PageHandler(log, null) );
	}

	public void end( IReportContent report ) throws BirtException {
		log.removePrefix('>');
		log.debug("end:", report);
		
		String reportTitle = report.getTitle();
		if( ( handlerState.getWb().getNumberOfSheets() == 1 ) 
				&& ! handlerState.firstSheetNamed 
				&& ( reportTitle != null )) {
			handlerState.getWb().setSheetName(0, reportTitle);
		}
		
		OutputStream outputStream = reportOutputStream;
		try {
			if( outputStream == null ) {
				if( ( reportOutputFilename != null ) && ! reportOutputFilename.isEmpty() ) {
					try {
						outputStream = new FileOutputStream( reportOutputFilename );
					} catch( IOException ex ) {
						log.warn( 0, "File \"" + reportOutputFilename + "\" cannot be opened for writing", ex);
						throw new BirtException( EmitterServices.getPluginName()
								, "Unable to open file (\"{}\") for writing"
								, new Object[] { reportOutputFilename }
								, null
								, ex 
								);
					}
				} 
			}
			handlerState.getWb().write(outputStream);
		} catch( Throwable ex ) {
			log.debug("ex:", ex.toString());
			ex.printStackTrace();
			
			throw new BirtException( EmitterServices.getPluginName()
					, "Unable to save file (\"{}\")"
					, new Object[] { reportOutputFilename }
					, null
					, ex 
					);
		} finally {
			if( reportOutputStream == null ) {
				try {
					outputStream.close();
				} catch( IOException ex ) {
					log.debug("ex:", ex.toString());
				}
			}
			handlerState = null;
			reportOutputFilename = null;			
			reportOutputStream = null;
		}
		
	}

	public void startPage( IPageContent page ) throws BirtException {
		log.addPrefix( 'P' );
		log.debug( handlerState, "startPage: " );
		handlerState.getHandler().startPage(handlerState,page);
	}
	public void endPage( IPageContent page ) throws BirtException {
		log.debug( handlerState, "endPage: " );
		handlerState.getHandler().endPage(handlerState,page);
		log.removePrefix( 'P' );
	}

	public void startTable( ITableContent table ) throws BirtException {
		log.addPrefix( 'T' );
		log.debug( handlerState, "startTable: " );
		handlerState.getHandler().startTable(handlerState,table);
	}
	public void endTable( ITableContent table ) throws BirtException {
		log.debug( handlerState, "endTable: " );
		handlerState.getHandler().endTable(handlerState,table);
		log.removePrefix( 'T' );
	}

	public void startTableBand( ITableBandContent band ) throws BirtException {
		log.addPrefix( 'B' );
		log.debug( handlerState, "startTableBand: " );
		handlerState.getHandler().startTableBand(handlerState,band);
	}
	public void endTableBand( ITableBandContent band ) throws BirtException {
		log.debug( handlerState, "endTableBand: " );
		handlerState.getHandler().endTableBand(handlerState,band);
		log.removePrefix( 'B' );
	}

	public void startRow( IRowContent row ) throws BirtException {
		log.addPrefix( 'R' );
		log.debug( handlerState, "startRow: " );
		handlerState.getHandler().startRow(handlerState,row);
	}
	public void endRow( IRowContent row ) throws BirtException {
		log.debug( handlerState, "endRow: " );
		handlerState.getHandler().endRow(handlerState,row);
		log.removePrefix( 'R' );
	}

	public void startCell( ICellContent cell ) throws BirtException {
		log.addPrefix( 'C' );
		log.debug( handlerState, "startCell: " );
		handlerState.getHandler().startCell(handlerState,cell);
	}
	public void endCell( ICellContent cell ) throws BirtException {
		log.debug( handlerState, "endCell: " );
		handlerState.getHandler().endCell(handlerState,cell);
		log.removePrefix( 'C' );
	}
	
	public void startList( IListContent list ) throws BirtException {
		log.addPrefix( 'L' );
		log.debug( handlerState, "startList: " );
		handlerState.getHandler().startList(handlerState,list);
	}
	public void endList( IListContent list ) throws BirtException {
		log.debug( handlerState, "endList: " );
		handlerState.getHandler().endList(handlerState,list);
		log.removePrefix( 'L' );
	}

	public void startListBand( IListBandContent listBand ) throws BirtException {
		log.addPrefix( 'B' );
		log.debug( handlerState, "startListBand: " );
		handlerState.getHandler().startListBand(handlerState,listBand);
	}
	public void endListBand( IListBandContent listBand ) throws BirtException {
		log.debug( handlerState, "endListBand: " );
		handlerState.getHandler().endListBand(handlerState,listBand);
		log.removePrefix( 'B' );
	}

	public void startContainer( IContainerContent container ) throws BirtException {
		log.addPrefix( 'O' );
		log.debug( handlerState, "startContainer: " );
		handlerState.getHandler().startContainer(handlerState,container);
	}
	public void endContainer( IContainerContent container ) throws BirtException {
		log.debug( handlerState, "endContainer: " );
		handlerState.getHandler().endContainer(handlerState,container);
		log.removePrefix( 'O' );
	}

	public void startText( ITextContent text ) throws BirtException {
		log.debug( handlerState, "startText: " );
		handlerState.getHandler().emitText(handlerState,text);
	}

	public void startData( IDataContent data ) throws BirtException {
		log.debug( handlerState, "startData: " );
		handlerState.getHandler().emitData(handlerState,data);
	}

	public void startLabel( ILabelContent label ) throws BirtException {
		log.debug( handlerState, "startLabel: " );
		handlerState.getHandler().emitLabel(handlerState,label);
	}
	
	public void startAutoText ( IAutoTextContent autoText ) throws BirtException {
		log.debug( handlerState, "startAutoText: " );
		handlerState.getHandler().emitAutoText(handlerState,autoText);
	}

	public void startForeign( IForeignContent foreign ) throws BirtException {
		log.debug( handlerState, "startForeign: " );
		handlerState.getHandler().emitForeign(handlerState,foreign);
	}

	public void startImage( IImageContent image ) throws BirtException {
		log.debug( handlerState, "startImage: " );
		handlerState.getHandler().emitImage(handlerState,image);
	}

	public void startContent( IContent content ) throws BirtException {
		log.addPrefix( 'N' );
		log.debug( handlerState, "startContent: " );
		handlerState.getHandler().startContent(handlerState,content);
	}
	public void endContent( IContent content) throws BirtException {
		log.debug( handlerState, "endContent: " );
		handlerState.getHandler().endContent(handlerState,content);
		log.removePrefix( 'N' );
	}
	
	public void startGroup( IGroupContent group ) throws BirtException {
		log.debug( handlerState, "startGroup: " );
		handlerState.getHandler().startGroup(handlerState,group);
	}
	public void endGroup( IGroupContent group ) throws BirtException {
		log.debug( handlerState, "endGroup: " );
		handlerState.getHandler().endGroup(handlerState,group);
	}

	public void startTableGroup( ITableGroupContent group ) throws BirtException {
		log.addPrefix( 'G' );
		log.debug( handlerState, "startTableGroup: " );
		handlerState.getHandler().startTableGroup(handlerState,group);
	}
	public void endTableGroup( ITableGroupContent group ) throws BirtException {
		log.debug( handlerState, "endTableGroup: " );
		handlerState.getHandler().endTableGroup(handlerState,group);
		log.removePrefix( 'G' );
	}

	public void startListGroup( IListGroupContent group ) throws BirtException {
		log.addPrefix( 'G' );
		log.debug( handlerState, "startListGroup: " );
		handlerState.getHandler().startListGroup(handlerState,group);
	}
	public void endListGroup( IListGroupContent group ) throws BirtException {
		log.debug( handlerState, "endListGroup: " );
		handlerState.getHandler().endListGroup(handlerState,group);
		log.removePrefix( 'G' );
	}
	
	
	
	
}
