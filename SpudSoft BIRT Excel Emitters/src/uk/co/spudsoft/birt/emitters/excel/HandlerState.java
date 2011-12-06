package uk.co.spudsoft.birt.emitters.excel;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.eclipse.birt.report.engine.api.IRenderOption;
import org.eclipse.birt.report.engine.api.ReportEngine;
import org.eclipse.birt.report.engine.emitter.IContentEmitter;

import uk.co.spudsoft.birt.emitters.excel.framework.Logger;
import uk.co.spudsoft.birt.emitters.excel.handlers.IHandler;

public class HandlerState {
	
	/**
	 * The emitter itself
	 */
	private IContentEmitter emitter;
	/**
	 * Logger.
	 */
	private Logger log;
	/**
	 * Set of functions for carrying out conversions between BIRT and POI. 
	 */
	private StyleManagerUtils smu;
	
	/**
	 * The current handler to pass on the processing to.
	 * Effectively this is the state machine for the emitter.
	 */
	private IHandler handler;
	
	/**
	 * The workbook being generated.
	 */
	private Workbook wb;
	/**
	 * Style cache, to enable reuse of styles between cells.
	 */
	private StyleManager sm;
	/**
	 * Render options
	 */
	private IRenderOption renderOptions;
	/**
	 * Report engine
	 */
	private ReportEngine reportEngine;
	
	/**
	 * Record whether a name was given to the first sheet
	 */
	public boolean firstSheetNamed;

	/**
	 * The current POI sheet being processed.
	 */
	public Sheet currentSheet;
	/**
	 * Collection of CellImage objects for the current sheet.
	 */
	public List<CellImage> images = new ArrayList<CellImage>();
	/**
	 * Possible name for the current sheet
	 */
	public String sheetName;
	/**
	 * The index of the row that should be created next
	 */
	public int rowNum;
	/**
	 * The index of the column in which the next data should begin
	 */
	public int colNum;
	/**
	 * The minimum row height required for this top level row
	 */
	public float requiredRowHeightInPoints;
	public int rowOffset;
	public int colOffset;

	/**
	 * Border overrides for the current row/table
	 */
	public List<AreaBorders> areaBorders = new ArrayList<AreaBorders>();
	
	/**
	 * Constructor
	 * @param log
	 * @param smu
	 * @param wb
	 * @param sm
	 */
	public HandlerState(IContentEmitter emitter, Logger log, StyleManagerUtils smu, Workbook wb, StyleManager sm, IRenderOption renderOptions) {
		super();
		this.emitter = emitter;
		this.log = log;
		this.smu = smu;
		this.wb = wb;
		this.sm = sm;
		this.renderOptions = renderOptions;
	}
	
	

	public IContentEmitter getEmitter() {
		return emitter;
	}

	public Logger getLog() {
		return log;
	}

	public StyleManagerUtils getSmu() {
		return smu;
	}

	public Workbook getWb() {
		return wb;
	}

	public StyleManager getSm() {
		return sm;
	}

	public IRenderOption getRenderOptions() {
		return renderOptions;
	}

	public ReportEngine getReportEngine() {
		return reportEngine;
	}

	public IHandler getHandler() {
		return handler;
	}

	public void setHandler(IHandler handler) {
		this.handler = handler;
		this.handler.notifyHandler(this);
	}
	
	public void insertBorderOverload(AreaBorders defn) {
		if( areaBorders == null ) {
			areaBorders = new ArrayList<AreaBorders>();
		}
		areaBorders.add( defn );
	}
	
	public void removeBorderOverload(AreaBorders defn) {
		if( areaBorders != null ) {
			areaBorders.remove(defn);
		}
	}
	
}
