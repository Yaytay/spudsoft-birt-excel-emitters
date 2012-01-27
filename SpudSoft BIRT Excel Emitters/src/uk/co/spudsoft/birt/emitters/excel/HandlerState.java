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
     * List of Current Spans
     * We could probably use CellRangeAdresses inside the sheet, but 
     * this way we keep the tests to a minimum.
     */
    public List<Area> rowSpans = new ArrayList<Area>();
	
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
	
	public boolean cellIsMergedWithBorders( int row, int column ) {
		if( areaBorders != null ) {
			for( AreaBorders areaBorder : areaBorders ) {
				if( ( areaBorder.isMergedCells ) 
						&& ( areaBorder.top == row )
						&& ( areaBorder.left == column ) ) {
					return true;
				}
			}
			
		}
		return false;
	}
	
	public boolean rowHasMergedCellsWithBorders( int row ) {
		if( areaBorders != null ) {
			for( AreaBorders areaBorder : areaBorders ) {
				if( ( areaBorder.isMergedCells ) 
						&& ( areaBorder.top <= row )
						&& ( areaBorder.bottom >= row ) ) {
					return true;
				}
			}			
		}
		return false;
	}
	
	public void addRowSpan(int rowX, int colX, int rowY, int colY) {
	    rowSpans.add(new Area(new Coordinate(rowX, colX), new Coordinate(rowY, colY)));
	}
	
    public int computeNumberSpanBefore(int row, int col) {
        int i = 0;
        for(Area a : rowSpans) {
        	// I'm now not removing passed spans, so do check a.y.row()
        	if( a.y.getRow() < row ) {
        		continue;
        	}
        	
            //Correct this col to know the real col number
            if(a.x.getCol() <= col) {
                col += (a.y.getCol() - a.x.getCol()) + 1;
            }
            if(row > a.x.getRow() //Span on first appearance is ok. 
                && a.x.getCol() <= col //This span is before this column
                ) {
                i += (a.y.getCol() - a.x.getCol()) + 1;
            }
        }
        return i;
    }
    
    public void clearRowSpans() {
    	rowSpans.clear();
    }
    
    public int findRowsSpanned( int rowX, int colX ) {
    	for( Area a : rowSpans ) {
    		if( ( a.x.getRow() == rowX ) && ( a.x.getCol() == colX ) ) {
    			return a.y.getRow() - a.x.getRow();
    		}
    	}
    	return 0;
    }
}
