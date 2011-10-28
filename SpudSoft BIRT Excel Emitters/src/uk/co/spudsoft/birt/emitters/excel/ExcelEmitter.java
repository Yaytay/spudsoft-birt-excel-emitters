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

package uk.co.spudsoft.birt.emitters.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.IAutoTextContent;
import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.content.IDataContent;
import org.eclipse.birt.report.engine.content.IImageContent;
import org.eclipse.birt.report.engine.content.ILabelContent;
import org.eclipse.birt.report.engine.content.IPageContent;
import org.eclipse.birt.report.engine.content.IReportContent;
import org.eclipse.birt.report.engine.content.IRowContent;
import org.eclipse.birt.report.engine.content.IStyle;
import org.eclipse.birt.report.engine.content.ITableContent;
import org.eclipse.birt.report.engine.emitter.ContentEmitterAdapter;
import org.eclipse.birt.report.engine.emitter.IEmitterServices;
import org.eclipse.birt.report.engine.ir.DimensionType;

import uk.co.spudsoft.birt.emitters.excel.framework.ExcelEmitterPlugin;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

/**
 * <p>
 * ExcelEmitter is the base class for the two Excel emitters in this bundle.
 * </p><p>
 * In theory ExcelEmitter is responsible for managing the tracking of the emitters state and 
 * translating that to POI objects.
 * In practice some noise has bled into the ExcelEmitter and it handles a little more than would be ideal.
 * </p>
 * @author Jim Talbut
 */
public abstract class ExcelEmitter extends ContentEmitterAdapter {
	
	/**
	 * Number of milliseconds in a day, to determine whether a given date is date/time/datetime
	 */
	private static final long oneDay = 24 * 60 * 60 * 1000;

	
	/**
	 * <p>
	 * CellImage is used to cache all the required data for inserting images so that they can be
	 * processed after all other spreadsheet contents has been inserted.
	 * </p><p>
	 * Processing images after all other spreadsheet contents means that the images will be unaffected
	 * by any column resizing that may be required.
	 * Images usually cause row resizing (the emitter never allows an image to spread onto multiple rows),
	 * but never cause column resizing.
	 * </p>
	 * 
	 * @author Jim Talbut
	 *
	 */
	protected class CellImage {
		Cell cell;
		int imageIdx;
		IImageContent image;
		public CellImage(Cell cell, int imageIdx, IImageContent image) {
			this.cell = cell;
			this.imageIdx = imageIdx;
			this.image = image;
		}
	}
	/**
	 * <p>
	 * Collection of CellImage objects for the current sheet.
	 * </p><p>
	 * Cleared (emptied) in endSheet().
	 * </p>
	 */
	protected List<CellImage> images = new ArrayList<CellImage>();

	/**
	 * <p>
	 * The workbook being generated.
	 * </p><p>
	 * This is set in start() and reset in end() and must not be set anywhere else.
	 * </p>
	 */
	protected Workbook wb;
	/**
	 * <p>
	 * Style cache, to enable reuse of styles between cells.
	 * </p><p>
	 * This is set in start() and reset in end() and must not be set anywhere else.
	 * </p>
	 */
	protected StyleManager sm;
	/**
	 * <p>
	 * Style stack, to allow cells to inherit properties from container elements.
	 * </p><p>
	 * This is set in start() and reset in end() and must not be set anywhere else.
	 * </p>
	 */
	protected StyleStack styleStack;
	
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
	 * Name of the file that the report is to be written to (for tracking only).
	 * </p><p>
	 * This is set in initialize() and reset in end() and must not be set anywhere else.
	 * </p>
	 */
	protected String reportOutputFilename;

	
	/**
	 * <p>
	 * The current POI sheet being processed.
	 * </p><p>
	 * Created in startSheet() and reset in endSheet().
	 * </p>
	 */
	protected Sheet currentSheet;
	/**
	 * <p>
	 * The (zero-based) index for the current sheet being processed.
	 * </p><p>
	 * Incremented in startSheet() and reset in start().
	 * </p>
	 */
	protected int sheetNum;
	/**
	 * <p>
	 * The drawing patriarch for any drawings on the page.
	 * </p><p>
	 * Created and cached in getDrawing, reset in endSheet().
	 * </p>
	 */
	protected Drawing currentDrawing;
	/**
	 * The current POI row being processed.
	 */
	protected Row currentRow;
	/**
	 * The (zero-based) index for the current row being processed.
	 */
	protected int rowNum;
	/**
	 * Flag to track whether a row has changed height, to prevent a row being shrunk more than once.
	 */
	protected boolean rowHeightChanged;
	/**
	 * The current POI cell being processed.
	 */
	protected Cell currentCell;
	/**
	 * The (zero-based) index for the current cell being processed.
	 */
	protected int colNum;
	/**
	 * Tracking for the row that is the start of the current table, to enable table borders to be processed in endTable().
	 */
	protected int tableStartRow;
	/**
	 * Count of the nested tables that have been started
	 */
	protected int nestedTableCount;
	/**
	 * Count of the nested rows that have been started
	 */
	protected int nestedRowCount;
	/**
	 * Count of the nested cells that have been started
	 */
	protected int nestedCellCount;
	/**
	 * The last value added to a cell
	 */
	protected Object lastValue;
	/**
	 * The last named table/grid seen, used to name sheets
	 */
	protected String lastTableName;
	/**
	 * Record whether a name was given to the first sheet
	 */
	protected boolean firstSheetNamed;
	
	/**
	 * Logger.
	 */
	protected Logger log;
	/**
	 * <p>
	 * Set of functions for carrying out conversions between BIRT and POI. 
	 * </p><p>
	 * Originally StyleManagerUtils was entirely static, but became virtual to support differences between HSSF and XSSF.
	 * </p>
	 */
	protected StyleManagerUtils smu;
	
	protected ExcelEmitter() {
		try {
			if( ExcelEmitterPlugin.getDefault() != null ) {
				log = ExcelEmitterPlugin.getDefault().getLogger();
			} else {
				log = new Logger( this.getClass().getPackage().getName() );
			}
			log.debug("ExcelEmitter");
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
	 * Returns the symbolic name for the plugin.
	 */
	protected String getSymbolicName() {
		if( ( ExcelEmitterPlugin.getDefault() != null ) && ( ExcelEmitterPlugin.getDefault().getBundle() != null ) ) {
			return ExcelEmitterPlugin.getDefault().getBundle().getSymbolicName();
		} else {
			return "uk.co.spudsoft.birt.emitters.excel";
		}
	}
	
	/**
	 * Sets the style manager utility object.
	 * Must be called immediately after the constructor (and cannot be made a constructor argument).
	 * @param smu
	 * The style manager utility object.
	 */
	protected void setStyleManagerUtils(StyleManagerUtils smu) {
		this.smu = smu;
	}
	
	/**
	 * Constructs a new workbook to be processed by the emitter.
	 * @return
	 * The new workbook.
	 */
	protected abstract Workbook createWorkbook();

	/**
	 * Create a new sheet.
	 * @param possibleTitle
	 * A possible title for the new sheet (may be null or empty).
	 */
	protected void startSheet( ) {
	    currentSheet = wb.createSheet();
	    ++sheetNum;
	    rowNum = 0;
	}
	
	/**
	 * Finalise the current sheet.
	 */
	protected void endSheet() {
		for( CellImage cellImage : images ) {
			processCellImage(cellImage);
		}
		images.clear();
		
		currentSheet = null;
		currentDrawing = null;
	}
	
	@Override
	public void start(IReportContent report) throws BirtException {
		log.addPrefix('>');
		log.info( 0, "start:" + report.toString(), null);
		super.start(report);
		
	    sheetNum = -1;
	    wb = createWorkbook();
	    styleStack = new StyleStack();	    
	    sm = new StyleManager(wb, styleStack, log, smu, report.getRoot().getCSSEngine());

	    nestedCellCount = 0;
	    nestedRowCount = 0;
	    nestedTableCount = 0;
	}

	@Override
	public void end(IReportContent report) throws BirtException {
		// endSheet();
		
		String reportTitle = report.getTitle();
		if( ( wb.getNumberOfSheets() == 1 ) 
				&& ! firstSheetNamed 
				&& ( reportTitle != null )) {
			wb.setSheetName(0, reportTitle);
		}
		
		log.removePrefix('>');
		log.debug("end:" + report.toString());
		try {
			wb.write(reportOutputStream);
		} catch( Throwable ex ) {
			log.debug("ex:" + ex.toString());
			ex.printStackTrace();
			
			throw new BirtException( getSymbolicName()
					, "Unable to save file (\"{}\")"
					, new Object[] { reportOutputFilename }
					, null
					, ex 
					);
		} finally {
			wb = null;
			styleStack = null;
			sm = null;
			reportOutputFilename = null;
			reportOutputStream = null;
		}
		super.end(report);
	}
	

	@SuppressWarnings("deprecation")
	@Override
	public void endCell(ICellContent cell) throws BirtException {

		log.removePrefix('C');
		log.debug("endCell");
		super.endCell(cell);
		--nestedCellCount;

		if( nestedCellCount == 0 ) {
			
			ICellContent cellContent = styleStack.pop(ICellContent.class);
			IStyle birtStyle = cellContent.getStyle();
			
			if( ( birtStyle.getNumberFormat() == null )
					&& ( birtStyle.getDateFormat() == null )
					&& ( birtStyle.getDateTimeFormat() == null )
					&& ( birtStyle.getTimeFormat() == null )
					&& ( lastValue != null )
					) {
				if( lastValue instanceof Date ) {
					long time = ((Date)lastValue).getTime();
					time = time - ((Date)lastValue).getTimezoneOffset() * 60000;
					if( time % oneDay == 0 ) {
						birtStyle.setDateFormat("Short Date");
					} else if( time < oneDay ) {
						birtStyle.setDateFormat("Short Time");
					} else {
						birtStyle.setDateFormat("General Date");
					}
				}
			}
			
			CellStyle cellStyle = sm.getStyle(cellContent);
			currentCell.setCellStyle(cellStyle);
			
			colNum += cell.getColSpan();
			currentCell = null;
			lastValue = null;
		}
	}


/*	@Override
	public void endContainer(IContainerContent container) throws BirtException {

		// log.debug("endContainer");
		super.endContainer(container);
	}
*/

/*	@Override
	public void endContent(IContent content) throws BirtException {

		log.debug("endContent");
		super.endContent(content);
	}
*/

/*	@Override
	public void endGroup(IGroupContent group) throws BirtException {

		log.debug("endGroup");
		super.endGroup(group);
	}
*/

/*	@Override
	public void endList(IListContent list) throws BirtException {

		log.debug("endList");
		super.endList(list);
	}
*/

/*	@Override
	public void endListBand(IListBandContent listBand) throws BirtException {

		log.debug("endListBand");
		super.endListBand(listBand);
	}
*/

/*	@Override
	public void endListGroup(IListGroupContent group) throws BirtException {

		log.debug("endListGroup");
		super.endListGroup(group);
	}
*/

	@Override
	public void startPage(IPageContent page) throws BirtException {
		log.addPrefix( 'P' );
		log.debug("startPage");
		super.startPage(page);
		if( ( nestedCellCount == 0 ) && ( nestedRowCount == 0 ) && ( nestedTableCount == 0 ) ) {
			startSheet();
			styleStack.push(page);
			log.debug("Page type: " + page.getPageType());
			
			if( page.getPageType() != null ) {
				setupPageSize(page);
			}
			
			prepareMarginDimensions(page);
		}
	}

	@Override
	public void endPage(IPageContent page) throws BirtException {
		log.removePrefix( 'P' );
		log.debug("endPage");
		super.endPage(page);
		if( ( nestedCellCount == 0 ) && ( nestedRowCount == 0 ) && ( nestedTableCount == 0 ) ) {
			styleStack.pop(IPageContent.class);
			if( lastTableName != null ) {
				log.debug("Attempting to name sheet " + sheetNum + "\"" + lastTableName + "\" ");
				boolean alreadyFound = false;
				for( int i = 0; i < sheetNum; ++i ) {
					if( wb.getSheetName(i).equals(lastTableName)) {
						alreadyFound = true;
					}
				}
				if(!alreadyFound) {
					wb.setSheetName(sheetNum, lastTableName);
				}
			} 
			endSheet();
		}
	}


	@Override
	public void endRow(IRowContent row) throws BirtException {

		log.removePrefix( 'R' );
		log.debug("endRow");
		super.endRow(row);
		--nestedRowCount;

		if( nestedRowCount == 0) {
			// Check whether the entire row should be deleted
			boolean blankRow = true;
			for(Iterator<Cell> iter = currentRow.cellIterator(); iter.hasNext(); ) {
				Cell cell = iter.next();
				if(! smu.cellIsEmpty(cell)) {
					blankRow = false;
					break;
				}
			}
			if(blankRow) {
				--rowNum;
				log.debug("Removing blank row");
				currentSheet.removeRow(currentRow);
			} else {
				DimensionType height = row.getHeight();
				if(height != null) {
					double points = height.convertTo(DimensionType.UNITS_PT);
					if( !rowHeightChanged || (points > currentRow.getHeightInPoints())) {
						rowHeightChanged = true;
						currentRow.setHeightInPoints((float)points);
					}
				}
				
				applyBordersToArea( 0, colNum - 1, rowNum, rowNum, row.getStyle() );
			}
	
			++rowNum;
			currentRow = null;
			styleStack.pop(IRowContent.class);
		}
	}
	
	/**
	 * Place a border around a region on the current sheet.
	 * This is used to apply borders to entire rows or entire tables.
	 * @param colStart
	 * The column marking the left-side boundary of the region.
	 * @param colEnd
	 * The column marking the right-side boundary of the region.
	 * @param rowStart
	 * The row marking the top boundary of the region.
	 * @param rowEnd
	 * The row marking the bottom boundary of the region.
	 * @param borderStyle
	 * The BIRT border style to apply to the region.
	 */
	private void applyBordersToArea( int colStart, int colEnd, int rowStart, int rowEnd, IStyle borderStyle ) {
		StringBuilder borderMsg = new StringBuilder();
		borderMsg.append( "applyBordersToArea [" ).append( colStart ).append( "," ).append( rowStart ).append( "]-[" ).append( colEnd ).append( "," ).append( rowEnd ).append( "]");
		
		String borderStyleBottom = borderStyle.getBorderBottomStyle();
		String borderWidthBottom = borderStyle.getBorderBottomWidth();
		String borderColourBottom = borderStyle.getBorderBottomColor();
		String borderStyleLeft = borderStyle.getBorderLeftStyle();
		String borderWidthLeft = borderStyle.getBorderLeftWidth();
		String borderColourLeft = borderStyle.getBorderLeftColor();
		String borderStyleRight = borderStyle.getBorderRightStyle();
		String borderWidthRight = borderStyle.getBorderRightWidth();
		String borderColourRight = borderStyle.getBorderRightColor();
		String borderStyleTop = borderStyle.getBorderTopStyle();
		String borderWidthTop = borderStyle.getBorderTopWidth();
		String borderColourTop = borderStyle.getBorderTopColor();
		
	 	borderMsg.append( ", Bottom:" ).append( borderStyleBottom ).append( "/" ).append( borderWidthBottom ).append( "/" + borderColourBottom );
		borderMsg.append( ", Left:" ).append( borderStyleLeft ).append( "/" ).append( borderWidthLeft ).append( "/" + borderColourLeft );
		borderMsg.append( ", Right:" ).append( borderStyleRight ).append( "/" ).append( borderWidthRight ).append( "/" ).append( borderColourRight );
		borderMsg.append( ", Top:" ).append( borderStyleTop ).append( "/" ).append( borderWidthTop ).append( "/" ).append( borderColourTop );
		log.debug( borderMsg.toString() );

		if( ( borderStyleBottom != null ) && ( borderWidthBottom == null ) ) {
			borderWidthBottom = "3pt";
		}
		if( ( borderStyleBottom != null ) && ( borderColourBottom == null ) ) {
			borderColourBottom = "rgb(0,0,0)";
		}
		
		if( ( borderStyleLeft != null ) && ( borderWidthLeft == null ) ) {
			borderWidthLeft = "3pt";
		}
		if( ( borderStyleLeft != null ) && ( borderColourLeft == null ) ) {
			borderColourLeft = "rgb(0,0,0)";
		}
		
		if( ( borderStyleRight != null ) && ( borderWidthRight == null ) ) {
			borderWidthRight = "3pt";
		}
		if( ( borderStyleRight != null ) && ( borderColourRight == null ) ) {
			borderColourRight = "rgb(0,0,0)";
		}
		
		if( ( borderStyleTop != null ) && ( borderWidthTop == null ) ) {
			borderWidthTop = "3pt";
		}
		if( ( borderStyleTop != null ) && ( borderColourTop == null ) ) {
			borderColourTop = "rgb(0,0,0)";
		}
		
		if( ( borderStyleBottom != null ) || ( borderWidthBottom != null ) || ( borderColourBottom != null ) 
				|| ( borderStyleLeft != null ) || ( borderWidthLeft != null ) || ( borderColourLeft != null )
				|| ( borderStyleRight != null ) || ( borderWidthRight != null ) || ( borderColourRight != null ) 
				|| ( borderStyleTop != null ) || ( borderWidthTop != null ) || ( borderColourTop != null ) 
				) {
			for( int row = rowStart; row <= rowEnd; ++row ) {
				Row styleRow = currentSheet.getRow(row);
				if( styleRow != null ) {
					for( int col = colStart; col <= colEnd; ++col ) {
						if( ( col == colStart ) || ( col == colEnd ) || ( row == rowStart ) || ( row == rowEnd ) ) {
							Cell styleCell = styleRow.getCell(col);
							if( styleCell == null ) {
								styleCell = styleRow.createCell(col);
							}
							if( styleCell != null ) {
								log.debug( "Applying border to cell [R" + styleCell.getRowIndex() + "C" + styleCell.getColumnIndex() + "]");
								CellStyle newStyle = sm.getStyleWithBorders( styleCell.getCellStyle()
										, ( (row == rowEnd) ? borderStyleBottom : null ), ( (row == rowEnd) ? borderWidthBottom : null ), ( (row == rowEnd) ? borderColourBottom : null )
										, ( (col == colStart) ? borderStyleLeft: null ), ( (col == colStart) ? borderWidthLeft: null ), ( (col == colStart) ? borderColourLeft: null )
										, ( (col == colEnd) ? borderStyleRight: null ), ( (col == colEnd) ? borderWidthRight: null ), ( (col == colEnd) ? borderColourRight: null )
										, ( (row == rowStart) ? borderStyleTop: null ), ( (row == rowStart) ? borderWidthTop: null ), ( (row == rowStart) ? borderColourTop: null )
										);
								styleCell.setCellStyle(newStyle);
							}
						}
					}
				}
			}
		}
	}


	@Override
	public void endTable(ITableContent table) throws BirtException {

		log.removePrefix( 'T' );
		log.debug("endTable");
		super.endTable(table);
		--nestedTableCount;
		if( nestedTableCount == 0) {
			
			// Drawings don't remain the same when columns are resized, so for now don't resize if there are any drawings
			if( this.currentDrawing == null ) {
				for( int col = 0; col < table.getColumnCount(); ++col ) {
					log.debug( "BIRT table column width: " + col + " = " + table.getColumn(col).getWidth());
					if( table.getColumn(col).getWidth() != null ) {
						currentSheet.setColumnWidth(col, smu.poiColumnWidthFromDimension(table.getColumn(col).getWidth()));
					}
				}
			}
	
			applyBordersToArea( 0, table.getColumnCount() - 1, tableStartRow, rowNum - 1, table.getStyle() );
			
			styleStack.pop(ITableContent.class);
		}
	}


/*	@Override
	public void endTableBand(ITableBandContent band) throws BirtException {

		log.debug("endTableBand");
		super.endTableBand(band);
	}
*/

/*	@Override
	public void endTableGroup(ITableGroupContent group) throws BirtException {

		log.debug("endTableGroup");
		super.endTableGroup(group);
	}
*/

	@Override
	public void initialize(IEmitterServices service) throws BirtException {

		log.debug("inintialize");
		reportOutputStream = service.getRenderOption().getOutputStream();
		reportOutputFilename = service.getRenderOption().getOutputFileName();
		
		super.initialize(service);
	}


	@Override
	public void startAutoText(IAutoTextContent autoText) throws BirtException {

		log.debug("startAutoText");		
		super.startAutoText(autoText);
	}


	@Override
	public void startCell(ICellContent cell) throws BirtException {

		log.addPrefix( 'C' );
		log.debug("startCell [R" + cell.getRow() + "C" + cell.getColumn() + "], span:" + cell.getColSpan() +", align:" + cell.getStyle().getTextAlign()
/*				+ ", \nstyle: " + StyleManagerUtils.birtStyleToString(cell.getStyle())
				+ ", \ninlineStyle: " + StyleManagerUtils.birtStyleToString(cell.getInlineStyle())
				+ ", \ncomputedStyle: " + StyleManagerUtils.birtStyleToString(cell.getComputedStyle())
*/				);
		super.startCell(cell);
		++nestedCellCount;
		if( nestedCellCount == 1 ) {
			currentCell = currentRow.createCell( cell.getColumn() );
			currentCell.setCellType(Cell.CELL_TYPE_BLANK);
					
			if(( cell.getColSpan() > 1 )||( cell.getRowSpan() > 1 )) {
				currentSheet.addMergedRegion( new CellRangeAddress( rowNum, rowNum + cell.getRowSpan() - 1
						, colNum, colNum + cell.getColSpan() - 1));
			}
			styleStack.push(cell);
		}
	}


/*	@Override
	public void startContainer(IContainerContent container)
			throws BirtException {

		// log.debug("startContainer");
		super.startContainer(container);
	}
*/

/*	@Override
	public void startContent(IContent content) throws BirtException {

		log.debug("startContent type:" + content.getContentType());
		super.startContent(content);
	}
*/

	@Override
	public void startData(IDataContent data) throws BirtException {

		log.debug("startData " + ( ( data != null ) && ( data.getValue() != null ) ? data.getValue().toString() + " (" + data.getValue().getClass().getCanonicalName() + ")" : "null" ) 
				// + ", style: " + StyleManagerUtils.birtStyleToString(data.getStyle())
				);
		super.startData(data);

		styleStack.mergeTop(data, ICellContent.class);
		Object value = data.getValue();
		
		setCurrentCellContents(value);
	}


	/**
	 * Set the contents of the current cell.
	 * If the current cell is empty this will format the cell optimally for the new value, if the current cell already has some contents this will simply append the text
	 * value to the current contents.
	 * @param value
	 * The value to put into the current cell.
	 */
	private void setCurrentCellContents(Object value) {
		switch( currentCell.getCellType() ) {
		case Cell.CELL_TYPE_BLANK:
			if( value instanceof Double ) {
				currentCell.setCellType(Cell.CELL_TYPE_NUMERIC);
				currentCell.setCellValue((Double)value);
				lastValue = value;
			} else if( value instanceof Integer ) {
				currentCell.setCellType(Cell.CELL_TYPE_NUMERIC);
				currentCell.setCellValue((Integer)value);				
				lastValue = value;
			} else if( value instanceof Long ) {
				currentCell.setCellType(Cell.CELL_TYPE_NUMERIC);
				currentCell.setCellValue((Long)value);				
				lastValue = value;
			} else if( value instanceof Date ) {
				currentCell.setCellType(Cell.CELL_TYPE_NUMERIC);
				currentCell.setCellValue((Date)value);
				lastValue = value;
			} else if( value instanceof Boolean ) {
				currentCell.setCellType(Cell.CELL_TYPE_BOOLEAN);
				currentCell.setCellValue(((Boolean)value).booleanValue());
				lastValue = value;
			} else if( value instanceof BigDecimal ) {
				currentCell.setCellType(Cell.CELL_TYPE_NUMERIC);
				currentCell.setCellValue(((BigDecimal)value).doubleValue());				
				lastValue = value;
			} else if( value instanceof String ) {
				currentCell.setCellType(Cell.CELL_TYPE_STRING);
				currentCell.setCellValue((String)value);				
				lastValue = value;
			} else if( value != null ){
				log.debug( "Un-handled data: " + ( value == null ? "<null>" : value.toString() ) );
			}
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			break;
		case Cell.CELL_TYPE_ERROR:
			break;
		case Cell.CELL_TYPE_FORMULA:
			break;
		case Cell.CELL_TYPE_NUMERIC:
			break;
		case Cell.CELL_TYPE_STRING:
			String newValue = currentCell.getStringCellValue() + " " + value.toString();
			currentCell.setCellValue( newValue );
			lastValue = newValue;
			break;
		}
	}


/*	@Override
	public void startForeign(IForeignContent foreign) throws BirtException {

		log.debug("startForeign");
		super.startForeign(foreign);
	}
*/

/*	@Override
	public void startGroup(IGroupContent group) throws BirtException {

		log.debug("startGroup");
		super.startGroup(group);
	}
*/

	@Override
	public void startImage(IImageContent image) throws BirtException {

		byte[] data = image.getData();
		log.debug("startImage: "
				+ "[" + image.getMIMEType() +"] "
				+ "{" + image.getWidth() + " x " + image.getHeight() +"} "
				+ ( data == null ? "(no data) " : "(" + data.length + " bytes) ")
				+ image.getURI());
		super.startImage(image);
		
		String mimeType = image.getMIMEType();
		if( ( data == null ) && ( image.getURI() != null ) ) {
			try {
				URL imageUrl = new URL( image.getURI() );
				URLConnection conn = imageUrl.openConnection();
				conn.connect();
				mimeType = conn.getContentType();
				int imageType = smu.poiImageTypeFromMimeType( mimeType, null );
				if( imageType == 0 ) {
					log.debug( "Unrecognised/unhandled image MIME type: " + mimeType );
				} else {
					data = smu.downloadImage(conn);
				}
			} catch( MalformedURLException ex ) {
				log.debug( ex.getClass().getName() + ": " + ex.getMessage() );
				ex.printStackTrace();
			} catch( IOException ex ) {
				log.debug( ex.getClass().getName() + ": " + ex.getMessage() );
				ex.printStackTrace();
			}
		}
		if( data != null ) {
			int imageType = smu.poiImageTypeFromMimeType( mimeType, data );
			if( imageType == 0 ) {
				log.debug( "Unrecognised/unhandled image MIME type: " + image.getMIMEType() );
			} else {
				int imageIdx = wb.addPicture( data, imageType );
				placeImageInCurrentCell( imageIdx, image );
			}
		}
	}
	
	/**
	 * Convert a horizontal position in a column (in mm) to a ClientAnchor DX position.
	 * @param width
	 * The position within the column.
	 * @param colWidth
	 * The width of the column.
	 * @return
	 * A value suitable for use as an argument to setDx2() on ClientAnchor.
	 */
	protected abstract int anchorDxFromMM( double width, double colWidth );
	/**
	 * Convert a vertical position in a row (in points) to a ClientAnchor DY position.
	 * @param height
	 * The position within the row.
	 * @param rowHeight
	 * The height of the row.
	 * @return
	 * A value suitable for use as an argument to setDy2() on ClientAnchor.	 * 
	 */
	protected abstract int anchorDyFromPoints( float height, float rowHeight );
	
	/**
	 * <p>
	 * Prepare to place an image in the current cell.
	 * </p><p>
	 * Now that images are post-processed in endSheet() this method simply prepares the target cell (if necessary)
	 * and records the information in the images List. 
	 * </p>
	 * @param imageIdx
	 * The index for the image to be placed (as returned by Workbook.addPicture).
	 * @param image
	 * The IImageContent information provided by BIRT.
	 */
	private void placeImageInCurrentCell( int imageIdx, IImageContent image ) {
		log.debug("Adding image " + imageIdx);
		Cell oldCell = currentCell;
		if( currentCell == null ) {
			currentRow = currentSheet.createRow(rowNum);
			++rowNum;			
			currentCell = currentRow.createCell( 0 );
			currentCell.setCellType(Cell.CELL_TYPE_BLANK);
		} else {
			styleStack.mergeTop(image, ICellContent.class);
		}
		
		images.add( new CellImage(currentCell, imageIdx, image) );

		if( oldCell == null ) {
			CellStyle cellStyle = sm.getStyle(image);
			currentCell.setCellStyle(cellStyle);
			
			currentCell = null;
			currentRow = null;
		}
	}

	/**
	 * <p>
	 * Process a CellImage from the images list and place the image on the sheet.
	 * </p><p>
	 * This involves changing the row height as necesssary and determining the column spread of the image.
	 * </p>
	 * @param cellImage
	 * The image to be placed on the sheet.
	 */
	private void processCellImage( CellImage cellImage ) {
		Cell cell = cellImage.cell;
		IImageContent image = cellImage.image;
		
		float ptHeight = cell.getRow().getHeightInPoints();
		if( cellImage.image.getHeight() != null ) {
			ptHeight = smu.fontSizeInPoints( image.getHeight().toString() );
			if( ptHeight > cell.getRow().getHeightInPoints()) {
				cell.getRow().setHeightInPoints( ptHeight );
			}
		}
		// Allow image to span multiple columns
		int endCol = cell.getColumnIndex();
        double lastColWidth = ClientAnchorConversions.widthUnits2Millimetres( (short)currentSheet.getColumnWidth( endCol ) );
        int dx = anchorDxFromMM( lastColWidth, lastColWidth );

        double mmWidth = 0.0;
        if( smu.isAbsolute(image.getWidth())) {
            mmWidth = image.getWidth().convertTo(DimensionType.UNITS_MM);
        } else if(smu.isPixels(image.getWidth())) {
            mmWidth = ClientAnchorConversions.pixels2Millimetres( image.getWidth().getMeasure() );
        }
        log.debug( "Image size: " + image.getWidth() + " translates as mmWidth = " + mmWidth );
        if( mmWidth > 0) {
            double mmAccumulatedWidth = 0;
            for( endCol = cell.getColumnIndex(); mmAccumulatedWidth < mmWidth; ++ endCol ) {
                lastColWidth = ClientAnchorConversions.widthUnits2Millimetres( (short)currentSheet.getColumnWidth( endCol ) );
                mmAccumulatedWidth += lastColWidth;
                log.debug( "lastColWidth = " + lastColWidth + "; mmAccumulatedWidth = " + mmAccumulatedWidth);
            }
            if( mmAccumulatedWidth > mmWidth ) {
                mmAccumulatedWidth -= lastColWidth;
                --endCol;
                double mmShort = mmWidth - mmAccumulatedWidth;
                dx = anchorDxFromMM( mmShort, lastColWidth );
            }
        }
			
		Drawing drawing = getDrawing();
	    
		// ClientAnchor anchor = wb.getCreationHelper().createClientAnchor();
		ClientAnchor anchor = wb.getCreationHelper().createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setRow1(cell.getRowIndex());
        anchor.setCol2(endCol);
        anchor.setRow2(cell.getRowIndex());
        anchor.setDx2(dx);
        anchor.setDy2( anchorDyFromPoints( ptHeight, cell.getRow().getHeightInPoints() ) );
        anchor.setAnchorType(ClientAnchor.MOVE_DONT_RESIZE);
	    drawing.createPicture(anchor, cellImage.imageIdx);
		
	}
	
	/**
	 * Get the drawing patriarch for the current sheet (creating it as necessary).
	 * @return
	 * The drawing patriarch.
	 */
	private Drawing getDrawing() {
		if( currentDrawing == null ) {
			currentDrawing = currentSheet.createDrawingPatriarch();
		}
		return currentDrawing;
	}

	@Override
	public void startLabel(ILabelContent label) throws BirtException {

		log.debug("startLabel " + label.getLabelText() + ", style: " + smu.birtStyleToString(label.getStyle()));
		super.startLabel(label);
		
		Cell oldCell = currentCell;
		if( currentCell == null ) {
			currentRow = this.currentSheet.createRow(rowNum);
			colNum = 0;
			++rowNum;			
			currentCell = currentRow.createCell( 0 );
			currentCell.setCellType(Cell.CELL_TYPE_BLANK);
		} else {
			styleStack.mergeTop(label, ICellContent.class);
		}
		setCurrentCellContents( label.getLabelText() );
		if( oldCell == null ) {
			CellStyle cellStyle = sm.getStyle(label);
			currentCell.setCellStyle(cellStyle);
			
			currentCell = null;
			currentRow = null;
		}
	}


/*	@Override
	public void startList(IListContent list) throws BirtException {

		log.debug("startList");
		super.startList(list);
	}
*/

/*	@Override
	public void startListBand(IListBandContent listBand) throws BirtException {

		log.debug("startListBand");
		super.startListBand(listBand);
	}
*/

/*	@Override
	public void startListGroup(IListGroupContent group) throws BirtException {

		log.debug("startListGroup");
		super.startListGroup(group);
	}
*/

	/**
	 * Set up the size of the sheet based upon the page definition from BIRT
	 * @param page
	 * The BIRT page.
	 */
	protected void setupPageSize(IPageContent page) {
		PrintSetup printSetup = currentSheet.getPrintSetup();
		printSetup.setPaperSize(smu.getPaperSizeFromString(page.getPageType()));
		if( page.getOrientation() != null ) {
			log.debug( "Orientation: " + page.getOrientation() );
			if( "landscape".equals(page.getOrientation())) {
				printSetup.setLandscape(true);
			}
		}
	}
	
	/**
	 * Prepare the margin dimensions on the current sheet as per the BIRT page.
	 * @param page
	 * The BIRT page.
	 */
	protected abstract void prepareMarginDimensions(IPageContent page);
	
	@Override
	public void startRow(IRowContent row) throws BirtException {

		log.addPrefix( 'R' );
		log.debug("startRow"
/*				+ ", \nstyle: " + StyleManagerUtils.birtStyleToString(row.getStyle())
				+ ", \ninlineStyle: " + StyleManagerUtils.birtStyleToString(row.getInlineStyle())
				+ ", \ncomputedStyle: " + StyleManagerUtils.birtStyleToString(row.getComputedStyle())
*/				);
		super.startRow(row);
		++nestedRowCount;
		if( nestedRowCount == 1) {
			currentRow = this.currentSheet.createRow(rowNum);
			colNum = 0;
			styleStack.push(row);
			rowHeightChanged = false;
		}
	}


	@Override
	public void startTable(ITableContent table) throws BirtException {

		log.addPrefix( 'T' );
		log.debug("startTable, style: " + smu.birtStyleToString(table.getStyle()));;
		super.startTable(table);
		++nestedTableCount;
		if( nestedTableCount == 1 ) {
			styleStack.push(table);
			tableStartRow = rowNum;
		}
		String tableName = table.getName();
		if( tableName != null ) {
			lastTableName = tableName;
			if( sheetNum == 1 ) {
				firstSheetNamed = true;
			}
		}
	}


/*	@Override
	public void startTableBand(ITableBandContent band) throws BirtException {

		log.debug("startTableBand");
		super.startTableBand(band);
	}
*/

/*	@Override
	public void startTableGroup(ITableGroupContent group) throws BirtException {

		log.debug("startTableGroup");
		super.startTableGroup(group);
	}
*/

/*	@Override
	public void startText(ITextContent text) throws BirtException {

		prefix();
		log.debug("startText " + text.getText());
		super.startText(text);
	}
*/
	
	
	
	

}
