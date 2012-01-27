package uk.co.spudsoft.birt.emitters.excel.handlers;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.IRowContent;
import org.eclipse.birt.report.engine.ir.DimensionType;
import org.eclipse.birt.report.model.api.util.DimensionUtil;

import uk.co.spudsoft.birt.emitters.excel.AreaBorders;
import uk.co.spudsoft.birt.emitters.excel.BirtStyle;
import uk.co.spudsoft.birt.emitters.excel.CellImage;
import uk.co.spudsoft.birt.emitters.excel.EmitterServices;
import uk.co.spudsoft.birt.emitters.excel.ExcelEmitter;
import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.StyleManagerUtils;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class AbstractRealTableRowHandler extends AbstractHandler {

	protected Row currentRow;
	protected int birtRowStartedAtPoiRow;
	protected int myRow;

	private BirtStyle rowStyle;
	private AreaBorders borderDefn;

	public AbstractRealTableRowHandler(Logger log, IHandler parent, IRowContent row) {
		super(log, parent, row);
	}

	@Override
	public void startRow(HandlerState state, IRowContent row) throws BirtException {
		birtRowStartedAtPoiRow = state.rowNum;
		resumeRow(state);
	}

	@Override
	public void endRow(HandlerState state, IRowContent row) throws BirtException {
		interruptRow(state);
		if( row.getBookmark() != null ) {
			createName(state, prepareName( row.getBookmark() ), birtRowStartedAtPoiRow, 0, state.rowNum - 1, currentRow.getLastCellNum() - 1 );
		}
		
		state.setHandler(parent);
	}

	public void resumeRow(HandlerState state) {
		log.debug( "Resume row at ", state.rowNum );

		myRow = state.rowNum;
		currentRow = state.currentSheet.createRow( state.rowNum );
		state.requiredRowHeightInPoints = 0;		
		
		rowStyle = new BirtStyle( (IRowContent)element );
		borderDefn = AreaBorders.create( myRow, 0, ((IRowContent)element).getTable().getColumnCount() - 1, myRow, rowStyle );
		if( borderDefn != null ) {
			state.insertBorderOverload(borderDefn);
		}
	}
	
	public void interruptRow(HandlerState state) throws BirtException {
		log.debug( "Interrupt row at ", state.rowNum );

		boolean blankRow = EmitterServices.booleanOption( state.getRenderOptions(), null, ExcelEmitter.REMOVE_BLANK_ROWS, true );

		if( state.rowHasMergedCellsWithBorders( state.rowNum ) ) {
			for( AreaBorders areaBorder : state.areaBorders ) {
				if( ( areaBorder.isMergedCells ) 
						&& ( areaBorder.top <= state.rowNum )
						&& ( areaBorder.bottom >= state.rowNum ) ) {

					for( int column = areaBorder.left; column <= areaBorder.right; ++column ) {
						if( currentRow.getCell(column) == null ) {
							BirtStyle birtCellStyle = new BirtStyle( state.getSm().getCssEngine() );
							Cell cell = state.currentSheet.getRow(state.rowNum).createCell( column );
							state.getSmu().applyAreaBordersToCell(state.areaBorders, cell, birtCellStyle, state.rowNum, column);
							CellStyle cellStyle = state.getSm().getStyle(birtCellStyle);
							cell.setCellStyle(cellStyle);
						}
					}
				}
			}			
			blankRow = false;
		}
		
		
		if( blankRow ) {
			for(Iterator<Cell> iter = currentRow.cellIterator(); iter.hasNext(); ) {
				Cell cell = iter.next();
				if( ! StyleManagerUtils.cellIsEmpty(cell)) {
					blankRow = false;
					break;
				}
			}
		}
		if( blankRow ) {
			for( CellImage cellImage : state.images ) {
				if( cellImage.location.getRow() == state.rowNum ) {
					blankRow = false;
					break;
				}
			}
		}
        if( blankRow ) {
            if(state.computeNumberSpanBefore(state.rowNum, state.colNum) > 0) {
                //this row is part of a row span. Dont delete it.
                blankRow = false;
            }
        }
        if( blankRow ) {
        	if( ((IRowContent)element).getBookmark() != null ) {
        		blankRow = false;
        	}
        }
        
		if(blankRow || ( currentRow.getPhysicalNumberOfCells() == 0 )) {
			state.currentSheet.removeRow(currentRow);
		} else {
			DimensionType height = ((IRowContent)element).getHeight();
			if(height != null) {
				if( DimensionUtil.isAbsoluteUnit(height.getUnits())) {
					double points = height.convertTo(DimensionType.UNITS_PT);
					currentRow.setHeightInPoints((float)points);
				}
			}
			if( state.requiredRowHeightInPoints > currentRow.getHeightInPoints() ) {
				currentRow.setHeightInPoints( state.requiredRowHeightInPoints );
			}
			
			state.rowNum += 1;
		}
		
		if( borderDefn != null ) {
			state.removeBorderOverload(borderDefn);
			borderDefn = null;
		}
	}
	
		
}
