package uk.co.spudsoft.birt.emitters.excel.handlers;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.IRowContent;
import org.eclipse.birt.report.engine.ir.DimensionType;
import org.eclipse.birt.report.model.api.util.DimensionUtil;

import uk.co.spudsoft.birt.emitters.excel.BirtStyle;
import uk.co.spudsoft.birt.emitters.excel.CellImage;
import uk.co.spudsoft.birt.emitters.excel.EmitterServices;
import uk.co.spudsoft.birt.emitters.excel.ExcelEmitter;
import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.StyleManagerUtils;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class AbstractRealTableRowHandler extends AbstractHandler {

	protected Row currentRow;
	protected int myRow;

	public AbstractRealTableRowHandler(Logger log, IHandler parent, IRowContent row) {
		super(log, parent, row);
	}

	@Override
	public void startRow(HandlerState state, IRowContent row) throws BirtException {
		resumeRow(state);
	}

	@Override
	public void endRow(HandlerState state, IRowContent row) throws BirtException {
		interruptRow(state);				
		state.setHandler(parent);
	}

	public void resumeRow(HandlerState state) {
		log.debug( "Resume row at " + state.rowNum );
		myRow = state.rowNum;
		currentRow = state.currentSheet.createRow( state.rowNum );
		state.requiredRowHeightInPoints = 0;		
	}
	
	public void interruptRow(HandlerState state) throws BirtException {
		log.debug( "Interrupt row at " + state.rowNum );
		boolean blankRow = EmitterServices.booleanOption( state.getRenderOptions(), ExcelEmitter.REMOVE_BLANK_ROWS, true );
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
		
		if(blankRow || ( currentRow.getPhysicalNumberOfCells() == 0 )) {
			log.debug( "currentRow.getPhysicalNumberOfCells() == " + currentRow.getPhysicalNumberOfCells() );
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
			
			state.getSmu().applyBordersToArea( state.getSm()
					, state.currentSheet
					, 0
					, ((IRowContent)element).getTable().getColumnCount() - 1
					, state.rowNum
					, state.rowNum
					, new BirtStyle( (IRowContent)element ) );
			
			state.rowNum += 1;
		}
	}
	
		
}