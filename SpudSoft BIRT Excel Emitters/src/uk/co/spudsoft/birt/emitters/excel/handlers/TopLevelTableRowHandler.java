package uk.co.spudsoft.birt.emitters.excel.handlers;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.content.IRowContent;
import org.eclipse.birt.report.engine.ir.DimensionType;
import org.eclipse.birt.report.model.api.util.DimensionUtil;

import uk.co.spudsoft.birt.emitters.excel.BirtStyle;
import uk.co.spudsoft.birt.emitters.excel.EmitterServices;
import uk.co.spudsoft.birt.emitters.excel.ExcelEmitter;
import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.StyleManagerUtils;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class TopLevelTableRowHandler extends AbstractHandler {

	Row currentRow;

	public TopLevelTableRowHandler(Logger log, IHandler parent, IRowContent row) {
		super(log, parent, row);
	}

	@Override
	public void startRow(HandlerState state, IRowContent row) throws BirtException {
		currentRow = state.currentSheet.createRow( state.rowNum );
		state.colNum = 0;
		state.requiredRowHeightInPoints = 0;
	}

	@Override
	public void endRow(HandlerState state, IRowContent row) throws BirtException {
		
		boolean blankRow = EmitterServices.booleanOption( state.getRenderOptions(), ExcelEmitter.REMOVE_BLANK_ROWS, true );
		for(Iterator<Cell> iter = currentRow.cellIterator(); iter.hasNext(); ) {
			Cell cell = iter.next();
			if( ! StyleManagerUtils.cellIsEmpty(cell)) {
				blankRow = false;
				break;
			}
		}
		if(blankRow) {
			state.currentSheet.removeRow(currentRow);
		} else {
			DimensionType height = row.getHeight();
			if(height != null) {
				if( DimensionUtil.isAbsoluteUnit(height.getUnits())) {
					double points = height.convertTo(DimensionType.UNITS_PT);
					currentRow.setHeightInPoints((float)points);
				}
			}
			if( state.requiredRowHeightInPoints > currentRow.getHeightInPoints() ) {
				currentRow.setHeightInPoints( state.requiredRowHeightInPoints );
			}
			
			state.getSmu().applyBordersToArea( state.getSm(), state.currentSheet, 0, row.getTable().getColumnCount() - 1, state.rowNum, state.rowNum, new BirtStyle( row ) );
			++state.rowNum;
		}
		state.handler = parent;
	}

	@Override
	public void startCell(HandlerState state, ICellContent cell) throws BirtException {
		state.handler = new TopLevelTableCellHandler(state.getEmitter(), log, this, cell);
		state.handler.startCell(state, cell);
	}

	
	
}
