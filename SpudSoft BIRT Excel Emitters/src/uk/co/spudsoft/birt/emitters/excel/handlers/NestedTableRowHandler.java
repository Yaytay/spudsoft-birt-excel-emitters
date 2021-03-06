package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.content.IRowContent;

import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class NestedTableRowHandler extends AbstractRealTableRowHandler {

	public NestedTableRowHandler(Logger log, IHandler parent, IRowContent row, int startCol) {
		super(log, parent, row, startCol);
	}

	@Override
	public void startRow(HandlerState state, IRowContent row) throws BirtException {
		log.debug( "startRow called with colOffset = ", startCol );
		super.startRow(state, row);
	}

	@Override
	public void startCell(HandlerState state, ICellContent cell) throws BirtException {
		log.debug( "startCell called with colOffset = ", startCol );
		state.setHandler(new NestedTableCellHandler(state.getEmitter(), log, this, cell, startCol));
		state.getHandler().startCell(state, cell);
	}

	@Override
	public void endRow(HandlerState state, IRowContent row) throws BirtException {
		super.endRow(state, row);
	}

	@Override
	protected boolean isNested() {
		return true;
	}
	
}
