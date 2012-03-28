package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.content.IRowContent;

import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class NestedTableRowHandler extends AbstractRealTableRowHandler {

	private int colOffset;

	public NestedTableRowHandler(Logger log, IHandler parent, IRowContent row, int colOffset) {
		super(log, parent, row);
		this.colOffset = colOffset;
	}

	@Override
	public void startRow(HandlerState state, IRowContent row) throws BirtException {
		log.debug( "startRow called with colOffset = ", colOffset );
		super.startRow(state, row);
		state.colNum = colOffset;
	}

	@Override
	public void startCell(HandlerState state, ICellContent cell) throws BirtException {
		log.debug( "startCell called with colOffset = ", colOffset );
		state.setHandler(new NestedTableCellHandler(state.getEmitter(), log, this, cell, colOffset));
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
