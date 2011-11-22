package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.content.IRowContent;

import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class NestedTableRowHandler extends AbstractRealTableRowHandler {

	public NestedTableRowHandler(Logger log, IHandler parent, IRowContent row) {
		super(log, parent, row);
	}

	@Override
	public void startRow(HandlerState state, IRowContent row) throws BirtException {
		super.startRow(state, row);
		++state.rowOffset;
		state.colNum = 0;
	}

	@Override
	public void startCell(HandlerState state, ICellContent cell) throws BirtException {
		state.setHandler(new NestedTableCellHandler(state.getEmitter(), log, this, cell));
		state.getHandler().startCell(state, cell);
	}
}
