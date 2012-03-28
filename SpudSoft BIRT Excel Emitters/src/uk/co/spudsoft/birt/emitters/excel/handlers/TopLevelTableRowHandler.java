package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.content.IRowContent;

import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class TopLevelTableRowHandler extends AbstractRealTableRowHandler {

	public TopLevelTableRowHandler(Logger log, IHandler parent, IRowContent row) {
		super(log, parent, row);
	}
	
	@Override
	public void startRow(HandlerState state, IRowContent row) throws BirtException {
		super.startRow(state, row);
		state.colNum = 0;
		state.rowOffset = 0;
	}

	@Override
	public void startCell(HandlerState state, ICellContent cell) throws BirtException {
		state.setHandler(new TopLevelTableCellHandler(state.getEmitter(), log, this, cell));
		state.getHandler().startCell(state, cell);
	}
	
	@Override
	protected boolean isNested() {
		return false;
	}
}
