package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.IRowContent;
import org.eclipse.birt.report.engine.content.ITableContent;

import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class TopLevelTableHandler extends AbstractRealTableHandler {
	
	public TopLevelTableHandler(Logger log,IHandler parent, ITableContent table) {
		super(log, parent, table);
	}
	
	@Override
	public void startTable(HandlerState state, ITableContent table) throws BirtException {
		super.startTable(state, table);
		String name = table.getName();
		if( ( name != null ) && ! name.isEmpty() ) {
			state.sheetName = name;
		}
	}

	@Override
	public void startRow(HandlerState state, IRowContent row) throws BirtException {
		state.setHandler(new TopLevelTableRowHandler(log, this, row));
		state.getHandler().startRow(state, row);
	}
	
}
