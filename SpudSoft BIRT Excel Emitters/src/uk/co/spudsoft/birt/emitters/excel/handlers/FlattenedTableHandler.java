package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.IRowContent;
import org.eclipse.birt.report.engine.content.ITableContent;

import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class FlattenedTableHandler extends AbstractHandler {
	
	private CellContentHandler contentHandler;

	public FlattenedTableHandler(CellContentHandler contentHandler, Logger log, IHandler parent, ITableContent table) {
		super(log, parent, table);
		this.contentHandler = contentHandler;
	}

	@Override
	public void startTable(HandlerState state, ITableContent table) throws BirtException {
		if( ( state.sheetName == null ) || state.sheetName.isEmpty() ) {
			String name = table.getName();
			if( ( name != null ) && ! name.isEmpty() ) {
				state.sheetName = name;
			}
		}
	}

	@Override
	public void endTable(HandlerState state, ITableContent table) throws BirtException {
		state.handler = this.parent;
	}

	@Override
	public void startRow(HandlerState state, IRowContent row) throws BirtException {
		state.handler = new FlattenedTableRowHandler(contentHandler, log, this, row);
		state.handler.startRow(state, row);
	}
	
}
