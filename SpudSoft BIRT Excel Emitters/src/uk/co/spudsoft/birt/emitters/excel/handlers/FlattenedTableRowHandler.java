package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.content.IRowContent;

import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class FlattenedTableRowHandler extends AbstractHandler {

	private CellContentHandler contentHandler;

	public FlattenedTableRowHandler(CellContentHandler contentHandler, Logger log, IHandler parent, IRowContent row) {
		super(log, parent, row);
		this.contentHandler = contentHandler;
	}

	@Override
	public void startRow(HandlerState state, IRowContent row) throws BirtException {
		contentHandler.lastCellContentsWasBlock = true;
	}

	@Override
	public void endRow(HandlerState state, IRowContent row) throws BirtException {
		contentHandler.lastCellContentsWasBlock = true;
		state.setHandler(parent);
	}

	@Override
	public void startCell(HandlerState state, ICellContent cell) throws BirtException {
		state.setHandler(new FlattenedTableCellHandler(contentHandler, log, this, cell));
		state.getHandler().startCell(state, cell);
	}
	
	

}
