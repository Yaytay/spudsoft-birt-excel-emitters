package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.ITableBandContent;
import org.eclipse.birt.report.engine.content.ITableContent;
import org.eclipse.birt.report.engine.content.ITableGroupContent;

import uk.co.spudsoft.birt.emitters.excel.BirtStyle;
import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class AbstractRealTableHandler extends AbstractHandler implements ITableHandler {

	protected int startRow;

	public AbstractRealTableHandler(Logger log, IHandler parent, ITableContent table) {
		super(log, parent, table);
	}

	@Override
	public int getColumnCount() {
		return ((ITableContent)this.element).getColumnCount();
	}

	@Override
	public void startTable(HandlerState state, ITableContent table) throws BirtException {
		startRow =  state.rowNum;
		for( int col = 0; col < table.getColumnCount(); ++col ) {
			log.debug( "BIRT table column width: " + col + " = " + table.getColumn(col).getWidth());
			if( table.getColumn(col).getWidth() != null ) {
				state.currentSheet.setColumnWidth(col, state.getSmu().poiColumnWidthFromDimension(table.getColumn(col).getWidth()));
			}
		}
	}

	@Override
	public void endTable(HandlerState state, ITableContent table) throws BirtException {
		state.setHandler(parent);

		state.getSmu().applyBordersToArea( state.getSm(), state.currentSheet, 0, table.getColumnCount() - 1, startRow, state.rowNum - 1, new BirtStyle( table ) );
	}
	
	@Override
	public void startTableBand(HandlerState state, ITableBandContent band) throws BirtException {
	}

	@Override
	public void endTableBand(HandlerState state, ITableBandContent band) throws BirtException {
	}

	@Override
	public void startTableGroup(HandlerState state, ITableGroupContent group) throws BirtException {
	}

	@Override
	public void endTableGroup(HandlerState state, ITableGroupContent group) throws BirtException {
	}

}
