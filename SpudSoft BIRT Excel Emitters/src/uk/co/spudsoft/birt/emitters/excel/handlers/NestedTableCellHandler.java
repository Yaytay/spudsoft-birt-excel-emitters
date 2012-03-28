package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.emitter.IContentEmitter;

import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class NestedTableCellHandler extends AbstractRealTableCellHandler {

	public NestedTableCellHandler(IContentEmitter emitter, Logger log, IHandler parent, ICellContent cell, int colOffset) {
		super(emitter, log, parent, cell);
		column = cell.getColumn() + colOffset;
	}
}
