package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.emitter.IContentEmitter;

import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class TopLevelTableCellHandler extends AbstractRealTableCellHandler {
	
	public TopLevelTableCellHandler(IContentEmitter emitter, Logger log, IHandler parent, ICellContent cell) {
		super(emitter, log, parent, cell);
	}

}
