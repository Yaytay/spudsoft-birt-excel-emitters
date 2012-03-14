package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.eclipse.birt.report.engine.content.IListContent;

import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class NestedListHandler extends TopLevelListHandler {
	
	public NestedListHandler(Logger log, IHandler parent, IListContent list) {
		super(log, parent, list);
	}

}
