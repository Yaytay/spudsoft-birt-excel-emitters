package uk.co.spudsoft.birt.emitters.excel.handlers;

import java.util.Map;
import java.util.Stack;

import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.IRowContent;
import org.eclipse.birt.report.engine.content.ITableContent;
import org.eclipse.birt.report.engine.content.ITableGroupContent;
import org.eclipse.birt.report.engine.ir.Expression;
import org.eclipse.birt.report.engine.ir.ReportElementDesign;

import uk.co.spudsoft.birt.emitters.excel.EmitterServices;
import uk.co.spudsoft.birt.emitters.excel.ExcelEmitter;
import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class TopLevelTableHandler extends AbstractRealTableHandler {
	
	private Stack<Integer> groupStarts;
	
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

	@Override
	public void startTableGroup(HandlerState state, ITableGroupContent group) throws BirtException {
		if( groupStarts == null ) {
			groupStarts = new Stack<Integer>();
		}
		groupStarts.push(state.rowNum);
	}

	@Override
	public void endTableGroup(HandlerState state, ITableGroupContent group) throws BirtException {
		int start = groupStarts.pop();
		if( start < state.rowNum - 2 ) {
			
			boolean disableGrouping = false;
			
			// Report user props and context first
			Map<String,Expression> userProperties = ((ITableContent)this.element).getReportContent().getDesign().getUserProperties();
			if( EmitterServices.booleanOption( state.getRenderOptions(), userProperties, ExcelEmitter.DISABLE_GROUPING, false ) ) {
				disableGrouping = true;
			}
			
			// Then table user props
			Object generatorObject = ((ITableContent)this.element).getGenerateBy();
			if( generatorObject instanceof ReportElementDesign ) {
				ReportElementDesign generatorDesign = (ReportElementDesign)generatorObject;
				userProperties = generatorDesign.getUserProperties(); 

				if( EmitterServices.booleanOption( null, userProperties, ExcelEmitter.DISABLE_GROUPING, false ) ) {
					disableGrouping = true;
				}
			}		

			if( ! disableGrouping ) {
				state.currentSheet.groupRow(start, state.rowNum - 2);
			}
		}
	}
	
}
