package uk.co.spudsoft.birt.emitters.excel.handlers;

import java.util.Map;

import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.IListBandContent;
import org.eclipse.birt.report.engine.content.IListContent;
import org.eclipse.birt.report.engine.content.IListGroupContent;
import org.eclipse.birt.report.engine.ir.Expression;
import org.eclipse.birt.report.engine.ir.ReportElementDesign;

import uk.co.spudsoft.birt.emitters.excel.AreaBorders;
import uk.co.spudsoft.birt.emitters.excel.BirtStyle;
import uk.co.spudsoft.birt.emitters.excel.EmitterServices;
import uk.co.spudsoft.birt.emitters.excel.ExcelEmitter;
import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class AbstractRealListHandler extends AbstractHandler {

	protected int startRow;
	
	private IListGroupContent currentGroup;
	private IListBandContent currentBand;
	
	private AreaBorders borderDefn;

	public AbstractRealListHandler(Logger log, IHandler parent, IListContent list) {
		super(log, parent, list);
	}

	@Override
	public void startList(HandlerState state, IListContent list) throws BirtException {
		startRow =  state.rowNum;
	}
	
	@Override
	public void endList(HandlerState state, IListContent list) throws BirtException {
		state.setHandler(parent);

		int endRow = state.rowNum - 1;
		int colStart = 0;
		int colEnd = 0;
		
		for( int row = startRow; row < endRow; ++row ) {
			int lastColInRow = state.currentSheet.getRow(row).getLastCellNum() - 1;
			if( lastColInRow > colEnd ) {
				colEnd = lastColInRow;
			}
		}
		
		state.getSmu().applyBordersToArea( state.getSm(), state.currentSheet, colStart, colEnd, startRow, endRow, new BirtStyle( list ) );
		
		if( borderDefn != null ) {
			state.removeBorderOverload(borderDefn);
		}
		
		Map<String,Expression> userProperties = null;
		Object generatorObject = list.getGenerateBy();
		if( generatorObject instanceof ReportElementDesign ) {
			ReportElementDesign generatorDesign = (ReportElementDesign)generatorObject;
			userProperties = generatorDesign.getUserProperties(); 
		}		
		
		if( list.getBookmark() != null ) {
			createName(state, prepareName( list.getBookmark() ), startRow, 0, state.rowNum - 1, 0);
		}
		
		if( EmitterServices.booleanOption( null, userProperties, ExcelEmitter.DISPLAYFORMULAS_PROP, false ) ) {
			state.currentSheet.setDisplayFormulas(true);
		}
		if( ! EmitterServices.booleanOption( null, userProperties, ExcelEmitter.DISPLAYGRIDLINES_PROP, true ) ) {
			state.currentSheet.setDisplayGridlines(false);
		}
		if( ! EmitterServices.booleanOption( null, userProperties, ExcelEmitter.DISPLAYROWCOLHEADINGS_PROP, true ) ) {
			state.currentSheet.setDisplayRowColHeadings(false);
		}
		if( ! EmitterServices.booleanOption( null, userProperties, ExcelEmitter.DISPLAYZEROS_PROP, true ) ) {
			state.currentSheet.setDisplayZeros(false);
		}
	}

	@Override
	public void startListBand(HandlerState state, IListBandContent band) throws BirtException {
		currentBand = band;
	}

	@Override
	public void endListBand(HandlerState state, IListBandContent band) throws BirtException {
		currentBand = null;
	}

	@Override
	public void startListGroup(HandlerState state, IListGroupContent group) throws BirtException {
		currentGroup = group;
	}

	@Override
	public void endListGroup(HandlerState state, IListGroupContent group) throws BirtException {
		currentGroup = null;
	}

}
