package uk.co.spudsoft.birt.emitters.excel.handlers;

import java.util.Map;

import org.apache.poi.ss.util.SheetUtil;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.ITableBandContent;
import org.eclipse.birt.report.engine.content.ITableContent;
import org.eclipse.birt.report.engine.content.ITableGroupContent;
import org.eclipse.birt.report.engine.ir.DimensionType;
import org.eclipse.birt.report.engine.ir.Expression;
import org.eclipse.birt.report.engine.ir.ReportElementDesign;

import uk.co.spudsoft.birt.emitters.excel.AreaBorders;
import uk.co.spudsoft.birt.emitters.excel.BirtStyle;
import uk.co.spudsoft.birt.emitters.excel.EmitterServices;
import uk.co.spudsoft.birt.emitters.excel.ExcelEmitter;
import uk.co.spudsoft.birt.emitters.excel.FilteredSheet;
import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class AbstractRealTableHandler extends AbstractHandler implements ITableHandler {

	protected int startRow;
	protected int startDetailsRow = -1;
	protected int endDetailsRow;
	
	private ITableGroupContent currentGroup;
	private ITableBandContent currentBand;
	
	private BirtStyle tableStyle;
	private AreaBorders borderDefn;

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
			DimensionType width = table.getColumn(col).getWidth();
			log.debug( "BIRT table column width: ", col, " = ", width);
			if( width != null ) {
				int newWidth = state.getSmu().poiColumnWidthFromDimension(width);
				int oldWidth = state.currentSheet.getColumnWidth(col);
				if( ( oldWidth == 256 * state.currentSheet.getDefaultColumnWidth() ) || ( newWidth > oldWidth ) ) {
					state.currentSheet.setColumnWidth(col, newWidth);
				}
			}
		}
		
		tableStyle = new BirtStyle( table );
		borderDefn = AreaBorders.create( -1, 0, table.getColumnCount() - 1, startRow, tableStyle );
		if( borderDefn != null ) {
			state.insertBorderOverload(borderDefn);
		}
	}
	
	private static boolean extractForceAutoColWidth(ITableContent table) {
		Object generatorObject = table.getGenerateBy();
		if( generatorObject instanceof ReportElementDesign ) {
			ReportElementDesign generatorDesign = (ReportElementDesign)generatorObject;
			Map<String,Expression> userProps = generatorDesign.getUserProperties(); 
			if( userProps != null ) {
				Expression expression = userProps.get( ExcelEmitter.FORCEAUTOCOLWIDTHS_PROP );
				if( expression instanceof Expression.Constant ) {
					Expression.Constant constant = (Expression.Constant)expression;
					return EmitterServices.booleanOption(constant.getValue(), false);
				}
			}
		}
		return false;
	}	

	@Override
	public void endTable(HandlerState state, ITableContent table) throws BirtException {
		state.setHandler(parent);

		state.getSmu().applyBottomBorderToRow( state.getSm(), state.currentSheet, 0, table.getColumnCount() - 1, state.rowNum - 1, new BirtStyle( table ) );
		
		if( borderDefn != null ) {
			state.removeBorderOverload(borderDefn);
		}
		
		// state.getSmu().applyBordersToArea( state.getSm(), state.currentSheet, 0, table.getColumnCount() - 1, startRow, state.rowNum - 1, new BirtStyle( table ) );

		log.debug( "Details rows from ", startDetailsRow, " to ", endDetailsRow );
		
		if( ( startDetailsRow > 0 ) && ( endDetailsRow > startDetailsRow ) ) {
			boolean forceAutoColWidths = extractForceAutoColWidth( table );
			for( int col = 0; col < table.getColumnCount(); ++col ) {
				int oldWidth = state.currentSheet.getColumnWidth(col);
				if( forceAutoColWidths || ( oldWidth == 256 * state.currentSheet.getDefaultColumnWidth() ) ) {
					FilteredSheet filteredSheet = new FilteredSheet( state.currentSheet, startDetailsRow, Math.min(endDetailsRow, startDetailsRow + 12) );
			        double calcWidth = SheetUtil.getColumnWidth( filteredSheet, col, false );

			        if (calcWidth > 1.0) {
			        	calcWidth *= 256;
			            int maxColumnWidth = 255*256; // The maximum column width for an individual cell is 255 characters
			            if (calcWidth > maxColumnWidth) {
			            	calcWidth = maxColumnWidth;
			            }
			            if( calcWidth > oldWidth ) {
			            	state.currentSheet.setColumnWidth( col, (int)(calcWidth) );
			            }
			        }
				}
			}
		}
		
		if( table.getBookmark() != null ) {
			createName(state, prepareName( table.getBookmark() ), startRow, 0, state.rowNum - 1, table.getColumnCount() - 1);
		}
		
	}

	@Override
	public void startTableBand(HandlerState state, ITableBandContent band) throws BirtException {
		if( ( band.getBandType() == ITableBandContent.BAND_DETAIL ) && ( startDetailsRow < 0 ) ) {
			startDetailsRow = state.rowNum;
		}
		currentBand = band;
	}

	@Override
	public void endTableBand(HandlerState state, ITableBandContent band) throws BirtException {
		if( band.getBandType() == ITableBandContent.BAND_DETAIL ) {
			endDetailsRow = state.rowNum - 1;
		}
		currentBand = null;
	}

	@Override
	public void startTableGroup(HandlerState state, ITableGroupContent group) throws BirtException {
		currentGroup = group;
	}

	@Override
	public void endTableGroup(HandlerState state, ITableGroupContent group) throws BirtException {
		currentGroup = null;
	}

}
