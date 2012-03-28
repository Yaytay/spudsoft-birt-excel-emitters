package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.IAutoTextContent;
import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.content.IContainerContent;
import org.eclipse.birt.report.engine.content.IDataContent;
import org.eclipse.birt.report.engine.content.IForeignContent;
import org.eclipse.birt.report.engine.content.IImageContent;
import org.eclipse.birt.report.engine.content.ILabelContent;
import org.eclipse.birt.report.engine.content.IListContent;
import org.eclipse.birt.report.engine.content.ITableContent;
import org.eclipse.birt.report.engine.content.ITextContent;
import org.eclipse.birt.report.engine.content.impl.CellContent;
import org.eclipse.birt.report.engine.css.engine.StyleConstants;
import org.eclipse.birt.report.engine.emitter.IContentEmitter;
import org.eclipse.birt.report.engine.ir.CellDesign;
import org.eclipse.birt.report.engine.ir.GridItemDesign;
import org.eclipse.birt.report.engine.layout.pdf.util.HTML2Content;

import uk.co.spudsoft.birt.emitters.excel.Coordinate;
import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class AbstractRealTableCellHandler extends CellContentHandler {

	protected int column;
	private AbstractRealTableRowHandler parentRow;
	private boolean containsTable;
	
	public AbstractRealTableCellHandler(IContentEmitter emitter, Logger log, IHandler parent, ICellContent cell) {
		super(emitter, log, parent, cell);
		column = cell.getColumn();
	}

	@Override
	public void notifyHandler(HandlerState state) {
		if( parentRow != null ) {
			parentRow.resumeRow(state);
			resumeCell(state);
			parentRow = null;
		}
	}

	@Override
	public void startCell(HandlerState state, ICellContent cell) throws BirtException {
		log.debug( "Cell - " 
				+ "BIRT[" 
					+ cell.getRow()
					+ ( cell.getRowSpan() > 1 ? "-" + (cell.getRow() + cell.getRowSpan()-1) : "" )
					+ ","
					+ cell.getColumn() 
					+ ( cell.getColSpan() > 1 ? "-" + (cell.getColumn() + cell.getColSpan()-1) : "" )
					+ "]"
				+ " state[" 
					+ state.rowNum
					+ ","
					+ state.colNum 
					+ "]"
				);
		resumeCell(state);
	}
	
	@Override
	public void endCell(HandlerState state, ICellContent cell) throws BirtException {
		if( cell.getBookmark() != null ) {
			createName(state, prepareName( cell.getBookmark() ), state.rowNum, state.colNum, state.rowNum, state.colNum);
		}

		interruptCell(state, ! containsTable );
	}
	
	public void resumeCell(HandlerState state) {
	}
	
	public void interruptCell(HandlerState state, boolean includeFormatOnly) throws BirtException {
		
		if( state == null ) {
			System.err.println( "state == null" );
		} else if( state.currentSheet == null ) {
			System.err.println( "state.currentSheet == null" );
		} else if( state.currentSheet.getRow(state.rowNum) == null ) {
			System.err.println( "state.currentSheet.getRow(" + state.rowNum + ") == null" );			
		}

		if( ( lastValue != null ) || includeFormatOnly ) {
			Cell currentCell = state.currentSheet.getRow(state.rowNum).getCell( column );
			if( currentCell == null ) {
				log.debug( "Creating cell[", state.rowNum, ",", column, "]");
				currentCell = state.currentSheet.getRow(state.rowNum).createCell( column );
			}
					
			ICellContent cell = (ICellContent)element; 
			
			if(( cell.getColSpan() > 1 )||( cell.getRowSpan() > 1 )) {

                int endRow = state.rowNum + cell.getRowSpan() - 1;
				int endCol = state.colNum + cell.getColSpan() - 1;
                
                if(cell.getRowSpan() > 1) {
                	log.debug( "Adding row span [", state.rowNum, ",", state.colNum, "] to [", endRow, ",", endCol, "]" );
                    state.addRowSpan(state.rowNum, state.colNum, endRow, endCol);
                }
				
                int offset = state.computeNumberSpanBefore(state.rowNum, state.colNum);
                log.debug( "Offset for [", state.rowNum, ",", state.colNum, "] calculated as ", offset);
                log.debug( "Merging [", state.rowNum, ",", state.colNum + offset, "] to [", endRow, ",", endCol + offset, "]" );
                log.debug( "Should be merging ? [", state.rowNum, ",", column, "] to [", endRow, ",", column + cell.getColSpan() - 1, "]" );
                // CellRangeAddress newMergedRegion = new CellRangeAddress( state.rowNum, endRow, state.colNum + offset, endCol + offset );
                CellRangeAddress newMergedRegion = new CellRangeAddress( state.rowNum, endRow, column, column + cell.getColSpan() - 1 );
                state.currentSheet.addMergedRegion( newMergedRegion );
                                
				colSpan = cell.getColSpan();
			}
	
			endCellContent(state, cell, lastElement, currentCell);
		}
		
		if( state.cellIsMergedWithBorders( state.rowNum, column ) ) {
			int absoluteColumn = column;
			++state.colNum;
			--colSpan;
			while( colSpan > 0 ) {
				++absoluteColumn;
				log.debug( "Creating cell[", state.rowNum, ",", absoluteColumn, "]");
				Cell currentCell = state.currentSheet.getRow(state.rowNum).createCell( absoluteColumn );
				endCellContent(state, null, null, currentCell);
				++state.colNum;
				--colSpan;
			}
		} else {
			state.colNum += colSpan;
		}
		
		state.setHandler(parent);
	}
	
	
	@Override
	public void startContainer(HandlerState state, IContainerContent container) throws BirtException {
		// log.debug( "Container display = " + getStyleProperty( container, StyleConstants.STYLE_DISPLAY, "block") );
		if( ! "inline".equals( getStyleProperty( container, StyleConstants.STYLE_DISPLAY, "block") ) ) {
			lastCellContentsWasBlock = true;
		}
	}
	
	@Override
	public void endContainer(HandlerState state, IContainerContent container) throws BirtException {
		// log.debug( "Container display = " + getStyleProperty( container, StyleConstants.STYLE_DISPLAY, "block") );
		if( ! "inline".equals( getStyleProperty( container, StyleConstants.STYLE_DISPLAY, "block") ) ) {
			lastCellContentsWasBlock = true;
		}
	}

	@Override
	public void startTable(HandlerState state, ITableContent table) throws BirtException {
		int colSpan = ((ICellContent)element).getColSpan();
		int rowSpan = ((ICellContent)element).getRowSpan();
		ITableHandler tableHandler = getAncestor(ITableHandler.class);
		if( ( tableHandler != null ) 
				&& ( tableHandler.getColumnCount() == colSpan )
				&& ( 1 == ( (CellDesign)( (CellContent)table.getParent() ).getGenerateBy()).getContentCount() )
				) {
			
			containsTable = true;
			parentRow = getAncestor(AbstractRealTableRowHandler.class);
			interruptCell(state, false);
			parentRow.interruptRow(state);
			
			state.setHandler(new NestedTableHandler(log, this, table, rowSpan));
			state.getHandler().startTable(state, table);
		} else if( ( tableHandler != null ) 
				&& ( table.getGenerateBy() instanceof GridItemDesign )
				&& ( ((GridItemDesign)table.getGenerateBy()).getColumnCount() == colSpan )
				) {
			
			containsTable = true;
			parentRow = getAncestor(AbstractRealTableRowHandler.class);
			interruptCell(state, false);
			
			removeMergedCell( state, state.rowNum, state.colNum );
			
			NestedTableHandler nestedTableHandler = new NestedTableHandler(log, this, table, rowSpan);
			nestedTableHandler.setInserted(true);
			state.setHandler(nestedTableHandler);
			state.getHandler().startTable(state, table);

		} else {
			state.setHandler(new FlattenedTableHandler(this, log, this, table));
			state.getHandler().startTable(state, table);
		}
	}
	
	@Override
	public void startList(HandlerState state, IListContent list) throws BirtException {
		int colSpan = ((ICellContent)element).getColSpan();
		ITableHandler tableHandler = getAncestor(ITableHandler.class);
		if( ( tableHandler != null ) 
				&& ( tableHandler.getColumnCount() == colSpan )
				&& ( 1 == ( (CellDesign)( (CellContent)list.getParent() ).getGenerateBy()).getContentCount() )
				) {
			
			containsTable = true;
			parentRow = getAncestor(AbstractRealTableRowHandler.class);
			interruptCell(state, false);
			parentRow.interruptRow(state);
			
			state.colNum = column;
			state.setHandler(new NestedListHandler(log, this, list));
			state.getHandler().startList(state, list);
		} else {
			state.setHandler(new FlattenedListHandler(this, log, this, list));
			state.getHandler().startList(state, list);
		}
	}

	@Override
	public void emitText(HandlerState state, ITextContent text) throws BirtException {
		String textText = text.getText();
		log.debug( "text:", textText );
		emitContent(state,text,textText, ( ! "inline".equals( getStyleProperty(text, StyleConstants.STYLE_DISPLAY, "block") ) ) );
	}

	@Override
	public void emitData(HandlerState state, IDataContent data) throws BirtException {
		emitContent(state,data,data.getValue(), ( ! "inline".equals( getStyleProperty(data, StyleConstants.STYLE_DISPLAY, "block") ) ) );
	}

	@Override
	public void emitLabel(HandlerState state, ILabelContent label) throws BirtException {
		// String labelText = ( label.getLabelText() != null ) ? label.getLabelText() : label.getText();
		String labelText = ( label.getText() != null ) ? label.getText() : label.getLabelText();
		log.debug( "labelText:", labelText );
		emitContent(state,label,labelText, ( ! "inline".equals( getStyleProperty(label, StyleConstants.STYLE_DISPLAY, "block") ) ));
	}

	@Override
	public void emitAutoText(HandlerState state, IAutoTextContent autoText) throws BirtException {
		emitContent(state,autoText,autoText.getText(), ( ! "inline".equals( getStyleProperty(autoText, StyleConstants.STYLE_DISPLAY, "block") ) ) );
	}

	@Override
	public void emitForeign(HandlerState state, IForeignContent foreign) throws BirtException {

		log.debug( "Handling foreign content of type ", foreign.getRawType() );
		if ( IForeignContent.HTML_TYPE.equalsIgnoreCase( foreign.getRawType( ) ) )
		{
			HTML2Content.html2Content( foreign );
			contentVisitor.visitChildren( foreign, null );			
		}
	}

	@Override
	public void emitImage(HandlerState state, IImageContent image) throws BirtException {
		boolean imageCanSpan = false;
		
		int colSpan = ((ICellContent)element).getColSpan();
		ITableHandler tableHandler = getAncestor(ITableHandler.class);
		if( ( tableHandler != null ) 
				&& ( tableHandler.getColumnCount() == colSpan )
				&& ( 1 == ( (CellDesign)( (CellContent)image.getParent() ).getGenerateBy()).getContentCount() )
				) {
			imageCanSpan = true;
		}
		recordImage(state, new Coordinate( state.rowNum, column ), image, imageCanSpan);
		lastElement = image;
	}
			
}
