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
import org.eclipse.birt.report.engine.content.ITableContent;
import org.eclipse.birt.report.engine.content.ITextContent;
import org.eclipse.birt.report.engine.css.engine.StyleConstants;
import org.eclipse.birt.report.engine.emitter.IContentEmitter;
import org.eclipse.birt.report.engine.layout.pdf.util.HTML2Content;

import uk.co.spudsoft.birt.emitters.excel.Coordinate;
import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class TopLevelTableCellHandler extends CellContentHandler {
	
	int column;

	public TopLevelTableCellHandler(IContentEmitter emitter, Logger log, IHandler parent, ICellContent cell) {
		super(emitter, log, parent, cell);
	}

	@Override
	public void startCell(HandlerState state, ICellContent cell) throws BirtException {
		column = cell.getColumn();
	}
	
	@Override
	public void endCell(HandlerState state, ICellContent cell) throws BirtException {
		
		Cell currentCell = state.currentSheet.getRow(state.rowNum).createCell( column );
		currentCell.setCellType(Cell.CELL_TYPE_BLANK);
				
		if(( cell.getColSpan() > 1 )||( cell.getRowSpan() > 1 )) {
			int endRow = state.rowNum + cell.getRowSpan() - 1;
			int endCol = state.colNum + cell.getColSpan() - 1;
			
/*			System.out.println( "addMergedRegion( "
					+ "" + state.rowNum 
					+ ", " + endRow
					+ ", " + state.colNum
					+ ", " + endCol
					+ " )" 
					+ " [ " + cell.getRowSpan() + " & " + cell.getColSpan() + " ]" 
					);
*/			state.currentSheet.addMergedRegion( new CellRangeAddress( state.rowNum, endRow, state.colNum, endCol ) );
			colSpan = cell.getColSpan();
		}

		endCellContent(state, cell, lastElement, currentCell);

		state.colNum += cell.getColSpan();
		
		state.handler = this.parent;
	}
	
	
	@Override
	public void startContainer(HandlerState state, IContainerContent container) throws BirtException {
		log.debug( "Container display = " + getStyleProperty( container, StyleConstants.STYLE_DISPLAY, "block") );
		if( ! "inline".equals( getStyleProperty( container, StyleConstants.STYLE_DISPLAY, "block") ) ) {
			lastCellContentsWasBlock = true;
		}
	}
	
	@Override
	public void endContainer(HandlerState state, IContainerContent container) throws BirtException {
		log.debug( "Container display = " + getStyleProperty( container, StyleConstants.STYLE_DISPLAY, "block") );
		if( ! "inline".equals( getStyleProperty( container, StyleConstants.STYLE_DISPLAY, "block") ) ) {
			lastCellContentsWasBlock = true;
		}
	}

	@Override
	public void startTable(HandlerState state, ITableContent table) throws BirtException {
		state.handler = new FlattenedTableHandler(this, log, this, table);
		state.handler.startTable(state, table);
	}

	@Override
	public void emitText(HandlerState state, ITextContent text) throws BirtException {
		log.debug( "text:" + text.getText() );
		emitContent(state,text,text.getText(), ( ! "inline".equals( getStyleProperty(text, StyleConstants.STYLE_DISPLAY, "block") ) ) );
	}

	@Override
	public void emitData(HandlerState state, IDataContent data) throws BirtException {
		emitContent(state,data,data.getValue(), ( ! "inline".equals( getStyleProperty(data, StyleConstants.STYLE_DISPLAY, "block") ) ) );
	}

	@Override
	public void emitLabel(HandlerState state, ILabelContent label) throws BirtException {
		String labelText = ( label.getLabelText() != null ) ? label.getLabelText() : label.getText();
		log.debug( "labelText:" + labelText );
		emitContent(state,label,labelText, ( ! "inline".equals( getStyleProperty(label, StyleConstants.STYLE_DISPLAY, "block") ) ));
	}

	@Override
	public void emitAutoText(HandlerState state, IAutoTextContent autoText) throws BirtException {
		emitContent(state,autoText,autoText.getText(), ( ! "inline".equals( getStyleProperty(autoText, StyleConstants.STYLE_DISPLAY, "block") ) ) );
	}

	@Override
	public void emitForeign(HandlerState state, IForeignContent foreign) throws BirtException {

		log.debug( "Handling foreign content of type " + foreign.getRawType() );
		if ( IForeignContent.HTML_TYPE.equalsIgnoreCase( foreign.getRawType( ) ) )
		{
			HTML2Content.html2Content( foreign );
			contentVisitor.visitChildren( foreign, null );			
		}
	}

	@Override
	public void emitImage(HandlerState state, IImageContent image) throws BirtException {
		recordImage(state, new Coordinate( state.rowNum, column ), image, false);
		lastElement = image;
	}
		
}
