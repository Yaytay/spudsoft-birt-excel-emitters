package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.IAutoTextContent;
import org.eclipse.birt.report.engine.content.IDataContent;
import org.eclipse.birt.report.engine.content.IForeignContent;
import org.eclipse.birt.report.engine.content.IImageContent;
import org.eclipse.birt.report.engine.content.ILabelContent;
import org.eclipse.birt.report.engine.content.IPageContent;
import org.eclipse.birt.report.engine.content.ITableContent;
import org.eclipse.birt.report.engine.content.ITextContent;
import org.eclipse.birt.report.engine.ir.DimensionType;

import uk.co.spudsoft.birt.emitters.excel.CellImage;
import uk.co.spudsoft.birt.emitters.excel.ClientAnchorConversions;
import uk.co.spudsoft.birt.emitters.excel.Coordinate;
import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.StyleManagerUtils;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class PageHandler extends AbstractHandler {

	public PageHandler(Logger log, IPageContent page) {
		super(log, null, page);
	}

	private void setupPageSize(HandlerState state, IPageContent page) {
		PrintSetup printSetup = state.currentSheet.getPrintSetup();
		printSetup.setPaperSize(state.getSmu().getPaperSizeFromString(page.getPageType()));
		if( page.getOrientation() != null ) {
			log.debug( "Orientation: " + page.getOrientation() );
			if( "landscape".equals(page.getOrientation())) {
				printSetup.setLandscape(true);
			}
		}
	}
	
	@Override
	public void startPage(HandlerState state, IPageContent page) throws BirtException {
	    state.currentSheet = state.getWb().createSheet();
		log.debug("Page type: " + page.getPageType());
		
		if( page.getPageType() != null ) {
			setupPageSize(state, page);
		}
		
		state.getSmu().prepareMarginDimensions(state.currentSheet, page);
	}

	@Override
	public void endPage(HandlerState state, IPageContent page) throws BirtException {
		if( state.sheetName != null ) {
			log.debug("Attempting to name sheet " + ( state.getWb().getNumberOfSheets() - 1 ) + "\"" + state.sheetName + "\" ");
			boolean alreadyFound = false;
			for( int i = 0; i < state.getWb().getNumberOfSheets() - 1; ++i ) {
				if( state.getWb().getSheetName(i).equals(state.sheetName)) {
					alreadyFound = true;
				}
			}
			if(!alreadyFound) {
				state.getWb().setSheetName(state.getWb().getNumberOfSheets() - 1, state.sheetName);
			}
			state.sheetName = null;
		} 

		Drawing drawing = null;
		if( ! state.images.isEmpty() ) {
			drawing = state.currentSheet.createDrawingPatriarch();
		}
		for( CellImage cellImage : state.images ) {
			processCellImage(state,drawing,cellImage);
		}
		state.images.clear();
		state.rowNum = 0;
		state.colNum = 0;
		
		state.currentSheet = null;
	}

	/**
	 * <p>
	 * Process a CellImage from the images list and place the image on the sheet.
	 * </p><p>
	 * This involves changing the row height as necesssary and determining the column spread of the image.
	 * </p>
	 * @param cellImage
	 * The image to be placed on the sheet.
	 */
	private void processCellImage( HandlerState state, Drawing drawing, CellImage cellImage ) {
		Coordinate location = cellImage.location;
		
		Cell cell = state.currentSheet.getRow( location.getRow() ).getCell( location.getCol() );

		IImageContent image = cellImage.image;		
		
		StyleManagerUtils smu = state.getSmu();
		float ptHeight = cell.getRow().getHeightInPoints();
		if( image.getHeight() != null ) {
			ptHeight = smu.fontSizeInPoints( image.getHeight().toString() );
		}

		// Get image width
		int endCol = cell.getColumnIndex();
        double lastColWidth = ClientAnchorConversions.widthUnits2Millimetres( (short)state.currentSheet.getColumnWidth( endCol ) )
        		+ 2.0;
        int dx = smu.anchorDxFromMM( lastColWidth, lastColWidth );
        double mmWidth = 0.0;
        if( smu.isAbsolute(image.getWidth())) {
            mmWidth = image.getWidth().convertTo(DimensionType.UNITS_MM);
        } else if(smu.isPixels(image.getWidth())) {
            mmWidth = ClientAnchorConversions.pixels2Millimetres( image.getWidth().getMeasure() );
        }
		// Allow image to span multiple columns
		if(cellImage.spanColumns) {
	        log.debug( "Image size: " + image.getWidth() + " translates as mmWidth = " + mmWidth );
	        if( mmWidth > 0) {
	            double mmAccumulatedWidth = 0;
	            for( endCol = cell.getColumnIndex(); mmAccumulatedWidth < mmWidth; ++ endCol ) {
	                lastColWidth = ClientAnchorConversions.widthUnits2Millimetres( (short)state.currentSheet.getColumnWidth( endCol ) )
	                		+ 2.0;
	                mmAccumulatedWidth += lastColWidth;
	                log.debug( "lastColWidth = " + lastColWidth + "; mmAccumulatedWidth = " + mmAccumulatedWidth);
	            }
	            if( mmAccumulatedWidth > mmWidth ) {
	                mmAccumulatedWidth -= lastColWidth;
	                --endCol;
	                double mmShort = mmWidth - mmAccumulatedWidth;
	                dx = smu.anchorDxFromMM( mmShort, lastColWidth );
	            }
	        }
		} else {
			// Adjust the height to fit the aspect ratio caused by the column width
			float widthRatio = (float)(mmWidth / lastColWidth);
			ptHeight = ptHeight / widthRatio;
		}

		if( ptHeight > cell.getRow().getHeightInPoints()) {
			cell.getRow().setHeightInPoints( ptHeight );
		}
		
		// ClientAnchor anchor = wb.getCreationHelper().createClientAnchor();
		ClientAnchor anchor = state.getWb().getCreationHelper().createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setRow1(cell.getRowIndex());
        anchor.setCol2(endCol);
        anchor.setRow2(cell.getRowIndex());
        anchor.setDx2(dx);
        anchor.setDy2( smu.anchorDyFromPoints( ptHeight, cell.getRow().getHeightInPoints() ) );
        anchor.setAnchorType(ClientAnchor.MOVE_DONT_RESIZE);
	    drawing.createPicture(anchor, cellImage.imageIdx);
	}

	@Override
	public void startTable(HandlerState state, ITableContent table) throws BirtException {
		state.setHandler(new TopLevelTableHandler(log,this,table));
		state.getHandler().startTable(state, table);
	}

	@Override
	public void emitText(HandlerState state, ITextContent text) throws BirtException {
		state.setHandler(new TopLevelContentHandler(state.getEmitter(), log, this));
		state.getHandler().emitText(state, text);
	}

	@Override
	public void emitData(HandlerState state, IDataContent data) throws BirtException {
		state.setHandler(new TopLevelContentHandler(state.getEmitter(), log, this));
		state.getHandler().emitData(state, data);
	}

	@Override
	public void emitLabel(HandlerState state, ILabelContent label) throws BirtException {
		state.setHandler(new TopLevelContentHandler(state.getEmitter(), log, this));
		state.getHandler().emitLabel(state, label);
	}

	@Override
	public void emitAutoText(HandlerState state, IAutoTextContent autoText) throws BirtException {
		state.setHandler(new TopLevelContentHandler(state.getEmitter(), log, this));
		state.getHandler().emitAutoText(state, autoText);
	}

	@Override
	public void emitForeign(HandlerState state, IForeignContent foreign) throws BirtException {
		state.setHandler(new TopLevelContentHandler(state.getEmitter(), log, this));
		state.getHandler().emitForeign(state, foreign);
	}

	@Override
	public void emitImage(HandlerState state, IImageContent image) throws BirtException {
		state.setHandler(new TopLevelContentHandler(state.getEmitter(), log, this));
		state.getHandler().emitImage(state, image);
	}
	
	
	
	
}
