package uk.co.spudsoft.birt.emitters.excel.handlers;

import java.io.IOException;
import java.math.BigDecimal;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.content.IContent;
import org.eclipse.birt.report.engine.content.IImageContent;
import org.eclipse.birt.report.engine.css.engine.StyleConstants;
import org.eclipse.birt.report.engine.css.engine.value.css.CSSConstants;
import org.eclipse.birt.report.engine.emitter.IContentEmitter;
import org.eclipse.birt.report.engine.ir.DimensionType;
import org.eclipse.birt.report.engine.layout.emitter.Image;
import org.eclipse.birt.report.engine.presentation.ContentEmitterVisitor;

import uk.co.spudsoft.birt.emitters.excel.BirtStyle;
import uk.co.spudsoft.birt.emitters.excel.CellImage;
import uk.co.spudsoft.birt.emitters.excel.ClientAnchorConversions;
import uk.co.spudsoft.birt.emitters.excel.Coordinate;
import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.RichTextRun;
import uk.co.spudsoft.birt.emitters.excel.StyleManager;
import uk.co.spudsoft.birt.emitters.excel.StyleManagerUtils;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class CellContentHandler extends AbstractHandler {

	/**
	 * Number of milliseconds in a day, to determine whether a given date is date/time/datetime
	 */
	private static final long oneDay = 24 * 60 * 60 * 1000;
	
	/**
	 * The last value added to a cell
	 */
	protected Object lastValue;
	/**
	 * The BIRT element that provided the lastValue
	 */
	protected IContent lastElement;
	/**
	 * List of font changes for a single cell.
	 */
	protected List<RichTextRun> richTextRuns = new ArrayList<RichTextRun>();
	/**
	 * When having to join multiple text blocks together, track whether they are block or inline display
	 */
	protected boolean lastCellContentsWasBlock;
	/**
	 * When having to join multiple text blocks together, track whether they need more of a gap between them (basically used for flattened table cells)
	 */
	protected boolean lastCellContentsRequiresSpace;
	/**
	 * The span of the current cell
	 */
	protected int colSpan;
	/**
	 * Visitor to enable processing of child elements created for foreign (HTML) elements.
	 */
	protected ContentEmitterVisitor contentVisitor;
	/** 
	 * Override the cell alignment to this instead, unless zero
	 */
	protected String preferredAlignment;
	

	public CellContentHandler(IContentEmitter emitter, Logger log, IHandler parent, ICellContent cell) {
		super(log, parent, cell);
		contentVisitor = new ContentEmitterVisitor( emitter );
		colSpan = 1;
	}

	@Override
	public void startCell(HandlerState state, ICellContent cell) throws BirtException {
	}

	/**
	 * Finish processing for the current (real) cell.
	 * @param element
	 * The element that signifies the end of the cell (this may not be an ICellContent object if the 
	 * cell is created for a label or text outside of a table). 
	 */
	protected void endCellContent(HandlerState state, ICellContent birtCell, IContent element, Cell cell ) {
		StyleManager sm = state.getSm();
		StyleManagerUtils smu = state.getSmu();
		
		BirtStyle birtCellStyle = null;
		if( birtCell != null ) {
			birtCellStyle = new BirtStyle( birtCell );
			if( element != null ) {
				birtCellStyle.overlay( element );
			}
		} else if( element != null ) {
			birtCellStyle = new BirtStyle( element );			
		}
		if( preferredAlignment != null ) {
			birtCellStyle.setString( StyleConstants.STYLE_TEXT_ALIGN, preferredAlignment );
		}
		if( CSSConstants.CSS_TRANSPARENT_VALUE.equals(birtCellStyle.getString(StyleConstants.STYLE_BACKGROUND_COLOR))) {
			if( parent != null ) {
				birtCellStyle.setString( StyleConstants.STYLE_BACKGROUND_COLOR, parent.getBackgroundColour() );
			}
		}
			
		if( lastValue != null ) {
			if( lastValue instanceof String ) {
				String lastString = (String)lastValue;

				smu.correctFontColorIfBackground( birtCellStyle );
				for( RichTextRun run  : richTextRuns ) {
					run.font = smu.correctFontColorIfBackground( sm.getFontManager(), state.getWb(), birtCellStyle, run.font ); 
				}
				
				if( ! richTextRuns.isEmpty() ) {
					RichTextString rich = smu.createRichTextString( lastString );
					int runStart = richTextRuns.get(0).startIndex;
					Font lastFont = richTextRuns.get(0).font;
					for( int i = 0; i < richTextRuns.size(); ++i ) {
						RichTextRun run = richTextRuns.get(i);
						log.debug( "Run: " + run.startIndex + " font :" + run.font.toString().replace( "\n", "" ) ); 
						if( ! lastFont.equals( run.font ) ) {
							log.debug("Applying " + runStart + " - " + run.startIndex );
							rich.applyFont(runStart, run.startIndex, lastFont);
							runStart = run.startIndex;
							lastFont = richTextRuns.get(i).font;						
						}
					}
					
					log.debug("Finalising with " + runStart + " - " + lastString.length() );
					rich.applyFont(runStart, lastString.length(), lastFont);

					setCellContents( cell, rich );
				} else {
					setCellContents( cell, lastString );
				}

				if( lastString.contains("\n") ) {
					if( ! CSSConstants.CSS_NOWRAP_VALUE.equals( lastElement.getStyle().getWhiteSpace() ) ) {
						birtCellStyle.setString( StyleConstants.STYLE_WHITE_SPACE, CSSConstants.CSS_PRE_VALUE );
					}
				}					
				if( ! richTextRuns.isEmpty() ) {
					birtCellStyle.setString( StyleConstants.STYLE_VERTICAL_ALIGN, CSSConstants.CSS_TOP_VALUE );
				}
				if( preferredAlignment != null ) {
					log.debug( "preferredAlignment = " + preferredAlignment );
					birtCellStyle.setString( StyleConstants.STYLE_TEXT_ALIGN, preferredAlignment );
				}
				
			} else {
				setCellContents( cell, lastValue );
			}
		}
		
		setCellStyle(sm, cell, birtCellStyle, lastValue);

		if( ( colSpan > 1 ) && ( ( lastValue instanceof String ) || ( lastValue instanceof RichTextString ) ) ) {
			Font defaultFont = state.getWb().getFontAt(cell.getCellStyle().getFontIndex());
			double cellWidth = spanWidthMillimetres( state.currentSheet, cell.getColumnIndex(), cell.getColumnIndex() + colSpan - 1 );
			float cellDesiredHeight = smu.calculateTextHeightPoints( cell.getStringCellValue(), defaultFont, cellWidth, richTextRuns ); 
			if( cellDesiredHeight > state.requiredRowHeightInPoints ) {
				state.requiredRowHeightInPoints = cellDesiredHeight;
			}
		}
			
		lastValue = null;
		lastElement = null;
		richTextRuns.clear();
	}
	
	/**
	 * Calculate the width of a set of columns, in millimetres.
	 * @param startCol
	 * The first column to consider (inclusive).
	 * @param endCol
	 * The last column to consider (inclusive).
	 * @return
	 * The sum of the widths of all columns between startCol and endCol (inclusive) in millimetres.
	 */
	private double spanWidthMillimetres( Sheet sheet, int startCol, int endCol ) {
		short result = 0;
		for ( int columnIndex = startCol; columnIndex <= endCol; ++columnIndex ) {
			result += sheet.getColumnWidth(columnIndex);
		}
		return ClientAnchorConversions.widthUnits2Millimetres( result );
	}

		
	/**
	 * Set the contents of an empty cell.
	 * This should now be the only way in which a cell value is set (cells should not be modified). 
	 * @param value
	 * The value to set.
	 * @param element
	 * The BIRT element supplying the value, used to set the style of the cell.
	 */
	private <T> void setCellContents(Cell cell, Object value) {
		if( value instanceof Double ) {
			cell.setCellType(Cell.CELL_TYPE_NUMERIC);
			cell.setCellValue((Double)value);
			lastValue = value;
		} else if( value instanceof Integer ) {
			cell.setCellType(Cell.CELL_TYPE_NUMERIC);
			cell.setCellValue((Integer)value);				
			lastValue = value;
		} else if( value instanceof Long ) {
			cell.setCellType(Cell.CELL_TYPE_NUMERIC);
			cell.setCellValue((Long)value);				
			lastValue = value;
		} else if( value instanceof Date ) {
			cell.setCellType(Cell.CELL_TYPE_NUMERIC);
			cell.setCellValue((Date)value);
			lastValue = value;
		} else if( value instanceof Boolean ) {
			cell.setCellType(Cell.CELL_TYPE_BOOLEAN);
			cell.setCellValue(((Boolean)value).booleanValue());
			lastValue = value;
		} else if( value instanceof BigDecimal ) {
			cell.setCellType(Cell.CELL_TYPE_NUMERIC);
			cell.setCellValue(((BigDecimal)value).doubleValue());				
			lastValue = value;
		} else if( value instanceof String ) {
			cell.setCellType(Cell.CELL_TYPE_STRING);
			cell.setCellValue((String)value);				
			lastValue = value;
		} else if( value instanceof RichTextString ) {
			cell.setCellType(Cell.CELL_TYPE_STRING);
			cell.setCellValue((RichTextString)value);				
			lastValue = value;
		} else if( value != null ){
			log.debug( "Unhandled data: " + ( value == null ? "<null>" : value.toString() ) );
			cell.setCellType(Cell.CELL_TYPE_STRING);
			cell.setCellValue(value.toString());				
			lastValue = value;
		}
	}

	/**
	 * Set the style of the current cell based on the style of a BIRT element.
	 * @param element
	 * The BIRT element to take the style from.
	 */
	@SuppressWarnings("deprecation")
	private void setCellStyle( StyleManager sm, Cell cell, BirtStyle birtStyle, Object value ) {
		
		if( ( StyleManagerUtils.getNumberFormat(birtStyle) == null )
				&& ( StyleManagerUtils.getDateFormat(birtStyle) == null )
				&& ( StyleManagerUtils.getDateTimeFormat(birtStyle) == null )
				&& ( StyleManagerUtils.getTimeFormat(birtStyle) == null )
				&& ( value != null )
				) {
			if( value instanceof Date ) {
				long time = ((Date)value).getTime();
				time = time - ((Date)value).getTimezoneOffset() * 60000;
				if( time % oneDay == 0 ) {
					StyleManagerUtils.setDateFormat( birtStyle, "Short Date", null);
				} else if( time < oneDay ) {
					StyleManagerUtils.setTimeFormat( birtStyle, "Short Time", null);
				} else {
					StyleManagerUtils.setDateTimeFormat( birtStyle, "General Date", null);
				}
			}
		}
		
		CellStyle cellStyle = sm.getStyle(birtStyle);
		cell.setCellStyle(cellStyle);
	}
	
	private String preferredAlignment( BirtStyle elementStyle ) {
		String newAlign = elementStyle.getString(StyleConstants.STYLE_TEXT_ALIGN);
		if( newAlign == null ) {
			newAlign = CSSConstants.CSS_LEFT_VALUE;
		} 
		if( preferredAlignment == null ) {
			return newAlign;
		}
		if( CSSConstants.CSS_LEFT_VALUE.equals(newAlign) ) {
			return newAlign;
		} else if( CSSConstants.CSS_RIGHT_VALUE.equals(newAlign) ) {
			if( CSSConstants.CSS_CENTER_VALUE.equals(preferredAlignment) ) {
				return newAlign;
			} else {
				return preferredAlignment;
			}
		} else {
			return preferredAlignment;
		}
	}
	
	/**
	 * Set the contents of the current cell.
	 * If the current cell is empty this will format the cell optimally for the new value, if the current cell already has some contents this will simply append the text
	 * value to the current contents.
	 * @param value
	 * The value to put into the current cell.
	 */ 
	protected void emitContent(HandlerState state, IContent element, Object value, boolean asBlock) {
		if( value == null ) {
			return ;
		}
		if( lastValue == null ) {
			lastValue = value;
			lastElement = element;
			lastCellContentsWasBlock = asBlock;
			log.debug( "value == " + value );
			log.debug( "lastCellContentsWasBlock == " + lastCellContentsWasBlock );
			return ;
		} else {
			log.debug( "value == " + value );
			log.debug( "lastCellContentsWasBlock == " + lastCellContentsWasBlock );
		}
		
		StyleManager sm = state.getSm();

		// Both to be improved to include formatting
		String oldValue = lastValue.toString();
		String newComponent = value.toString();
		
		if( lastCellContentsWasBlock 
				&& ! newComponent.startsWith("\n") 
				&& ! oldValue.endsWith("\n") ) {
			oldValue = oldValue + "\n";
			lastCellContentsWasBlock = false;
		}
		if( lastCellContentsRequiresSpace 
				&& ! newComponent.startsWith("\n") 
				&& ! oldValue.endsWith("\n") ) {
			oldValue = oldValue + " ";
			lastCellContentsRequiresSpace = false;
		}

		String newValue = oldValue + newComponent;
		lastValue = newValue;
		
		if( element != null ) {
			BirtStyle elementStyle = new BirtStyle(element);
			Font newFont = sm.getFontManager().getFont( elementStyle );
			richTextRuns.add(new RichTextRun(oldValue.length(), newFont));

			String newAlignment = preferredAlignment(elementStyle);
			log.debug( "preferredAlignment changing from " + preferredAlignment + " to " + newAlignment );
			preferredAlignment = newAlignment;
		}
		
		lastCellContentsWasBlock = asBlock;
	}

	public void recordImage(HandlerState state, Coordinate location, IImageContent image, boolean spanColumns) throws BirtException {
		byte[] data = image.getData();
		log.debug("startImage: "
				+ "[" + image.getMIMEType() +"] "
				+ "{" + image.getWidth() + " x " + image.getHeight() +"} "
				+ ( data == null ? "(no data) " : "(" + data.length + " bytes) ")
				+ image.getURI());
		
		StyleManagerUtils smu = state.getSmu();
		Workbook wb = state.getWb();
		String mimeType = image.getMIMEType();
		if( ( data == null ) && ( image.getURI() != null ) ) {
			try {
				URL imageUrl = new URL( image.getURI() );
				URLConnection conn = imageUrl.openConnection();
				conn.connect();
				mimeType = conn.getContentType();
				int imageType = smu.poiImageTypeFromMimeType( mimeType, null );
				if( imageType == 0 ) {
					log.debug( "Unrecognised/unhandled image MIME type: " + mimeType );
				} else {
					data = smu.downloadImage(conn);
				}
			} catch( MalformedURLException ex ) {
				log.debug( ex.getClass().getName() + ": " + ex.getMessage() );
				ex.printStackTrace();
			} catch( IOException ex ) {
				log.debug( ex.getClass().getName() + ": " + ex.getMessage() );
				ex.printStackTrace();
			}
		}
		if( data != null ) {
			int imageType = smu.poiImageTypeFromMimeType( mimeType, data );
			if( imageType == 0 ) {
				log.debug( "Unrecognised/unhandled image MIME type: " + image.getMIMEType() );
			} else {
				int imageIdx = wb.addPicture( data, imageType );
				
				if( ( image.getHeight() == null ) || ( image.getWidth() == null ) ) {
					Image birtImage = new Image();
					birtImage.setInput( data );
					birtImage.check();
					log.debug( "Calculated image dimensions "
							+ birtImage.getWidth() + " (@" + birtImage.getPhysicalWidthDpi() + "dpi=" + birtImage.getPhysicalWidthInch() + "in) x "
							+ birtImage.getHeight() + " (@" + birtImage.getPhysicalHeightDpi() + "dpi=" + birtImage.getPhysicalHeightInch() + "in)"
							);
					if( image.getWidth() == null ) {
						DimensionType Width = new DimensionType( 
								( birtImage.getPhysicalWidthInch() > 0 ) ? birtImage.getPhysicalWidthInch() : birtImage.getWidth() / 96.0
										, "in" );
						image.setWidth( Width );
					}
					if( image.getHeight() == null ) {
						DimensionType Height = new DimensionType( 
								( birtImage.getPhysicalHeightInch() > 0 ) ? birtImage.getPhysicalHeightInch() : birtImage.getHeight() / 96.0
										, "in" );
						image.setHeight( Height );
					}
				}
				
				state.images.add( new CellImage(location, imageIdx, image, spanColumns) );
				lastElement = image;
			}
		}
	}
	
}
