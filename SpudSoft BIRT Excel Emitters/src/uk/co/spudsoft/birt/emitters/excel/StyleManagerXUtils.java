/********************************************************************************
* (C) Copyright 2011, by James Talbut.
*
*   This program is free software: you can redistribute it and/or modify
*   it under the terms of the GNU General Public License as published by
*   the Free Software Foundation, either version 3 of the License, or
*   (at your option) any later version.
*
*   This program is distributed in the hope that it will be useful,
*   but WITHOUT ANY WARRANTY; without even the implied warranty of
*   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
*   GNU General Public License for more details.
*
*   You should have received a copy of the GNU General Public License
*   along with this program.  If not, see <http://www.gnu.org/licenses/>.
*
*   [Java is a trademark or registered trademark of Sun Microsystems, Inc.
*   in the United States and other countries.]
********************************************************************************/

package uk.co.spudsoft.birt.emitters.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import org.eclipse.birt.report.engine.content.IPageContent;
import org.eclipse.birt.report.engine.content.IStyle;
import org.eclipse.birt.report.engine.css.dom.AreaStyle;
import org.eclipse.birt.report.engine.css.engine.StyleConstants;
import org.eclipse.birt.report.engine.css.engine.value.css.CSSConstants;
import org.eclipse.birt.report.engine.ir.DimensionType;
import org.eclipse.birt.report.model.api.util.ColorUtil;
import org.w3c.dom.css.CSSValue;

import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

/**
 * StyleManagerXUtils is an extension of the StyleManagerUtils to provide XSSFWorkbook specific functionality.
 * @author Jim Talbut
 *
 */
public class StyleManagerXUtils extends StyleManagerUtils {

	private static Factory factory = new StyleManagerUtils.Factory() {
		@Override
		public StyleManagerUtils create(Logger log) {
			return new StyleManagerXUtils(log);
		}
	};
	
	public static Factory getFactory() {
		return factory;
	}

	/**
	 * @param log
	 * Logger used by StyleManagerXUtils to record anything of interest.
	 */
	public StyleManagerXUtils(Logger log) {
		super(log);
	}

	@Override
	public RichTextString createRichTextString(String value) {
		XSSFRichTextString result = new XSSFRichTextString(value);
		return result;
	}
	
	/**
	 * Converts a BIRT border style into a POI BorderStyle.
	 * @param birtBorder
	 * The BIRT border style.
	 * @param width
	 * The width of the border as understood by BIRT.
	 * @return
	 * A POI BorderStyle object.
	 */
	private BorderStyle poiBorderStyleFromBirt( String birtBorder, String width ) {
		if( "none".equals(birtBorder) ) {
			return BorderStyle.NONE;
		}
		double pxWidth = 3.0;
		if( CSSConstants.CSS_THIN_VALUE.equals( width ) ) {
			pxWidth = 1.0;
		} else if( CSSConstants.CSS_MEDIUM_VALUE.equals( width ) ) {
			pxWidth = 3.0;
		} else if( CSSConstants.CSS_THICK_VALUE.equals( width ) ) {
			pxWidth = 4.0;
		} else {
			DimensionType dim = DimensionType.parserUnit( width );
			if( dim != null ) {
				if( "px".equals(dim.getUnits()) ) {
					pxWidth = dim.getMeasure();
				}
			} 
		}
		if( "solid".equals(birtBorder) ) {
			if( pxWidth < 2.9 ) {
				return BorderStyle.THIN;
			} else if( pxWidth < 3.1 ) {
				return BorderStyle.MEDIUM;
			} else {
				return BorderStyle.THICK;
			}
		} else if( "dashed".equals(birtBorder) ) {
			if( pxWidth < 2.9 ) {
				return BorderStyle.DASHED;
			} else {
				return BorderStyle.MEDIUM_DASHED;
			}
		} else if( "dotted".equals(birtBorder) ) {
			return BorderStyle.DOTTED;
		} else if( "double".equals(birtBorder) ) {
			return BorderStyle.DOUBLE;
		}

		log.debug( "Border style \"" + birtBorder + "\" is not recognised." );
		return BorderStyle.NONE;
	}

	@Override
	public void applyBorderStyle(Workbook workbook, CellStyle style, BorderSide side, CSSValue colour, CSSValue borderStyle, CSSValue width) {
		if( ( colour != null ) || ( borderStyle != null ) || ( width != null ) ) {
			String colourString = colour == null ? "rgb(0,0,0)" : colour.getCssText();
			String borderStyleString = borderStyle == null ? "solid" : borderStyle.getCssText();
			String widthString = width == null ? "medium" : width.getCssText();
			
			if( style instanceof XSSFCellStyle ) {
				XSSFCellStyle xStyle = (XSSFCellStyle)style;
				
				BorderStyle xBorderStyle = poiBorderStyleFromBirt(borderStyleString, widthString);
				XSSFColor xBorderColour = getXColour(colourString);
				if(xBorderStyle != BorderStyle.NONE) {
					switch( side ) {
					case TOP:
						xStyle.setBorderTop(xBorderStyle);
						xStyle.setTopBorderColor(xBorderColour);
						// log.debug( "Top border: " + xStyle.getBorderTop() + " / " + xStyle.getTopBorderXSSFColor().getARGBHex() );
						break;
					case LEFT:
						xStyle.setBorderLeft(xBorderStyle);
						xStyle.setLeftBorderColor(xBorderColour);
						// log.debug( "Left border: " + xStyle.getBorderLeft() + " / " + xStyle.getLeftBorderXSSFColor().getARGBHex() );
						break;
					case RIGHT:
						xStyle.setBorderRight(xBorderStyle);
						xStyle.setRightBorderColor(xBorderColour);
						// log.debug( "Right border: " + xStyle.getBorderRight() + " / " + xStyle.getRightBorderXSSFColor().getARGBHex() );
						break;
					case BOTTOM:
						xStyle.setBorderBottom(xBorderStyle);
						xStyle.setBottomBorderColor(xBorderColour);
						// log.debug( "Bottom border: " + xStyle.getBorderBottom() + " / " + xStyle.getBottomBorderXSSFColor().getARGBHex() );
						break;
					}
				}
			}
		}
	}
	
	private XSSFColor getXColour(String colour) {
		int[] rgbInt = ColorUtil.getRGBs(colour);
		if( rgbInt == null ) {
			return null;
		}
		byte[] rgbByte = { (byte)-1, (byte)rgbInt[0], (byte)rgbInt[1], (byte)rgbInt[2] };
		// System.out.println( "The X colour for " + colour + " is [ " + rgbByte[0] + "," + rgbByte[1] + "," + rgbByte[2] + "," + rgbByte[3] + "]" );
		XSSFColor result = new XSSFColor( rgbByte );
		return result;		
	}

	@Override
	public void addColourToFont(Workbook workbook, Font font, String colour) {
		if(colour == null) {
			return ;
		}
		if(IStyle.TRANSPARENT_VALUE.equals(colour)) {
			return ;
		}
		if(font instanceof XSSFFont) {
			log.debug("Colour " + colour);
			XSSFFont xFont = (XSSFFont)font;
			XSSFColor xColour = getXColour(colour);
			
			log.debug("XColour " + xColour.getARGBHex());
			if(xColour != null) {
				xFont.setColor(xColour);
			}
		}
	}
	
	@Override
	public void addBackgroundColourToStyle(Workbook workbook, CellStyle style, String colour) {
		if(colour == null) {
			return ;
		}
		if(IStyle.TRANSPARENT_VALUE.equals(colour)) {
			return ;
		}
		if(style instanceof XSSFCellStyle) {
			XSSFCellStyle cellStyle = (XSSFCellStyle)style;
			XSSFColor xColour = getXColour(colour);
			if(xColour != null) {
				cellStyle.setFillForegroundColor(xColour);
				cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			}
		}
	}
	
	@Override
	public Font correctFontColorIfBackground( FontManager fm, Workbook wb, BirtStyle birtStyle, Font font ) {
		CSSValue bgColour = birtStyle.getProperty( StyleConstants.STYLE_BACKGROUND_COLOR );
		int bgRgb[] = parseColour( bgColour == null ? null : bgColour.getCssText(), "white" );

		XSSFColor colour = ((XSSFFont)font).getXSSFColor();
		int fgRgb[] = rgbOnly( colour.getARgb() );
		if( ( fgRgb[0] == 255 ) && ( fgRgb[1] == 255 ) && ( fgRgb[2] == 255 ) ) {
			fgRgb[0]=fgRgb[1]=fgRgb[2]=0;
		} else if( ( fgRgb[0] == 0 ) && ( fgRgb[1] == 0 ) && ( fgRgb[2] == 0 ) ) {
			fgRgb[0]=fgRgb[1]=fgRgb[2]=255;
		}

		if( ( bgRgb[ 0 ] == fgRgb[ 0 ] ) && ( bgRgb[ 1 ] == fgRgb[ 1 ] ) && ( bgRgb[ 2 ] == fgRgb[ 2 ] ) ) {
			
			IStyle addedStyle = new AreaStyle( fm.getCssEngine() );
			addedStyle.setColor( contrastColour( bgRgb ) );
			
			return fm.getFontWithExtraStyle( font, addedStyle );
		} else {
			return font;
		}
	}

	@Override
	public int anchorDxFromMM( double widthMM, double colWidthMM ) {
        return (int)(widthMM * 36000); 
	}
	
	@Override
	public int anchorDyFromPoints( float height, float rowHeight ) {
		return (int)( height * XSSFShape.EMU_PER_POINT );
	}

	@Override
	public void prepareMarginDimensions(Sheet sheet, IPageContent page) {
		double headerHeight = 0.0;
		double footerHeight = 0.0;
		if( page.getHeaderHeight() != null ) {
			headerHeight = page.getHeaderHeight().convertTo(DimensionType.UNITS_IN);
			sheet.setMargin(Sheet.HeaderMargin, headerHeight);
		}
		if( page.getFooterHeight() != null ) {
			footerHeight = page.getFooterHeight().convertTo(DimensionType.UNITS_IN);
			sheet.setMargin(Sheet.FooterMargin, footerHeight);
		}
		if( page.getMarginBottom() != null ) {
			sheet.setMargin(Sheet.BottomMargin, footerHeight + page.getMarginBottom().convertTo(DimensionType.UNITS_IN));
		}
		if( page.getMarginLeft() != null ) {
			sheet.setMargin(Sheet.LeftMargin, page.getMarginLeft().convertTo(DimensionType.UNITS_IN));
		}
		if( page.getMarginRight() != null ) {
			sheet.setMargin(Sheet.RightMargin, page.getMarginRight().convertTo(DimensionType.UNITS_IN));
		}
		if( page.getMarginTop() != null ) {
			sheet.setMargin(Sheet.TopMargin, headerHeight + page.getMarginTop().convertTo(DimensionType.UNITS_IN));
		}
	}
	
}
