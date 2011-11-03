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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import org.eclipse.birt.report.engine.content.IStyle;
import org.eclipse.birt.report.engine.css.dom.AreaStyle;
import org.eclipse.birt.report.engine.css.engine.CSSEngine;
import org.eclipse.birt.report.engine.ir.DimensionType;
import org.eclipse.birt.report.model.api.util.ColorUtil;

import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

/**
 * StyleManagerXUtils is an extension of the StyleManagerUtils to provide XSSFWorkbook specific functionality.
 * @author Jim Talbut
 *
 */
public class StyleManagerXUtils extends StyleManagerUtils {

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
		DimensionType dim = DimensionType.parserUnit( width );
		double pxWidth = 3.0;
		if( dim != null ) {
			if( "px".equals(dim.getUnits()) ) {
				pxWidth = dim.getMeasure();
			}
		} 
		log.debug( "Border width (" + birtBorder + "/" + width + "): " + dim + " == " + pxWidth + "px." );
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
	public void applyBorderStyle(Workbook workbook, CellStyle style, BorderSide side, String colour, String borderStyle, String width) {
		if( ( colour != null ) && ( borderStyle != null ) && ( width != null ) ) {
			if( style instanceof XSSFCellStyle ) {
				XSSFCellStyle xStyle = (XSSFCellStyle)style;
				
				BorderStyle xBorderStyle = poiBorderStyleFromBirt(borderStyle, width);
				XSSFColor xBorderColour = getXColour(colour);
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
			
	private boolean fontColorOk( CellStyle cellStyle, XSSFColor fontColour ) {
		if( cellStyle.getFillForegroundColorColor() == null ) {
			if( fontColour == null ) {
				return true;
			}
			byte[] argb = fontColour.getARgb();
			// Note that this handling is explicitly wrong, because POI is trying to compensate for Excel handling colours back to front
			if( ( argb[1] != 0 ) && ( argb[2] != 0 ) && ( argb[3] != 0 ) ) {
				return true; 
			}
		} else if( ! cellStyle.getFillForegroundColorColor().equals(fontColour) ) {
			return true; 
		}
		
		return false;
	}
	
	private IStyle prepareBirtStyleWithSafeColour(CSSEngine cssEngine, XSSFColor colour) {
		byte[] argb = colour.getARgb();
		IStyle addedStyle = new AreaStyle( cssEngine );
		// Note that this handling is explicitly wrong, because POI is trying to compensate for Excel handling colours back to front
		if( ( argb[0] == 0 ) && ( argb[1] == 0 ) && ( argb[2] == 0 ) && ( argb[3] == 0 ) ) {
			addedStyle.setColor("rgb(255, 255, 255)");
		} else {
			addedStyle.setColor("rgb(0, 0, 0)");
		}
		return addedStyle;
	}
	
	public Font correctFontColorIfBackground( FontManager fm, CellStyle cellStyle, Font font ) {
		XSSFColor colour = ((XSSFFont)font).getXSSFColor();

		if( fontColorOk(cellStyle, colour)) {
			return font;
		}
		
		IStyle addedStyle = prepareBirtStyleWithSafeColour(fm.getCssEngine(), colour);
		
		return fm.getFontWithExtraStyle( font, addedStyle );
	}

	@Override
	public void correctFontColorIfBackground(StyleManager sm, Cell cell) {
		XSSFColor colour = ((XSSFCell)cell).getCellStyle().getFont().getXSSFColor();

		if( fontColorOk(cell.getCellStyle(), colour)) {
			return ;
		}
		
		IStyle addedStyle = prepareBirtStyleWithSafeColour(sm.getCssEngine(), colour);
		
		cell.setCellStyle( sm.getStyleWithExtraStyle( cell.getCellStyle(), addedStyle ) );
	}
	
	
}
