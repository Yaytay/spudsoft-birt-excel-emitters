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

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import org.eclipse.birt.report.engine.content.IPageContent;
import org.eclipse.birt.report.engine.content.IStyle;
import org.eclipse.birt.report.engine.css.dom.AreaStyle;
import org.eclipse.birt.report.engine.css.engine.StyleConstants;
import org.eclipse.birt.report.engine.ir.DimensionType;
import org.eclipse.birt.report.model.api.util.ColorUtil;
import org.w3c.dom.css.CSSValue;

import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

/**
 * StyleManagerHUtils is an extension of the StyleManagerUtils to provide HSSFWorkbook specific functionality.
 * @author Jim Talbut
 *
 */
public class StyleManagerHUtils extends StyleManagerUtils {

	private static Factory factory = new StyleManagerUtils.Factory() {
		@Override
		public StyleManagerUtils create(Logger log) {
			return new StyleManagerHUtils(log);
		}
	};
	
	public static Factory getFactory() {
		return factory;
	}

	
	/**
	 * @param log
	 * Logger used by StyleManagerHUtils to record anything of interest.
	 */
	public StyleManagerHUtils(Logger log) {
		super(log);
	}
	
	@Override
	public RichTextString createRichTextString(String value) {
		return new HSSFRichTextString(value);
	}

	/**
	 * Converts a BIRT border style into a POI border style (short constant defined in CellStyle).
	 * @param birtBorder
	 * The BIRT border style.
	 * @param width
	 * The width of the border as understood by BIRT.
	 * @return
	 * One of the CellStyle BORDER constants.
	 */
	private short poiBorderStyleFromBirt( String birtBorder, String width ) {
		if( "none".equals(birtBorder) ) {
			return CellStyle.BORDER_NONE;
		}
		DimensionType dim = DimensionType.parserUnit( width );
		double pxWidth = 3.0;
		if( ( dim != null ) && ( "px".equals(dim.getUnits()) ) ){
			pxWidth = dim.getMeasure();
		}
		if( "solid".equals(birtBorder) ) {
			if( pxWidth < 2.9 ) {
				return CellStyle.BORDER_THIN;
			} else if( pxWidth < 3.1 ) {
				return CellStyle.BORDER_MEDIUM;
			} else {
				return CellStyle.BORDER_THICK;
			}
		} else if( "dashed".equals(birtBorder) ) {
			if( pxWidth < 2.9 ) {
				return CellStyle.BORDER_DASHED;
			} else {
				return CellStyle.BORDER_MEDIUM_DASHED;
			}
		} else if( "dotted".equals(birtBorder) ) {
			return CellStyle.BORDER_DOTTED;
		} else if( "double".equals(birtBorder) ) {
			return CellStyle.BORDER_DOUBLE;
		} else if( "none".equals(birtBorder) ) {
			return CellStyle.BORDER_NONE;
		}

		log.debug( "Border style \"", birtBorder, "\" is not recognised" );
		return CellStyle.BORDER_NONE;
	}
	
	/**
	 * Get an HSSFPalette index for a workbook that closely approximates the passed in colour.
	 * @param workbook
	 * The workbook for which the colour is being sought.
	 * @param colour
	 * The colour, in the form "rgb(<i>r</i>, <i>g</i>, <i>b</i>)".
	 * @return
	 * The index into the HSSFPallete for the workbook for a colour that approximates the passed in colour.
	 */
	private short getHColour( HSSFWorkbook workbook, String colour ) {
		int[] rgbInt = ColorUtil.getRGBs(colour);
		if( rgbInt == null ) {
			return 0;
		}
		
		byte[] rgbByte = new byte[] { (byte)rgbInt[0], (byte)rgbInt[1], (byte)rgbInt[2] };
		HSSFPalette palette = workbook.getCustomPalette();
		
		HSSFColor result = palette.findColor(rgbByte[0], rgbByte[1], rgbByte[2]);
		if( result ==  null) {
			result = palette.findSimilarColor(rgbByte[0], rgbByte[1], rgbByte[2]);
		}
		return result.getIndex();
	}

	@Override
	public void applyBorderStyle(Workbook workbook, CellStyle style, BorderSide side, CSSValue colour, CSSValue borderStyle, CSSValue width) {
		if( ( colour == null ) || ( borderStyle != null ) || ( width != null ) ) {
			String colourString = colour == null ? "rgb(0,0,0)" : colour.getCssText();
			String borderStyleString = borderStyle == null ? "solid" : borderStyle.getCssText();
			String widthString = width == null ? "medium" : width.getCssText();

			if( style instanceof HSSFCellStyle ) {
				HSSFCellStyle hStyle = (HSSFCellStyle)style;
				
				short hBorderStyle = poiBorderStyleFromBirt(borderStyleString, widthString);
				short colourIndex = getHColour((HSSFWorkbook)workbook, colourString);
				if( colourIndex > 0 ) {
					if(hBorderStyle != CellStyle.BORDER_NONE) {
						switch( side ) {
						case TOP:
							hStyle.setBorderTop(hBorderStyle);
							hStyle.setTopBorderColor(colourIndex);
							// log.debug( "Top border: " + xStyle.getBorderTop() + " / " + xStyle.getTopBorderXSSFColor().getARGBHex() );
							break;
						case LEFT:
							hStyle.setBorderLeft(hBorderStyle);
							hStyle.setLeftBorderColor(colourIndex);
							// log.debug( "Left border: " + xStyle.getBorderLeft() + " / " + xStyle.getLeftBorderXSSFColor().getARGBHex() );
							break;
						case RIGHT:
							hStyle.setBorderRight(hBorderStyle);
							hStyle.setRightBorderColor(colourIndex);
							// log.debug( "Right border: " + xStyle.getBorderRight() + " / " + xStyle.getRightBorderXSSFColor().getARGBHex() );
							break;
						case BOTTOM:
							hStyle.setBorderBottom(hBorderStyle);
							hStyle.setBottomBorderColor(colourIndex);
							// log.debug( "Bottom border: " + xStyle.getBorderBottom() + " / " + xStyle.getBottomBorderXSSFColor().getARGBHex() );
							break;
						}
					}
				}
			}
		}
	}	

	@Override
	public void addColourToFont(Workbook workbook, Font font, String colour) {
		if(colour == null) {
			return ;
		}
		if(IStyle.TRANSPARENT_VALUE.equals(colour)) {
			return ;
		}
		if(font instanceof HSSFFont) {
			HSSFFont hFont = (HSSFFont)font;
			short colourIndex = getHColour((HSSFWorkbook)workbook, colour);
			if( colourIndex > 0 ) {
				hFont.setColor(colourIndex);
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
		if(style instanceof HSSFCellStyle) {
			HSSFCellStyle cellStyle = (HSSFCellStyle)style;
			short colourIndex = getHColour((HSSFWorkbook)workbook, colour);
			if( colourIndex > 0 ) {
				cellStyle.setFillForegroundColor(colourIndex);
				cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			}
		}
	}

	@Override
	public Font correctFontColorIfBackground(FontManager fm, Workbook wb, BirtStyle birtStyle, Font font) {
		HSSFPalette palette = ((HSSFWorkbook)wb).getCustomPalette();		
		
		CSSValue bgColour = birtStyle.getProperty( StyleConstants.STYLE_BACKGROUND_COLOR );
		int bgRgb[] = parseColour( bgColour == null ? null : bgColour.getCssText(), "white" );

		short fgRgb[] = HSSFColor.BLACK.triplet;
		if( ( font != null ) && ( font.getColor() != Short.MAX_VALUE ) ) {
			fgRgb = palette.getColor(font.getColor()).getTriplet();
		}
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
        return (int)( 1023.0 * widthMM / colWidthMM );
	}
	
	@Override
	public int anchorDyFromPoints( float height, float rowHeight ) {
        return (int)( 255.0 * height / rowHeight );
	}

	@Override
	public void prepareMarginDimensions(Sheet sheet, IPageContent page) {
		double headerHeight = 0.0;
		double footerHeight = 0.0;
		if( page.getHeaderHeight() != null ) {
			headerHeight = page.getHeaderHeight().convertTo(DimensionType.UNITS_IN);
			sheet.getPrintSetup().setHeaderMargin(headerHeight);
		}
		if( page.getFooterHeight() != null ) {
			footerHeight = page.getFooterHeight().convertTo(DimensionType.UNITS_IN);
			sheet.getPrintSetup().setFooterMargin(footerHeight);
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
