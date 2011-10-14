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

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import org.eclipse.birt.report.engine.content.IStyle;
import org.eclipse.birt.report.engine.content.IStyledElement;
import org.eclipse.birt.report.engine.css.dom.AbstractStyle;
import org.eclipse.birt.report.engine.css.dom.AreaStyle;
import org.eclipse.birt.report.engine.css.engine.CSSEngine;
import org.w3c.dom.css.CSSValue;

import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

/**
 * StyleManager is a cache of POI CellStyles to enable POI CellStyles to be reused based upon their BIRT styles.
 * @author Jim Talbut
 *
 */
public class StyleManager {
	
	/**
	 * StylePair maintains the relationship between a BIRT style and a POI style.
	 * @author Jim Talbut
	 *
	 */
	private class StylePair {
		public IStyle birtStyle;
		public CellStyle poiStyle;
		
		public StylePair(IStyle birtStyle, CellStyle poiStyle) {
			this.birtStyle = birtStyle;
			this.poiStyle = poiStyle;
		}
	}
	
	private Workbook workbook;
	private FontManager fm;
	private List<StylePair> styles = new ArrayList<StylePair>();
	private StyleStack styleStack;
	private StyleManagerUtils smu;
	private CSSEngine cssEngine;
	private Logger log;

	/**
	 * @param workbook
	 * The workbook for which styles are being tracked.
	 * @param styleStack
	 * A style stack, to allow cells to inherit properties from container elements.
	 * @param log
	 * Logger to be used during processing.
	 * @param smu
	 * Set of functions for carrying out conversions between BIRT and POI. 
	 * @param cssEngine
	 * BIRT CSS Engine for creating BIRT styles. 
	 */
	public StyleManager(Workbook workbook, StyleStack styleStack, Logger log, StyleManagerUtils smu, CSSEngine cssEngine) {
		this.workbook = workbook;
		this.fm = new FontManager(workbook, smu);
		this.styleStack = styleStack;
		this.log = log;
		this.smu = smu;
		this.cssEngine = cssEngine;
	}
	
	/**
	 * Merge appropriate styles from the styleStack with the current style.
	 * <BR>
	 * At this time the only property that is merged is background colour. 
	 * @param style
	 * The style to be merged with the styleStack.
	 * @return
	 * The style after merging (this is not a new instance).
	 */
	private IStyle mergeStyles(IStyle style) {
		CSSValue bgColor = style.getProperty(IStyle.STYLE_BACKGROUND_COLOR);
		// System.err.println( "bgColor: " + ( bgColor == null ? "<null>" : bgColor.toString() ) );
		if((bgColor == null) || IStyle.TRANSPARENT_VALUE.equals(bgColor)) {
			bgColor = styleStack.getProperty(IStyle.STYLE_BACKGROUND_COLOR);
			// System.err.println( "stack bgColor: " + ( bgColor == null ? "<null>" : bgColor.toString() ) );
			style.setProperty(IStyle.STYLE_BACKGROUND_COLOR, bgColor);
		}
		return style;
	}
	
	/**
	 * Test whether two BIRT styles are equivalent, as far as the attributes understood by POI are concerned.
	 * <br/>
	 * Every attribute tested in this method must be used in the construction of the CellStyle in createStyle.
	 * @param style1
	 * The first BIRT style to be compared.
	 * @param style2
	 * The second BIRT style to be compared.
	 * @return
	 * true if style1 and style2 would produce identical CellStyles if passed to createStyle.
	 */
	private boolean stylesEquivalent(IStyle style1, IStyle style2) {
		// Alignment
		if(!StyleManagerUtils.objectsEqual(style1.getTextAlign(), style2.getTextAlign())) {
			return false;
		}
		// Font
		if(!FontManager.fontsEquivalent(style1, style2)) {
			return false;
		}
		// Background colour
		if(!StyleManagerUtils.objectsEqual(style1.getBackgroundColor(), style2.getBackgroundColor())) {
			return false;
		}
		// Top border
		if( !StyleManagerUtils.objectsEqual(style1.getBorderTopStyle(), style2.getBorderTopStyle())
			|| !StyleManagerUtils.objectsEqual(style1.getBorderTopWidth(), style2.getBorderTopWidth())
			|| !StyleManagerUtils.objectsEqual(style1.getBorderTopColor(), style2.getBorderTopColor())) {
			return false;
		}
		// Left border
		if( !StyleManagerUtils.objectsEqual(style1.getBorderLeftStyle(), style2.getBorderLeftStyle())
			|| !StyleManagerUtils.objectsEqual(style1.getBorderLeftWidth(), style2.getBorderLeftWidth())
			|| !StyleManagerUtils.objectsEqual(style1.getBorderLeftColor(), style2.getBorderLeftColor())) {
			return false;
		}
		// Right border
		if( !StyleManagerUtils.objectsEqual(style1.getBorderRightStyle(), style2.getBorderRightStyle())
			|| !StyleManagerUtils.objectsEqual(style1.getBorderRightWidth(), style2.getBorderRightWidth())
			|| !StyleManagerUtils.objectsEqual(style1.getBorderRightColor(), style2.getBorderRightColor())) {
			return false;
		}
		// Bottom border
		if( !StyleManagerUtils.objectsEqual(style1.getBorderBottomStyle(), style2.getBorderBottomStyle())
			|| !StyleManagerUtils.objectsEqual(style1.getBorderBottomWidth(), style2.getBorderBottomWidth())
			|| !StyleManagerUtils.objectsEqual(style1.getBorderBottomColor(), style2.getBorderBottomColor())) {
			return false;
		}
		// Number format
		if( !StyleManagerUtils.objectsEqual(style1.getNumberFormat(), style2.getNumberFormat())
			|| !StyleManagerUtils.objectsEqual(style1.getDateFormat(), style2.getDateFormat())
			|| !StyleManagerUtils.objectsEqual(style1.getDateTimeFormat(), style2.getDateTimeFormat())
			|| !StyleManagerUtils.objectsEqual(style1.getTimeFormat(), style2.getTimeFormat()) ){
			return false;
		}
		
		return true;
	}
	
	/**
	 * Create a new POI CellStyle based upon a BIRT style.
	 * @param birtStyle
	 * The BIRT style to base the CellStyle upon.
	 * @return
	 * The CellStyle whose attributes are described by the BIRT style. 
	 */
	private CellStyle createStyle(IStyle birtStyle) {
		log.debug( "Creating style" );
		
		CellStyle poiStyle = workbook.createCellStyle();
		// Alignment
		poiStyle.setAlignment(smu.poiAlignmentFromBirtAlignment(birtStyle.getTextAlign()));
		// Font
		Font font = fm.getFont(birtStyle);
		if( font != null ) {
			poiStyle.setFont(font);
		}
		// Background colour
		smu.addBackgroundColourToStyle(workbook, poiStyle, birtStyle.getBackgroundColor());
		// Top border 
		smu.applyBorderStyle(workbook, poiStyle, BorderSide.TOP, birtStyle.getBorderTopColor(), birtStyle.getBorderTopStyle(), birtStyle.getBorderTopWidth());
		// Left border 
		smu.applyBorderStyle(workbook, poiStyle, BorderSide.LEFT, birtStyle.getBorderLeftColor(), birtStyle.getBorderLeftStyle(), birtStyle.getBorderLeftWidth());
		// Right border 
		smu.applyBorderStyle(workbook, poiStyle, BorderSide.RIGHT, birtStyle.getBorderRightColor(), birtStyle.getBorderRightStyle(), birtStyle.getBorderRightWidth());
		// Bottom border 
		smu.applyBorderStyle(workbook, poiStyle, BorderSide.BOTTOM, birtStyle.getBorderBottomColor(), birtStyle.getBorderBottomStyle(), birtStyle.getBorderBottomWidth());
		// Number format
		smu.applyNumberFormat(workbook, birtStyle, poiStyle);

		styles.add(new StylePair(birtStyle, poiStyle));
		return poiStyle;
	}

	/**
	 * Get a CellStyle matching the BIRT style, either from the cache or creating a new one.
	 * @param element
	 * The BIRT element that has a style to be copied.
	 * @return
	 * A POI CellStyle containing attributes defined by the BIRT element.
	 */
	public CellStyle getStyle( IStyledElement element ) {
		IStyle birtStyle = element.getComputedStyle();
		return getStyle( birtStyle );
	}
	
	private CellStyle getStyle( IStyle birtStyle ) {
		if( birtStyle == null ) {
			return null;
		}
		
		birtStyle = mergeStyles(birtStyle);
		for(StylePair stylePair : styles) {
			if(stylesEquivalent(birtStyle, stylePair.birtStyle)) {
				return stylePair.poiStyle;
			}
		}
		
		return createStyle(birtStyle);		
	}
	
	private IStyle birtStyleFromCellStyle( CellStyle source ) {
		for(StylePair stylePair : styles) {
			if( source.equals(stylePair.poiStyle) ) {
				AreaStyle styleCopy = new AreaStyle( (AbstractStyle)stylePair.birtStyle );
				return styleCopy;
			}
		}
		
		return new AreaStyle( cssEngine );
	}

	/**
	 * Given a POI CellStyle, add border definitions to it and obtain a CellStyle (from the cache or newly created) based upon that.
	 * @param source
	 * The POI CellStyle to form the base style.
	 * @param borderStyleBottom
	 * The BIRT style of the bottom border.
	 * @param borderWidthBottom
	 * The BIRT with of the bottom border.
	 * @param borderColourBottom
	 * The BIRT colour of the bottom border.
	 * @param borderStyleLeft
	 * The BIRT style of the left border.
	 * @param borderWidthLeft
	 * The BIRT width of the left border.
	 * @param borderColourLeft
	 * The BIRT colour of the left border.
	 * @param borderStyleRight
	 * The BIRT width of the right border.
	 * @param borderWidthRight
	 * The BIRT colour of the right border.
	 * @param borderColourRight
	 * The BIRT style of the right border.
	 * @param borderStyleTop
	 * The BIRT style of the top border.
	 * @param borderWidthTop
	 * The BIRT width of the top border.
	 * @param borderColourTop
	 * The BIRT colour of the top border.
	 * @return
	 * A POI CellStyle equivalent to the source CellStyle with all the defined borders added to it.
	 */
	public CellStyle getStyleWithBorders( CellStyle source
			, String borderStyleBottom, String borderWidthBottom, String borderColourBottom 
			, String borderStyleLeft, String borderWidthLeft, String borderColourLeft 
			, String borderStyleRight, String borderWidthRight, String borderColourRight 
			, String borderStyleTop, String borderWidthTop, String borderColourTop 
			) {

		IStyle birtStyle = birtStyleFromCellStyle( source );
		if( borderStyleBottom != null ) {
			birtStyle.setBorderBottomStyle( borderStyleBottom );
		}
		if( borderWidthBottom != null ) {
			birtStyle.setBorderBottomWidth( borderWidthBottom );
		}
		if( borderColourBottom != null ) {
			birtStyle.setBorderBottomColor( borderColourBottom );			
		}
		if( borderStyleLeft != null ) {
			birtStyle.setBorderLeftStyle( borderStyleLeft );
		}
		if( borderWidthLeft != null ) {
			birtStyle.setBorderLeftWidth( borderWidthLeft );
		}
		if( borderColourLeft != null ) {
			birtStyle.setBorderLeftColor( borderColourLeft );			
		}
		if( borderStyleRight != null ) {
			birtStyle.setBorderRightStyle( borderStyleRight );
		}
		if( borderWidthRight != null ) {
			birtStyle.setBorderRightWidth( borderWidthRight );
		}
		if( borderColourRight != null ) {
			birtStyle.setBorderRightColor( borderColourRight );			
		}
		if( borderStyleTop != null ) {
			birtStyle.setBorderTopStyle( borderStyleTop );
		}
		if( borderWidthTop != null ) {
			birtStyle.setBorderTopWidth( borderWidthTop );
		}
		if( borderColourTop != null ) {
			birtStyle.setBorderTopColor( borderColourTop );			
		}
		CellStyle newStyle = getStyle( birtStyle );
		return newStyle;
	}
}
