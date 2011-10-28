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

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.eclipse.birt.report.engine.content.IStyle;

/**
 * FontManager is a cache of fonts to enable POI Fonts to be reused based upon their BIRT styles.
 * @author Jim Talbut
 *
 */
public class FontManager {
	
	/**
	 * FontPair maintains the relationship between a BIRT style and a POI font.
	 * @author Jim Talbut
	 *
	 */
	private class FontPair {
		public IStyle birtStyle;
		public Font poiFont;
		
		public FontPair(IStyle birtStyle, Font poiFont) {
			this.birtStyle = birtStyle;
			this.poiFont = poiFont;
		}
	}
	
	private Workbook workbook;
	private StyleManagerUtils smu;
	private List<FontPair> fonts = new ArrayList<FontPair>();

	/**
	 * @param workbook
	 * The workbook for which fonts are being tracked.
	 * @param smu
	 * The StyleManagerUtils instance that will be used in the comparison of styles and manipulation of colours.
	 */
	public FontManager(Workbook workbook, StyleManagerUtils smu) {
		this.workbook = workbook;
		this.smu = smu;
	}
	
	/**
	 * Test whether two BIRT styles are equivalent, as far as their font definitions are concerned.
	 * <br/>
	 * Every attribute tested in this method must be used in the construction of the font in createFont.
	 * @param style1
	 * The first BIRT style to be compared.
	 * @param style2
	 * The second BIRT style to be compared.
	 * @return
	 * true if style1 and style2 would produce identical Fonts if passed to createFont.
	 */
	public static boolean fontsEquivalent(IStyle style1, IStyle style2) {
		// Family
		if(!StyleManagerUtils.objectsEqual(style1.getFontFamily(), style2.getFontFamily())) {
			return false;
		}
		if( style1.getFontFamily() == null ) {
			return true;
		}
		// Size
		if(!StyleManagerUtils.objectsEqual(style1.getFontSize(), style2.getFontSize())) {
			return false;
		}
		// Weight
		if(!StyleManagerUtils.objectsEqual(style1.getFontWeight(), style2.getFontWeight())) {
			return false;
		}
		// Italic
		if(!StyleManagerUtils.objectsEqual(style1.getFontStyle(), style2.getFontStyle())) {
			return false;
		}
		// Colour
		if(!StyleManagerUtils.objectsEqual(style1.getColor(), style2.getColor())) {
			return false;
		}
		
		return true;
	}
	
	/**
	 * Create a new POI Font based upon a BIRT style.
	 * @param birtStyle
	 * The BIRT style to base the Font upon.
	 * @return
	 * The Font whose attributes are described by the BIRT style. 
	 */
	private Font createFont(IStyle birtStyle) {
		Font font = workbook.createFont();
		
		// Family
		String fontName = smu.poiFontNameFromBirt(birtStyle.getFontFamily());
		if( fontName == null ) {
			fontName = "Calibri";
		}
		font.setFontName(fontName);
		// Size
		short fontSize = smu.fontSizeInPoints(birtStyle.getFontSize());
		if(fontSize > 0) {
			font.setFontHeightInPoints(fontSize);
		}
		// Weight
		short fontWeight = smu.poiFontWeightFromBirt(birtStyle.getFontWeight());
		if(fontWeight > 0) {
			font.setBoldweight(fontWeight);
		}
		// Style
		if("italic".equals(birtStyle.getFontStyle()) || "oblique".equals(birtStyle.getFontStyle())) {
			font.setItalic(true);
		}
		// Colour
		smu.addColourToFont(workbook, font, birtStyle.getColor());
				
		fonts.add(new FontPair(birtStyle, font));
		return font;
	}
	
	/**
	 * Get a Font matching the BIRT style, either from the cache or by creating a new one.
	 * @param birtStyle
	 * The BIRT style to base the Front upon.
	 * @return
	 * A Font whose attributes are described by the BIRT style. 
	 */
	Font getFont( IStyle birtStyle ) {
		if( birtStyle == null ) {
			return null;
		}
		
		if( birtStyle.getFontFamily() == null ) {
			return null;
		}
		
		for(FontPair fontPair : fonts) {
			if(fontsEquivalent(birtStyle, fontPair.birtStyle)) {
				return fontPair.poiFont;
			}
		}
		
		return createFont(birtStyle);
	}
}
