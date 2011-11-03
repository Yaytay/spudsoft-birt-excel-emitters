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
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.Workbook;
import org.eclipse.birt.report.engine.content.IStyle;
import org.eclipse.birt.report.engine.css.dom.AreaStyle;
import org.eclipse.birt.report.engine.css.engine.CSSEngine;
import org.eclipse.birt.report.engine.css.engine.value.css.CSSConstants;
import org.w3c.dom.css.CSSValue;

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
	private Font defaultFont = null;
	private CSSEngine cssEngine;

	/**
	 * @param workbook
	 * The workbook for which fonts are being tracked.
	 * @param smu
	 * The StyleManagerUtils instance that will be used in the comparison of styles and manipulation of colours.
	 */
	public FontManager(CSSEngine cssEngine, Workbook workbook, StyleManagerUtils smu) {
		this.cssEngine = cssEngine;
		this.workbook = workbook;
		this.smu = smu;
	}

	/**
	 * Obtain the CSS Engine known by this font manager.
	 */
	CSSEngine getCssEngine() {
		return cssEngine;
	}
	
	/**
	 * Remove quotes surrounding a string.
	 * @param family
	 * The string that may be surrounded by double quotes.
	 * @return
	 * family, without any surrounding double quotes.
	 */
	private static String cleanupQuotes( String family ) {
		if( ( family == null ) || family.isEmpty() ) {
			return family;
		}
		if( family.startsWith( "\"" ) && family.endsWith( "\"" ) ) {
			String newFamily = family.substring(1, family.length()-1);
			return newFamily;
		}
		return family;
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
		if(!StyleManagerUtils.objectsEqual(cleanupQuotes(style1.getFontFamily()), cleanupQuotes(style2.getFontFamily()))) {
			return false;
		}
		// Size
		if(!StyleManagerUtils.objectsEqual(cleanupQuotes(style1.getFontSize()), cleanupQuotes(style2.getFontSize()))) {
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
		// Underline
		if(!StyleManagerUtils.objectsEqual(style1.getProperty(IStyle.STYLE_TEXT_UNDERLINE), style2.getProperty(IStyle.STYLE_TEXT_UNDERLINE))) {
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
		String fontName = smu.poiFontNameFromBirt(cleanupQuotes(birtStyle.getFontFamily()));
		if( fontName == null ) {
			fontName = "Calibri";
		}
		font.setFontName(fontName);
		// Size
		short fontSize = smu.fontSizeInPoints(cleanupQuotes(birtStyle.getFontSize()));
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
		// Underline
		if( ( birtStyle.getProperty(IStyle.STYLE_TEXT_UNDERLINE) != null )
			&& CSSConstants.CSS_UNDERLINE_VALUE.equals(birtStyle.getProperty(IStyle.STYLE_TEXT_UNDERLINE).getCssText())) {
			font.setUnderline(FontUnderline.SINGLE.getByteValue());
		}
		// Colour
		smu.addColourToFont(workbook, font, birtStyle.getColor());
				
		fonts.add(new FontPair(birtStyle, font));
		return font;
	}
	
	/**
	 * <p>
	 * Return the default font for the workbook.
	 * </p><p>
	 * At this stage this is hardcoded to return Calibri 11pt, but it could be changed to either pull a value from POI
	 * or to have a parameterised value set from the emitter (via the constructor).
	 * @return
	 * A Font object representing the default to use when no other options are available.
	 */
	private Font getDefaultFont() {
		if( defaultFont == null ) {
			defaultFont = workbook.createFont();
			defaultFont.setFontName("Calibri");
			defaultFont.setFontHeightInPoints((short)11);
		}
		return defaultFont;
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
			return getDefaultFont();
		}
		
		if( ( birtStyle.getFontFamily() == null )
				&& ( birtStyle.getFontSize() == null )
				&& ( birtStyle.getFontWeight() == null )
				&& ( birtStyle.getFontStyle() == null )
				&& ( birtStyle.getColor() == null )
				) {
			return getDefaultFont();
		}
		
		for(FontPair fontPair : fonts) {
			if(fontsEquivalent(birtStyle, fontPair.birtStyle)) {
				return fontPair.poiFont;
			}
		}
		
		return createFont(birtStyle);
	}
	
	private IStyle birtStyleFromFont( Font source ) {
		for(FontPair fontPair : fonts) {
			if( source.equals(fontPair.poiFont) ) {
				return smu.copyBirtStyle( fontPair.birtStyle );
			}
		}
		
		return new AreaStyle( cssEngine );
	}
	
	/**
	 * Return a POI font created by combining a POI font with a BIRT style, where the BIRT style overrides the values in the POI font.
	 * @param source
	 * The POI font that represents the base font.
	 * @param birtExtraStyle
	 * The BIRT style to overlay on top of the POI style.
	 * @return
	 * A POI font representing the combination of source and birtExtraStyle.
	 */
	public Font getFontWithExtraStyle( Font source, IStyle birtExtraStyle ) {

		IStyle birtStyle = birtStyleFromFont( source );
		
		for(int i = 0; i < IStyle.NUMBER_OF_STYLE; ++i ) {
			CSSValue value = birtExtraStyle.getProperty( i );
			if( value != null ) {
				birtStyle.setProperty( i , value );
			}
		}

		Font newFont = getFont(birtStyle);
		return newFont;
	}
	
	
}
