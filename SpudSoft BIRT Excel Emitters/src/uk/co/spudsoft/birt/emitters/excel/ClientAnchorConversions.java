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

/**
 * <p>
 * ClientAnchorConversions provides a small set of functions for converting the values used with ClientAnchors.
 * </p><p>
 * This class is very heavily based on the ConvertImageUnits class from the POI examples.
 * The differences between that class and this are:
 * <ol>
 * <li>This class contains only the functionality that I need.</li>
 * <li>This class contains no public static values, only methods.</li>
 * </ol>
 * <p> 
 * @author Jim Talbut
 *
 */
public class ClientAnchorConversions {
	
    // Constants that defines how many pixels and points there are in a
    // millimetre. These values are required for the conversion algorithm.
    private static final double PIXELS_PER_MILLIMETRES = 3.78;         // MB
    private static final short EXCEL_COLUMN_WIDTH_FACTOR = 256;
    private static final int UNIT_OFFSET_LENGTH = 7;
    private static final int[] UNIT_OFFSET_MAP = new int[] { 0, 36, 73, 109, 146, 182, 219 };

    /**
     * Convert a measure in column width units (1/256th of a character) to a measure in millimetres.
     * <BR>
     * Makes assumptions about font size and relevant DPI.
     * @param widthUnits
     * The size in width units.
     * @return
     * The size in millimetres.
     */
    public static double widthUnits2Millimetres( int widthUnits ) {
        int pixels = (widthUnits / EXCEL_COLUMN_WIDTH_FACTOR) * UNIT_OFFSET_LENGTH;
        int offsetWidthUnits = widthUnits % EXCEL_COLUMN_WIDTH_FACTOR;
        pixels += Math.round(offsetWidthUnits / ((float) EXCEL_COLUMN_WIDTH_FACTOR / UNIT_OFFSET_LENGTH));
        return pixels / PIXELS_PER_MILLIMETRES;		
	}
	
	/** 
	 * Convert a measure of millimetres to width units.
	 * @param millimetres
	 * The size in millimetres.
	 * @return
	 * The size in width units.
	 */
	public static int millimetres2WidthUnits(double millimetres) {
		int pixels = (int)(millimetres * PIXELS_PER_MILLIMETRES);
		short widthUnits = (short) (EXCEL_COLUMN_WIDTH_FACTOR * (pixels / UNIT_OFFSET_LENGTH));
        widthUnits += UNIT_OFFSET_MAP[(pixels % UNIT_OFFSET_LENGTH)];
        return widthUnits;
    }

    /**
     * Convert a measure of pixels to millimetres (for column widths).
     * @param pixels
     * The size in pixels.
     * @return
     * The size in millimetres.
     */
	public static double pixels2Millimetres( double pixels ) {
		return pixels / PIXELS_PER_MILLIMETRES;
	}
	
	/**
	 * Convert a measure of millimetres to pixels (for column widths)
	 * @param millimetres
	 * The size in millimetres.
	 * @return
	 * The size in pixels.
	 */
	public static int millimetres2Pixels( double millimetres ) {
        return (int)(millimetres * PIXELS_PER_MILLIMETRES );
	}
	
}
