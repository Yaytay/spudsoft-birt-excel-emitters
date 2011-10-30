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

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.net.URLConnection;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import org.eclipse.birt.report.engine.content.IStyle;
import org.eclipse.birt.report.engine.css.dom.AbstractStyle;
import org.eclipse.birt.report.engine.css.dom.AreaStyle;
import org.eclipse.birt.report.engine.css.engine.CSSEngine;
import org.eclipse.birt.report.engine.css.engine.value.DataFormatValue;
import org.eclipse.birt.report.engine.ir.DimensionType;
import org.w3c.dom.css.CSSValue;

import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

/**
 * <p>
 * StyleManagerUtils contains methods implementing the details of converting BIRT styles to POI styles.
 * </p><p>
 * StyleManagerUtils is abstract to support a small number of methods that require HSSF/XSSF specific implementations.
 * 
 * @author Jim Talbut
 *
 */
public abstract class StyleManagerUtils {
	
	   protected static String cssProperties[] = {
		     "margin-left"
		   , "margin-right"
		   , "margin-top"
		   , "DATA_FORMAT"
		   , "border-right-color"
		   , "direction"
		   , "border-top-width"
		   , "padding-left"
		   , "border-right-width"
		   , "padding-bottom"
		   , "padding-top"
		   , "NUMBER_ALIGN"
		   , "padding-right"
		   , "CAN_SHRINK"
		   , "border-top-color"
		   , "background-repeat"
		   , "margin-bottom"
		   , "background-width"
		   , "background-height"
		   , "border-right-style"
		   , "border-bottom-color"
		   , "text-indent"
		   , "line-height"
		   , "border-bottom-width"
		   , "text-align"
		   , "background-color"
		   , "color"
		   , "overflow"
		   , "TEXT_LINETHROUGH"
		   , "border-left-color"
		   , "widows"
		   , "border-left-width"
		   , "border-bottom-style"
		   , "font-weight"
		   , "font-variant"
		   , "text-transform"
		   , "white-space"
		   , "TEXT_OVERLINE"
		   , "vertical-align"
		   , "BACKGROUND_POSITION_X"
		   , "border-left-style"
		   , "VISIBLE_FORMAT"
		   , "MASTER_PAGE"
		   , "orphans"
		   , "font-size"
		   , "font-style"
		   , "border-top-style"
		   , "page-break-before"
		   , "SHOW_IF_BLANK"
		   , "background-image"
		   , "BACKGROUND_POSITION_Y"
		   , "word-spacing"
		   , "background-attachment"
		   , "TEXT_UNDERLINE"
		   , "display"
		   , "font-family"
		   , "letter-spacing"
		   , "page-break-inside"
		   , "page-break-after"
	   };		

	protected Logger log;	
	
	/**
	 * @param log
	 * The Logger to use for any information reports to be made.
	 */
	public StyleManagerUtils(Logger log) {
		this.log = log;
	}

	/**
	 * Compare two objects in a null-safe manner.
	 * @param lhs
	 * The first object to compare.
	 * @param rhs
	 * The second object to compare.
	 * @return
	 * true is both objects are null or lhs.equals(rhs), otherwise false.
	 */
	public static boolean objectsEqual(Object lhs, Object rhs) {
		return (lhs == null) ? (rhs == null) : lhs.equals(rhs);  
	}
	
	/**
	 * Convert a BIRT text alignment string into a POI CellStyle constant.
	 * @param alignment
	 * The BIRT alignment string.
	 * @return
	 * One of the CellStyle.ALIGN* constants.
	 */
	public short poiAlignmentFromBirtAlignment(String alignment) {
		if("left".equals(alignment)) {
			return CellStyle.ALIGN_LEFT;
		}
		if("right".equals(alignment)) {
			return CellStyle.ALIGN_RIGHT;
		}
		if("center".equals(alignment)) {
			return CellStyle.ALIGN_CENTER;
		}
		return CellStyle.ALIGN_GENERAL;
	}
	
	/**
	 * Convert a BIRT font size string (either a dimensioned string or "xx-small" - "xx-large") to a point size. 
	 * @param fontSize
	 * The BIRT font size.
	 * @return
	 * An appropriate size in points.
	 */
	public short fontSizeInPoints(String fontSize) {
		if( fontSize == null ) {
			return 0;
		}
		if("xx-small".equals(fontSize)) {
			return 6;
		} else if("x-small".equals(fontSize)) {
			return 8;
		} else if("small".equals(fontSize)) {
			return 10;
		} else if("medium".equals(fontSize)) {
			return 12;
		} else if("large".equals(fontSize)) {
			return 14;
		} else if("x-large".equals(fontSize)) {
			return 18;
		} else if("xx-large".equals(fontSize)) {
			return 24;
		} else if("smaller".equals(fontSize)) {
			return 10;
		} else if("larger".equals(fontSize)) {
			return 14;
		}
		
		DimensionType dim = DimensionType.parserUnit(fontSize, "pt");
		log.debug( "fontSize: \"" + fontSize + "\", parses as: \"" + dim.toString() + "\" (" + dim.getMeasure() + " " + dim.getUnits() + ")");
		if(DimensionType.UNITS_PX.equals(dim.getUnits())) {
			log.debug( ", cannot convert px, so returning " + dim.getMeasure() );
			return (short)dim.getMeasure();
		} else {
			double points = dim.convertTo(DimensionType.UNITS_PT);
			log.debug( ", converts as: \"" + points + "pt\"");
			return (short)points;
		}
	}
	
	/**
	 * Obtain a POI column width from a BIRT DimensionType. 
	 * @param dim
	 * The BIRT dimension, which must be in absolute units.
	 * @return
	 * The column with in width units, or zero if a suitable conversion could not be performed.
	 */
	public int poiColumnWidthFromDimension( DimensionType dim ) {
		if (dim != null) {
			double mmWidth = dim.convertTo( "mm" );
			int result = ClientAnchorConversions.millimetres2WidthUnits(mmWidth);
			log.debug( "Column width in mm: " + mmWidth + "; converted result: " + result );			
			return result;
		} else {
			return 0;
		}
	}
	
	/**
	 * Object a POI font weight from a BIRT string.
	 * @param fontWeight
	 * The font weight as understood by BIRT.
	 * @return
	 * One of the Font.BOLDWEIGHT_* constants.
	 */
	public short poiFontWeightFromBirt(String fontWeight) {
		if(fontWeight == null) {
			return 0;
		}
		if("bold".equals(fontWeight)) {
			return Font.BOLDWEIGHT_BOLD;
		}
		return Font.BOLDWEIGHT_NORMAL;
	}
	
	/**
	 * Convert a BIRT font name into a system font name.
	 * <br>
	 * Just returns the passed in name unless that is a known family name ("serif" or "sans-serif").
	 * @param fontName
	 * The font name from BIRT.
	 * @return
	 * A real font name.
	 */
	public String poiFontNameFromBirt(String fontName) {
		if("serif".equals(fontName)) {
			return "Times New Roman";
		} else if("sans-serif".equals(fontName)) {
			return "Arial";
		}
		return fontName;
	}
	
	/**
	 * <p>
	 * Add a colour (specified as "rgb(<i>r</i>, <i>g</i>, <i>b</i>)") to a Font.
	 * </p><p>
	 * In the current implementations the XSSF implementation will always produce exactly the right colour,
	 * whilst the HSSF implementation takes the best approximation from the current palette.
	 * @param workbook
	 * The workbook in which the Font is to be used, needed to obtain the colour palette.
	 * @param font
	 * The font to which the colour is to be added.
	 * @param colour
	 * The colour to add.
	 */
	public abstract void addColourToFont(Workbook workbook, Font font, String colour);
	
	/**
	 * <p>
	 * Add a colour (specified as "rgb(<i>r</i>, <i>g</i>, <i>b</i>)") as the background colour of a CellStyle.
	 * </p><p>
	 * In the current implementations the XSSF implementation will always produce exactly the right colour,
	 * whilst the HSSF implementation takes the best approximation from the current palette.
	 * @param workbook
	 * The workbook in which the Font is to be used, needed to obtain the colour palette.
	 * @param style
	 * The style to which the colour is to be added.
	 * @param colour
	 * The colour to add.
	 */
	public abstract void addBackgroundColourToStyle(Workbook workbook, CellStyle style, String colour);
	
	/**
	 * Convert a BIRT style to a string for debug purposes.
	 * @param style
	 * The BIRT style.
	 * @return
	 * A string representing all the configured values in the BIRT style.
	 */
	public String birtStyleToString(IStyle style) {
		StringBuilder result = new StringBuilder();
		if(!style.isEmpty()) {
			for( int i = 0; i < IStyle.NUMBER_OF_STYLE; ++i ) {				
				CSSValue val = style.getProperty( i );
				if( val != null ) {
					try {
						result.append(cssProperties[i]).append(':').append(val.getCssText()).append("; ");
					} catch(Exception ex) {
						result.append(cssProperties[i]).append(":{").append(ex.getMessage()).append("}; ");						
					}
				}
			}
		}
		return result.toString();
	}
	
	/**
	 * Check whether a cell is empty and unformatted.
	 * @param cell
	 * The cell to consider.
	 * @return
	 * true is the cell is empty and has no style or has no background fill.
	 */
	public boolean cellIsEmpty(Cell cell) {
		if( cell.getCellType() != Cell.CELL_TYPE_BLANK ) {
			return false;
		}
		CellStyle cellStyle = cell.getCellStyle();
		if( cellStyle == null ) {
			return true;
		}
		if( cellStyle.getFillPattern() == CellStyle.NO_FILL ) {
			return true;
		}
		return false;		
	}
	
	/**
	 * Apply a BIRT border style to one side of a POI CellStyle.
	 * @param workbook
	 * The workbook that contains the cell being styled.
	 * @param style
	 * The POI CellStyle that is to have the border applied to it. 
	 * @param side
	 * The side of the border that is to be applied.<br>
	 * Note that although this value is from XSSFCellBorder it is equally valid for HSSFCellStyles.
	 * @param colour
	 * The colour for the new border.
	 * @param borderStyle
	 * The BIRT style for the new border.
	 * @param width
	 * The width of the new border.
	 */
	public abstract void applyBorderStyle(Workbook workbook, CellStyle style, BorderSide side, String colour, String borderStyle, String width);
	
	/**
	 * <p>
	 * Convert a MIME string into a Workbook.PICTURE* constant.
	 * </p><p>
	 * In some cases BIRT fails to submit a MIME string, in which case this method falls back to basic data signatures for JPEG and PNG images.
	 * <p>
	 * @param mimeType
	 * The MIME type.
	 * @param data
	 * The image data to consider if no recognisable MIME type is provided.
	 * @return
	 * A Workbook.PICTURE* constant.
	 */
	public int poiImageTypeFromMimeType( String mimeType, byte[] data ) {
		if( "image/jpeg".equals(mimeType) ) {
			return Workbook.PICTURE_TYPE_JPEG;
		} else if( "image/png".equals(mimeType) ) {
			return Workbook.PICTURE_TYPE_PNG;
		} else {
			if( null != data ) {
				log.debug( "Data bytes: "
						+ " " + Integer.toHexString( data[0] ).toUpperCase()  
						+ " " + Integer.toHexString( data[1] ).toUpperCase()  
						+ " " + Integer.toHexString( data[2] ).toUpperCase()
						+ " " + Integer.toHexString( data[3] ).toUpperCase()
						);
				if( ( data.length > 2 )
						&& ( data[0] == (byte)0xFF)
						&& ( data[1] == (byte)0xD8) 
						&& ( data[2] == (byte)0xFF)
						) {
					return Workbook.PICTURE_TYPE_JPEG;
				}
				if( ( data.length > 4 )
						&& ( data[0] == (byte)0x89)
						&& ( data[1] == (byte)'P') 
						&& ( data[2] == (byte)'N') 
						&& ( data[3] == (byte)'G') 
						) {
					return Workbook.PICTURE_TYPE_PNG;
				}
			} 
			return 0;
		}
	}
	
	/**
	 * Read an InputStream in full and put the results into a byte[].
	 * <br>
	 * This is needed by the emitter to handle images accessed by URL.
	 * @param stream
	 * The InputStream to read.
	 * @param length
	 * The length of the InputStream
	 * @return
	 * A byte array containing the contents of the InputStream.
	 * @throws IOException
	 */
	public byte[] streamToByteArray( InputStream stream, int length ) throws IOException {
		ByteArrayOutputStream buffer;
		if( length > 0 ) {
			buffer = new ByteArrayOutputStream( length );
		} else {
			buffer = new ByteArrayOutputStream();
		}
	
		int nRead;
		byte[] data = new byte[16384];
	
		while ((nRead = stream.read(data, 0, data.length)) != -1) {
		  buffer.write(data, 0, nRead);
		}
	
		buffer.flush();
	
		return buffer.toByteArray();
	}
	
	/**
	 * Read an image from a URLConnection into a byte array.
	 * @param conn
	 * The URLConnection to provide the data.
	 * @return
	 * A byte array containing the data downloaded from the URL.
	 */
	public byte[] downloadImage( URLConnection conn ) {
		try {
			int contentLength = conn.getContentLength();
			InputStream imageStream = conn.getInputStream();
			try {
				return streamToByteArray( imageStream, contentLength );
			} finally {
				imageStream.close();
			}
		} catch( MalformedURLException ex ) {
			log.debug( ex.getClass().getName() + ": " + ex.getMessage() );
			return null;
		} catch( IOException ex ) {
			log.debug( ex.getClass().getName() + ": " + ex.getMessage() );
			return null;
		}
		
	}
	
	/**
	 * Convert a BIRT paper size string into a POI PrintSetup.*PAPERSIZE constant.
	 * @param name
	 * The paper size as a BIRT string.
	 * @return
	 * A POI PrintSetup.*PAPERSIZE constant.
	 */
	public short getPaperSizeFromString( String name ) {
		if( "a4".equals(name) ) {
			return PrintSetup.A4_PAPERSIZE;
		} else if( "a3".equals(name)) {
			return PrintSetup.A3_PAPERSIZE;
		} else if( "us-letter".equals(name)) {
			return PrintSetup.LETTER_PAPERSIZE;
		}
		
		return PrintSetup.A4_PAPERSIZE;
	}
	
	/**
	 * Check whether a DimensionType represents an absolute (physical) dimension.
	 * @param dim
	 * The DimensionType to consider.
	 * @return
	 * true if dim represents an absolute measurement.
	 */
	public boolean isAbsolute( DimensionType dim ) {
		if( dim == null ) {
			return false;
		}
		return DimensionType.UNITS_CM.equals(dim.getUnits())
				|| DimensionType.UNITS_IN.equals(dim.getUnits())
				|| DimensionType.UNITS_MM.equals(dim.getUnits())
				|| DimensionType.UNITS_PT.equals(dim.getUnits())
				;
	}
	
	/**
	 * Check whether a DimensionType represents pixels.
	 * @param dim
	 * The DimensionType to consider.
	 * @return
	 * true if dim represents pixels.
	 */
	public boolean isPixels( DimensionType dim ) {
		return (dim != null) && DimensionType.UNITS_PX.equals(dim.getUnits());
	}
	
	/**
	 * <p>
	 * Convert a BIRT number format to a POI data format.
	 * </p><p>
	 * There is no way this function is complete!  More special cases will be added as they are found.
	 * </p>
	 * @param birtFormat
	 * A string representing a number format in BIRT.
	 * @return
	 * A string representing a data format in Excel.
	 */
	private String poiNumberFormatFromBirt(String birtFormat) {
		if( "General Number".equalsIgnoreCase(birtFormat)) {
			return null;
		}
		birtFormat = birtFormat.replace("E00", "E+00");
		int brace = birtFormat.indexOf('{');
		if( brace >= 0 ) {
			birtFormat = birtFormat.substring(0, brace);
		}
		return birtFormat;
	}
	
	/**
	 * <p>
	 * Convert a BIRT date/time format to a POI data format.
	 * </p><p>
	 * This function is likely to be more complete than poiNumberFormatFromBirt, but it is still likely to have issues.
	 * More special cases will be added as they are found.
	 * </p>
	 * @param birtFormat
	 * A string representing a date/time format in BIRT.
	 * @return
	 * A string representing a data format in Excel.
	 */
	private String poiDateTimeFormatFromBirt(String birtFormat) {
        if ( "General Date".equalsIgnoreCase( birtFormat ) ) {
        	return "dd/MM/yyyy hh:mm";
        }
        if ( "Long Date".equalsIgnoreCase( birtFormat ) ) {
        	return "dddd, mmmm dd, yyyy";
        }
        if ( "Medium Date".equalsIgnoreCase( birtFormat ) ) {
        	return "ddd, dd mmm yyyy";
        }
        if ( "Short Date".equalsIgnoreCase( birtFormat ) ) {
        	return "yyyy-MM-dd";
        }
        if ( "Long Time".equalsIgnoreCase( birtFormat ) ) {
        	return "hh:mm:ss";
        }
        if ( "Medium Time".equalsIgnoreCase( birtFormat ) ) {
        	return "hh:mm";
        }
        if ( "Short Time".equalsIgnoreCase( birtFormat ) ) {
        	return "hh:mm";
        }
		return birtFormat;
	}
	
	/**
	 * Apply a BIRT number/date/time format to a POI CellStyle.
	 * @param workbook
	 * The workbook containing the CellStyle (needed to create a new DataFormat).
	 * @param birtStyle
	 * The BIRT style which may contain a number format.
	 * @param poiStyle
	 * The CellStyle that is to receive the number format.
	 */
	public void applyNumberFormat(Workbook workbook, IStyle birtStyle, CellStyle poiStyle) {
		String dataFormat = null;
		if( birtStyle.getNumberFormat() != null ) {
			log.debug( "BIRT number format == " + birtStyle.getNumberFormat());
			dataFormat = poiNumberFormatFromBirt(birtStyle.getNumberFormat());
		} else if( birtStyle.getDateTimeFormat() != null ) {
			log.debug( "BIRT date/time format == " + birtStyle.getDateTimeFormat());
			dataFormat = poiDateTimeFormatFromBirt(birtStyle.getDateTimeFormat());
		} else if( birtStyle.getTimeFormat() != null ) {
			log.debug( "BIRT time format == " + birtStyle.getTimeFormat());
			dataFormat = poiDateTimeFormatFromBirt(birtStyle.getTimeFormat());
		} else if( birtStyle.getDateFormat() != null ) {
			log.debug( "BIRT date format == " + birtStyle.getDateFormat());
			dataFormat = poiDateTimeFormatFromBirt(birtStyle.getDateFormat());
		}
		if( dataFormat != null ) {
			DataFormat format = workbook.createDataFormat();
			log.debug( "Setting POI data format to " + dataFormat);
			poiStyle.setDataFormat(format.getFormat(dataFormat));
		}
	}
		
	/**
	 * Create a new BIRT style that is the same as another BIRT style
	 */
	public IStyle copyBirtStyle( IStyle style ) {
		CSSEngine cssEngine = ((AbstractStyle)style).getCSSEngine();
		AreaStyle result = new AreaStyle( cssEngine );

		for(int i = 0; i < IStyle.NUMBER_OF_STYLE; ++i ) {
			CSSValue value = style.getProperty( i );
			if( value != null ) {
				if( value instanceof DataFormatValue ) {
					DataFormatValue dataValue = (DataFormatValue)value;
					DataFormatValue newValue = new DataFormatValue();
					newValue.setDateFormat( dataValue.getDatePattern(), dataValue.getDateLocale() );
					newValue.setDateTimeFormat( dataValue.getDateTimePattern(), dataValue.getDateTimeLocale() );
					newValue.setTimeFormat( dataValue.getTimePattern(), dataValue.getTimeLocale() );
					newValue.setNumberFormat( dataValue.getNumberPattern(), dataValue.getNumberLocale() );
					newValue.setStringFormat( dataValue.getStringPattern(), dataValue.getStringLocale() );
					value = newValue;
				}
				
 				result.setProperty( i , value );
 			}
		}
		
		return result;
	}
	
}
