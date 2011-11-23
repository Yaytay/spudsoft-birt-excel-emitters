package uk.co.spudsoft.birt.emitters.excel;

import java.util.BitSet;
import java.util.Map;

import org.eclipse.birt.report.engine.content.IContent;
import org.eclipse.birt.report.engine.content.IStyle;
import org.eclipse.birt.report.engine.css.dom.AbstractStyle;
import org.eclipse.birt.report.engine.css.engine.CSSEngine;
import org.eclipse.birt.report.engine.css.engine.StyleConstants;
import org.eclipse.birt.report.engine.css.engine.value.DataFormatValue;
import org.eclipse.birt.report.engine.css.engine.value.StringValue;
import org.eclipse.birt.report.engine.css.engine.value.css.CSSConstants;
import org.eclipse.birt.report.engine.ir.Expression;
import org.eclipse.birt.report.engine.ir.ReportElementDesign;
import org.w3c.dom.css.CSSValue;

public class BirtStyle {
	
	public static final int NUMBER_OF_STYLES = StyleConstants.NUMBER_OF_STYLE + 1;
	public static final int TEXT_ROTATION = StyleConstants.NUMBER_OF_STYLE;

	protected static final String cssProperties[] = {
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
		   
		   , "Rotation"
	   };		
	
	
	private IStyle elemStyle;
	private CSSValue[] propertyOverride;
	private CSSEngine cssEngine;
	
	public BirtStyle( CSSEngine cssEngine ) {
		this.cssEngine = cssEngine;
	}
	
	public BirtStyle(IContent element) {
		elemStyle = element.getComputedStyle();
		
		if( elemStyle instanceof AbstractStyle ) {
			cssEngine = ((AbstractStyle)elemStyle).getCSSEngine();
		} else {
			throw new IllegalStateException( "Unable to obtain CSSEngine from elemStyle: " + elemStyle );
		}
		
		String rotation = extractRotation(element);
		if( rotation != null ) {
			setString(TEXT_ROTATION, rotation);
		}
	}
	
	private static String extractRotation(IContent element) {
		Object generatorObject = element.getGenerateBy();
		if( generatorObject instanceof ReportElementDesign ) {
			ReportElementDesign generatorDesign = (ReportElementDesign)generatorObject;
			Map<String,Expression> userProps = generatorDesign.getUserProperties(); 
			if( userProps != null ) {
				Expression rotationExpression = userProps.get( ExcelEmitter.ROTATION_PROP );
				if( rotationExpression != null ) {
					return rotationExpression.getScriptText();
				}
			}
		}
		return null;
	}
	
	public void setProperty( int propIndex, CSSValue newValue ) {
		if( propertyOverride == null ) {
			propertyOverride = new CSSValue[ BirtStyle.NUMBER_OF_STYLES ];
		}
		propertyOverride[ propIndex ] = newValue; 
	}
	
	public CSSValue getProperty( int propIndex ) {
		if( ( propertyOverride != null )
				&& ( propertyOverride[ propIndex ] != null ) ) {
			return propertyOverride[ propIndex ];
		}
		if( ( elemStyle != null ) && ( propIndex < StyleConstants.NUMBER_OF_STYLE ) ) {
			return elemStyle.getProperty( propIndex );
		} else {
			return null;
		}
	}
	
	public void setString( int propIndex, String newValue ) {
		if( propertyOverride == null ) {
			propertyOverride = new CSSValue[ BirtStyle.NUMBER_OF_STYLES ];
		}
			
		if( propIndex < StyleConstants.NUMBER_OF_STYLE ) {
			propertyOverride[ propIndex ] = cssEngine.parsePropertyValue( propIndex , newValue );
		} else {
			propertyOverride[ propIndex ] = new StringValue( StringValue.CSS_STRING, newValue);
		}
	}
	
	public String getString( int propIndex ) {
		CSSValue value = getProperty( propIndex );
		if( value != null ) {
			return value.getCssText();
		} else {
			return null;
		}
	}

	@Override
	protected BirtStyle clone() {
		BirtStyle result = new BirtStyle(this.cssEngine);

		result.propertyOverride = new CSSValue[ BirtStyle.NUMBER_OF_STYLES ];
				
		for(int i = 0; i < NUMBER_OF_STYLES; ++i ) {
			CSSValue value = getProperty( i );
			if( value != null ) {
				if( value instanceof DataFormatValue ) {
					value = StyleManagerUtils.cloneDataFormatValue((DataFormatValue)value);
				}
				
 				result.propertyOverride[ i ] = value;
 			}
		}
		
		return result;
	}

	private static final BitSet SPECIAL_OVERLAY_PROPERTIES = PrepareSpecialOverlayProperties();
	
	private static BitSet PrepareSpecialOverlayProperties() {
		BitSet result = new BitSet( BirtStyle.NUMBER_OF_STYLES );
		result.set( StyleConstants.STYLE_BACKGROUND_COLOR );
		result.set( StyleConstants.STYLE_BORDER_BOTTOM_STYLE );
		result.set( StyleConstants.STYLE_BORDER_BOTTOM_WIDTH );
		result.set( StyleConstants.STYLE_BORDER_BOTTOM_COLOR );
		result.set( StyleConstants.STYLE_BORDER_LEFT_STYLE );
		result.set( StyleConstants.STYLE_BORDER_LEFT_WIDTH );
		result.set( StyleConstants.STYLE_BORDER_LEFT_COLOR );
		result.set( StyleConstants.STYLE_BORDER_RIGHT_STYLE );
		result.set( StyleConstants.STYLE_BORDER_RIGHT_WIDTH );
		result.set( StyleConstants.STYLE_BORDER_RIGHT_COLOR );
		result.set( StyleConstants.STYLE_BORDER_BOTTOM_STYLE );
		result.set( StyleConstants.STYLE_BORDER_BOTTOM_WIDTH );
		result.set( StyleConstants.STYLE_BORDER_BOTTOM_COLOR );
		result.set( StyleConstants.STYLE_DATA_FORMAT );		
		return result;
	}
	
	private void overlayBorder( IStyle style, int propStyle, int propWidth, int propColour ) {
		CSSValue ovlStyle = style.getProperty( propStyle );
		CSSValue ovlWidth = style.getProperty( propWidth );
		CSSValue ovlColour = style.getProperty( propColour );
		if( ( ovlStyle != null )
				&& ( ovlWidth != null )
				&& ( ovlColour != null ) 
				&& ( ! CSSConstants.CSS_NONE_VALUE.equals( ovlStyle.getCssText() ) ) ) {
			setProperty( propStyle, ovlStyle );
			setProperty( propWidth, ovlWidth );
			setProperty( propColour, ovlColour );
		}
	}
	
	public void overlay( IContent element ) {
		
		// System.out.println( "overlay: Before - " + StyleManagerUtils.birtStyleToString( this ) );
		
		IStyle style = element.getComputedStyle();
		for(int propIndex = 0; propIndex < StyleConstants.NUMBER_OF_STYLE; ++propIndex ) {
			if( ! SPECIAL_OVERLAY_PROPERTIES.get(propIndex) ) {
				CSSValue overlayValue = style.getProperty( propIndex );
				CSSValue localValue = getProperty( propIndex );
				if( ( overlayValue != null ) && ! overlayValue.equals( localValue ) ) {
					setProperty( propIndex, overlayValue );
				}
			}	
		}
		
		// Background colour, only overlay if not null and not transparent
		CSSValue overlayBgColour = style.getProperty( StyleConstants.STYLE_BACKGROUND_COLOR );
		CSSValue localBgColour = getProperty( StyleConstants.STYLE_BACKGROUND_COLOR );
		if( ( overlayBgColour != null ) 
				&& ( ! CSSConstants.CSS_TRANSPARENT_VALUE.equals( overlayBgColour.getCssText() ) )
				&& ( ! overlayBgColour.equals( localBgColour ) ) ) {
			setProperty( StyleConstants.STYLE_BACKGROUND_COLOR, overlayBgColour );
		}
		
		// Borders, only overlay if all three components are not null - and then overlay all three
		overlayBorder( style, StyleConstants.STYLE_BORDER_BOTTOM_STYLE, StyleConstants.STYLE_BORDER_BOTTOM_WIDTH, StyleConstants.STYLE_BORDER_BOTTOM_COLOR );
		overlayBorder( style, StyleConstants.STYLE_BORDER_LEFT_STYLE, StyleConstants.STYLE_BORDER_LEFT_WIDTH, StyleConstants.STYLE_BORDER_LEFT_COLOR );
		overlayBorder( style, StyleConstants.STYLE_BORDER_RIGHT_STYLE, StyleConstants.STYLE_BORDER_RIGHT_WIDTH, StyleConstants.STYLE_BORDER_RIGHT_COLOR );
		overlayBorder( style, StyleConstants.STYLE_BORDER_BOTTOM_STYLE, StyleConstants.STYLE_BORDER_BOTTOM_WIDTH, StyleConstants.STYLE_BORDER_BOTTOM_COLOR );
		
		// Data format
		CSSValue overlayDataFormat = style.getProperty( StyleConstants.STYLE_DATA_FORMAT );
		CSSValue localDataFormat = getProperty( StyleConstants.STYLE_DATA_FORMAT );
		if( ! StyleManagerUtils.dataFormatsEquivalent((DataFormatValue)overlayDataFormat, (DataFormatValue)localDataFormat) ) {
			setProperty( StyleConstants.STYLE_DATA_FORMAT, StyleManagerUtils.cloneDataFormatValue((DataFormatValue)overlayDataFormat) );
		}
		
		// Rotation
		String rotation = extractRotation(element);
		if( rotation != null ) {
			setString(TEXT_ROTATION, rotation);
		}
		
		// System.out.println( "overlay: After - " + StyleManagerUtils.birtStyleToString( this ) );
	}

	@Override
	public String toString() {
		StringBuilder result = new StringBuilder();
		for( int i = 0; i < NUMBER_OF_STYLES; ++i ) {				
			CSSValue val = getProperty( i );
			if( val != null ) {
				try {
					result.append(cssProperties[i]).append(':').append(val.getCssText()).append("; ");
				} catch(Exception ex) {
					result.append(cssProperties[i]).append(":{").append(ex.getMessage()).append("}; ");						
				}
			}
		}
		return result.toString();
	}
	
	

}
