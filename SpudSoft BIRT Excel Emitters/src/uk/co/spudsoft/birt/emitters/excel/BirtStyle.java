package uk.co.spudsoft.birt.emitters.excel;

import java.util.BitSet;

import org.eclipse.birt.report.engine.content.IStyle;
import org.eclipse.birt.report.engine.content.IStyledElement;
import org.eclipse.birt.report.engine.css.dom.AbstractStyle;
import org.eclipse.birt.report.engine.css.engine.CSSEngine;
import org.eclipse.birt.report.engine.css.engine.StyleConstants;
import org.eclipse.birt.report.engine.css.engine.value.DataFormatValue;
import org.eclipse.birt.report.engine.css.engine.value.css.CSSConstants;
import org.w3c.dom.css.CSSValue;

public class BirtStyle {

	IStyle elemStyle;
	CSSValue[] propertyOverride;
	CSSEngine cssEngine;
	
	public BirtStyle( CSSEngine cssEngine ) {
		this.cssEngine = cssEngine;
	}
	
	public BirtStyle(IStyledElement element) {
		elemStyle = element.getComputedStyle();
		
		if( elemStyle instanceof AbstractStyle ) {
			cssEngine = ((AbstractStyle)elemStyle).getCSSEngine();
		} else {
			throw new IllegalStateException( "Unable to obtain CSSEngine from elemStyle: " + elemStyle );
		}
	}
	
	public void setProperty( int propIndex, CSSValue newValue ) {
		if( propertyOverride == null ) {
			propertyOverride = new CSSValue[ StyleConstants.NUMBER_OF_STYLE ];
		}
		propertyOverride[ propIndex ] = newValue; 
	}
	
	public CSSValue getProperty( int propIndex ) {
		if( ( propertyOverride != null )
				&& ( propertyOverride[ propIndex ] != null ) ) {
			return propertyOverride[ propIndex ];
		}
		if( elemStyle != null ) {
			return elemStyle.getProperty( propIndex );
		} else {
			return null;
		}
	}
	
	public void setString( int propIndex, String newValue ) {
		if( propertyOverride == null ) {
			propertyOverride = new CSSValue[ StyleConstants.NUMBER_OF_STYLE ];
		}
			
		propertyOverride[ propIndex ] = cssEngine.parsePropertyValue( propIndex , newValue );
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

		result.propertyOverride = new CSSValue[ StyleConstants.NUMBER_OF_STYLE ];
				
		for(int i = 0; i < IStyle.NUMBER_OF_STYLE; ++i ) {
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

	private static BitSet SPECIAL_OVERLAY_PROPERTIES = PrepareSpecialOverlayProperties();
	
	private static BitSet PrepareSpecialOverlayProperties() {
		BitSet result = new BitSet( StyleConstants.NUMBER_OF_STYLE );
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
	
	public void overlay( IStyle style ) {
		
		// System.out.println( "overlay: Before - " + StyleManagerUtils.birtStyleToString( this ) );
		
		for(int propIndex = 0; propIndex < IStyle.NUMBER_OF_STYLE; ++propIndex ) {
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
		
		// System.out.println( "overlay: After - " + StyleManagerUtils.birtStyleToString( this ) );
	}
	
	

}
