package uk.co.spudsoft.birt.emitters.excel;

import org.eclipse.birt.report.engine.css.engine.StyleConstants;
import org.eclipse.birt.report.engine.css.engine.value.css.CSSConstants;
import org.w3c.dom.css.CSSValue;

public class AreaBorders {
	public int bottom;
	public int left;
	public int right;
	public int top;
	
	public CSSValue[] cssStyle = new CSSValue[4];
	public CSSValue[] cssWidth = new CSSValue[4];
	public CSSValue[] cssColour = new CSSValue[4];

	private AreaBorders(int bottom, int left, int right, int top,
			CSSValue[] cssStyle, CSSValue[] cssWidth, CSSValue[] cssColour) {
		this.bottom = bottom;
		this.left = left;
		this.right = right;
		this.top = top;
		this.cssStyle = cssStyle;
		this.cssWidth = cssWidth;
		this.cssColour = cssColour;
	}


	public static AreaBorders create(int bottom, int left, int right, int top, BirtStyle borderStyle) {
		
		CSSValue borderStyleBottom = borderStyle.getProperty( StyleConstants.STYLE_BORDER_BOTTOM_STYLE );
		CSSValue borderWidthBottom = borderStyle.getProperty( StyleConstants.STYLE_BORDER_BOTTOM_WIDTH );
		CSSValue borderColourBottom = borderStyle.getProperty( StyleConstants.STYLE_BORDER_BOTTOM_COLOR );
		CSSValue borderStyleLeft = borderStyle.getProperty( StyleConstants.STYLE_BORDER_LEFT_STYLE );
		CSSValue borderWidthLeft = borderStyle.getProperty( StyleConstants.STYLE_BORDER_LEFT_WIDTH );
		CSSValue borderColourLeft = borderStyle.getProperty( StyleConstants.STYLE_BORDER_LEFT_COLOR );
		CSSValue borderStyleRight = borderStyle.getProperty( StyleConstants.STYLE_BORDER_RIGHT_STYLE );
		CSSValue borderWidthRight = borderStyle.getProperty( StyleConstants.STYLE_BORDER_RIGHT_WIDTH );
		CSSValue borderColourRight = borderStyle.getProperty( StyleConstants.STYLE_BORDER_RIGHT_COLOR );
		CSSValue borderStyleTop = borderStyle.getProperty( StyleConstants.STYLE_BORDER_TOP_STYLE );
		CSSValue borderWidthTop = borderStyle.getProperty( StyleConstants.STYLE_BORDER_TOP_WIDTH );
		CSSValue borderColourTop = borderStyle.getProperty( StyleConstants.STYLE_BORDER_TOP_COLOR );
				
/*		borderMsg.append( ", Bottom:" ).append( borderStyleBottom ).append( "/" ).append( borderWidthBottom ).append( "/" + borderColourBottom );
		borderMsg.append( ", Left:" ).append( borderStyleLeft ).append( "/" ).append( borderWidthLeft ).append( "/" + borderColourLeft );
		borderMsg.append( ", Right:" ).append( borderStyleRight ).append( "/" ).append( borderWidthRight ).append( "/" ).append( borderColourRight );
		borderMsg.append( ", Top:" ).append( borderStyleTop ).append( "/" ).append( borderWidthTop ).append( "/" ).append( borderColourTop );
		log.debug( borderMsg.toString() );
*/
		if( ( borderStyleBottom == null ) || ( CSSConstants.CSS_NONE_VALUE.equals( borderStyleBottom.getCssText() ) )
				|| ( borderWidthBottom == null ) || ( "0".equals(borderWidthBottom.getCssText()) )
				|| ( borderColourBottom == null ) || ( CSSConstants.CSS_TRANSPARENT_VALUE.equals(borderColourBottom) ) ) {
				borderStyleBottom = null;
				borderWidthBottom = null;
				borderColourBottom = null;
		}

		if( ( borderStyleLeft == null ) || ( CSSConstants.CSS_NONE_VALUE.equals( borderStyleLeft.getCssText() ) )
				|| ( borderWidthLeft == null ) || ( "0".equals(borderWidthLeft.getCssText()) )
				|| ( borderColourLeft == null ) || ( CSSConstants.CSS_TRANSPARENT_VALUE.equals(borderColourLeft) ) ) {
				borderStyleLeft = null;
				borderWidthLeft = null;
				borderColourLeft = null;
		}

        if( ( borderStyleRight == null ) || ( CSSConstants.CSS_NONE_VALUE.equals( borderStyleRight.getCssText() ) )
				|| ( borderWidthRight == null ) || ( "0".equals(borderWidthRight.getCssText()) )
				|| ( borderColourRight == null ) || ( CSSConstants.CSS_TRANSPARENT_VALUE.equals(borderColourRight) ) ) {
				borderStyleRight = null;
				borderWidthRight = null;
				borderColourRight = null;
		}

		if( ( borderStyleTop == null ) || ( CSSConstants.CSS_NONE_VALUE.equals( borderStyleTop.getCssText() ) )
				|| ( borderWidthTop == null ) || ( "0".equals(borderWidthTop.getCssText()) )
				|| ( borderColourTop == null ) || ( CSSConstants.CSS_TRANSPARENT_VALUE.equals(borderColourTop) ) ) {
				borderStyleTop = null;
				borderWidthTop = null;
				borderColourTop = null;
		}

		if( ( ( bottom >= 0 ) && ( ( borderStyleBottom != null ) || ( borderWidthBottom != null ) || ( borderColourBottom != null ) ) ) 
				|| ( ( left >= 0 ) && ( ( borderStyleLeft != null ) || ( borderWidthLeft != null ) || ( borderColourLeft != null ) ) )
				|| ( ( right >= 0 ) && ( ( borderStyleRight != null ) || ( borderWidthRight != null ) || ( borderColourRight != null ) ) ) 
				|| ( ( top >= 0 ) && ( ( borderStyleTop != null ) || ( borderWidthTop != null ) || ( borderColourTop != null ) ) ) 
				) {
			CSSValue[] cssStyle = new CSSValue[] { borderStyleBottom, borderStyleLeft, borderStyleRight, borderStyleTop };
			CSSValue[] cssWidth = new CSSValue[] { borderWidthBottom, borderWidthLeft, borderWidthRight, borderWidthTop };
			CSSValue[] cssColour = new CSSValue[] { borderColourBottom, borderColourLeft, borderColourRight, borderColourTop };
			return new AreaBorders(bottom, left, right, top, cssStyle, cssWidth, cssColour);
		}
		return null;
	}
	
}
