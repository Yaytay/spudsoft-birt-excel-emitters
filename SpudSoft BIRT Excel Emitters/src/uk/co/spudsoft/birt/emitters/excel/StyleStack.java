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

import java.util.Stack;

import org.eclipse.birt.report.engine.content.IStyle;
import org.eclipse.birt.report.engine.content.IStyledElement;
import org.eclipse.birt.report.engine.css.engine.CSSEngine;
import org.eclipse.birt.report.engine.css.engine.StyleConstants;
import org.w3c.dom.css.CSSValue;

/**
 * 
 * StyleStack maintains a stack of BIRT styles so that can be applied hierarchically to elements.
 * @author Jim Talbut
 *
 */
public class StyleStack {
	
	Stack<IStyledElement> stack = new Stack<IStyledElement>();
	private CSSEngine cssEngine;
	
	public void setCssEngine( CSSEngine cssEngine ) {
		this.cssEngine = cssEngine;
	}
	
	/**
	 * Push a new BIRT styled element onto the stack.
	 * @param element
	 */
	public void push( IStyledElement element ) {
		stack.push( element );
	}
	
	/**
	 * Pop a styled element from the stack and, if it's of the right type, return it.
	 * @param clazz
	 * The expected type of the top item on the stack.
	 * @return
	 * The top item on the stack.
	 * @throws IllegalStateException
	 * If the top item on the stack is not an instance of clazz.
	 */
	@SuppressWarnings("unchecked")
	public <T> T pop( Class<T> clazz ) {
		IStyledElement element = stack.pop();
		assert(clazz.isInstance(element));
		if(clazz.isInstance(element)) {
			return (T)element;
		} else {
			throw new IllegalStateException( "The top element on the stack is of type " + element.getClass().getName() + " rather than the expected " + clazz.getName());
		}
	}
	
	
/*	private void dumpStack() {
		System.out.print( "Stack: ");
		for( IStyledElement element : stack ) {
			System.out.print( element.getClass().getSimpleName() );
			System.out.print( "[ " + element.getStyle().getBorderBottomColor() + "]" );
			System.out.print( "; " );
		}
		System.out.println();
	}
*/	
	private void mergeAllIfAny( IStyle destStyle, IStyle sourceStyle, int[] properties, String[] defaults ) {

		boolean any = false;
		for( int prop : properties ) {
			if( sourceStyle.getProperty( prop ) != null ) {
				any = true;
				break;
			}
		}
		
		if( any ) {
			for( int i = 0; i < properties.length; ++i ) {
				int prop = properties[ i ];
				CSSValue value = sourceStyle.getProperty( prop );
				if( value == null ) {
					destStyle.setProperty( prop, cssEngine.parsePropertyValue( prop, defaults[ i ] ) );
				} else {
					destStyle.setProperty( prop, value );
				}
			}			
		}
	}
	
	/**
	 * Merge the passed in styled element with the topmost element, as long as the top-most element is of the right type.
	 * @param element
	 * The element whose style is to be merged with the top-most element.
	 * @param clazz 
	 * The expected type of the top item on the stack.
	 * @throws IllegalStateException
	 * If the top item on the stack is not an instance of clazz.
	 */
	public <T> void mergeTop( IStyledElement element, Class<T> clazz ) {
		// System.out.println( "mergeTop [" +element.getClass().getSimpleName() + "]" + "[ " + element.getStyle().getBorderBottomColor() + "]" );
		// dumpStack();
		IStyledElement topElement = stack.lastElement();
/*		if( ! clazz.isInstance(topElement) ) {
			throw new IllegalStateException( "The top element on the stack is of type " + topElement.getClass().getName() + " rather than the expected " + clazz.getName());
		}
*/		
/*		IStyle birtStyle = topElement.getComputedStyle();
		System.out.println( "B4>>: " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_TOP_STYLE) + "[ " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_TOP_COLOR) + " & " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_TOP_WIDTH) + " ], " 
				+ birtStyle.getProperty(StyleConstants.STYLE_BORDER_RIGHT_STYLE) + "[ " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_RIGHT_COLOR) + " & " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_RIGHT_WIDTH) + " ], "
				+ birtStyle.getProperty(StyleConstants.STYLE_BORDER_BOTTOM_STYLE) + "[ " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_BOTTOM_COLOR) + " & " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_BOTTOM_WIDTH) + " ], "
				+ birtStyle.getProperty(StyleConstants.STYLE_BORDER_LEFT_STYLE) + "[ " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_LEFT_COLOR) + " & " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_LEFT_WIDTH) + " ], "
				);
*/
		
		IStyle topStyle = topElement.getStyle();
		IStyle sourceStyle = element.getStyle();
		for(int i = 0; i < IStyle.NUMBER_OF_STYLE; ++i ) {
			CSSValue value = sourceStyle.getProperty( i );
			if( value != null ) {
				topStyle.setProperty( i , value );
			}
		}
		
		// Special handling for borders, to permit defaults with null value values
		mergeAllIfAny( topStyle, sourceStyle, new int[] { StyleConstants.STYLE_BORDER_TOP_STYLE, StyleConstants.STYLE_BORDER_TOP_COLOR, StyleConstants.STYLE_BORDER_TOP_WIDTH }, new String[] { "solid", "rgb(0,0,0)", "medium" } );
		mergeAllIfAny( topStyle, sourceStyle, new int[] { StyleConstants.STYLE_BORDER_RIGHT_STYLE, StyleConstants.STYLE_BORDER_RIGHT_COLOR, StyleConstants.STYLE_BORDER_RIGHT_WIDTH }, new String[] { "solid", "rgb(0,0,0)", "medium" } );
		mergeAllIfAny( topStyle, sourceStyle, new int[] { StyleConstants.STYLE_BORDER_BOTTOM_STYLE, StyleConstants.STYLE_BORDER_BOTTOM_COLOR, StyleConstants.STYLE_BORDER_BOTTOM_WIDTH }, new String[] { "solid", "rgb(0,0,0)", "medium" } );
		mergeAllIfAny( topStyle, sourceStyle, new int[] { StyleConstants.STYLE_BORDER_LEFT_STYLE, StyleConstants.STYLE_BORDER_LEFT_COLOR, StyleConstants.STYLE_BORDER_LEFT_WIDTH }, new String[] { "solid", "rgb(0,0,0)", "medium" } );
		

/*		birtStyle = topElement.getComputedStyle();
		System.out.println( "AFTER>>: " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_TOP_STYLE) + "[ " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_TOP_COLOR) + " & " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_TOP_WIDTH) + " ], " 
				+ birtStyle.getProperty(StyleConstants.STYLE_BORDER_RIGHT_STYLE) + "[ " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_RIGHT_COLOR) + " & " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_RIGHT_WIDTH) + " ], "
				+ birtStyle.getProperty(StyleConstants.STYLE_BORDER_BOTTOM_STYLE) + "[ " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_BOTTOM_COLOR) + " & " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_BOTTOM_WIDTH) + " ], "
				+ birtStyle.getProperty(StyleConstants.STYLE_BORDER_LEFT_STYLE) + "[ " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_LEFT_COLOR) + " & " + birtStyle.getProperty(StyleConstants.STYLE_BORDER_LEFT_WIDTH) + " ], "
				);
*/
	}
	
	public IStyledElement top() {
		return stack.lastElement();
	}
	
}
