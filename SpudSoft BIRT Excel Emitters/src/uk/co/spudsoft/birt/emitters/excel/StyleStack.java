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

import java.util.ListIterator;
import java.util.Stack;

import org.eclipse.birt.report.engine.content.IStyle;
import org.eclipse.birt.report.engine.content.IStyledElement;
import org.w3c.dom.css.CSSValue;

/**
 * 
 * StyleStack maintains a stack of BIRT styles so that can be applied hierarchically to elements.
 * @author Jim Talbut
 *
 */
public class StyleStack {
	
	Stack<IStyledElement> stack = new Stack<IStyledElement>();

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
	
	/**
	 * Get a property from the stack of styles, returning the top-most value that is set.
	 * @param property
	 * The property to be returned.
	 * @return
	 * The value of the property from the top-most item on the stack to specify it.
	 */
	public CSSValue getProperty1( int property ) {
		for(ListIterator<IStyledElement> iter = stack.listIterator(stack.size()); iter.hasPrevious();){
			IStyledElement element = iter.previous();
			CSSValue value = element.getStyle().getProperty( property );
			if( value != null ) {
				return value;
			}
		}
		return null;
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
	public <T> void mergeTop1( IStyledElement element, Class<T> clazz ) {
		IStyledElement target = stack.lastElement();
		
		IStyle targetStyle = target.getStyle();
		IStyle sourceStyle = element.getStyle();
		for(int i = 0; i < IStyle.NUMBER_OF_STYLE; ++i ) {
			CSSValue value = sourceStyle.getProperty( i );
			if( value != null ) {
				targetStyle.setProperty( i , value );
			}
		}
	}
	
}
