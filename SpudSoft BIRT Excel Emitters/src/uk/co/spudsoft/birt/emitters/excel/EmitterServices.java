package uk.co.spudsoft.birt.emitters.excel;

import java.util.Map;

import org.eclipse.birt.report.engine.api.ITaskOption;
import org.eclipse.birt.report.engine.ir.Expression;

import uk.co.spudsoft.birt.emitters.excel.framework.ExcelEmitterPlugin;

public class EmitterServices {

	/**
	 * Convert an Object to a boolean, with quite a few options about the class of the Object. 
	 * @param options
	 * The task options to extract the value from.
	 * @param name
	 * The name of the value to extract from options.
	 * @param defaultValue
	 * Value to return if value is null.
	 * @return
	 * true if value in some way represents a boolean TRUE value.
	 */
	public static boolean booleanOption( ITaskOption options, Map<String,Expression> userProperties, String name, boolean defaultValue ) {
		boolean result = defaultValue;
		Object value = null;
		
		if( options != null ) {
			value = options.getOption(name);
		}
		if( userProperties != null ) {
			Expression expression = userProperties.get(name);
			if( expression instanceof Expression.Constant ) {
				Expression.Constant constant = (Expression.Constant)expression;
				value = constant.getValue();
			}
		}
		
		if( value != null ) {
			result = booleanOption(value, defaultValue);
		}
		
		return result;
	}
	
	/**
	 * Convert an Object to a boolean, with quite a few options about the class of the Object. 
	 * @param value
	 * A value that can be of any type.
	 * @param defaultValue
	 * Value to return if value is null.
	 * @return
	 * true if value in some way represents a boolean TRUE value.
	 */
	public static boolean booleanOption( Object value, boolean defaultValue ) {
		if( value != null ) {
			if( value instanceof Boolean ) {
				return ((Boolean)value).booleanValue();
			}
			if( value instanceof Number ) {
				return ((Number)value).doubleValue() != 0.0;
			}
			if( value != null ) {
				return Boolean.parseBoolean(value.toString());
			}
		}
		return defaultValue;
	}
	
	
	/**
	 * Returns the symbolic name for the plugin.
	 */
	public static String getPluginName() {
		if( ( ExcelEmitterPlugin.getDefault() != null ) && ( ExcelEmitterPlugin.getDefault().getBundle() != null ) ) {
			return ExcelEmitterPlugin.getDefault().getBundle().getSymbolicName();
		} else {
			return "uk.co.spudsoft.birt.emitters.excel";
		}
	}
	

}
