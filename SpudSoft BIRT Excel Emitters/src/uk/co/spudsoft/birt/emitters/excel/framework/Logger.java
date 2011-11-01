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

package uk.co.spudsoft.birt.emitters.excel.framework;

import org.eclipse.core.runtime.ILog;
import org.eclipse.core.runtime.IStatus;
import org.eclipse.core.runtime.Status;
import org.slf4j.LoggerFactory;

/**
 * The Logger for the SpudSoft BIRT Excel Emitter.
 * <br/>
 * In a standard eclipse environment the Logger wraps the eclipse ILog and discards debug messages.
 * In the BIRT runtime environment the Logger uses slf4j and is dependent upon the whatever slf4j implementation is in use.
 * <p>
 * The Logger maintains a stack of characters as a prefix applied to any debug log.
 * This is used to track the start/end of items reported by BIRT.
 * </p>  
 * @author Jim Talbut
 *
 */
public class Logger {
	
	private ILog eclipseLog;
	private String pluginId;
	private org.slf4j.Logger backupLog;  
	private StringBuilder prefix = new StringBuilder();
	private boolean debug;
	
	/**
	 * Constructor used to initialise the slf4j logger.
	 * @param pluginId
	 * The plugin ID used to identify the logger with slf4j. 
	 */
	public Logger( String pluginId ) {
		this.backupLog = LoggerFactory.getLogger("uk.co.spudsoft.birt.emitters.excel");
	}
	
	/** 
	 * Constructor used to initialise the eclipse ILog.
	 * @param log
	 * The eclipse ILog.
	 * @param pluginId
	 * The plugin ID used in IStatus messages.
	 */
	Logger( ILog log, String pluginId ) {
		this.eclipseLog = log;
		this.pluginId = pluginId;
	}
	
	/**
	 * Set the debug state of the logger.
	 * @param debug
	 * When true and run within Equinox debug statements are output to the console.
	 */
	public void setDebug( boolean debug ) {
		this.debug = debug;
	}

	/**
	 * Add a new character to the prefix stack.
	 * @param c
	 * Character to add to the prefix stack.
	 */
	public void addPrefix( char c ) {
		prefix.append( c );
	}

	/**
	 * Remove a character from the prefix stack, if the appropriate character is at the top of the stack.
	 * @param c
	 * Character to remove from the prefix stack.
	 * @throws IllegalStateException
	 * If the prefix at the top of the prefix stack does not match c.
	 */
	public void removePrefix( char c ) {
		int length = prefix.length();
		char old = prefix.charAt( length - 1 );
		if(old != c ) {
			throw new IllegalStateException( "Old prefix (" + old + ") does not match that expected (" + c + "), whole prefix is \"" + prefix + "\"" );
		}
		prefix.setLength( length - 1);
	}
	
	/** 
	 * Log a message with debug severity. 
	 * @param message
	 * The message to log.
	 */
	public void debug( String message ) {
		if( eclipseLog != null ) {
			if( debug ) {
				System.out.println( prefix.toString() + " " + message );
			}
		} else {
			backupLog.debug( prefix.toString() + " " + message );
		}
	}
	
	/**
	 * Log a message with info severity.
	 * @param code
	 * The message code.
	 * @param message
	 * The message to log.
	 * @param exception
	 * Any exception associated with the log.
	 */
	public void info( int code, String message, Throwable exception ) {
		if( eclipseLog != null ) {
			log( IStatus.INFO, code, message, exception );
		} else {
			backupLog.info(message, exception);
		}
	}
	
	/**
	 * Log a message with warn severity.
	 * @param code
	 * The message code.
	 * @param message
	 * The message to log.
	 * @param exception
	 * Any exception associated with the log.
	 */
	public void warn( int code, String message, Throwable exception ) {
		if( eclipseLog != null ) {
			log( IStatus.WARNING, code, message, exception );
		} else {
			backupLog.warn(message, exception);
		}
	}
	
	/**
	 * Log a message with error severity.
	 * @param code
	 * The message code.
	 * @param message
	 * The message to log.
	 * @param exception
	 * Any exception associated with the log.
	 */
	public void error( int code, String message, Throwable exception ) {
		if( eclipseLog != null ) {
			log( IStatus.ERROR, code, message, exception );
		} else {
			backupLog.error(message, exception);
		}
	}
	
	private void log( int severity, int code, String message, Throwable exception ) {
		if( eclipseLog != null ) {
			IStatus record = new Status( severity, pluginId, code, message, exception);
			eclipseLog.log(record);
		}
	}
	
}
