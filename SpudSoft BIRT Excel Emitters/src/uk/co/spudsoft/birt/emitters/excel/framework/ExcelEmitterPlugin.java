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

import org.eclipse.core.runtime.Plugin;
import org.osgi.framework.BundleContext;

/**
 * ExcelEmitterPlugin represents the SpudSoft Excel emitter.
 * <br/>
 * This is the activator for the plugin and is necessary to capture the eclipse log.
 * <br/>
 * Note that the BIRT runtime is not OSGi and does not activate the bundle, so getDefault always returns null when used in that environment.
 * 
 * @author Jim Talbut
 *
 */
public class ExcelEmitterPlugin extends Plugin {

	private static ExcelEmitterPlugin plugin;

	/**
	 * Get the plugin, if it has been activated.
	 * @return
	 * The plugin, if it has been activated, or null otherwise. 
	 */
	public static ExcelEmitterPlugin getDefault() {
		return plugin;
	}
	
	private Logger logger;

	/*
	 * (non-Javadoc)
	 * @see org.osgi.framework.BundleActivator#start(org.osgi.framework.BundleContext)
	 */
	public void start(BundleContext bundleContext) throws Exception {
		super.start(bundleContext);
		plugin = this;
		logger = new Logger(getLog(), bundleContext.getBundle().getSymbolicName());		
	}

	/*
	 * (non-Javadoc)
	 * @see org.osgi.framework.BundleActivator#stop(org.osgi.framework.BundleContext)
	 */
	public void stop(BundleContext bundleContext) throws Exception {
		super.stop(bundleContext);
	}
	
	/**
	 * Get the logger.
	 * @return
	 * The logger.
	 */
	public Logger getLogger() {
		return logger;
	}

}
