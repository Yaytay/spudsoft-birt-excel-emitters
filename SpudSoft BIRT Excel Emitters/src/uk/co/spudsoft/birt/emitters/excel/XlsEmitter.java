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

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * XlsEmitter is the leaf class for implementing the ExcelEmitter with HSSFWorkbook.
 * @author Jim Talbut
 *
 */
public class XlsEmitter extends ExcelEmitter {

	/**
	 */
	public XlsEmitter() {
		super(StyleManagerHUtils.getFactory());
		log.debug("Constructed XlsEmitter");
	}
	
	@Override
	public String getOutputFormat() {
		return "xls";
	}

	@Override
	protected Workbook createWorkbook() {
		return new HSSFWorkbook();
	}
	
}
