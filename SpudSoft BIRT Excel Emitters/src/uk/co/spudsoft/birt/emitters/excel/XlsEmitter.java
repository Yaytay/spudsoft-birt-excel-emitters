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
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.eclipse.birt.report.engine.content.IPageContent;
import org.eclipse.birt.report.engine.ir.DimensionType;

/**
 * XlsEmitter is the leaf class for implementing the ExcelEmitter with HSSFWorkbook.
 * @author Jim Talbut
 *
 */
public class XlsEmitter extends ExcelEmitter {

	/**
	 */
	public XlsEmitter() {
		super();
		setStyleManagerUtils(new StyleManagerHUtils(super.log));
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
	
	@Override
	protected int anchorDxFromMM( double widthMM, double colWidthMM ) {
        return (int)( 1023.0 * widthMM / colWidthMM );
	}
	
	@Override
	protected int anchorDyFromPoints( float height, float rowHeight ) {
        return (int)( 255.0 * height / rowHeight );
	}

	@Override
	protected void prepareMarginDimensions(IPageContent page) {
		if( page.getMarginBottom() != null ) {
			currentSheet.setMargin(Sheet.BottomMargin, page.getMarginBottom().convertTo(DimensionType.UNITS_IN));
		}
		if( page.getMarginLeft() != null ) {
			currentSheet.setMargin(Sheet.LeftMargin, page.getMarginLeft().convertTo(DimensionType.UNITS_IN));
		}
		if( page.getMarginRight() != null ) {
			currentSheet.setMargin(Sheet.RightMargin, page.getMarginRight().convertTo(DimensionType.UNITS_IN));
		}
		if( page.getMarginTop() != null ) {
			currentSheet.setMargin(Sheet.TopMargin, page.getMarginTop().convertTo(DimensionType.UNITS_IN));
		}
	}
	
}
