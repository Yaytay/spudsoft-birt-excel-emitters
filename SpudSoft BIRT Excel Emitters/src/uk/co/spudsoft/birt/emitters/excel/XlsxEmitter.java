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

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.birt.report.engine.content.IPageContent;
import org.eclipse.birt.report.engine.ir.DimensionType;

/**
 * XlsxEmitter is the leaf class for implementing the ExcelEmitter with XSSFWorkbook.
 * @author Jim Talbut
 *
 */
public class XlsxEmitter extends ExcelEmitter {
	
	/**
	 */
	public XlsxEmitter() {
		super();
		setStyleManagerUtils(new StyleManagerXUtils(super.log));
		log.debug("Constructed XlsxEmitter");
	}

	@Override
	public String getOutputFormat() {
		return "xlsx";
	}

	@Override
	protected Workbook createWorkbook() {
		return new XSSFWorkbook();
	}

	@Override
	protected int anchorDxFromMM( double widthMM, double colWidthMM ) {
        return (int)(widthMM * 36000); 
	}
	
	@Override
	protected int anchorDyFromPoints( float height, float rowHeight ) {
		return (int)( height * XSSFShape.EMU_PER_POINT );
	}

	@Override
	protected void prepareMarginDimensions(IPageContent page) {
		if( page.getHeaderHeight() != null ) {
			currentSheet.setMargin(Sheet.HeaderMargin, page.getHeaderHeight().convertTo(DimensionType.UNITS_IN));
		}
		if( page.getFooterHeight() != null ) {
			currentSheet.setMargin(Sheet.FooterMargin, page.getFooterHeight().convertTo(DimensionType.UNITS_IN));
		}
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
