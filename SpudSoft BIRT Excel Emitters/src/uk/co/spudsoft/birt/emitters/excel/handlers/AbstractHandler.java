package uk.co.spudsoft.birt.emitters.excel.handlers;

import org.eclipse.birt.core.exception.BirtException;
import org.eclipse.birt.report.engine.content.IAutoTextContent;
import org.eclipse.birt.report.engine.content.ICellContent;
import org.eclipse.birt.report.engine.content.IContainerContent;
import org.eclipse.birt.report.engine.content.IContent;
import org.eclipse.birt.report.engine.content.IDataContent;
import org.eclipse.birt.report.engine.content.IForeignContent;
import org.eclipse.birt.report.engine.content.IGroupContent;
import org.eclipse.birt.report.engine.content.IImageContent;
import org.eclipse.birt.report.engine.content.ILabelContent;
import org.eclipse.birt.report.engine.content.IListBandContent;
import org.eclipse.birt.report.engine.content.IListContent;
import org.eclipse.birt.report.engine.content.IListGroupContent;
import org.eclipse.birt.report.engine.content.IPageContent;
import org.eclipse.birt.report.engine.content.IRowContent;
import org.eclipse.birt.report.engine.content.IStyledElement;
import org.eclipse.birt.report.engine.content.ITableBandContent;
import org.eclipse.birt.report.engine.content.ITableContent;
import org.eclipse.birt.report.engine.content.ITableGroupContent;
import org.eclipse.birt.report.engine.content.ITextContent;
import org.eclipse.birt.report.engine.css.engine.value.css.CSSConstants;
import org.w3c.dom.css.CSSValue;

import uk.co.spudsoft.birt.emitters.excel.HandlerState;
import uk.co.spudsoft.birt.emitters.excel.framework.Logger;

public class AbstractHandler implements IHandler {

	Logger log;
	IHandler parent;
	IStyledElement element;
	
	public AbstractHandler(Logger log,IHandler parent,IStyledElement element) {
		this.log = log;
		this.parent = parent;
		this.element = element;
	}
	
	@Override
	public IHandler getParent() {
		return parent;
	}
	
	@Override
	public String getBackgroundColour() {
		if( element != null ) {
			String elemColour = element.getComputedStyle().getBackgroundColor();
			if( ( elemColour != null ) && ! CSSConstants.CSS_TRANSPARENT_VALUE.equals( elemColour ) ) {
				return elemColour;
			}
		}
		if( parent != null ) {
			return parent.getBackgroundColour();
		}
		return CSSConstants.CSS_TRANSPARENT_VALUE;
	}
	
	protected static String getStyleProperty( IStyledElement element, int property, String defaultValue ) {
		CSSValue value = element.getComputedStyle().getProperty(property);
		if( value != null ) {
			return value.getCssText();
		} else {
			return defaultValue;
		}
	}
	
	public void startPage(HandlerState state, IPageContent page) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".startPage" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}
	public void endPage(HandlerState state, IPageContent page) throws BirtException{
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".endPage" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void startTable(HandlerState state, ITableContent table) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".startTable" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}
	public void endTable(HandlerState state, ITableContent table) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".endTable" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void startTableBand(HandlerState state, ITableBandContent band) throws BirtException {
		// NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".startTableBand" );
		// log.error(0, "Method not implemented", ex);
		// throw ex;
	}
	public void endTableBand(HandlerState state, ITableBandContent band) throws BirtException {
		// NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".endTableBand" );
		// log.error(0, "Method not implemented", ex);
		// throw ex;
	}

	public void startRow(HandlerState state, IRowContent row) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".startRow" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}
	public void endRow(HandlerState state, IRowContent row) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".endRow" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void startCell(HandlerState state, ICellContent cell) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".startCell" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}
	public void endCell(HandlerState state, ICellContent cell) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".endCell" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void startList(HandlerState state, IListContent list) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".startList" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}
	public void endList(HandlerState state, IListContent list) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".endList" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void startListBand(HandlerState state, IListBandContent listBand) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".startListBand" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}
	public void endListBand(HandlerState state, IListBandContent listBand) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".endListBand" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void startContainer(HandlerState state, IContainerContent container) throws BirtException {
		// NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".startContainer" );
		// log.error(0, "Method not implemented", ex);
		// throw ex;
	}
	public void endContainer(HandlerState state, IContainerContent container) throws BirtException {
		// NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".endContainer" );
		// log.error(0, "Method not implemented", ex);
		// throw ex;
	}

	public void startContent(HandlerState state, IContent content) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".startContent" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}
	public void endContent(HandlerState state, IContent content) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".endContent" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void startGroup(HandlerState state, IGroupContent group) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".startGroup" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}
	public void endGroup(HandlerState state, IGroupContent group) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".endGroup" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void startTableGroup(HandlerState state, ITableGroupContent group) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".startTableGroup" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}
	public void endTableGroup(HandlerState state, ITableGroupContent group) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".endTableGroup" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void startListGroup(HandlerState state, IListGroupContent group) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".startListGroup" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}
	public void endListGroup(HandlerState state, IListGroupContent group) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".endListGroup" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void emitText(HandlerState state, ITextContent text) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".emitText" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void emitData(HandlerState state, IDataContent data) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".emitData" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void emitLabel(HandlerState state, ILabelContent label) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".emitLabel" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void emitAutoText(HandlerState state, IAutoTextContent autoText) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".emitAutoText" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void emitForeign(HandlerState state, IForeignContent foreign) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".emitForeign" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

	public void emitImage(HandlerState state, IImageContent image) throws BirtException {
		NoSuchMethodError ex = new NoSuchMethodError( "Method not implemented: " + this.getClass().getSimpleName() + ".emitImage" );
		log.error(0, "Method not implemented", ex);
		throw ex;
	}

}
