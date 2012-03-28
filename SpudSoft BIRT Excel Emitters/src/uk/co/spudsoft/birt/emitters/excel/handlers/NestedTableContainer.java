package uk.co.spudsoft.birt.emitters.excel.handlers;


public interface NestedTableContainer extends IHandler {

	public void addNestedTable( NestedTableHandler nestedTableHandler );
	public boolean rowHasNestedTable( int rowNum );
	public int extendRowBy( int rowNum );
	
}
