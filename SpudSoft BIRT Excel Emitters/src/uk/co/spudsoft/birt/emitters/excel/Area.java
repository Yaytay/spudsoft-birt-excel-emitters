package uk.co.spudsoft.birt.emitters.excel;

public class Area {
    Coordinate x;
    Coordinate y;
    
    public Area(Coordinate x, Coordinate y) {
        this.x = x;
        this.y = y;
    }

	public Coordinate getX() {
		return x;
	}

	public Coordinate getY() {
		return y;
	}
    
}
