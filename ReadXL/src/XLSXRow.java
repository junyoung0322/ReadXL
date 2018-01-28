import java.util.HashMap;

public class XLSXRow {
	private int iRowNumber;
	private HashMap<Integer,XLSXCell> mCells = new HashMap<Integer,XLSXCell>();
	private int iTotalColumns = 0;

	public XLSXRow(int RowNumber) {
		iRowNumber = RowNumber;
	}

	public int getRowNumber() {
		return iRowNumber;
	}

	public int getTotalColumns() {
		return iTotalColumns;
	}

	public void addCell(XLSXCell Cell) {
		mCells.put(new Integer(Cell.getColumnNumber()),Cell);
		if (Cell.getColumnNumber() >= iTotalColumns)
			iTotalColumns = Cell.getColumnNumber() + 1;
	}

	public XLSXCell getCellAt(int ColumnNumber) {
		if (ColumnNumber < iTotalColumns) {
			Integer key = new Integer(ColumnNumber);
			if (mCells.containsKey(key))
				return mCells.get(key);
		}
		return new XLSXCell(null,ColumnNumber,"",false);
	}
}