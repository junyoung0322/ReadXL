import java.util.HashMap;

public class XLSXSheet {
	private String sSheetName;
	private HashMap<Integer,XLSXRow> mRows = new HashMap<Integer,XLSXRow>();
	private int iTotalRows = 0;

	public XLSXSheet(String SheetName) {
		sSheetName = SheetName;
	}

	public String getSheetName() {
		return sSheetName;
	}

	public int getTotalRows() {
		return iTotalRows;
	}

	public void addRow(XLSXRow Row) {
		mRows.put(new Integer(Row.getRowNumber()),Row);
		if (Row.getRowNumber() >= iTotalRows)
			iTotalRows = Row.getRowNumber() + 1;
	}

	public void addCell(int RowIndex,XLSXCell Cell) {
		Integer key = new Integer(RowIndex);
		if (mRows.containsKey(key))
			mRows.get(key).addCell(Cell);
		else {
			XLSXRow row = new XLSXRow(RowIndex);
			addRow(row);
			row.addCell(Cell);
		}
	}

	public XLSXRow getRowAt(int RowNumber) {
		if (RowNumber < iTotalRows) {
			Integer key = new Integer(RowNumber);
			if (mRows.containsKey(key))
				return mRows.get(key);
		}
		return new XLSXRow(RowNumber);
	}
}