import java.util.ArrayList;
import java.util.List;

public class XLSXWorkbook {
	private List<XLSXSheet> sSheets = new ArrayList<XLSXSheet>();
	public int getTotalSheets() {
		return sSheets.size();
	}

	public int addSheet(String SheetName) {
		int index = sSheets.size();
		sSheets.add(new XLSXSheet(SheetName));
		return index;
	}
	public XLSXSheet getSheet(String SheetName) {
		for (XLSXSheet sheet : sSheets) {
			if (sheet.getSheetName().compareToIgnoreCase(SheetName) == 0)
				return sheet;
		}
		return null;
	}

	public XLSXSheet getSheetAt(int SheetIndex) throws IndexOutOfBoundsException {
		if (SheetIndex < 0 || SheetIndex >= sSheets.size())
			throw new IndexOutOfBoundsException("Sheet index is invalid");

		return sSheets.get(SheetIndex);
	}

	public void addCell(int SheetIndex,int RowIndex,XLSXCell Cell) throws IndexOutOfBoundsException {
		if (SheetIndex < 0 || SheetIndex >= sSheets.size())
			throw new IndexOutOfBoundsException("Sheet index is invalid");

		sSheets.get(SheetIndex).addCell(RowIndex,Cell);
	}
}
