public class XLSXCell {
	private XLSXSharedStrings oSharedStringLookup;
	private int iColumnNumber;
	private String iContents;
	private int iSharedIndex = -1;

	public XLSXCell(XLSXSharedStrings SharedStringLookup,int ColumnNumber, String Contents,boolean isSharedString) {
		oSharedStringLookup = SharedStringLookup;
		iColumnNumber = ColumnNumber;
		iContents = Contents == null ? "" : Contents;

		if (isSharedString && iContents.length() > 0) {
			boolean isNumber = true;
			for (int count = 0; count < iContents.length(); count++) {
				if (!Character.isDigit(iContents.charAt(count))) {
					isNumber = false;
					break;
				}
			}
			if (isNumber)
				iSharedIndex = Integer.parseInt(iContents);
		}
	}

	public int getColumnNumber() {
		return iColumnNumber;
	}

	public String getContents() {
		if (iSharedIndex >= 0)
			return oSharedStringLookup.getSharedString(iSharedIndex);

		return iContents;
	}
}