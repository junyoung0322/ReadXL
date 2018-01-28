import java.util.HashMap;

public class XLSXSharedStrings {
	private HashMap<Integer,String> mSharedStrings = new HashMap<Integer,String>();
	public void addSharedString(int Index,String SharedString) {
		mSharedStrings.put(new Integer(Index),SharedString);
	}
	public String getSharedString(int Index) {
		Integer iKey = new Integer(Index);
		if (mSharedStrings.containsKey(iKey))
			return mSharedStrings.get(iKey);

		return "";
	}
}
