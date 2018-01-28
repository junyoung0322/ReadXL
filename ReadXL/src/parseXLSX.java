import java.util.Date;

public class parseXLSX {
	/**
	 * @param args
	 * @throws Exception
	 */
	public static void main(String [] args) throws Exception {
		System.out.println(new Date());
		final String sSlash = "/";
		final String sFileName = sSlash + "file" + sSlash + "data.xlsx";
		StringBuilder sPathSb = new StringBuilder();
		String sProjectPath = null;

		String sClassPath = parseXLSX.class.getResource("").getPath();
		String[] sPaths = sClassPath.split("/");
		int iCount = 1;
		for (int i = 1; i < sPaths.length-1; i++) {
			if (iCount != 1) {
				sPathSb.append(sSlash);
			}
			sPathSb.append(sPaths[i]);
			iCount++;
		}
		sProjectPath = sPathSb.toString();

		System.out.println(sProjectPath + sFileName);
		XLSXWorkbook oWorkbook = (new XLSXContentReader(sProjectPath + sFileName)).getWorkbook();

		for (int curSheet = 0; curSheet < oWorkbook.getTotalSheets(); curSheet++) {
			XLSXSheet oSheet = oWorkbook.getSheetAt(curSheet);

			int iNameLen = 10;
			int iScoreLen = 8;
			int iTotalSLen = iScoreLen + 2;

			printn(" 名前" , iNameLen); System.out.print(" | ");
			printn("国語", iScoreLen); printn("数学", iScoreLen); printn("英語", iScoreLen); printn("社会", iScoreLen); printn("理科", iScoreLen);
			printn("合計点", iTotalSLen);
			for (int y = 3; y < oSheet.getTotalRows(); y++) {
				XLSXRow row = oSheet.getRowAt(y);
				System.out.println("");
				boolean isName = true;
				StringBuilder sNameSb = new StringBuilder();
				int iTotalScore = 0;
				for (int x = 1; x < row.getTotalColumns(); x++) {
					XLSXCell cell = row.getCellAt(x);
					if (isName) {
						sNameSb.append(" ").append(cell.getContents());
						if (x == 2) {
							isName = false;

							String sName = sNameSb.toString();
							printn(sName, iNameLen);
							System.out.print(" | ");
						}
					}
					else {
						String sScore = cell.getContents();
						iTotalScore += Integer.parseInt(sScore);
						printn(cell.getContents(), iScoreLen);
					}
				}
				printn(String.valueOf(iTotalScore), iTotalSLen);
			}
		}
	}
	static void printn(String s, int len) {
		int l = 0;
		for (int i = 0; i < s.length(); i++) {
			char c = s.charAt(i);
			if ((c <= '\u007e') || (c == '\u00a5') || (c == '\u203e') || (c >= '\uff61' && c <= '\uff9f')) {
				l++;
				if (l > len) {
					l--;
					break;
				}
			} else {
				l += 2;
				if (l > len) {
					l -= 2;
					break;
				}
			}
			System.out.print(c);
		}
		if (l < len) {
			System.out.printf("%" + (len - l) + "s", "");
		}
	}
}

