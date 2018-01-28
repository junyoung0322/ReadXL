import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

public class XLSXContentReader {
	private String WORKSHEETPATH = "xl/worksheets/";
	private String SHAREDSTRINGPATH = "xl/sharedstrings.xml";
	private XLSXWorkbook oWorkbook = new XLSXWorkbook();
	private XLSXSharedStrings oSharedStrings = new XLSXSharedStrings();

	public XLSXContentReader(String Filename) throws IOException {
		this(new File(Filename));
	}

	public XLSXContentReader(File File) throws IOException {
		this(new FileInputStream(File));
	}

	public XLSXWorkbook getWorkbook() {
		return oWorkbook;
	}

	public XLSXContentReader(InputStream Stream) throws IOException {
		ZipInputStream oZip = new ZipInputStream(Stream);

		ZipEntry oEntry;
		while ((oEntry = oZip.getNextEntry()) != null) {
			if (oEntry.isDirectory())
				continue;

			if (oEntry.getName().compareToIgnoreCase(SHAREDSTRINGPATH) == 0) {
				ByteArrayOutputStream buffer = new ByteArrayOutputStream(8192);
				while (true) {
					int b = oZip.read();
					if (b == -1)
						break;
					buffer.write(b);
				}
				parseSharedStrings(buffer.toByteArray());
				continue;
			}

			if (!oEntry.getName().toLowerCase().startsWith(WORKSHEETPATH))
				continue;

			String sFilename = oEntry.getName().substring(WORKSHEETPATH.length());
			if (sFilename.contains("/"))
				continue;

			ByteArrayOutputStream buffer = new ByteArrayOutputStream(8192);
			while (true) {
				int b = oZip.read();
				if (b == -1)
					break;
				buffer.write(b);
			}
			parseSpreadsheet(sFilename, buffer.toByteArray());
		}
		oZip.close();
	}

	private void parseSharedStrings(byte [] Contents) {
		try
		{
			DocumentBuilder dBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
			Document dom = dBuilder.parse(new ByteArrayInputStream(Contents));

			int index = 0;
			NodeList shared = dom.getElementsByTagName("si");
			for (int current = 0; current < shared.getLength(); current++) {
				Element si = (Element)shared.item(current);

				NodeList lTextData = si.getElementsByTagName("t");
				if (lTextData.getLength() > 0) {
					Element el = (Element)lTextData.item(0);
					if ( el.getFirstChild() != null ) {
						String sValue = el.getFirstChild().getNodeValue();
						oSharedStrings.addSharedString(index++,sValue);
					} else {
						oSharedStrings.addSharedString(index++,null);
					}
				}
			}
		} catch (Exception ee) {
			System.out.println(ee.getMessage());
		}
	}

	private void parseSpreadsheet(String Filename,byte [] Contents) {
		String sSheetName = Filename;
		if (Filename.contains("."))
			sSheetName = Filename.substring(0,Filename.lastIndexOf("."));

		int iSheetIndex = oWorkbook.addSheet(sSheetName);

		try {
			DocumentBuilder dBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
			Document dom = dBuilder.parse(new ByteArrayInputStream(Contents));


			NodeList lSheetData = dom.getElementsByTagName("sheetData");
			for (int iCurrent = 0; iCurrent < lSheetData.getLength(); iCurrent++)
				parseSheetData((Element)lSheetData.item(iCurrent),iSheetIndex);
		} catch (Exception ee) {
			System.out.println(ee.getMessage());
		}
	}

	private void parseSheetData(Element SheetData,int SheetIndex) throws Exception {
		NodeList lRowData = SheetData.getElementsByTagName("row");
		for (int curRow = 0; curRow < lRowData.getLength(); curRow++) {
			Element row = (Element)lRowData.item(curRow);

			try {
				int iRowIndex = Integer.parseInt(row.getAttribute("r")) - 1;

				NodeList lColData = row.getElementsByTagName("c");
				for (int curCol = 0; curCol < lColData.getLength(); curCol++) {
					Element col = (Element)lColData.item(curCol);
					int iColIndex = DecodeColumnNumber(col.getAttribute("r"));

					boolean isSharedString = false;
					String sTypeString = col.getAttribute("t");
					if (sTypeString != null && sTypeString.compareToIgnoreCase("s") == 0)
						isSharedString = true;

					NodeList lValueData = col.getElementsByTagName("v");
					if (lValueData.getLength() == 0)
						lValueData = col.getElementsByTagName("t");
					if (lValueData.getLength() > 0) {
						Element el = (Element)lValueData.item(0);
						String sValue = el.getFirstChild().getNodeValue();
						oWorkbook.addCell(SheetIndex,iRowIndex,
								new XLSXCell(oSharedStrings,iColIndex,sValue,isSharedString));
					}
				}
			} catch (Exception ee) {
				System.out.println(ee.getMessage());
			}
		}
	}

	private int DecodeColumnNumber(String Attribute) {
		if (Character.isLetter(Attribute.charAt(1))) {
			int iMultiplier = (Character.toUpperCase(Attribute.charAt(0)) - 'A') + 1;
			int iColumn = Character.toUpperCase(Attribute.charAt(1)) - 'A';

			return iMultiplier * 26 + iColumn;
		}

		return Character.toUpperCase(Attribute.charAt(0)) - 'A';
	}
}