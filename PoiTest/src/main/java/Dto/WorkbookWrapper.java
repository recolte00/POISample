package Dto;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.xml.sax.SAXException;


/**
 * using inlineStr instead of sharedStrings
 *
 */
public class WorkbookWrapper {

	public static final String STYLE_DATE = "STYLE_DATE";

	/** template Zip */
	private ZipFile templateZip;

	/** template */
	private XSSFWorkbook templateWb;

	/** Map */
	Map<String, File> substituteMap = new HashMap<String, File>();

	public WorkbookWrapper(File file) throws FileNotFoundException, IOException {
		super();
		this.templateZip = new ZipFile(file);
		this.templateWb = new XSSFWorkbook(new FileInputStream(file));
	}

	public WorkbookWrapper(XSSFWorkbook wb) throws IOException {
		super();
		this.templateWb = wb;

		File templateFile = File.createTempFile("template", "xlsx");
		templateFile.deleteOnExit();
		templateWb.write(new FileOutputStream(templateFile));
		this.templateZip = new ZipFile(templateFile);
	}


	public void write(OutputStream os) throws IOException {
		ZipUtil.substitute(templateZip, substituteMap, os);
	}

	private String getSheetXmlName(String sheetName) {
		return "sheet" + (templateWb.getSheetIndex(sheetName) + 1);
	}

	/**
	 * 蠖楢ｩｲsheet縺ｮXML繧奪ocument蠖｢蠑上〒蜿門ｾ励☆繧�
	 * @param sheetName
	 * @return
	 * @throws ZipException
	 * @throws IOException
	 * @throws ParserConfigurationException
	 * @throws SAXException
	 */
	private Document getSheetXML(String sheetName) throws ZipException, IOException, ParserConfigurationException, SAXException {
		XSSFSheet sheet = templateWb.getSheet(sheetName);
		if (sheet == null) {
			return null;
		}
		return ZipUtil.getXmlDocument(templateZip, getEntry(sheet));
	}

	/**
	 * sheet縺九ｉzip譖ｸ霎ｼ縺ｿ逕ｨ縺ｮEntry繧貞叙蠕励☆繧�
	 * @param sheet
	 * @return
	 */
	private String getEntry(XSSFSheet sheet) {
		return sheet.getPackagePart().getPartName().getName().substring(1);
	}

	/**
	 * sheet縺ｫdatas繧呈嶌縺崎ｾｼ繧�<br/>
	 * header陦後′縺ゅｌ縺ｰ谿九＠縺ｦ縲∝ｾ檎ｶ壹↓霑ｽ蜉�縺吶ｋ<br/>
	 * header陦後�ｮ谺｡縺ｮ陦後�ｮ譖ｸ蠑上ｒ繧ｳ繝斐�ｼ縺励※菴ｿ逕ｨ縺吶ｋ<br/>
	 * @param sheetName
	 * @param prepareData
	 * @throws IOException
	 * @throws SAXException
	 * @throws ParserConfigurationException
	 * @throws TransformerException
	 */
	public void writeSheet(String sheetName, List<? extends XlsxWritable> datas) throws IOException, ParserConfigurationException, SAXException, TransformerException {
		Document sheetXml = getSheetXML(sheetName);
		if (sheetXml == null) {
			throw new IllegalArgumentException("No Such Sheet");
		}
		XlsxWriter.addDataToSheet(sheetXml, datas);
		substituteMap.put(getEntry(templateWb.getSheet(sheetName)), ZipUtil.createTempFileFromDocument(getSheetXmlName(sheetName), ".xml", sheetXml));
	}
}
