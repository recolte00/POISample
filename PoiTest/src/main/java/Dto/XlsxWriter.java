package Dto;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;


public class XlsxWriter {


	/**
	 * sheet縺ｮxml縺ｫ繝�繝ｼ繧ｿ繧定ｿｽ蜉�縺吶ｋ
	 * header陦後′縺ゅｌ縺ｰ谿九＠縺ｦ縲∝ｾ檎ｶ壹↓霑ｽ蜉�縺吶ｋ
	 * header陦後�ｮ谺｡縺ｮ陦後�ｮ譖ｸ蠑上ｒ繧ｳ繝斐�ｼ縺励※菴ｿ逕ｨ縺吶ｋ
	 * header陦後�ｮ谺｡縺ｮ陦後↓蠑上′蜈･蜉帙＆繧後※縺�繧句�ｴ蜷医�∝推陦後↓繧ｳ繝斐�ｼ縺吶ｋ
	 */
	@SuppressWarnings("unchecked")
	public static void addDataToSheet(Document sheetXml, List<? extends XlsxWritable> datas) {
		Node sheetDataNode = sheetXml.getDocumentElement().getElementsByTagName("sheetData").item(0);
		int startRowNumber = getStartRowNumber(sheetDataNode);

		Node startRowNode = getRowNode(sheetXml, sheetDataNode, startRowNumber);

		// 蜈磯�ｭ�ｼ�header髯､縺擾ｼ芽｡後↓蟇ｾ縺吶ｋ譖ｸ蠑�
		Map<Integer, Integer> styleMap = getStyleMap(startRowNode);

		// 蛻怜�ｨ菴薙↓蟇ｾ縺吶ｋ譖ｸ蠑擾ｼ育ｯ�蝗ｲ險ｭ螳夲ｼ�
		Map<Integer, Integer> colStyleMap = getColStyleMap(sheetXml.getDocumentElement().getElementsByTagName("cols").item(0));

		// 蜈磯�ｭ�ｼ�header髯､縺擾ｼ芽｡後↓蟇ｾ縺吶ｋ謨ｰ蠑�
		Map<Integer, String> functionMap = getFunctionMap(startRowNode);

		// header陦後ｒ谿九＠縺ｦ蜈ｨ陦後ｒ荳�蠎ｦ蜑企勁
		XmlUtil.removeAllChilderenWithoutHeader(sheetDataNode, startRowNumber);

		for (int row = 0; row < datas.size(); row++) {
			// 陦後�ｮ菴懈��
			Node rowNode  = createRowNode(sheetXml, sheetDataNode, startRowNumber + row);

			// 蜷�蛻励�ｮ譖ｸ霎ｼ縺ｿ
			Map<Integer, Object> map = datas.get(row).getMap();
			for (Integer col : getKeyList(map.keySet(), functionMap.keySet())) {
				Node cellNode = createCellNode(sheetXml, startRowNumber + row, col, map.get(col), styleMap.get(col), colStyleMap.get(col));
				// 蠑上�ｮ繧ｳ繝斐�ｼ
				cellNode = addFuntion(sheetXml, cellNode, FunctionUtil.convertCellReferencesRow(functionMap.get(col), startRowNumber, startRowNumber + row), startRowNumber +row, col);
				if (cellNode != null) {
					rowNode.appendChild(cellNode);
				}
			}
			sheetDataNode.appendChild(rowNode);
		}
	}

	/**
	 * Cell縺ｫfunction繧定ｿｽ蜉�縺吶ｋ
	 * @param sheetXml
	 * @param cellNode null縺ｮ蝣ｴ蜷医�，ell Node繧団reate縺励※蠑上ｒ霑ｽ蜉�縺吶ｋ
	 * @param functionStr null縺ｾ縺溘�ｯ遨ｺ縺ｮ蝣ｴ蜷医�∝ｼ墓焚繧偵◎縺ｮ縺ｾ縺ｾ霑斐☆
	 * @param row
	 * @param col
	 * @return
	 */
	private static Node addFuntion(Document sheetXml, Node cellNode, String functionStr, int row, int col) {
		if (functionStr == null || functionStr.isEmpty()) {
			return cellNode;
		}
		if (cellNode != null) {
			return FunctionUtil.addFunctionStr(cellNode, functionStr);
		}
		Node newCellNode = createCell(sheetXml, row, col, null, null, null);
		return FunctionUtil.addFunctionStr(newCellNode, functionStr);
	}

	/**
	 * keySets繧痴ort貂医�ｮArrayList縺ｫ螟画鋤縺吶ｋ�ｼ�XML譖ｸ縺崎ｾｼ縺ｿ譎ゅ↓蟾ｦ蛻励°繧蛾��縺ｫ蜃ｦ逅�縺吶ｋ蠢�隕√′縺ゅｋ縺溘ａ�ｼ�
	 * @param keySets
	 * @return
	 */
	@SafeVarargs
	private static List<Integer> getKeyList(Set<Integer> ... keySets) {
		Set<Integer> key = new TreeSet<Integer>();
		for (Set<Integer> set : keySets) {
			for (Integer integer : set) {
				key.add(integer);
			}
		}
		List<Integer> list = new ArrayList<Integer>();
		list.addAll(key);
		return list;
	}

	/**
	 * cols縺ｫ螳夂ｾｩ縺輔ｌ縺ｦ縺�繧虐tyle繧偵�｀ap<蛻礼分蜿ｷ, style index>縺ｫ螟画鋤縺吶ｋ
	 * @param cols
	 * @return
	 */
	private static Map<Integer, Integer> getColStyleMap(Node cols) {
		Map<Integer, Integer> colStyleMap = new HashMap<Integer, Integer>();
		if (cols == null) {
			return colStyleMap;
		}

		NodeList childNodes = cols.getChildNodes();
		for (int i = 0; i < childNodes.getLength(); i++) {
			Node cellNode = childNodes.item(i);
			String minStr = XmlUtil.getAttributeValue(cellNode, "min");
			String maxStr = XmlUtil.getAttributeValue(cellNode, "max");
			String style = XmlUtil.getAttributeValue(cellNode, "style");
			if (style == null || style.isEmpty()) {
				continue;
			}
			int[] colNoArray = getColNoArray(minStr, maxStr);
			for (int colNo : colNoArray) {
				colStyleMap.put(colNo, Integer.parseInt(style));
			}
		}
		return colStyleMap;
	}

	private static int[] getColNoArray(String minStr, String maxStr) {
		int min = Integer.parseInt(minStr);
		int max = Integer.parseInt(maxStr);
		int size = max - min + 1;
		int[] result = new int[size];
		for (int i = 0; i < size; i++) {
			result[i] = min + i;
		}
		return result;
	}

	/**
	 * 謖�螳夊｡後�ｮStyle繧偵�｀ap<蛻礼分蜿ｷ, style index>縺ｫ螟画鋤縺吶ｋ
	 * @param rowNode null縺ｮ蝣ｴ蜷医�∫ｩｺ縺ｮMap繧定ｿ斐☆
	 * @return
	 */
	public static Map<Integer, Integer> getStyleMap(Node rowNode) {
		Map<Integer, Integer> styleMap = new HashMap<Integer, Integer>();
		if (rowNode == null) {
			return styleMap;
		}

		NodeList childNodes = rowNode.getChildNodes();
		for (int i = 0; i < childNodes.getLength(); i++) {
			Node cellNode = childNodes.item(i);
			String r = XmlUtil.getAttributeValue(cellNode, "r");
			if (r == null || r.isEmpty()) {
				continue;
			}
			String s = XmlUtil.getAttributeValue(cellNode, "s");
			if (s == null || s.isEmpty()) {
				continue;
			}
			styleMap.put((int) getColumnIndex(r), Integer.parseInt(s));
		}
		return styleMap;
	}

	/**
	 * 謖�螳夊｡後�ｮ蠑上ｒ縲｀ap<蛻礼分蜿ｷ, 蠑上�ｮ譁�蟄怜��>縺ｫ螟画鋤縺吶ｋ
	 * @param rowNode null縺ｮ蝣ｴ蜷医�∫ｩｺ縺ｮMap繧定ｿ斐☆
	 * @return
	 */
	public static Map<Integer, String> getFunctionMap(Node rowNode) {
		Map<Integer, String> functionMap = new HashMap<Integer, String>();
		if (rowNode == null) {
			return functionMap;
		}

		NodeList childNodes = rowNode.getChildNodes();
		for (int i = 0; i < childNodes.getLength(); i++) {
			Node cellNode = childNodes.item(i);
			String r = XmlUtil.getAttributeValue(cellNode, "r");
			String f = FunctionUtil.getFunctionStr(cellNode);
			if (f != null) {
				functionMap.put((int) getColumnIndex(r), f);
			}
		}
		return functionMap;
	}

	/**
	 * row繧貞叙蠕励☆繧九�∝ｭ伜惠縺励↑縺�蝣ｴ蜷医�ｯnull繧定ｿ斐☆
	 * @param sheetXml
	 * @param sheetDataNode
	 * @param rowNumber
	 * @return
	 */
	public static Node getRowNode(Document sheetXml, Node sheetDataNode, int rowNumber) {
		NodeList rows = sheetDataNode.getChildNodes();
		for (int i = 0; i < rows.getLength(); i++) {
			Node node = rows.item(i);
			String rValue = XmlUtil.getAttributeValue(node, "r");
			if (rValue == null) {
				continue;
			}
			if (rValue.equals(Integer.toString(rowNumber + 1))) {
				return node;
			}
		}
		return null;
	}

	/**
	 * row繧剃ｽ懈�舌☆繧�
	 * @param sheetXml
	 * @param sheetDataNode
	 * @param rowNumber
	 * @return
	 */
	public static Node createRowNode(Document sheetXml, Node sheetDataNode, int rowNumber) {
		Element newNode = sheetXml.createElement("row");
		newNode.setAttribute("r", Integer.toString(rowNumber + 1));
		return newNode;
	}

	/**
	 * 蜈ｨ蛻励′遨ｺ縺ｾ縺溘�ｯ蠑上�ｮ陦後ｒ譛�蛻昴�ｮ陦後→蛻､譁ｭ縺吶ｋ
	 * 遨ｺ縺ｮ繧ｷ繝ｼ繝医�ｮ蝣ｴ蜷医�ｯ縲�0繧定ｿ斐☆
	 * @param sheetDataNode
	 * @return
	 */
	public static int getStartRowNumber(Node sheetDataNode) {
		NodeList rows = sheetDataNode.getChildNodes();
		for(int i = 0; i < rows.getLength(); i++) {
			if (FunctionUtil.isEmptyNode(rows.item(i))) {
				return i;
			}
		}
		return rows.getLength();
	}


	/**
	 * (0, 0)竍�(A1)
	 * @param rowIndex
	 * @param columnIndex
	 * @return
	 */
	private static String getCellName(int rowIndex, int columnIndex) {
		return new CellReference(rowIndex, columnIndex).formatAsString();
	}

	/**
	 * (A2)竍�(1)
	 * @param cellReference
	 * @return
	 */
	private static short getColumnIndex(String cellReference) {
		return new CellReference(cellReference).getCol();
	}


	/**
	 * @param row
	 * @param col
	 * @param value
	 * @param styleIndex
	 * @param styleIndexCol
	 * @return
	 */
	public static Node createCellNode(Document sheetXml, int row, int col, Object value, Integer styleIndex, Integer styleIndexCol) {
		if (value == null) {
			if (styleIndex == null) {
				return null;
			} else {
				return createCell(sheetXml, row, col, null, styleIndex, styleIndexCol);
			}
		}
		if (value instanceof String) {
			return createStringCellNode(sheetXml, row, col, (String) value, styleIndex, styleIndexCol);
		} else if (value instanceof Number) {
			return createNumberCellNode(sheetXml, row, col, (Number) value, styleIndex, styleIndexCol);
		} else if (value instanceof BigDecimal) {
			return createNumberCellNode(sheetXml, row, col, (BigDecimal) value, styleIndex, styleIndexCol);
		} else if (value instanceof Date) {
			return createDateCellNode(sheetXml, row, col, (Date) value, styleIndex, styleIndexCol);
		} else if (value instanceof Calendar) {
			return createDateCellNode(sheetXml, row, col, ((Calendar) value).getTime(), styleIndex, styleIndexCol);
		} else {
			return createStringCellNode(sheetXml, row, col, value.toString(), styleIndex, styleIndexCol);
		}
	}

	private static Element createCell(Document sheetXml, int row, int col, String attributeT, Integer styleIndex, Integer styleIndexCol) {
		Element colNode = sheetXml.createElement("c");
		colNode.setAttribute("r", getCellName(row, col));
		if (attributeT != null && !attributeT.isEmpty()) {
			colNode.setAttribute("t", attributeT);
		}
		if (styleIndex != null) {
			colNode.setAttribute("s", styleIndex.toString());
		} else if (styleIndexCol != null) {
			colNode.setAttribute("s", styleIndexCol.toString());
		}
		return colNode;
	}

	private static Node createStringCellNode(Document sheetXml, int row, int col, String value, Integer styleIndex, Integer styleIndexCol) {
		Element colNode = createCell(sheetXml, row, col, "inlineStr", styleIndex, styleIndexCol);

		Element isNode = sheetXml.createElement("is");
		colNode.appendChild(isNode);

		Element tNode = sheetXml.createElement("t");
		isNode.appendChild(tNode);
		tNode.appendChild(sheetXml.createTextNode(value));

		return colNode;
	}

	private static Node createNumberCellNode(Document sheetXml, int row, int col, Number value, Integer styleIndex, Integer styleIndexCol) {
		Element colNode = createCell(sheetXml, row, col, "n", styleIndex, styleIndexCol);

		Element vNode = sheetXml.createElement("v");
		colNode.appendChild(vNode);

		vNode.appendChild(sheetXml.createTextNode(Double.toString(value.doubleValue())));

		return colNode;
	}

	private static Node createNumberCellNode(Document sheetXml, int row, int col, BigDecimal value, Integer styleIndex, Integer styleIndexCol) {
		return createNumberCellNode(sheetXml, row, col, value.doubleValue(), styleIndex, styleIndexCol);
	}

	private static Node createDateCellNode(Document sheetXml, int row, int col, Date value, Integer styleIndex, Integer styleIndexCol) {
		Element colNode = createCell(sheetXml, row, col, "n", styleIndex, styleIndexCol);

		Element vNode = sheetXml.createElement("v");
		colNode.appendChild(vNode);

		vNode.appendChild(sheetXml.createTextNode(Double.toString(DateUtil.getExcelDate(value))));

		return colNode;
	}
}
