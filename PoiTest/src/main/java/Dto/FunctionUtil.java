package Dto;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.util.CellReference;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * Excel蠑上�ｮ繧ｳ繝斐�ｼ逕ｨ縺ｮUtil繧ｯ繝ｩ繧ｹ
 */
public class FunctionUtil {

	private static final String REG_CELL_REFERENCE = "\\$?[A-Z]+\\$?[0-9]+|\\$?[A-Z]+:\\$?[A-Z]+|\\$?[0-9]+:\\$?[0-9]+";
	private static final Pattern PATTERN_CELL_REFERENCE = Pattern.compile(REG_CELL_REFERENCE);
	private static final String REG_STR = "\".*\"";
	private static final Pattern PATTERN_STR = Pattern.compile(REG_STR);


	/**
	 * 蠑上ｒ蜷ｫ繧�繧ｻ繝ｫ縺ｮ蝣ｴ蜷医�∝ｼ上ｒ霑斐☆
	 * 縺昴ｌ莉･螟悶�ｮ蝣ｴ蜷医�ｯ縲］ull繧定ｿ斐☆
	 * @param cellNode
	 * @return
	 */
	public static String getFunctionStr(Node cellNode) {
		NodeList nodeList = cellNode.getChildNodes();
		for (int i = 0; i < nodeList.getLength(); i++) {
			Node n = nodeList.item(i);
			if (n.getNodeName().equals("f")) {
				return n.getChildNodes().item(0).getTextContent();
			}
		}
		return null;
	}

	/**
	 * 繧ｻ繝ｫ縺ｫ蠑上ｒ霑ｽ蜉�縺吶ｋ
	 * @param cellNode
	 * @param functionStr
	 */
	public static Node addFunctionStr(Node cellNode, String functionStr) {
		Document document = cellNode.getOwnerDocument();
		Node fNode = document.createElement("f");
		Node textNode = document.createTextNode(functionStr);
		fNode.appendChild(textNode);
		cellNode.insertBefore(fNode, cellNode.getFirstChild());
		return cellNode;
	}


	/**
	 * 蠖楢ｩｲNode縺悟�､縺檎ｩｺ縺ｾ縺溘�ｯ蠑上�ｮ縺ｿ繧貞性繧�Node縺ｮ縺ｿ縺九ｉ讒区�舌＆繧後ｋ蝣ｴ蜷医�》rue繧定ｿ斐☆縲�
	 * @param node
	 * @return
	 */
	public static boolean isEmptyNode(Node node) {
		NodeList nodeList = node.getChildNodes();
		for (int i = 0; i < nodeList.getLength(); i++) {
			Node childNode = nodeList.item(i);
			if (getFunctionStr(childNode) != null) {
				continue;
			} else if (!childNode.hasChildNodes()) {
				String value = childNode.getNodeValue();
				if (value != null && !value.isEmpty()) {
					return false;
				}
			} else {
				if (!isEmptyNode(childNode)) {
					return false;
				}
			}
		}
		return true;
	}

	/**
	 * 蠑上�ｮ繧ｻ繝ｫ蜿ら�ｧ繧貞､画鋤縺吶ｋ�ｼ郁｡後�ｮ遘ｻ蜍包ｼ�<br/>
	 * {@link #convertCellReferences(String, int, int, int, int)}繧貞他縺ｶ
	 * @param originalStr null縺ｮ蝣ｴ蜷医�］ull繧定ｿ斐☆
	 * @param srcRow 繧ｻ繝ｫ蜿ら�ｧ蜿門ｾ励そ繝ｫ縺ｮ蛻礼分蜿ｷ
	 * @param destRow 險ｭ螳壼�医そ繝ｫ縺ｮ蛻礼分蜿ｷ
	 * @return
	 */
	public static String convertCellReferencesRow(String originalStr, int srcRow, int destRow) {
		return convertCellReferences(originalStr, 0, srcRow, 0, destRow);
	}

	/**
	 * 蠑上�ｮ繧ｻ繝ｫ蜿ら�ｧ繧貞､画鋤縺吶ｋ
	 * @param originalStr null縺ｮ蝣ｴ蜷医�］ull繧定ｿ斐☆
	 * @param srcCol
	 * @param srcRow
	 * @param destCol
	 * @param destRow
	 * @return
	 */
	public static String convertCellReferences(String originalStr, int srcCol, int srcRow, int destCol, int destRow) {
		if (originalStr == null || originalStr.isEmpty()) {
			return originalStr;
		}
		String resultStr = "";
		// 譁�蟄怜�励ｒ蜑肴婿縺九ｉ鬆�縺ｫ蜃ｦ逅�
		while(!originalStr.isEmpty()) {
			Matcher strMatcher = PATTERN_STR.matcher(originalStr);
			Matcher referenceMatcher = PATTERN_CELL_REFERENCE.matcher(originalStr);
			if (!referenceMatcher.find()) {
				// 繧ｻ繝ｫ蜿ら�ｧ繧貞性縺ｾ縺ｪ縺�蝣ｴ蜷医�√◎縺ｮ縺ｾ縺ｾ霑斐☆
				resultStr += originalStr;
				return resultStr;
			} else if (!strMatcher.find() || (referenceMatcher.start() < strMatcher.start())) {
				// 繧ｻ繝ｫ蜿ら�ｧ繧貞性繧�縺梧枚蟄怜�励ｒ蜷ｫ縺ｾ縺ｪ縺�蝣ｴ蜷医�√ｂ縺励￥縺ｯ縲√そ繝ｫ蜿ら�ｧ縺梧枚蟄怜�励ｈ繧雁燕縺ｫ縺ゅｋ蝣ｴ蜷医�∵怙蛻昴�ｮ繧ｻ繝ｫ蜿ら�ｧ繧貞､画鋤縺励※蜃ｦ逅�繧堤ｶ壹￠繧�
				int referenceStart = referenceMatcher.start();
				resultStr += originalStr.substring(0, referenceStart);
				originalStr = originalStr.substring(referenceStart);
				resultStr += convertCellReference(referenceMatcher.group(), srcCol, srcRow, destCol, destRow);
				originalStr = originalStr.replaceFirst(REG_CELL_REFERENCE, "");
			} else {
				// 譁�蟄怜�励′繧ｻ繝ｫ蜿ら�ｧ繧医ｊ蜑阪↓縺ゅｋ蝣ｴ蜷医�∵枚蟄怜�励ｒ遘ｻ縺励※蜃ｦ逅�繧堤ｶ壹￠繧�
				resultStr += originalStr.substring(0, strMatcher.end(0));
				originalStr = originalStr.substring(strMatcher.end(0));
			}
		}
		return resultStr;

	}

	/**
	 * 繧ｳ繝斐�ｼ蜈�繧ｻ繝ｫ縺九ｉ繧ｳ繝斐�ｼ蜈医そ繝ｫ縺ｫ繧ｻ繝ｫ蜿ら�ｧ繧偵さ繝斐�ｼ縺励◆蝣ｴ蜷医�ｮ縲√さ繝斐�ｼ蠕後�ｮ蜿ら�ｧ繧定ｿ斐☆
	 * ex. ("A1", 1, 1, 2, 2) => "B2"
	 * @param cellReference
	 * @param srcCol
	 * @param srcRow
	 * @param destCol
	 * @param destRow
	 * @return
	 */
	public static String convertCellReference(String cellReference, int srcCol, int srcRow, int destCol, int destRow) {
		if (cellReference.contains(":")) {
			// 遽�蝗ｲ縺ｮ蝣ｴ蜷医�∝�榊ｸｰ逧�縺ｫ蜃ｦ逅�
			String[] range = cellReference.split(":");
			return convertCellReference(range[0], srcCol, srcRow, destCol, destRow) + ":" + convertCellReference(range[1], srcCol, srcRow, destCol, destRow);
		}
		CellReference cr = new CellReference(cellReference);
		String col;
		if (cr.getCol() == -1 ) {
			col = "";
		} else if (cr.isColAbsolute()) {
			col = "$" + CellReference.convertNumToColString(cr.getCol());
		} else {
			col = CellReference.convertNumToColString(cr.getCol() - srcCol + destCol);
		}

		String row;
		if (cr.getRow() == -1) {
			row = "";
		} else if (cr.isRowAbsolute()) {
			row = "$" + (cr.getRow() + 1);
		} else {
			row = Integer.toString(cr.getRow() - srcRow + destRow + 1);
		}
		return col + row;
	}
}
