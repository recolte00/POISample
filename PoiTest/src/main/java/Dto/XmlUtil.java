package Dto;
import java.util.ArrayList;
import java.util.List;

import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 *
 */

/**
 * @author kazuhiro1-wada
 *
 */
public class XmlUtil {

	/**
	 * 蠖楢ｩｲNode縺悟�､縺檎ｩｺ縺ｮNode縺ｮ縺ｿ縺九ｉ讒区�舌＆繧後ｋ蝣ｴ蜷医�》rue繧定ｿ斐☆縲�
	 * @param node
	 * @return
	 */
	public static boolean isEmptyNode(Node node) {
		NodeList nodeList = node.getChildNodes();
		for (int i = 0; i < nodeList.getLength(); i++) {
			Node childNode = nodeList.item(i);
			if (!childNode.hasChildNodes()) {
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
	 * @param node
	 * @param attributeKey
	 * @return
	 */
	public static String getAttributeValue(Node node, String attributeKey) {
		NamedNodeMap attributes = node.getAttributes();
		Node attribute = attributes.getNamedItem(attributeKey);
		if (attribute == null) {
			return null;
		}
		return attribute.getNodeValue();
	}

	/**
	 * 蜈磯�ｭ蟄占ｦ∫ｴ�繧呈ｮ九＠縺ｦ縲∝ｭ占ｦ∫ｴ�繧貞炎髯､縺吶ｋ
	 * @param node
	 */
	public static void removeAllChildren(Node node) {
		removeAllChilderenWithoutHeader(node, 0);
	}

	/**
	 * 蜈磯�ｭ縺九ｉ謖�螳壹�ｮ謨ｰ縺ｮ蟄占ｦ∫ｴ�繧呈ｮ九＠縺ｦ縲∝ｭ占ｦ∫ｴ�繧貞炎髯､縺吶ｋ
	 * @param node
	 * @param headerCount 谿九☆蟄占ｦ∫ｴ�謨ｰ
	 */
	public static void removeAllChilderenWithoutHeader(Node node, int remainChildCount) {
		NodeList childNodes = node.getChildNodes();
		List<Node> removeNodeList = new ArrayList<Node>();
		for (int i = remainChildCount; i < childNodes.getLength(); i++) {
			removeNodeList.add(childNodes.item(i));
		}

		for(Node childNode : removeNodeList) {
			node.removeChild(childNode);
		}

	}
}
