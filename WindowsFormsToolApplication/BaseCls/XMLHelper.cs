using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Xml;
using System.Xml.Linq;

using System.Text.RegularExpressions;//正则
using System.Xml.XPath;//xpath


//xml操作工具类
namespace XMLHelper
{
    public static class XMLHelper
    {

        /// <summary>
        /// 复制节点方法，等等(参数需要修改）
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="height">单元格高度</param>
        /// <param name="width">单元格宽度</param>
        /// <returns>返回是否为合并单元格的布尔(Boolean)值</returns>
        public static bool CopyNodeByID(string nodeid, out int height, out int width)
        {
            height = 0;
            width = 0;

            return true;
        }
        #region method1 试用此方法
        /// <summary>
        /// The C# version of Mark Miller。Makes the X path. Use a format like //configuration/appSettings/add[@key='name']/@value
        /// </summary>
        /// <param name="doc">The doc.</param>
        /// <param name="xpath">The xpath.</param>
        /// <returns></returns>
        public static XNode createNodeFromXPath(XElement elem, string xpath)
        {
            // Create a new Regex object
            //用正则把分割符号 / 过滤出来
            Regex r = new Regex(@"/*([a-zA-Z0-9_\.\-]+)(\[@([a-zA-Z0-9_\.\-]+)='([^']*)'\])?|/@([a-zA-Z0-9_\.\-]+)");

            xpath = xpath.Replace("\"", "'");
            // Find matches
            Match m = r.Match(xpath);

            XNode currentNode = elem;
            StringBuilder currentPath = new StringBuilder();

            while (m.Success)
            {
                String currentXPath = m.Groups[0].Value;    // "/configuration" or "/appSettings" or "/add"
                String elementName = m.Groups[1].Value;     // "configuration" or "appSettings" or "add"
                String filterName = m.Groups[3].Value;      // "" or "key"
                String filterValue = m.Groups[4].Value;     // "" or "name"
                String attributeName = m.Groups[5].Value;   // "" or "value"

                StringBuilder builder = currentPath.Append(currentXPath);
                String relativePath = builder.ToString();
                XNode newNode = (XNode)elem.XPathSelectElement(relativePath);

                if (newNode == null)
                {
                    if (!string.IsNullOrEmpty(attributeName))
                    {
                        ((XElement)currentNode).Attribute(attributeName).Value = "";
                        newNode = (XNode)elem.XPathEvaluate(relativePath);
                    }
                    else if (!string.IsNullOrEmpty(elementName))
                    {
                        XElement newElem = new XElement(elementName);
                        if (!string.IsNullOrEmpty(filterName))
                        {
                            newElem.Add(new XAttribute(filterName, filterValue));
                        }

                        ((XElement)currentNode).Add(newElem);
                        newNode = newElem;
                    }
                    else
                    {
                        throw new FormatException("The given xPath is not supported " + relativePath);
                    }
                }

                currentNode = newNode;
                m = m.NextMatch();
            }

            // Assure that the node is found or created
            if (elem.XPathEvaluate(xpath) == null)
            {
                throw new FormatException("The given xPath cannot be created " + xpath);
            }

            return currentNode;
        }
        #endregion method1

        #region method2
        /// <summary>
        ///  For XDocument Supports attribute creation
        /// </summary>
        /// <param name="document">The doc.</param>
        /// <param name="xpath">The xpath.</param>
        /// <returns></returns>
        public static XDocument CreateElement(XDocument document, string xpath)
        {
            if (string.IsNullOrEmpty(xpath))
                throw new InvalidOperationException("Xpath must not be empty");

            var xNodes = Regex.Matches(xpath, @"\/[^\/]+").Cast<Match>().Select(it => it.Value).ToList();
            if (!xNodes.Any())
                throw new InvalidOperationException("Invalid xPath");

            var parent = document.Root;
            var currentNodeXPath = "";
            foreach (var xNode in xNodes)
            {
                currentNodeXPath += xNode;
                var nodeName = Regex.Match(xNode, @"(?<=\/)[^\[]+").Value;
                var existingNode = parent.XPathSelectElement(currentNodeXPath);
                if (existingNode != null)
                {
                    parent = existingNode;
                    continue;
                }

                var attributeNames =
                  Regex.Matches(xNode, @"(?<=@)([^=]+)\=([^]]+)")
                        .Cast<Match>()
                        .Select(it =>
                        {
                            var groups = it.Groups.Cast<Group>().ToList();
                            return new { AttributeName = groups[1].Value, AttributeValue = groups[2].Value };
                        });

                parent.Add(new XElement(nodeName, attributeNames.Select(it => new XAttribute(it.AttributeName, it.AttributeValue)).ToArray()));
                parent = parent.Descendants().Last();
            }
            return document;
        }
        #endregion method2


        #region method3
        /// <summary>
        ///  通过xpath生成xml          eg: Set(doc, "/configuration/appSettings/add[@key='Server']/@value", "foobar");
        /// </summary>
        /// <param name="document">The doc.</param>
        /// <param name="xpath">The xpath.</param>
        /// <returns></returns>
        public static void Set(XmlDocument doc, string xpath, string value)
        {
            if (doc == null) throw new ArgumentNullException("doc");
            if (string.IsNullOrEmpty(xpath)) throw new ArgumentNullException("xpath");

            XmlNodeList nodes = doc.SelectNodes(xpath);
            if (nodes.Count > 1) throw new FormatException("Xpath '" + xpath + "' was not found multiple times!");
            else if (nodes.Count == 0) createXPath(doc, xpath).InnerText = value;//创建节点并赋值
            else nodes[0].InnerText = value;//xpath路径下有节点 进行赋值
        }


        static XmlNode createXPath(XmlDocument doc, string xpath)
        {
            XmlNode node = doc;
            foreach (string part in xpath.Substring(1).Split('/'))
            {
                XmlNodeList nodes = node.SelectNodes(part);
                if (nodes.Count > 1) throw new FormatException("Xpath '" + xpath + "' was not found multiple times!");
                else if (nodes.Count == 1) { node = nodes[0]; continue; }

                if (part.StartsWith("@"))
                {
                    var anode = doc.CreateAttribute(part.Substring(1));
                    node.Attributes.Append(anode);
                    node = anode;
                }
                else
                {
                    string elName, attrib = null;
                    if (part.Contains("["))
                    {
                        part.SplitOnce("[", out elName, out attrib);
                        if (!attrib.EndsWith("]")) throw new FormatException("Unsupported XPath (missing ]): " + part);
                        attrib = attrib.Substring(0, attrib.Length - 1);
                    }
                    else elName = part;

                    XmlNode next = doc.CreateElement(elName);
                    node.AppendChild(next);
                    node = next;

                    if (attrib != null)
                    {
                        if (!attrib.StartsWith("@")) throw new FormatException("Unsupported XPath attrib (missing @): " + part);
                        string name, value;
                        attrib.Substring(1).SplitOnce("='", out name, out value);
                        if (string.IsNullOrEmpty(value) || !value.EndsWith("'")) throw new FormatException("Unsupported XPath attrib: " + part);
                        value = value.Substring(0, value.Length - 1);
                        var anode = doc.CreateAttribute(name);
                        anode.Value = value;
                        node.Attributes.Append(anode);
                    }
                }
            }
            return node;
        }


        public static void SplitOnce(this string value, string separator, out string part1, out string part2)
        {
            if (value != null)
            {
                int idx = value.IndexOf(separator);
                if (idx >= 0)
                {
                    part1 = value.Substring(0, idx);
                    part2 = value.Substring(idx + separator.Length);
                }
                else
                {
                    part1 = value;
                    part2 = null;
                }
            }
            else
            {
                part1 = "";
                part2 = null;
            }
        }
        #endregion method3


    }
}
