using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Xml;
using System.Xml.Linq;

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
    }
}
