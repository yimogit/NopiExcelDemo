using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace YimoFramework.ExcelExport
{
    /// <summary>
    /// Excel数据列
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelColumn<T>
    {
        /// <summary>
        /// 列名称
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// 链接地址
        /// </summary>
        public Func<T, Object> HrefDelegate { get; set; }

        /// <summary>
        /// 列头样式
        /// </summary>
        public Action<ICellStyle> HeaderStyle { get; set; }

        /// <summary>
        ///  数据样式
        /// </summary>
        public Action<ICellStyle> BodyStyle { get; set; }

        /// <summary>
        /// 数据格式化字符串
        /// </summary>
        public String Format { get; set; }

        /// <summary>
        /// 自定义列头样式
        /// </summary>
        public Action<ICell> CustomHeader { get; set; }

        /// <summary>
        /// 自定义单元格的数据
        /// </summary>
        public Func<T, Object> ColumnDelegate { get; set; }

        /// <summary>
        /// 自定义单元格的数据呈现
        /// </summary>
        public Action<T, ICell> CustomRenderer { get; set; }

        /// <summary>
        /// 判断是否是一个集合
        /// </summary>
        public bool IsCollection { get; set; }
    }
}
