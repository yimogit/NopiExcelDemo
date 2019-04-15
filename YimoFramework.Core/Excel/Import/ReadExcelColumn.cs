using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YimoFramework.ExcelImport
{
    /// <summary>
    /// 读取Excel工作表数据列
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ReadExcelColumn<T> where T : class
    {
        /// <summary>
        /// 列索引，从0开始
        /// </summary>
        public Int32 ColumnIndex { get; set; }

        /// <summary>
        /// 列名称
        /// </summary>
        public String ColumnName { get; set; }

        /// <summary>
        /// 自定义读取工作表中一行数据的表达式
        /// </summary>
        public Action<T, IExcelDataRow> CustomDelegate { get; set; }

        /// <summary>
        /// 自定义读取单元格数据的表达式
        /// </summary>
        public Action<T, String> CustomEvaluater { get; set; }

        public ReadExcelColumn()
        {
            this.ColumnIndex = -1;
        }
    }
}
