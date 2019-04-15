using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;

namespace YimoFramework.ExcelImport
{
    /// <summary>
    /// 读取Excel数据列构造器
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ReadExcelColumnBuilder<T> : IReadExcelColumnBuilder<T>, IRootReadExcelColumnBuilder<T>, IEnumerable<ReadExcelColumn<T>>
       where T : class
    {
        //读取Excel工作表的数据列集合
        private readonly List<ReadExcelColumn<T>> columns = new List<ReadExcelColumn<T>>();
        //当前工作表读取数据列
        private ReadExcelColumn<T> currentColumn;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ReadExcelColumn<T> this[int index]
        {
            get
            {
                return columns[index];
            }
        }

        #region 构造函数
        public ReadExcelColumnBuilder()
        {
        }
        #endregion

        #region IEnumerable<ReadExcelColumn<T>> 成员

        public IEnumerator<ReadExcelColumn<T>> GetEnumerator()
        {
            return this.columns.GetEnumerator();
        }

        #endregion

        #region IEnumerable 成员

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.columns.GetEnumerator();
        }

        #endregion

        #region IRootReadExcelColumnBuilder<T> 成员
        /// <summary>
        /// 生成工作表的数据列读取器
        /// </summary>
        /// <param name="expression">读取单元格数据的表达式</param>
        /// <param name="index">按数据列索引读取</param>
        /// <returns></returns>
        public IReadExcelColumnBuilder<T> For(Action<T, string> expression, int index)
        {
            currentColumn = new ReadExcelColumn<T>()
            {
                ColumnIndex = index,
                CustomEvaluater = expression,
            };
            this.columns.Add(currentColumn);
            return this;
        }

        /// <summary>
        /// 生成工作表的数据列读取器
        /// </summary>
        /// <param name="expression">读取单元格数据的表达式</param>
        /// <param name="name">列标题名称</param>
        /// <returns></returns>
        public IReadExcelColumnBuilder<T> For(Action<T, string> expression, string name)
        {
            currentColumn = new ReadExcelColumn<T>()
            {
                ColumnName = name,
                CustomEvaluater = expression,
            };
            this.columns.Add(currentColumn);
            return this;
        }

        /// <summary>
        /// 生成工作表的数据列读取器
        /// </summary>
        /// <param name="expression">读取工作表一行数据的表达式</param>
        /// <returns></returns>
        public IReadExcelColumnBuilder<T> For(Action<T, IExcelDataRow> expression)
        {
            currentColumn = new ReadExcelColumn<T>()
            {
                CustomDelegate = expression
            };
            this.columns.Add(currentColumn);
            return this;
        }

        #endregion
    }
}
