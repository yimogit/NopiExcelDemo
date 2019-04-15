using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace YimoFramework.ExcelImport
{
    /// <summary>
    /// DataRow的包装类
    /// </summary>
    public class DataRowWrapper : IExcelDataRow, IEnumerable<String>, IDisposable
    {
        private Int32 m_rowIndex;
        private DataRow m_originalData;

        #region 构造器
        /// <summary>
        /// 
        /// </summary>
        /// <param name="originalData"></param>
        public DataRowWrapper(DataRow originalData, Int32 rowIndex)
        {
            if (null == originalData)
            {
                throw new ArgumentNullException("originalData");
            }
            this.m_originalData = originalData;
            this.m_rowIndex = rowIndex;
        }
        #endregion

        #region IExcelDataRow 成员
        /// <summary>
        /// 原始数据行
        /// </summary>
        public DataRow OriginalDataRow
        {
            get { return m_originalData; }
        }

        public String this[Int32 columnIndex]
        {
            get
            {
                if (this.IsNull(columnIndex))
                {
                    return null;
                }
                Object value = m_originalData[columnIndex];
                return value == null ? null : value.ToString();
            }
        }

        public String this[String columnName]
        {
            get
            {
                if (this.IsNull(columnName))
                {
                    return null;
                }
                Object value = m_originalData[columnName];
                return value == null ? null : value.ToString();
            }
        }

        /// <summary>
        /// 获取一个值，该值指示位于指定索引处的列是否包含空值。
        /// </summary>
        /// <param name="columnIndex">列的从零开始的索引。</param>
        /// <returns> 如果列包含空值，则为 true；否则为 false</returns>
        public Boolean IsNull(Int32 columnIndex)
        {

            return m_originalData.IsNull(columnIndex);
        }

        /// <summary>
        /// 获取一个值，该值指示位于指定索引处的列是否包含空值。
        /// </summary>
        /// <param name="columnName">列的名称</param>
        /// <returns> 如果列包含空值，则为 true；否则为 false</returns>
        public Boolean IsNull(String columnName)
        {
            return m_originalData.IsNull(columnName);
        }

        /// <summary>
        /// 第几行
        /// </summary>
        public Int32 RowIndex
        {
            get { return m_rowIndex; }
        }

        /// <summary>
        /// 数据项大小
        /// </summary>
        public Int32 ItemCount
        {
            get { return m_originalData.ItemArray.Length; }
        }

        #endregion

        #region IEnumerable<String> 成员

        public IEnumerator<String> GetEnumerator()
        {
            foreach (var item in m_originalData.ItemArray)
            {
                yield return (String)item;
            }
        }

        #endregion

        #region IEnumerable 成员

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        #endregion

        #region IDisposable 成员

        public void Dispose()
        {
            m_originalData = null;
        }

        #endregion
    }
}
