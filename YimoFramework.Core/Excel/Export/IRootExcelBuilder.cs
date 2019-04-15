using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace YimoFramework.ExcelExport
{
    /// <summary>
    /// Excel 构建器，用于生成多个工作表
    /// </summary>
    public interface IRootExcelBuilder
    {
        /// <summary>
        /// 获取或设置 文档属性 委托
        /// </summary>
        Action<DocumentSummaryInformation> DocumentProperty { get; set; }

        /// <summary>
        /// 获取或设置 文档摘要属性 委托
        /// </summary>
        Action<SummaryInformation> SummaryProperty { get; set; }

        /// <summary>
        /// 获取工作薄
        /// </summary>
        HSSFWorkbook Workbook { get;  }

        /// <summary>
        /// 获取当前工作表
        /// </summary>
        ISheet CurrentSheet { get; }

        /// <summary>
        /// 生成Excel的工作表
        /// </summary>
        /// <param name="sheetName">工作表名称，不能重复</param>
        /// <param name="dataSource">数据源</param>
        /// <param name="columns">工作表中的数据列构造器</param>
        /// <returns></returns>
        IRootExcelBuilder Sheet<T>(String sheetName, IEnumerable<T> dataSource, Action<IRootExcelColumnBuilder<T>> columns)
            where T : class;
    }
}
