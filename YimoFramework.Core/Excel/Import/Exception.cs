using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace YimoFramework.ExcelImport
{
    /// <summary>
    /// 处理单元格信息时的异常
    /// </summary>
    public class ProcessCellDataException : ExcelReaderException
    {
        public ProcessCellDataException(String message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    /// <summary>
    /// 处理Excel时的异常信息
    /// </summary>
    public class ProcessExcelException : ExcelReaderException
    {
        public ProcessExcelException(String message)
            : base(message)
        {
        }

        public ProcessExcelException(String message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    /// <summary>
    /// 工作表为空异常信息
    /// </summary>
    public class SheetTableNullException : ExcelReaderException
    {
        public SheetTableNullException(String message)
            : base(message)
        {
        }
    }

    /// <summary>
    /// 验证工作表的异常信息
    /// </summary>
    public class ValidateSheetException : ExcelReaderException
    {
        public ValidateSheetException(String message)
            : base(message)
        {
        }
    }

    /// <summary>
    /// 验证信息列的异常信息
    /// </summary>
    public class ValidateColumnException : ExcelReaderException
    {
        public ValidateColumnException(String message)
            : base(message)
        {
        }
    }

    /// <summary>
    /// Excel读取器的异常
    /// </summary>
    public abstract class ExcelReaderException : Exception
    {
        protected ExcelReaderException(String message)
            : base(message)
        {
        }

        protected ExcelReaderException(String message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
