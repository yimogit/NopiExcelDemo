# NopiExcelDemo
使用Nopi导入Excel的例子

> NOPI版本：2.3.0,依赖于NPOI的SharpZipLib版本：0.86,经测试适用于.net4.0+

## 遇到的几个问题

1. NOPI中的`IWorkbook`接口：xls使用`HSSFWorkbook`类实现，xlsx使用`XSSFWorkbook`类实现           
2. 日期转换，判断`row.GetCell(j).CellType == NPOI.SS.UserModel.CellType.Numeric && HSSFDateUtil.IsCellDateFormatted(row.GetCell(j)`        
不能直接使用`row.GetCell(j).DateCellValue`,这玩意会直接抛出异常来~ 
## 功能
1. 将文件流转换为DataTable 
2. 文件上传导入
3. 本地路径读取导入