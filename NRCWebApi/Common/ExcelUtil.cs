using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NRCWebApi.Dto;
using System.Data;
using System.IO;
using System.Net;
using System.Reflection;

namespace NRCWebApi.Common
{
    /// <summary>
    /// Excel操作
    /// </summary>
    public class ExcelUtil
    {
        #region 导出

        /// <summary>
        /// 自动生成xlsx文件并导出xlsx文件流
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <param name="entities">数据实体</param>
        /// <param name="dicColumns">列对应关系,如Name->姓名</param>
        /// <param name="title">标题</param>
        /// <returns></returns>
        public static byte[] ExportExcel<T>(List<T> entities, List<NPOIRequestDto> dicColumns, string title = null)
        {
            if (entities.Count <= 0)
            {
                return null;
            }
            //HSSFWorkbook => xls
            //XSSFWorkbook => xlsx
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("data");//名称自定义

            return ExportExcelHelp<T>(workbook, sheet, entities, dicColumns, title);
        }


        /// <summary>
        /// 模板xlsx文件并导出xlsx文件流
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <param name="filePath">模板路径</param>
        /// <param name="entities">数据实体</param>
        /// <param name="dicColumns">列对应关系和顺序,如Name->姓名</param>
        /// <param name="title">标题</param>
        /// <returns></returns>
        public static byte[] ExportExcel<T>(string filePath, List<T> entities, List<string> dicColumns)
        {
            if (entities.Count <= 0)
            {
                return null;
            }
            //HSSFWorkbook => xls
            //XSSFWorkbook => xlsx

            FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = new XSSFWorkbook(file);
            ISheet sheet = workbook.GetSheetAt(0);

            // IRow? cellsColumn = null;
            IRow? cellsData = null;
            //获取实体属性名
            PropertyInfo[] properties = entities[0].GetType().GetProperties();
            int cellsIndex = 0;

            //列名的顺序
            // cellsColumn = sheet.CreateRow(cellsIndex);
            int index = 0;
            Dictionary<string, int> columns = new Dictionary<string, int>();
            foreach (var item in dicColumns)
            {
                // cellsColumn.CreateCell(index).SetCellValue(item);
                columns.Add(item, index);
                index++;
            }


            cellsIndex += 1;


            //边框数据
            ICellStyle style = workbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;

            //实体数据
            foreach (var item in entities)
            {
                cellsData = sheet.CreateRow(cellsIndex);
                for (int i = 0; i < properties.Length; i++)
                {
                    if (!dicColumns.Exists(m => m == properties[i].Name)) continue;
                    //这里可以也根据数据类型做不同的赋值，也可以根据不同的格式参考上面的ICellStyle设置不同的样式
                    object[] entityValues = new object[properties.Length];
                    entityValues[i] = properties[i].GetValue(item);
                    //获取对应列下标
                    var dataTemp = dicColumns.FirstOrDefault(m => m == properties[i].Name);
                    index = columns[dataTemp];

                    ICell cell = cellsData.CreateCell(index);
                    cell.SetCellValue(entityValues[i].ToString());

                    //添加边框样式
                    cell.CellStyle = style;

                }
                cellsIndex++;
            }

            byte[]? buffer = null;
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                buffer = ms.ToArray();
            }

            return buffer;
        }

        private static byte[] ExportExcelHelp<T>(IWorkbook workbook, ISheet sheet, List<T> entities, List<NPOIRequestDto> dicColumns, string title = null)
        {

            IRow? cellsColumn = null;
            IRow? cellsData = null;
            //获取实体属性名
            PropertyInfo[] properties = entities[0].GetType().GetProperties();
            int cellsIndex = 0;
            //标题
            if (!string.IsNullOrEmpty(title))
            {
                ICellStyle style = workbook.CreateCellStyle();
                //边框  
                style.BorderBottom = BorderStyle.Dotted;
                style.BorderLeft = BorderStyle.Hair;
                style.BorderRight = BorderStyle.Hair;
                style.BorderTop = BorderStyle.Dotted;
                //水平对齐  
                //style.Alignment = HorizontalAlignment.Center;

                //垂直对齐  
                //style.VerticalAlignment = VerticalAlignment.Center;

                //设置字体
                IFont font = workbook.CreateFont();
                font.FontHeightInPoints = 10;
                font.FontName = "微软雅黑";
                style.SetFont(font);



                IRow cellsTitle = sheet.CreateRow(0);
                //设置值
                cellsTitle.CreateCell(0).SetCellValue(title);
                //设置样式
                cellsTitle.RowStyle = style;
                //设置行高
                cellsTitle.HeightInPoints = 30;
                //合并单元格
                sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 0, dicColumns.Count - 1));

                //设置合并后style
                var cell = sheet.GetRow(0).GetCell(0);
                //设置style
                ICellStyle cellstyle = workbook.CreateCellStyle();
                cellstyle.VerticalAlignment = VerticalAlignment.Center; //水平对齐
                cellstyle.Alignment = HorizontalAlignment.Center; //垂直对齐
                cell.CellStyle = cellstyle;

                cellsIndex = 1;
            }
            //列名
            cellsColumn = sheet.CreateRow(cellsIndex);
            int index = 0;
            Dictionary<string, int> columns = new Dictionary<string, int>();
            foreach (var item in dicColumns)
            {
                cellsColumn.CreateCell(index).SetCellValue(item.TitleName);
                columns.Add(item.TitleName, index);

                sheet.SetColumnWidth(index, 256 * item.Width);

                index++;


            }
            cellsIndex += 1;
            //数据
            foreach (var item in entities)
            {
                cellsData = sheet.CreateRow(cellsIndex);
                for (int i = 0; i < properties.Length; i++)
                {
                    if (!dicColumns.Exists(m => m.DataName == properties[i].Name)) continue;
                    //这里可以也根据数据类型做不同的赋值，也可以根据不同的格式参考上面的ICellStyle设置不同的样式
                    object[] entityValues = new object[properties.Length];
                    entityValues[i] = properties[i].GetValue(item);
                    //获取对应列下标
                    var dataTemp = dicColumns.FirstOrDefault(m => m.DataName == properties[i].Name);
                    index = columns[dataTemp.TitleName];
                    cellsData.CreateCell(index).SetCellValue(entityValues[i].ToString());
                }
                cellsIndex++;
            }
;

            byte[]? buffer = null;
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                buffer = ms.ToArray();
            }

            return buffer;

        }

        #endregion


        #region 导入


        /// <summary>
        /// excel导入成DataTable
        /// </summary>
        /// <param name="ExcelFileStream">文件流</param>
        /// <param name="SheetIndex">开始的sheet</param>
        /// <param name="HeaderRowIndex">从第行开始读取</param>
        /// <param name="fileExt">文件名，带.</param>
        /// <returns>返回DataTable</returns>
        public static DataTable ExcelToTable(Stream ExcelFileStream, int SheetIndex, int HeaderRowIndex, string fileExt)
        {
            DataTable dt = new DataTable();
            IWorkbook workbook;
            //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
            if (fileExt.ToLower() == ".xlsx")
            {
                workbook = new XSSFWorkbook(ExcelFileStream);
            }
            else if (fileExt.ToLower() == ".xls")
            {
                workbook = new HSSFWorkbook(ExcelFileStream);
            }
            else
            {
                throw new Exception("文件格式不正确");
            }

            ISheet sheet = workbook.GetSheetAt(SheetIndex);

            //表头  
            IRow header = sheet.GetRow(HeaderRowIndex);
            List<int> columns = new List<int>();
            //填充列明
            for (int i = 0; i < header.LastCellNum; i++)
            {
                object obj = GetValueType(header.GetCell(i));
                if (obj == null || obj.ToString() == string.Empty)
                {
                    dt.Columns.Add(new DataColumn("Columns" + i.ToString()));
                }
                else
                {
                    dt.Columns.Add(new DataColumn(obj.ToString()));
                }

                columns.Add(i);
            }
            //数据  
            for (int i = HeaderRowIndex + 1; i <= sheet.LastRowNum; i++)
            {
                DataRow dr = dt.NewRow();
                bool hasValue = false;
                foreach (int j in columns)
                {
                    dr[j] = GetValueType(sheet.GetRow(i).GetCell(j));
                    if (dr[j] != null && dr[j].ToString() != string.Empty)
                    {
                        hasValue = true;
                    }
                }
                if (hasValue)
                {
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        /// <summary>
        /// 获取单元格类型
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static object GetValueType(ICell cell)
        {
            if (cell == null)
            {
                return null;
            }

            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:  
                    return null;
                case CellType.Boolean: //BOOLEAN:  
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:  
                    return cell.NumericCellValue;
                case CellType.String: //STRING:  
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:  
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:  
                default:
                    return "=" + cell.CellFormula;
            }
        }


        /// <summary>
        /// 将excel导入到list
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="ExcelFileStream">Stream 文件流</param>
        /// <param name="dict">转换的Dictionary：例如 ("昵称","nickname")xlxs中的名称：字段的名称</param>
        /// <returns></returns>
        public static List<T> ExcelToList<T>(Stream ExcelFileStream, int SheetIndex, int HeaderRowIndex, string fileExt, Dictionary<string, string> dict) where T : class, new()
        {
            List<T> ts = new List<T>();
            IWorkbook workbook = null;
            T t = new T();

            //XSSFWorkbook 适用XLSX格式，HSSFWorkbook 适用XLS格式
            if (fileExt.ToLower() == ".xlsx")
            {
                workbook = new XSSFWorkbook(ExcelFileStream);
            }
            else if (fileExt.ToLower() == ".xls")
            {
                workbook = new HSSFWorkbook(ExcelFileStream);
            }
            else
            {
                throw new Exception("文件格式不正确");
            }

            ISheet sheet = workbook.GetSheetAt(SheetIndex);

            List<string> listName = new List<string>();
            try
            {   // 获得此模型的公共属性
                var propertys = t.GetType().GetProperties().ToList();

                if (workbook != null)
                {
                    sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet

                    if (sheet != null)
                    {
                        int rowCount = sheet.LastRowNum;//总行数
                        if (rowCount > 0)
                        {
                            IRow firstRow = sheet.GetRow(HeaderRowIndex);//第一行
                            int cellCount = firstRow.LastCellNum;//列数
                            //循环列数
                            for (int i = 0; i < cellCount; i++)
                            {
                                //循环需要转换的值
                                foreach (var item in dict)
                                {
                                    if (item.Key.Equals(firstRow.GetCell(i).StringCellValue))
                                    {
                                        //替换表头
                                        firstRow.GetCell(i).SetCellValue(item.Value);
                                    }
                                }
                                //获取已经替换的表头
                                var s = firstRow.GetCell(i).StringCellValue;
                                //添加到listname
                                listName.Add(s);
                            }

                            //循环行
                            for (int i = HeaderRowIndex + 1; i <= rowCount; i++)
                            {
                                t = new T();
                                IRow currRow = sheet.GetRow(i);//第i行
                                //循环列
                                for (int k = 0; k < cellCount; k++)
                                {   //取值
                                    object value = null;
                                    if (currRow.GetCell(k) != null)
                                    {
                                        // firstRow.GetCell(0).SetCellType(CellType.String);
                                        currRow.GetCell(k).SetCellType(CellType.String);
                                        value = currRow.GetCell(k).StringCellValue;
                                    }
                                    else
                                    {
                                        continue;
                                    }

                                    var Name = string.Empty;
                                    //获取第表头的值
                                    Name = listName[k];
                                    //循环属性
                                    foreach (var pi in propertys)
                                    {
                                        if (pi.Name.Equals(Name))
                                        {
                                            //获取属性类型名称
                                            var s = pi.PropertyType.Name;

                                            //如果非空，则赋给对象的属性
                                            if (value != DBNull.Value)
                                            {
                                                //判断属性的类型(可以自行添加)
                                                switch (s)
                                                {
                                                    case "Guid":
                                                        pi.SetValue(t, new Guid(value.ToString()), null);
                                                        break;

                                                    case "Int32":
                                                        pi.SetValue(t, value.ToString() == "" ? 0 : Convert.ToInt32(value.ToString()), null);
                                                        break;

                                                    case "Decimal":
                                                        pi.SetValue(t, value.ToString() == "" ? 0 : Convert.ToDecimal(value.ToString()), null);
                                                        break;

                                                    case "DateTime":
                                                        pi.SetValue(t, Convert.ToDateTime(value.ToString()), null);
                                                        break;

                                                    case "Double":
                                                        pi.SetValue(t, value.ToString() == "" ? 0 : Convert.ToDouble(value.ToString()), null);
                                                        break;

                                                    case "String":
                                                        pi.SetValue(t, value, null);
                                                        break;

                                                    default:
                                                        break;
                                                }
                                            }
                                        }
                                    }
                                }
                                //对象添加到泛型集合中
                                ts.Add(t);
                            }
                        }
                    }
                }
                return ts;
            }
            catch (Exception ex)
            {
                if (ExcelFileStream != null)
                {
                    ExcelFileStream.Close();
                }
                return null;
            }
        }

        #endregion
    }
}
