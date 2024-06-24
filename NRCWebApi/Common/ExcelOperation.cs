using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;

namespace NRCWebApi.Common
{
    /// <summary>
    /// Excel操作类
    /// </summary>
    public static class ExcelOperation
    {
        #region//Public-Method
        /// <summary>
        /// 获取excel列名(列名为空也会返回一个空列名)
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="sheetIndex">表下标</param>
        /// <param name="fieldRowIndex">从第几行开始</param>
        /// <returns></returns>
        public static List<string> GetCNames(string path, int sheetIndex = 0, int fieldRowIndex = 0)
        {
            //列名
            List<string> cNames = new List<string>();
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    //
                    ISheet sheet = null;
                    //获取工作表,默认获取第一张
                    if (workbook.NumberOfSheets > sheetIndex && workbook.NumberOfSheets > 0)
                    {
                        //获取excel表格
                        sheet = workbook.GetSheetAt(sheetIndex);
                    }
                    else
                    {
                         new Exception("索引超出界限或该文件表格为空!");
                       
                    }
                    //是否已经获取工作表，如果没有则直接返回
                    if (sheet == null)
                    {
                        return null;
                    }
                    #region//获取表头,文本类型
                    //获取列名(默认获取第一行)
                    if (fieldRowIndex > sheet.LastRowNum)//判断行标是否超界
                    {
                        return null;
                    }
                    IRow headRow = sheet.GetRow(fieldRowIndex);
                    //列数
                    int cellcount = headRow.LastCellNum;
                    for (int i = 0; i < cellcount; i++)
                    {
                        //获取行中的i个元素
                        ICell cell = headRow.GetCell(i);
                        //将列名添加至数组
                        cNames.Add(cell.ToString());
                    }
                    #endregion
                }
                return cNames;
            }
            catch
            {
                return null;
            }

        }
        /// <summary>
        /// 获取excel列名
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="sheetName">表名</param>
        /// <param name="fieldRowIndex">从第几行开始</param>
        /// <returns></returns>
        public static List<string> GetCNames(string path, string sheetName, int fieldRowIndex = 0)
        {
            //列名
            List<string> cNames = new List<string>();
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    //
                    ISheet sheet = null;
                    //获取工作表
                    if (string.IsNullOrWhiteSpace(sheetName))
                    {
                        //sheet = workbook.GetSheetAt(0);
                        //System.Windows.Forms.MessageBox.Show(sheetName + "表不存在!");
                    }
                    else
                    {
                        //获取excel表格
                        sheet = workbook.GetSheet(sheetName);
                    }
                    //是否已经获取工作表，如果没有则直接返回
                    if (sheet == null)
                    {
                        return null;
                    }


                    #region//获取表头,文本类型
                    //获取列名所在行(默认获取第一行)
                    if (fieldRowIndex > sheet.LastRowNum)//判断行标是否超界
                    {
                        return null;
                    }
                    IRow headRow = sheet.GetRow(fieldRowIndex);
                    //列数
                    int cellcount = headRow.LastCellNum;
                    for (int i = 0; i < cellcount; i++)
                    {
                        //获取行中的i个元素
                        ICell cell = headRow.GetCell(i);
                        cNames.Add(cell.ToString());
                    }
                    #endregion
                }
                return cNames;
            }
            catch
            {
                return null;
            }



        }

        /// <summary>
        /// 从Excel中获取Datatable
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="sheetIndex">表下标</param>
        /// <param name="fieldRowIndex">从第几行开始</param>
        /// <returns></returns>
        public static DataTable GetDataTable(string path, int sheetIndex = 0, int fieldRowIndex = 0)
        {
            //保存获取表格的数据
            DataTable dt = new DataTable();
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    //
                    ISheet sheet = null;
                    //获取工作表,默认获取第一张
                    if (workbook.NumberOfSheets > sheetIndex && workbook.NumberOfSheets > 0)
                    {
                        //获取工作表
                        sheet = workbook.GetSheetAt(sheetIndex);
                    }
                    else
                    {
                       // System.Windows.Forms.MessageBox.Show("索引超出界限或该文件表格为空！");
                    }
                    //是否已经获取工作表，如果没有则直接返回
                    if (sheet == null)
                    {
                        return null;
                    }

                    #region//获取表头,文本类型
                    //默认获取第一行
                    if (fieldRowIndex > sheet.LastRowNum)//判断行标是否超界
                    {
                       // System.Windows.Forms.MessageBox.Show("行索引超出界限！");
                        return null;
                    }
                    //获取表头行
                    IRow headRow = sheet.GetRow(fieldRowIndex);
                    //列数
                    int cellcount = headRow.LastCellNum;
                    for (int i = 0; i < cellcount; i++)
                    {
                        DataColumn dc = new DataColumn();
                        //获取行中的i个元素
                        ICell cell = headRow.GetCell(i);
                        if (cell != null)
                        {
                            dc.ColumnName = cell.ToString();
                            //获取cell的数据类型
                            dc.DataType = Type.GetType("System.String");
                        }
                        else
                        {
                            dt.Columns.Add("");
                            dc.DataType = Type.GetType("System.String");
                        }
                        //添加列
                        if (!dt.Columns.Contains(dc.ColumnName))//判断列是否已经存在
                        {
                            dt.Columns.Add(dc);
                        }
                        else
                        {

                        }

                    }
                    #endregion

                    #region//获取表格内容
                    for (int i = (fieldRowIndex + 1); i <= sheet.LastRowNum; i++)
                    {
                        //获取工作表的一行
                        IRow row = sheet.GetRow(i);
                        DataRow dataRow = dt.NewRow();
                        for (int j = row.FirstCellNum; j < cellcount; j++)
                        {
                            if (row.GetCell(j) != null)
                            {
                                //判断单元格是否为日期格式(如果Excel有重复行会导致DataRow索引出界)
                                if (row.GetCell(j).CellType == NPOI.SS.UserModel.CellType.Numeric && HSSFDateUtil.IsCellDateFormatted(row.GetCell(j)))
                                {
                                    if (row.GetCell(j).DateCellValue.Year > 1000)
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue.ToString();
                                    }
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j).ToString();
                                    }
                                }
                                else
                                {
                                    dataRow[j] = row.GetCell(j).ToString();
                                }
                            }
                        }
                        dt.Rows.Add(dataRow);
                    }
                    #endregion

                    return dt;
                }
            }
            catch
            {
                return null;
            }
        }
        /// <summary>
        /// 获取Excel中的所有表格
        /// </summary>
        /// <param name="path"></param>
        /// <param name="fieldRowIndex"></param>
        /// <returns></returns>
        public static List<DataTable> GetDataTables(string path, int fieldRowIndex = 0)
        {
            List<DataTable> tables = new List<DataTable>();
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    for (int n = 0; n < workbook.NumberOfSheets; n++)
                    {
                        //
                        ISheet sheet = workbook.GetSheetAt(n);//获取表格
                        string sheetName = sheet.SheetName;//表名
                        DataTable dt = new DataTable();
                        dt.TableName = sheetName;
                        #region//获取表头,文本类型
                        //默认获取第一行
                        if (fieldRowIndex > sheet.LastRowNum)//判断行标是否超界
                        {
                            //System.Windows.Forms.MessageBox.Show("行索引超出界限！");
                            return null;
                        }
                        //获取表头行
                        IRow headRow = sheet.GetRow(fieldRowIndex);
                        //列数
                        int cellcount = headRow.LastCellNum;
                        for (int i = 0; i < cellcount; i++)
                        {
                            DataColumn dc = new DataColumn();
                            //获取行中的i个元素
                            ICell cell = headRow.GetCell(i);
                            if (cell != null)
                            {
                                dc.ColumnName = cell.ToString();
                                //获取cell的数据类型
                                dc.DataType = Type.GetType("System.String");
                            }
                            else
                            {
                                dt.Columns.Add("");
                                dc.DataType = Type.GetType("System.String");
                            }
                            //添加列
                            if (!dt.Columns.Contains(dc.ColumnName))//判断列是否已经存在
                            {
                                dt.Columns.Add(dc);
                            }
                            else
                            {
                            }
                        }
                        #endregion
                        bool isGoOn = false;
                        for (int i = (fieldRowIndex + 1); i <= sheet.LastRowNum; i++)
                        {
                            //获取工作表的一行
                            IRow row = sheet.GetRow(i);
                            for (int c = row.FirstCellNum; c < row.Cells.Count; c++)
                            {
                                if (row.GetCell(c).ToString().Trim() != "")
                                {
                                    isGoOn = true;
                                    break;
                                }
                            }
                            //判断是否继续获取
                            if (isGoOn is false)
                            {
                                tables.Add(dt);
                                break;
                            }

                            DataRow dataRow = dt.NewRow();
                            for (int j = row.FirstCellNum; j < cellcount; j++)
                            {
                                if (row.GetCell(j) != null)
                                {
                                    //判断单元格是否为日期格式(如果Excel有重复行会导致DataRow索引出界)
                                    if (row.GetCell(j).CellType == NPOI.SS.UserModel.CellType.Numeric && HSSFDateUtil.IsCellDateFormatted(row.GetCell(j)))
                                    {
                                        if (row.GetCell(j).DateCellValue.Year > 1000)
                                        {
                                            dataRow[j] = row.GetCell(j).DateCellValue.ToString();
                                        }
                                        else
                                        {
                                            dataRow[j] = row.GetCell(j).ToString();
                                        }
                                    }
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j).ToString();
                                    }
                                }
                            }
                            dt.Rows.Add(dataRow);
                        }
                        if (isGoOn is true)
                        {
                            tables.Add(dt);
                        }

                    }
                    return tables;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// 从Excel中获取Datatable
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="sheetName">表名</param>
        /// <param name="fieldRowIndex">从第几行开始</param>
        /// <returns></returns>
        public static DataTable GetDataTable(string path, string sheetName, int fieldRowIndex = 0)
        {
            //新建表格保存excel数据
            DataTable dt = new DataTable();
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    //
                    ISheet sheet = null;
                    //获取工作表
                    if (string.IsNullOrWhiteSpace(sheetName))
                    {
                        //sheet = workbook.GetSheetAt(0);
                       // System.Windows.Forms.MessageBox.Show(sheetName + "表不存在!");
                    }
                    else
                    {
                        sheet = workbook.GetSheet(sheetName);
                    }
                    //是否已经获取工作表，如果没有则直接返回
                    if (sheet == null)
                    {
                        return null;
                    }


                    #region//获取表头,文本类型
                    //默认获取第一行
                    if (fieldRowIndex > sheet.LastRowNum)//判断行标是否超界
                    {
                       // System.Windows.Forms.MessageBox.Show("行索引超出界限！");
                        return null;
                    }
                    IRow headRow = sheet.GetRow(fieldRowIndex);
                    //列数
                    int cellcount = headRow.LastCellNum;
                    for (int i = 0; i < cellcount; i++)
                    {
                        DataColumn dc = new DataColumn();
                        //获取行中的i个元素
                        ICell cell = headRow.GetCell(i);
                        if (cell != null)
                        {
                            dc.ColumnName = cell.ToString();
                            //获取cell的数据类型
                            //dc.DataType = Type.GetType("System.String");
                            dc.DataType = GetType(cell.CellType);
                        }
                        else
                        {
                            dt.Columns.Add("");
                            dc.DataType = GetType(cell.CellType);
                        }
                        dt.Columns.Add(dc);
                    }
                    #endregion

                    #region//获取表格内容
                    for (int i = (fieldRowIndex + 1); i <= sheet.LastRowNum; i++)
                    {
                        //获取工作表的某一行
                        IRow row = sheet.GetRow(i);
                        DataRow dataRow = dt.NewRow();
                        for (int j = row.FirstCellNum; j < cellcount; j++)
                        {
                            if (row.GetCell(j) != null)
                            {
                                //判断单元格是否为日期格式
                                if (row.GetCell(j).CellType == NPOI.SS.UserModel.CellType.Numeric && HSSFDateUtil.IsCellDateFormatted(row.GetCell(j)))
                                {
                                    if (row.GetCell(j).DateCellValue.Year > 1000)
                                    {
                                        dataRow[j] = row.GetCell(j).DateCellValue.ToString();
                                    }
                                    else
                                    {
                                        dataRow[j] = row.GetCell(j).ToString();
                                    }
                                }
                                else
                                {
                                    dataRow[j] = row.GetCell(j).ToString();
                                }
                            }
                        }
                        //添加列
                        dt.Rows.Add(dataRow);
                    }
                    #endregion

                    return dt;
                }
            }
            catch
            {
                return null;
            }
        }
        /// <summary>
        /// 更新数据
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="sheetIndex">表下标</param>
        /// <param name="fieldRowIndex">表头从第几行开始</param>
        /// <param name="rowIndex">第几条数据</param>
        /// <param name="values">Key(列名)-值</param>
        /// <returns></returns>
        public static DataTable UpDataExcel(string path, Dictionary<string, object> values, int sheetIndex = 0, int fieldRowIndex = 0, int rowIndex = 0)
        {
            //保存更新后的表
            DataTable dt = new DataTable();
            try
            {
                //获取excel表
                dt = GetDataTable(path, sheetIndex, fieldRowIndex);
                //判断表是否为空,且rowIndex不能超出索引
                if (dt != null && dt.Rows.Count > 0 && rowIndex < dt.Rows.Count)
                {
                    //遍历字典
                    foreach (KeyValuePair<string, object> keyValue in values)
                    {
                        if (dt.Columns.Contains(keyValue.Key))
                        {
                            dt.Rows[rowIndex][keyValue.Key] = keyValue.Value;
                        }
                        else
                        {
                            continue;

                            //提醒窗口，字段不存在的提醒是否终止该操作(但是已经修改过的数据还在)
                            //System.Windows.Forms.DialogResult dr = System.Windows.Forms.MessageBox.Show("不包含列:" + keyValue.Key, "是否跳过继续", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                            //if (dr == System.Windows.Forms.DialogResult.Yes)
                            //{
                            //    //点确定的代码
                            //    continue;
                            //}
                            //else
                            //{
                            //    //点取消的代码
                            //    break;
                            //}
                        }
                    }
                }
                else
                {
                   // System.Windows.Forms.MessageBox.Show("表为空表或者目标索引出界！\n数据更新失败！");
                }
                return dt;
            }
            catch
            {
                return null;
            }

        }

        /// <summary>
        /// 将表格存储为Excel
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="dt">表</param>
        /// <param name="sheetName">sheet表名</param>
        /// <param name="hasFieldRow">是否转表名</param>
        /// <returns></returns>
        public static bool SaveAsExcel(string path, DataTable dt, string sheetName, bool hasFieldRow = true)
        {
            try
            {
                //获取excel的文件类型
                ExcelType type = GetExcelFileType(path).Value;
                IWorkbook workbook;
                //新建工作目录
                switch (type)
                {
                    case ExcelType.xlsx:
                        workbook = new XSSFWorkbook();
                        break;
                    default:
                        workbook = new HSSFWorkbook();
                        break;
                }
                ISheet sheet;
                //表名
                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    //创建工作表
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    int i = 1;
                    while (true)
                    {
                        //判断表名是否存在
                        if (workbook.GetSheetIndex("sheet" + i) == -1)
                        {
                            //创建表
                            sheet = workbook.CreateSheet("sheet" + i);
                            break;
                        }
                        else
                        {
                            i++;
                        }
                    }
                }
                //创建表是否成功
                if (sheet == null)
                {
                    return false;
                }
                #region//添加表头
                if (hasFieldRow)//判断是否添加表头
                {
                    IRow headRow = sheet.CreateRow(0);
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        //获取单元格
                        ICell cell = headRow.CreateCell(i);
                        //设置单元格数据类型
                        cell.SetCellType(CellType.String);
                        //单元格赋值
                        cell.SetCellValue(dt.Columns[i].ColumnName);
                    }
                }
                #endregion

                #region//添加数据
                //添加表头为1，不添加为0
                int starRow;
                if (hasFieldRow)
                {
                    starRow = 1;
                }
                else
                {
                    starRow = 0;
                }
                //遍历行
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //获取第i行
                    DataRow dr = dt.Rows[i];
                    //创建行
                    IRow cells = sheet.CreateRow(i + starRow);
                    //遍历列
                    for (int j = 0; j < dr.ItemArray.Length; j++)
                    {
                        //创建单元格
                        ICell cell = cells.CreateCell(j);
                        //设置单元格数据类型
                        cell.SetCellType(CellType.String);
                        //设置单元格值
                        cell.SetCellValue(dr.ItemArray[j].ToString());
                    }
                }
                #endregion
                //保存至Path
                bool success = Export(workbook, path);
                return success;
            }
            catch
            {
                return false;
            }

        }
        /// <summary>
        /// 将表格存储为Excel
        /// </summary>
        /// <param name="path">保存路径</param>
        /// <param name="dt">DataTable</param>
        /// <param name="tagRowIndex">标红的行号</param>
        /// <param name="sheetName"></param>
        /// <param name="hasFieldRow"></param>
        /// <returns></returns>
        public static bool SaveAsExcelWithTag(string path, DataTable dt, List<int> tagRowIndex, string sheetName, bool hasFieldRow = true)
        {
            try
            {
                //获取excel的文件类型
                ExcelType type = GetExcelFileType(path).Value;
                IWorkbook workbook;
                //新建工作目录
                switch (type)
                {
                    case ExcelType.xlsx:
                        workbook = new XSSFWorkbook();
                        break;
                    default:
                        workbook = new HSSFWorkbook();
                        break;
                }
                ISheet sheet;
                //表名
                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    //创建工作表
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    int i = 1;
                    while (true)
                    {
                        //判断表名是否存在
                        if (workbook.GetSheetIndex("sheet" + i) == -1)
                        {
                            //创建表
                            sheet = workbook.CreateSheet("sheet" + i);
                            break;
                        }
                        else
                        {
                            i++;
                        }
                    }
                }
                //创建表是否成功
                if (sheet == null)
                {
                    return false;
                }
                #region//添加表头
                if (hasFieldRow)//判断是否添加表头
                {
                    IRow headRow = sheet.CreateRow(0);
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        //获取单元格
                        ICell cell = headRow.CreateCell(i);
                        //设置单元格数据类型
                        cell.SetCellType(CellType.String);
                        //单元格赋值
                        cell.SetCellValue(dt.Columns[i].ColumnName);
                    }
                }
                #endregion

                #region//添加数据
                //添加表头为1，不添加为0
                int starRow;
                if (hasFieldRow)
                {
                    starRow = 1;
                }
                else
                {
                    starRow = 0;
                }
                ICellStyle style = workbook.CreateCellStyle();
                style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
                style.FillPattern = FillPattern.SolidForeground;
                //遍历行
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //获取第i行
                    DataRow dr = dt.Rows[i];
                    //判断是不是标注行
                    bool isTag = tagRowIndex.Contains(i) ? true : false;
                    //创建行
                    IRow cells = sheet.CreateRow(i + starRow);
                    //遍历列
                    for (int j = 0; j < dr.ItemArray.Length; j++)
                    {

                        //创建单元格
                        ICell cell = cells.CreateCell(j);
                        //设置单元格数据类型
                        cell.SetCellType(CellType.String);
                        //设置单元格值
                        cell.SetCellValue(dr.ItemArray[j].ToString());
                        //判断是否进行标注行
                        if (isTag is true)
                        {
                            cell.CellStyle = style;
                        }

                    }
                }
                #endregion
                //保存至Path
                bool success = Export(workbook, path);
                return success;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 多表存储至指定路径
        /// </summary>
        /// <param name="path"></param>
        /// <param name="dts"></param>
        /// <param name="hasFieldRow"></param>
        /// <returns></returns>
        public static bool SaveAsExcel(string path, List<DataTable> dts, bool hasFieldRow = true)
        {
            try
            {
                //获取excel的文件类型
                ExcelType type = GetExcelFileType(path).Value;
                IWorkbook workbook;
                //新建工作目录
                switch (type)
                {
                    case ExcelType.xlsx:
                        workbook = new XSSFWorkbook();
                        break;
                    default:
                        workbook = new HSSFWorkbook();
                        break;
                }
                foreach (DataTable dt in dts)
                {
                    ISheet sheet;
                    string sheetName = dt.TableName;
                    //表名
                    if (!string.IsNullOrWhiteSpace(sheetName))
                    {
                        //创建工作表
                        sheet = workbook.CreateSheet(sheetName);
                    }
                    else
                    {
                        int i = 1;
                        while (true)
                        {
                            //判断表名是否存在
                            if (workbook.GetSheetIndex("sheet" + i) == -1)
                            {
                                //创建表
                                sheet = workbook.CreateSheet("sheet" + i);
                                break;
                            }
                            else
                            {
                                i++;
                            }
                        }
                    }
                    //创建表是否成功
                    if (sheet == null)
                    {
                        return false;
                    }
                    #region//添加表头
                    if (hasFieldRow)//判断是否添加表头
                    {
                        IRow headRow = sheet.CreateRow(0);
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            //获取单元格
                            ICell cell = headRow.CreateCell(i);
                            //设置单元格数据类型
                            cell.SetCellType(CellType.String);
                            //单元格赋值
                            cell.SetCellValue(dt.Columns[i].ColumnName);
                        }
                    }
                    #endregion

                    #region//添加数据
                    //添加表头为1，不添加为0
                    int starRow;
                    if (hasFieldRow)
                    {
                        starRow = 1;
                    }
                    else
                    {
                        starRow = 0;
                    }
                    //遍历行
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        //获取第i行
                        DataRow dr = dt.Rows[i];
                        //创建行
                        IRow cells = sheet.CreateRow(i + starRow);
                        //遍历列
                        for (int j = 0; j < dr.ItemArray.Length; j++)
                        {
                            //创建单元格
                            ICell cell = cells.CreateCell(j);
                            //设置单元格数据类型
                            cell.SetCellType(CellType.String);
                            //设置单元格值
                            cell.SetCellValue(dr.ItemArray[j].ToString());
                        }
                    }
                    #endregion
                }
                //保存至Path
                bool success = Export(workbook, path);
                return success;
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// 往Excel里面插入DataTable
        /// </summary>
        /// <param name="path">excel路径</param>
        /// <param name="dt">表</param>
        /// <param name="aimRowIndex">从第几行开始</param>
        /// <param name="insertIndex">下标</param>
        /// <param name="hasFieldRow"></param>
        /// <returns></returns>
        public static bool InsertDataoExcel(string path, DataTable dt, int insertIndex, int aimRowIndex = 0, bool hasFieldRow = true)
        {
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    ISheet sheet = null;
                    int name_i = 1;
                    while (true)
                    {
                        if (workbook.GetSheetIndex("Sheet" + name_i) == -1)
                        {
                            //判断插入下标是否符合要求
                            if (insertIndex > (workbook.NumberOfSheets + 1))
                            {
                              //  System.Windows.Forms.MessageBox.Show("不能将表插入该索引下！");
                            }
                            else
                            {
                                //创建工作表
                                sheet = workbook.CreateSheet("Sheet" + name_i);
                                //设置表的索引下标
                                workbook.SetSheetOrder("Sheet" + name_i, insertIndex);
                            }
                            break;
                        }
                        else
                        {
                            name_i++;
                        }
                    }
                    //是否已经获取工作表，如果没有则直接返回
                    if (sheet == null)
                    {
                        return false;
                    }

                    #region//往表中追加Datatable（包括表头信息）
                    if (insertIndex < workbook.NumberOfSheets)//插入的下标不能在已有数据范围
                    {
                        if (hasFieldRow)
                        {
                            //添加表头
                            IRow headRow = sheet.CreateRow(0 + aimRowIndex);
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                //获取单元格
                                ICell cell = headRow.CreateCell(i);
                                //设置单元格数据类型
                                cell.SetCellType(CellType.String);
                                //单元格赋值
                                cell.SetCellValue(dt.Columns[i].ColumnName);
                            }
                        }
                        //添加表头starRow为1，不添加则为0
                        int starRow;
                        if (hasFieldRow)
                        {
                            starRow = 1;
                        }
                        else
                        {
                            starRow = 0;
                        }
                        //遍历行
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            //创建行
                            IRow cells = sheet.CreateRow(sheet.LastRowNum + starRow);
                            //遍历列
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                //创建单元格
                                ICell cell = cells.CreateCell(j);
                                //设置单元格属性
                                cell.SetCellType(CellType.String);
                                //单元格赋值
                                cell.SetCellValue(dt.Rows[i][j].ToString());
                            }
                        }
                    }
                    #endregion

                    //保存
                    bool success = Export(workbook, path);
                    return success;
                }
            }
            catch
            {
                return false;
            }

        }

        /// <summary>
        /// 往Excel里面插入DataTable
        /// </summary>
        /// <param name="path">excel路径</param>
        /// <param name="dt">表</param>
        /// <param name="sheetName">sheet表名</param>
        /// <param name="insertIndex">下标</param>
        /// <param name="hasFieldRow"></param>
        /// <returns></returns>
        public static bool InsertDataoExcel(string path, DataTable dt, string sheetName, int insertIndex, bool hasFieldRow = true)
        {
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    ISheet sheet = null;
                    if (workbook.GetSheetIndex(sheetName) == -1)
                    {
                        if (insertIndex > (workbook.NumberOfSheets + 1))
                        {
                          //  System.Windows.Forms.MessageBox.Show("不能将表插入该索引下！");
                        }
                        else
                        {
                            //创建表
                            sheet = workbook.CreateSheet(sheetName);
                            //设置表的索引下标
                            workbook.SetSheetOrder(sheetName, insertIndex);
                        }


                    }
                    else
                    {
                       // System.Windows.Forms.MessageBox.Show(sheetName + "表创建失败," + sheetName + "表已经存在!");
                        return false;
                    }

                    //是否已经获取工作表，如果没有则直接返回
                    if (sheet == null)
                    {
                        return false;
                    }

                    #region//往表中追加Datatable
                    if (insertIndex < workbook.NumberOfSheets)//插入的下标不能在已有数据范围
                    {
                        if (hasFieldRow)
                        {
                            //添加表头
                            IRow headRow = sheet.CreateRow(0);
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                //获取单元格
                                ICell cell = headRow.CreateCell(i);
                                //设置单元格数据类型
                                cell.SetCellType(CellType.String);
                                //单元格赋值
                                cell.SetCellValue(dt.Columns[i].ColumnName);
                            }
                        }
                        int starRow;
                        if (hasFieldRow)
                        {
                            starRow = 1;
                        }
                        else
                        {
                            starRow = 0;
                        }
                        //遍历行
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            //创建行
                            IRow cells = sheet.CreateRow(sheet.LastRowNum + starRow);
                            //遍历列
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                //创建单元格
                                ICell cell = cells.CreateCell(j);
                                //设置单元格属性
                                cell.SetCellType(CellType.String);
                                //单元格赋值
                                cell.SetCellValue(dt.Rows[i][j].ToString());
                            }
                        }
                    }
                    #endregion
                    //保存
                    bool success = Export(workbook, path);
                    return success;
                }
            }
            catch
            {
                return false;
            }

        }
        /// <summary>
        /// 往Excel里面插入Sheet
        /// </summary>
        /// <param name="aimpath">目标excel</param>
        /// <param name="aimSheetIndex">目标sheet下标</param>
        /// <param name="aimRowIndex">目标从第几行开始</param>
        /// <param name="path">源excel路径</param>
        /// <param name="sheetIndex">源excelsheet下标</param>
        /// <returns></returns>
        public static bool InsertDataoExcel(string aimpath, int aimSheetIndex, int aimRowIndex, string path, int sheetIndex)
        {
            try
            {
                //获取源目标的datatble
                DataTable dt = GetDataTable(path, sheetIndex);
                //将datatable插入到目标excel中
                bool success = InsertDataoExcel(aimpath, dt, aimSheetIndex, aimRowIndex);
                return success;
            }
            catch
            {
                return false;
            }

        }

        /// <summary>
        /// 追加记录DataTable(不含表头，将表内容直接追加进去至末尾)
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="sheetIndex">目标表下标</param>
        /// <param name="dt">表格</param>
        /// <returns></returns>
        public static bool AppendDataTableToExcel(string path, int sheetIndex, DataTable dt, bool hasFieldRow = true)
        {
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    //
                    ISheet sheet = null;
                    //获取工作表,默认获取第一张
                    if (workbook.NumberOfSheets > sheetIndex && workbook.NumberOfSheets > 0)
                    {
                        sheet = workbook.GetSheetAt(sheetIndex);
                    }
                    else
                    {
                      //  System.Windows.Forms.MessageBox.Show("索引超出界限或该文件表格为空！");
                    }
                    //是否已经获取工作表，如果没有则直接返回
                    if (sheet == null)
                    {
                        return false;
                    }
                    #region//往表中追加Datatable
                    if (hasFieldRow)//判断是否添加表头
                    {

                        //添加表头(当表为空时从0开始，不为空则加1)
                        IRow headRow = sheet.CreateRow(sheet.LastRowNum + (sheet.LastRowNum == 0 ? 0 : 1));
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            //获取单元格
                            ICell cell = headRow.CreateCell(i);
                            //设置单元格数据类型
                            cell.SetCellType(CellType.String);
                            //单元格赋值
                            cell.SetCellValue(dt.Columns[i].ColumnName.ToString());
                        }
                    }
                    //添加表头为1，不添加表头为0
                    int starRow;
                    if (hasFieldRow)
                    {
                        starRow = 1;
                    }
                    else
                    {
                        starRow = 0;
                    }
                    //遍历行
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        //创建行
                        IRow cells = sheet.CreateRow(sheet.LastRowNum + starRow);
                        //遍历列
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            //创建单元格
                            ICell cell = cells.CreateCell(j);
                            //设置单元格属性
                            cell.SetCellType(CellType.String);
                            //单元格赋值
                            cell.SetCellValue(dt.Rows[i][j].ToString());
                        }
                    }
                    #endregion
                    //保存
                    bool success = Export(workbook, path);
                    return success;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 追加记录DataRow
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="sheetIndex">目标表下标</param>
        /// <param name="dr">行数据</param>
        /// <returns></returns>
        public static bool AppendDataRowToExcel(string path, int sheetIndex, DataRow dr)
        {
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    //
                    ISheet sheet = null;
                    //获取工作表,默认获取第一张
                    if (workbook.NumberOfSheets > sheetIndex && workbook.NumberOfSheets > 0)
                    {
                        //获取工作表
                        sheet = workbook.GetSheetAt(sheetIndex);
                    }
                    else
                    {
                      //  System.Windows.Forms.MessageBox.Show("索引超出界限或该文件表格为空！");
                    }
                    //是否已经获取工作表，如果没有则直接返回
                    if (sheet == null)
                    {
                        return false;
                    }
                    #region//往表中追加Datatable(没有追加表头数据)
                    //创建行单元(当表为空时从0开始，不为空则加1)
                    IRow cells = sheet.CreateRow(sheet.LastRowNum + (sheet.LastRowNum == 0 ? 0 : 1));
                    //遍历列
                    for (int j = 0; j < dr.ItemArray.Length; j++)
                    {
                        //创建单元格
                        ICell cell = cells.CreateCell(j);
                        //设置单元格属性
                        cell.SetCellType(CellType.String);
                        //单元格赋值
                        cell.SetCellValue(dr.ItemArray[j].ToString());
                    }
                    #endregion
                    //保存
                    bool success = Export(workbook, path);
                    return success;
                }
            }
            catch
            {
                return false;
            }
        }


        /// <summary>
        /// 插入一个工作表
        /// </summary>
        /// <param name="aimpath">目标表路径</param>
        /// <param name="aimSheetIndex">目标表存放表下标</param>
        /// <param name="aimRowIndex">目标表第几行开始</param>
        /// <param name="path">插入表路径</param>
        /// <param name="sheetIndex">插入表的表下标</param>
        /// <returns></returns>
        public static bool InsetSheetToExcel(string aimpath, int aimSheetIndex, int aimRowIndex, string path, int sheetIndex)
        {
            try
            {
                //目标工作表
                ISheet aimSheet = CreateSheet(aimpath, aimSheetIndex);
                //插入的工作表
                ISheet pathISheet = GetExcelSheet(path, sheetIndex);
                //判断工作表是否为空
                if (aimSheet != null && pathISheet != null)
                {
                    for (int i = 0; i <= pathISheet.LastRowNum; i++)
                    {
                        //新建行
                        IRow cells = aimSheet.CreateRow(i);
                        //获取插入表行
                        IRow row = pathISheet.GetRow(i);
                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            //创建单元
                            ICell cell = cells.CreateCell(j);
                            //设置单元格类型
                            cell.SetCellType(CellType.String);
                            //单元格赋值
                            cell.SetCellValue(row.GetCell(j).ToString());
                        }
                    }
                    Export(aimSheet.Workbook, aimpath);
                    return true;
                }
            }
            catch
            {
                return false;
            }
            return false;
        }
        #endregion

        #region//Private-Method
        /// <summary>
        /// 获取Excel里面的工作表
        /// </summary>
        /// <param name="path">Excel路径</param>
        /// <param name="sheetIndex">工作表索引下标</param>
        /// <returns></returns>
        private static ISheet GetExcelSheet(string path, int sheetIndex)
        {
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    //
                    ISheet sheet = null;
                    //获取工作表,默认获取第一张
                    if (workbook.NumberOfSheets > sheetIndex && workbook.NumberOfSheets > 0)
                    {
                        //获取工作表
                        sheet = workbook.GetSheetAt(sheetIndex);
                    }
                    else
                    {
                       // System.Windows.Forms.MessageBox.Show("索引超出界限或该文件表格为空！");
                    }
                    //是否已经获取工作表，如果没有则直接返回
                    if (sheet == null)
                    {
                        return null;
                    }
                    else
                    {
                        return sheet;
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 获取Excel里面的工作表
        /// </summary>
        /// <param name="path">Excel路径</param>
        /// <param name="sheetIndex">工作表名称</param>
        /// <returns></returns>
        private static ISheet GetExcelSheet(string path, string sheetName)
        {
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    //
                    ISheet sheet = null;
                    //获取工作表,默认获取第一张
                    if (string.IsNullOrWhiteSpace(sheetName))
                    {
                        //sheet = workbook.GetSheetAt(0);
                       // System.Windows.Forms.MessageBox.Show(sheetName + "表不存在!");
                    }
                    else
                    {
                        sheet = workbook.GetSheet(sheetName);
                    }
                    //是否已经获取工作表，如果没有则直接返回
                    if (sheet == null)
                    {
                        return null;
                    }
                    else
                    {
                        return sheet;
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Excel创建一个工作表
        /// </summary>
        /// <param name="path"></param>
        /// <param name="insertIndex"></param>
        /// <returns></returns>
        private static ISheet CreateSheet(string path, int insertIndex)
        {
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;

                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    ISheet sheet = null;
                    int name_i = 1;
                    while (true)
                    {
                        if (workbook.GetSheetIndex("Sheet" + name_i) == -1)
                        {
                            //判断插入下标是否符合要求
                            if (insertIndex > (workbook.NumberOfSheets + 1))
                            {
                              //  System.Windows.Forms.MessageBox.Show("不能将表插入该索引下！");
                            }
                            else
                            {
                                //创建工作表
                                sheet = workbook.CreateSheet("Sheet" + name_i);
                                //设置表的索引下标
                                workbook.SetSheetOrder("Sheet" + name_i, insertIndex);
                            }
                            break;
                        }
                        else
                        {
                            name_i++;
                        }
                    }
                    //是否已经获取工作表，如果没有则直接返回
                    if (sheet == null)
                    {
                        return null;
                    }
                    else
                    {
                        return sheet;
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Excel创建一个工作表
        /// </summary>
        /// <param name="path"></param>
        /// <param name="insertIndex"></param>
        /// <returns></returns>
        private static ISheet CreateSheet(string path, string sheetName, int insertIndex = 0)
        {
            try
            {
                //获取文件流
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(path).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                    ISheet sheet = null;
                    if (workbook.GetSheetIndex(sheetName) == -1)
                    {
                        if (insertIndex > (workbook.NumberOfSheets + 1))
                        {
                          //  System.Windows.Forms.MessageBox.Show("不能将表插入该索引下！");
                        }
                        else
                        {
                            //创建表
                            sheet = workbook.CreateSheet(sheetName);
                            //设置表的索引下标
                            workbook.SetSheetOrder(sheetName, insertIndex);
                        }

                    }
                    else
                    {
                       // System.Windows.Forms.MessageBox.Show(sheetName + "表创建失败," + sheetName + "表已经存在!");
                    }
                    //是否已经获取工作表，如果没有则直接返回
                    if (sheet == null)
                    {
                        return null;
                    }
                    else
                    {
                        return sheet;
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// excel的类型
        /// </summary>
        private enum ExcelType
        {
            xlsx, xls,
        }

        /// <summary>
        /// 获取指定excel文件后缀的格式类型，Nullable为可控类型如Nullable<int32>
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static Nullable<ExcelType> GetExcelFileType(string fileName)
        {
            var ext = Path.GetExtension(fileName);
            if (!string.IsNullOrWhiteSpace(ext) && (ext.ToLower() == ".xls" || ext.ToLower() == ".xlsx"))
                return ext.ToLower() == ".xls" ? ExcelType.xls : ExcelType.xlsx;
            else
                return null;
        }
        /// <summary>
        /// 获取表格对应datatable的数据类型
        /// </summary>
        /// <param name="cellType"></param>
        /// <returns></returns>
        private static Type GetType(CellType cellType)
        {
            Type type = null;
            switch (cellType)
            {
                case CellType.Numeric:
                    type = System.Type.GetType("System.Double");
                    break;
                default:
                    type = System.Type.GetType("System.String");
                    break;
            }
            return type;
        }
        /// <summary>
        /// 获取表格对应datatable的数据类型
        /// </summary>
        /// <param name="cellType"></param>
        /// <returns></returns>
        private static CellType GetType(Type type)
        {
            CellType celltype = CellType.String;
            switch (type.ToString())
            {
                case "System.Int32":
                case "System.Float":
                case "System.Double":
                    celltype = CellType.Numeric;
                    break;
                default:
                    celltype = CellType.String;
                    break;
            }
            return celltype;
        }

        /// <summary>
        /// 导出至savePath
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="savePath"></param>
        private static bool Export(IWorkbook workbook, string savePath)
        {
            try
            {
                // 写入 ,创建其支持存储区为内存的流
                System.IO.MemoryStream ms = new System.IO.MemoryStream();
                //写入内存
                workbook.Write(ms);
                workbook = null;
                //
                FileStream fs = new FileStream(savePath, FileMode.Create, FileAccess.Write);
                byte[] data = ms.ToArray();
                fs.Write(data, 0, data.Length);
                fs.Flush();
                //关闭
                ms.Close();
                //释放
                ms.Dispose();
                fs.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 导出至savePath
        /// </summary>
        /// <param name="savePath"></param>
        private static bool ExcelSave(string savePath)
        {
            try
            {
                IWorkbook workbook;
                //获取文件流
                using (var stream = new FileStream(savePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    //获取excel的文件类型
                    ExcelType type = GetExcelFileType(savePath).Value;
                    //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                    switch (type)
                    {
                        case ExcelType.xlsx:
                            workbook = new XSSFWorkbook(stream);
                            break;
                        default:
                            workbook = new HSSFWorkbook(stream);
                            break;
                    }
                }
                if (workbook == null)
                {
                    return false;
                }
                // 写入 ,创建其支持存储区为内存的流
                System.IO.MemoryStream ms = new System.IO.MemoryStream();
                //写入内存
                workbook.Write(ms);
                workbook = null;
                //
                FileStream fs = new FileStream(savePath, FileMode.Create, FileAccess.Write);
                byte[] data = ms.ToArray();
                fs.Write(data, 0, data.Length);
                fs.Flush();
                //关闭
                ms.Close();
                //释放
                ms.Dispose();
                fs.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static IWorkbook GetWorkBook(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                IWorkbook workbook;
                //获取excel的文件类型
                ExcelType type = GetExcelFileType(path).Value;
                //通过不同的文件类型创建不同的读取接口(xls使用HSSFWorkbook类实现，xlsx使用XSSFWorkbook类实现)
                switch (type)
                {
                    case ExcelType.xlsx:
                        workbook = new XSSFWorkbook(stream);
                        break;
                    default:
                        workbook = new HSSFWorkbook(stream);
                        break;
                }
                return workbook;
            }

        }
        #endregion
    }
}
