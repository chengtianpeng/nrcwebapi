using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging.Abstractions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NRCWebApi.Common;
using NRCWebApi.Dto;
using NRCWebApi.Model;
using SqlSugar;
using System.Collections.Generic;
using System.Data;

namespace NRCWebApi.Controllers
{
    /// <summary>
    /// HR操作类
    /// </summary>
    [Route("api/[controller]/[Action]")]
    [ApiController]
    public class HRController : ControllerBase
    {
        private readonly IDictionary<ConnectionKey, ISqlSugarClient> _sqlSugarClients;
        private readonly ILogger<HRController> _logger;

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="sqlSugarClients"></param>
        /// <param name="logger"></param>
        public HRController(IDictionary<ConnectionKey, ISqlSugarClient> sqlSugarClients,
            ILogger<HRController> logger)
        {
            _sqlSugarClients = sqlSugarClients;
            _logger = logger;
        }


        /// <summary>
        /// 通过模板导出人员薪资
        /// </summary>
        /// <returns></returns>
        [HttpGet(Name = "QuerySalary")]
        public IActionResult QuerySalary([FromQuery] SalaryExcelRequest request)
        {
            if (request == null || string.IsNullOrEmpty(request.CompanyIds) || string.IsNullOrEmpty(request.AdminPassword))
            {
                return Ok("请求参数不能为空");
            }

            var db = _sqlSugarClients[ConnectionKey.HR];
            string password = db.Ado.SqlQuerySingle<string>("select UserPassword from sms_temp_BuiltInUser where LoginName='ADMIN'");
            if (request.AdminPassword != password)
            {

                return Ok("密码不正确");
            }

            //找到文件路径 find file path :模板文件
            string filePath = Path.Combine(Path.GetFullPath("Excel"), "PersonnelSalaryForExpatriates_Single.xls");

            //导出的文件
            string saveFilePath = Path.Combine(Path.GetFullPath("Temp"), $"Salary{DateTime.Now:yyyyMMdd}.xls");


            string CompanyIdStr = "";
            List<string> CompanyIdList = request.CompanyIds.Trim().Split(',').ToList();
            for (int i = 0; i < CompanyIdList.Count; i++)
            {
                if (i == CompanyIdList.Count - 1)
                {
                    //最后一条记录
                    CompanyIdStr += "'" + CompanyIdList[i] + "'";
                }
                else
                {
                    CompanyIdStr += "'" + CompanyIdList[i] + "',";
                }
            }

            string sql = string.Format(@"SELECT
              (CASE
                WHEN ld1.DeptName IS NULL THEN ld.DeptName
                ELSE ld1.DeptName
              END
              ) AS DeptName
             ,(CASE
                WHEN ld1.DeptName IS NULL THEN NULL
                ELSE ld.DeptName
              END
              ) AS UnitName
             ,cc.*
             ,cc.[W&E_Allow] AS WE_Allow
             ,us.Email
             ,us.FamilyValue
            FROM dbo.cms_data_Chinese_Compensation cc
                 LEFT JOIN dbo.Uv_LocalDept ld
                   ON cc.Department = ld.DeptName
                 LEFT JOIN dbo.Uv_LocalDept ld1
                   ON ld.PDeptId = ld1.DeptId
                     AND ld1.PDeptId <> 0
                 LEFT JOIN dbo.Uv_EmployeeDeptUnit us
                   ON cc.CompanyId = us.CompanyId
                ,dbo.cms_biz_SalaryData s
            WHERE s.SalaryTable = 'cms_data_Chinese_Compensation'
            AND s.YearMonth = cc.YearMonth
            AND s.Status IN (5, 6, 7)
            AND cc.YearMonth
            BETWEEN '{0}' AND '{1}'
            AND [cc].CompanyId IN ({2})
            ORDER BY cc.CompanyId,
            cc.YearMonth;", request.BegionMonth, request.EndMonth, CompanyIdStr);


            DataTable salary = db.Ado.GetDataTable(sql);

            int step = 320;
            int tag = 0;

            byte[] buf = new byte[] { };
            for (int i = 0; i < salary.Rows.Count; i++)
            {
                if (i % step == 0)
                {
                    tag = i / step;

                    // byte[] buf = new byte[] { };

                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        HSSFWorkbook hssfworkbook = new HSSFWorkbook(fs);

                        for (int j = tag * step; j < (tag + 1) * step && j < salary.Rows.Count; j++)
                        {
                            ISheet sheet = j == tag * step ? hssfworkbook.GetSheetAt(0) : hssfworkbook.CloneSheet(0);
                            //设置sheet的名称，会报错
                            //hssfworkbook.SetSheetName(j, salary.Rows[j]["Name"].ToString() + "_" + salary.Rows[j]["YearMonth"].ToString());

                            //向sheet中插入数据
                            //标题
                            string month = "";

                            switch (salary.Rows[j]["YearMonth"].ToString().Substring(4))
                            {
                                case "01": month = "Janvier"; break;
                                case "02": month = "Février"; break;
                                case "03": month = "Mars"; break;
                                case "04": month = "Avril"; break;
                                case "05": month = "Mai"; break;
                                case "06": month = "Juin"; break;
                                case "07": month = "Juillet"; break;
                                case "08": month = "Août"; break;
                                case "09": month = "Septembre"; break;
                                case "10": month = "Octobre"; break;
                                case "11": month = "Novembre"; break;
                                case "12": month = "Décembre"; break;
                            }

                            sheet.GetRow(2).GetCell(1).SetCellValue("BULLETIN DE PAIE DE " + month + " " + salary.Rows[j]["YearMonth"].ToString().Substring(0, 4));

                            //表头
                            sheet.GetRow(4).GetCell(4).SetCellValue(salary.Rows[j]["DeptName"].ToString());
                            sheet.GetRow(4).GetCell(8).SetCellValue(salary.Rows[j]["UnitName"].ToString());
                            sheet.GetRow(5).GetCell(4).SetCellValue(salary.Rows[j]["Name"].ToString());
                            sheet.GetRow(6).GetCell(4).SetCellValue(salary.Rows[j]["CompanyId"].ToString());
                            sheet.GetRow(6).GetCell(8).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Exchange_rate"]), 2));
                            sheet.GetRow(7).GetCell(4).SetCellValue(salary.Rows[j]["Post"].ToString());

                            //津贴
                            sheet.GetRow(10).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Current_Base_Salary"]), 2));
                            sheet.GetRow(11).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Housing_Allow"]), 2));
                            sheet.GetRow(12).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Telecommunication_Allow"]), 2));
                            sheet.GetRow(13).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["WE_Allow"]), 2));
                            sheet.GetRow(14).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Extra_Pay"]), 2));
                            sheet.GetRow(15).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Household_Allow"]), 2));
                            sheet.GetRow(16).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Gas_Allow"]), 2));
                            sheet.GetRow(17).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Job_Subsidies"]), 2));
                            sheet.GetRow(18).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Family_Allow"]), 2));
                            sheet.GetRow(19).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Education_Allow"]), 2));
                            sheet.GetRow(20).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Transportation_Allow"]), 2));
                            sheet.GetRow(21).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Special_Industry_Allow"]), 2));
                            sheet.GetRow(22).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Medical_Allow"]), 2));

                            //扣减项
                            sheet.GetRow(23).GetCell(8).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Cnps_Private"]), 2));
                            sheet.GetRow(24).GetCell(8).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["PIT_Month_FCFA"]), 2));
                            sheet.GetRow(25).GetCell(8).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["FIR_FCFA"]), 2));

                            //汇总
                            sheet.GetRow(26).GetCell(7).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Salary_Pre_Tax"]), 2));

                            double rate = Convert.ToDouble(salary.Rows[j]["Exchange_rate"]);
                            double totalTax = Convert.ToDouble(salary.Rows[j]["Cnps_Private"]) / rate + Convert.ToDouble(salary.Rows[j]["PIT_Month_FCFA"]) / rate + Convert.ToDouble(salary.Rows[j]["FIR_FCFA"]) / rate;
                            sheet.GetRow(26).GetCell(8).SetCellValue(Math.Round(totalTax, 2));

                            sheet.GetRow(27).GetCell(8).SetCellValue(Math.Round(Convert.ToDouble(salary.Rows[j]["Salary_Taxed"]), 2));


                            Console.WriteLine(salary.Rows[j]["Name"].ToString() + "_" + salary.Rows[j]["YearMonth"].ToString() + "_" + (j + 1).ToString());
                        }

                        //转为字节数组  
                        MemoryStream stream = new MemoryStream();
                        hssfworkbook.Write(stream);
                        buf = stream.ToArray();
                        fs.Close();
                    }



                    //if (System.IO.File.Exists(saveFilePath))
                    //{
                    //    System.IO.File.Delete(saveFilePath);
                    //}

                    ////保存为文件  
                    //using (FileStream fs = new FileStream(saveFilePath, FileMode.CreateNew, FileAccess.ReadWrite))
                    //{
                    //    fs.Write(buf, 0, buf.Length);
                    //    fs.Flush();
                    //    fs.Close();
                    //}
                }
                else
                {
                    continue;
                }
            }

            //  return Ok(saveFilePath);

            return File(buf, "application/octet-stream", $"Salary{DateTime.Now:yyyyMMdd}.xls");


        }
    }
}
