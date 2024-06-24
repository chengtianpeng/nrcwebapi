namespace NRCWebApi.Dto
{
    /// <summary>
    /// 导出中方薪资请求类
    /// </summary>
    public class SalaryExcelRequest
    {
        /// <summary>
        /// 开始月份如：202401
        /// </summary>
        public string BegionMonth { get; set; } = "202401";
        /// <summary>
        /// 开始月份：如：202402
        /// </summary>
        public string EndMonth { get; set; } = "202402";
        /// <summary>
        /// 人员字符串如：12344,43434
        /// </summary>
        public string CompanyIds { get; set; } = "";
        /// <summary>
        /// Hr系统Admin的密码
        /// </summary>
        public string AdminPassword { get; set; } = "";
    }
}
