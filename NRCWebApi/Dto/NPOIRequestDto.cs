namespace NRCWebApi.Dto
{
    /// <summary>
    /// NopI请求类
    /// </summary>
    public class NPOIRequestDto
    {

        /// <summary>
        /// 对应数据的字段名称
        /// </summary>
        public string DataName { get; set; }
        /// <summary>
        /// xlsx的名称
        /// </summary>
        public string TitleName { get; set; }
        /// <summary>
        /// 列宽
        /// </summary>
        public int Width { get; set; } = 20;
    }
}
