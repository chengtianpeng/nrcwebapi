namespace NRCWebApi.Dto
{
    /// <summary>
    /// 密码返回参数
    /// </summary>
    public class UpdateEmailPasswordResponseDto
    {
        public int errcode { get; set; }
        public string errmsg { get; set; }
        public string access_token { get; set; }
        public int expires_in { get; set; }
    }
}
