namespace NRCWebApi.Dto
{

    /// <summary>
    /// 修改邮箱密码请求参数
    /// </summary>
    public class UpdateEmailPasswordRequestDto
    {

        /// <summary>
        /// tdnrc邮箱地址
        /// </summary>
        public string EmailAddriss { get; set; }
        /// <summary>
        /// 邮箱的新密码,必须同时包含大小写字母和数字，长度8-32位，不包含账户信息，账户若已绑定手机或微信，需员工修改密码
        /// </summary>
        public string Password { get; set; }
        /// <summary>
        /// 校验密码：moss的那个密码
        /// </summary>
        public string PermissionPass { get; set; }
    }
}
