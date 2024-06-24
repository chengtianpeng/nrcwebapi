namespace NRCWebApi.Dto
{
    /// <summary>
    /// OA 用户信息DTO
    /// </summary>
    public class sys_user_infoDTO
    {

        /// <summary>
        /// 部门名称
        /// </summary>
        public string dept_name { get; set; }

        /// <summary>
        /// 用户账号
        /// </summary>
        public string? User_name { get; set; }
        /// <summary>
        /// 用户昵称
        /// </summary>

        public string nick_name { get; set; }
        /// <summary>
        /// 用户邮箱
        /// </summary>
        public string email { get; set; }
        /// <summary>
        /// 手机号码
        /// </summary>
        public string phonenumber { get; set; }
    }
}
