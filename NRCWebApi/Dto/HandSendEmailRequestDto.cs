namespace NRCWebApi.Dto
{
    /// <summary>
    /// 手动发送邮件请求类
    /// </summary>
    public class HandSendEmailRequestDto
    {
        /// <summary>
        /// 发送人，itequipment@tdnrc.com 或者 pms@tdnrc.com
        /// </summary>
        public string MsgFrom { get; set; }

        /// <summary>
        /// 收件人，多个以,分割
        /// </summary>
        public string MsgTo { get; set; }

        /// <summary>
        /// 抄送人，多个以,分割
        /// </summary>
        public string Cc { get; set; } = "";

        /// <summary>
        /// 密送人，多个以,分割
        /// </summary>
        public string Bc { get; set; } = "";

        /// <summary>
        /// 标题
        /// </summary>
        public string MsgTitle { get; set; }

        /// <summary>
        /// 内容  
        /// </summary>
        public string MsgContent { get; set; }
        /// <summary>
        /// 校验密码：moss的那个密码
        /// </summary>
        public string PermissionPass { get; set; }
    }
}
