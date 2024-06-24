namespace NRCWebApi.Dto
{
    public class ReturnMsg
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="success"></param>
        /// <param name="message"></param>
        public ReturnMsg()
        {
            AdSuccess = false;
            BindSuccess = false;
            Message = "";
        }


        public string UserName { get; set; }


        public bool AdSuccess { get; set; }

        /// <summary>
        /// 返回结果
        /// </summary>
        public bool BindSuccess { get; set; }

        /// <summary>
        /// 提示消息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 电话号码
        /// </summary>
        public string PhoneNumber { get; set; }

        public List<Msgmethod> msgmethods { get; set; } = new List<Msgmethod>();

        /// <summary>
        /// 邮件
        /// </summary>
        public string Email { get; set; }
    }

    public class Msgmethod
    {

        public string label { get; set; }

        public string value { get; set; }

    }
}
