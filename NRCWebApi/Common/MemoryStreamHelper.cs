namespace NRCWebApi.Common
{
    /// <summary>
    /// 文件流重写
    /// </summary>
    public class MemoryStreamHelper : MemoryStream
    {
        /// <summary>
        /// 构造方法
        /// </summary>
        public MemoryStreamHelper()
        {
            AllowClose = true;
        }

        /// <summary>
        /// 是否允许关闭
        /// </summary>
        public bool AllowClose { get; set; }

        /// <summary>
        /// 重新关闭
        /// </summary>
        public override void Close()
        {
            if (AllowClose)
                base.Close();
        }

    }
}
