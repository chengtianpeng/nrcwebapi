namespace NRCWebApi.Common
{
    /// <summary>
    /// 公共帮助类
    /// </summary>
    public static class CommonHelper
    {
        /// <summary>
        /// 获取当前时间秒数
        /// </summary>
        /// <returns></returns>
        public static long GetSecondTimeStamp()
        {
            DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1)); // 当地时区
            long duration = (long)(DateTime.Now - startTime).TotalMilliseconds / 1000; // 相差毫秒秒数
            return duration;
        }

        /// <summary>
        /// 将 Stream 转成 byte[]
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static byte[] StreamToBytes(Stream stream)
        {
            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);
            // 设置当前流的位置为流的开始
            stream.Seek(0, SeekOrigin.Begin);
            return bytes;

        }
    }
}
