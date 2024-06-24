namespace NRCWebApi.Dto.appsettings
{
    /// <summary>
    /// 获取配置
    /// </summary>
    public class RedisServer
    {
        /// <summary>
        /// 数据库连接
        /// </summary>
        public string Connection { get; set; }
        /// <summary>
        /// 实例名
        /// </summary>
        public string InstanceName { get; set; }
        /// <summary>
        /// DB数据库
        /// </summary>
        public int DefaultDB { get; set; }

    }
}
