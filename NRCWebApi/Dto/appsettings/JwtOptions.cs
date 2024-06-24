namespace NRCWebApi.Dto.appsettings
{
    /// <summary>
    /// jwt
    /// </summary>
    public class JwtOptions
    {
        /// <summary>
        /// 签发者
        /// </summary>
        public string Isuser { get; set; }

        /// <summary>
        /// 接收者
        /// </summary>
        public string Audience { get; set; }

        /// <summary>
        /// 加密密钥
        /// </summary>
        public string SecurityKey { get; set; }

        /// <summary>
        /// 过期时间
        /// </summary>
        public int ExpireSeconds { get; set; }
    }
}
