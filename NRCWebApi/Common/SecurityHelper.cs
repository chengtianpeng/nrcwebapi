using System.Security.Cryptography;
using System.Text;

namespace NRCWebApi.Common
{
    /// <summary>
    /// 密码帮助类
    /// </summary>
    public static class SecurityHelper
    {
        /// <summary>
        /// 字符串转MD5
        /// </summary>
        /// <param name="Plaintext"></param>
        /// <returns></returns>
        public static string ToMD5String(string Plaintext)
        {
            #region 代码实现
            string tempString = string.Empty;
            string ciphertext = string.Empty;
            //实例化一个MD5对像
            System.Security.Cryptography.MD5 Instance = System.Security.Cryptography.MD5.Create();
            byte[] Buffer = Instance.ComputeHash(Encoding.Default.GetBytes(Plaintext));
            for (int i = 0; i < Buffer.Length; i++)
            {
                tempString = Buffer[i].ToString("X");
                if (tempString.Length == 1)
                {
                    tempString = "0" + tempString;
                }
                ciphertext += tempString;
            }
            return ciphertext;
            #endregion
        }


        /// <summary>
        /// 生成一个 32 字节的密钥
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static byte[] JWTEncrypt(string str)
        {
            byte[] key = new byte[32];
            using (var generator = new Rfc2898DeriveBytes(str, 16))
            {
                key = generator.GetBytes(32);
            }
            return key;
            
        }
    }
}
