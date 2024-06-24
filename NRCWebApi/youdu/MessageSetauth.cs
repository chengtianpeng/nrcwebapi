
using Newtonsoft.Json;
using System.Security.Cryptography;
using System.Text;

namespace NRCWebApi.youdu
{
    public class MessageSetauth
    {

        string _userId { get; set; }
        string _pwd { get; set; }
        public MessageSetauth(string userId, string pwd)
        {
            _userId = userId;
            _pwd = tomd5(pwd);
        }

        private string tomd5(string sourceText)
        {
            using (MD5 md5 = MD5.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(sourceText);
                byte[] hashBytes = md5.ComputeHash(inputBytes);
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < hashBytes.Length; i++)
                {
                    sb.Append(hashBytes[i].ToString("x2"));  // 将每个字节转成16进制形式
                }

                return sb.ToString();
            }
        }

        public string ToJson()
        {
            var jsonObj = new Dictionary<string, object>();
            jsonObj["userId"] = _userId;
            jsonObj["authType"] = 0;
            jsonObj["passwd"] = _pwd;


            return JsonConvert.SerializeObject(jsonObj, Formatting.Indented);
        }


    }
}
