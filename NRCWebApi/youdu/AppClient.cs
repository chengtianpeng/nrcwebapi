using Newtonsoft.Json;
using NRCWebApi.Common;
using NRCWebApi.youdu.AES;
using System.Net.Http.Headers;
using System.Text;

namespace NRCWebApi.youdu
{
    public class AppClient
    {


        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="address">目标服务器地址，"ip:port"的形式</param>
        /// <param name="buin">企业号</param>
        /// <param name="appId">AppId</param>
        /// <param name="encodingAesKey">encodingAesKey</param>
        public AppClient(string address, int buin, string appId, string encodingAesKey)
        {
            if (address.Length == 0 || buin == 0 || appId.Length == 0 || encodingAesKey.Length == 0)
            {
                throw new ArgumentException();
            }
            m_addr = address;
            m_buin = buin;
            m_appId = appId;
            m_crypto = new AESCrypto(appId, encodingAesKey);
        }

        private class Token
        {
            public string accessToken;
            public long activeTime;
            public int expireIn;

            public Token()
            {
            }

            public Token(string token, long activeTime, int expire)
            {
                this.accessToken = token;
                this.activeTime = activeTime;
                this.expireIn = expire;
            }
        }

        private AESCrypto m_crypto;
        private Token m_tokenInfo;

        private string m_addr;
        private int m_buin;
        private string m_appId;

        public string Addr
        {
            get
            {
                return m_addr;
            }
        }

        public int Buin
        {
            get
            {
                return m_buin;
            }
        }

        public string AppId
        {
            get
            {
                return m_appId;
            }
        }


        private string apiGetToken()
        {
            return EntAppApi.SCHEME + m_addr + EntAppApi.API_GET_TOKEN;
        }



        private string apiSetauth()
        {
            return EntAppApi.SCHEME + m_addr + EntAppApi.API_setauth + "?accessToken=" + m_tokenInfo.accessToken;
        }


        private string apiState()
        {
            return EntAppApi.SCHEME + m_addr + EntAppApi.API_state + "?accessToken=" + m_tokenInfo.accessToken;
        }


        public async Task<bool> IsYDUserExist(string userid)
        {
            if (m_tokenInfo == null)
            {
                m_tokenInfo = await this.getToken();
            }
            long endTime = m_tokenInfo.activeTime + m_tokenInfo.expireIn;
            if (endTime <= CommonHelper.GetSecondTimeStamp())
            {
                m_tokenInfo = await this.getToken();
            }
            bool exist = false;
            try
            {
                var Client = new HttpClient();
                var response = await Client.GetAsync(this.apiState() + "&userId=" + userid);
                string result = await response.Content.ReadAsStringAsync();

                if (result != null)
                {
                    Resmsg resmsg = JsonConvert.DeserializeObject<Resmsg>(result);
                    if ((resmsg.errcode == "0") && (resmsg.errmsg == "ok"))
                    {
                        exist = true;
                    }

                }
            }
            catch (Exception e)
            {
                Console.WriteLine("获取有度用户信息失败|" + e.ToString() + Environment.NewLine);
            }
            return exist;
        }





        private async Task<Token> getToken()
        {
            try
            {
                var now = CommonHelper.GetSecondTimeStamp();
                var timestamp = AESCrypto.ToBytes(string.Format("{0}", now));
                var encryptTime = m_crypto.Encrypt(timestamp);
                var param = new Dictionary<string, object>()
                {
                    { "buin",  m_buin },
                    { "appId", m_appId },
                    { "encrypt" , encryptTime }
                };
                var Client = new HttpClient();
                string content = JsonConvert.SerializeObject(param);
                var bufferdata = Encoding.UTF8.GetBytes(content);
                var byteContent = new ByteArrayContent(bufferdata);
                byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                var response = await Client.PostAsync(this.apiGetToken(), byteContent).ConfigureAwait(false);
                string result = await response.Content.ReadAsStringAsync();

                if (result != null)
                {
                    Resmsg resmsg = JsonConvert.DeserializeObject<Resmsg>(result);
                    var buffer = Encoding.UTF8.GetString(m_crypto.Decrypt(resmsg.encrypt));
                    Token tokenInfo = JsonConvert.DeserializeObject<Token>(buffer);
                    tokenInfo.activeTime = CommonHelper.GetSecondTimeStamp();
                    return tokenInfo;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception e)
            {

                Console.WriteLine("获取有度token失败|" + e.ToString() + Environment.NewLine);
            }

            return null;
        }



        /// <summary>
        /// 修改用户密码
        /// </summary>
        /// <param name="msg"></param>
        /// <exception cref="HttpRequestException"></exception>
        /// <exception cref="UnexpectedException"></exception>
        public async Task<bool> ChangePWD(MessageSetauth msg)
        {
            if (m_tokenInfo == null)
            {
                m_tokenInfo = await this.getToken();
            }
            long endTime = m_tokenInfo.activeTime + m_tokenInfo.expireIn;
            if (endTime <= CommonHelper.GetSecondTimeStamp())
            {
                m_tokenInfo = await this.getToken();
            }

            bool issucess = false;
            try
            {
                Console.WriteLine(msg.ToJson());
                var cipherText = m_crypto.Encrypt(AESCrypto.ToBytes(msg.ToJson()));
                var param = new Dictionary<string, object>()
                {
                    { "buin", m_buin },
                    { "appId", m_appId },
                    { "encrypt", cipherText }
                };
                var Client = new HttpClient();

                string content = JsonConvert.SerializeObject(param);
                var bufferdata = Encoding.UTF8.GetBytes(content);
                var byteContent = new ByteArrayContent(bufferdata);
                byteContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");
                var response = await Client.PostAsync(this.apiSetauth(), byteContent).ConfigureAwait(false);
                string result = await response.Content.ReadAsStringAsync();


                Resmsg resmsg = JsonConvert.DeserializeObject<Resmsg>(result);

                if (resmsg.errmsg.Equals("ok"))
                {
                    issucess = true;
                }
                else
                {
                    Console.WriteLine("修改有度密码失败|" + resmsg.errcode + Environment.NewLine);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("修改有度密码异常|" + e.ToString() + Environment.NewLine);
            }
            return issucess;
        }

    }
}
