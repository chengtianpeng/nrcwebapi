using System.Net.Security;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace NRCWebApi.Common
{
    /// <summary>
    /// http
    /// </summary>
    public class WebRequestHelper
    {
        /// <summary>
        /// 设置cookie
        /// </summary>
        private CookieContainer cookie;

        /// <summary>
        /// 是否允许重定向
        /// </summary>
        private bool allowAutoRedirect = true;

        /// <summary>
        /// contentType
        /// </summary>
        private string contentType = "application/json";

        /// <summary>
        /// accept
        /// </summary>
        private string accept = "*/*";

        /// <summary>
        /// 过期时间
        /// </summary>
        private int time = 5000;

        /// <summary>
        /// 设置请求过期时间（单位：毫秒）（默认：5000）
        /// </summary>
        /// <param name="time"></param>
        public void SetTimeOut(int time)
        {
            this.time = time;
        }

        /// <summary>
        /// 设置accept(默认:*/*)
        /// </summary>
        /// <param name="accept"></param>
        public void SetAccept(string accept)
        {
            this.accept = accept;
        }

        /// <summary>
        /// 设置contentType(默认:application/x-www-form-urlencoded)
        /// </summary>
        /// <param name="contentType"></param>
        public void SetContentType(string contentType)
        {
            this.contentType = contentType;
        }

        /// <summary>
        /// 设置Cookie
        /// </summary>
        /// <param name="cookie"></param>
        public void SetCookie(CookieContainer cookie)
        {
            this.cookie = cookie;
        }
        /// <summary>
        /// 是否允许重定向(默认:true)
        /// </summary>
        /// <param name="allowAutoRedirect"></param>
        public void SetIsAllowAutoRedirect(bool allowAutoRedirect)
        {
            this.allowAutoRedirect = allowAutoRedirect;
        }

        /// <summary>
        /// post请求返回html
        /// </summary>
        /// <param name="url"></param>
        /// <param name="postDataStr"></param>
        /// <returns></returns>
        public string HttpPost(string url, string postdata)
        {

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.AllowAutoRedirect = allowAutoRedirect;
            request.Method = "POST";
            request.Accept = accept;
            request.ContentType = this.contentType;
            request.Timeout = time;
            byte[] data = Encoding.UTF8.GetBytes(postdata);
            request.ContentLength = data.Length;

            if (cookie != null)
                request.CookieContainer = cookie; //cookie信息由CookieContainer自行维护
            Stream myRequestStream = request.GetRequestStream();

            myRequestStream.Write(data, 0, data.Length);
            myRequestStream.Close();
            HttpWebResponse response = null;
            try
            {
                this.SetCertificatePolicy();
                response = (HttpWebResponse)request.GetResponse();
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
            //获取重定向地址
            //string url1 = response.Headers["Location"];
            if (response != null)
            {
                Stream myResponseStream = response.GetResponseStream();
                StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
                string retString = myStreamReader.ReadToEnd();
                myStreamReader.Close();
                myResponseStream.Close();
                return retString;
            }
            else
            {
                return null; //post请求返回为空
            }
        }
        /// <summary>
        /// get请求获取返回的html
        /// </summary>
        /// <param name="url">无参URL</param>
        /// <param name="querydata">参数</param>
        /// <returns></returns>
        public string HttpGet(string url, Dictionary<string, string> querydata)
        {
            if (querydata != null && querydata.Count > 0)
            {
                url += "?" + string.Join("&", querydata.Select(it => it.Key + "=" + it.Value));
            }

            return HttpGet(url);
        }
        /// <summary>
        /// get请求获取返回的html
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public string HttpGet(string url)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            request.ContentType = "text/html;charset=UTF-8";
            request.CookieContainer = cookie;
            request.Accept = this.accept;
            request.Timeout = time;
            this.SetCertificatePolicy();
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            // response.Cookies = cookie.GetCookies(response.ResponseUri);
            Stream myResponseStream = response.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
            string retString = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();
            return retString;
        }


        //注册证书验证回调事件，在请求之前注册
        private void SetCertificatePolicy()
        {
            ServicePointManager.ServerCertificateValidationCallback
                       += RemoteCertificateValidate;
        }
        /// <summary> 
        /// 远程证书验证，固定返回true
        /// </summary> 
        private static bool RemoteCertificateValidate(object sender, X509Certificate cert,
            X509Chain chain, SslPolicyErrors error)
        {
            return true;
        }
    }
}
