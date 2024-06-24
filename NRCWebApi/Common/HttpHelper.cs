using Newtonsoft.Json;
using NPOI.SS.Formula.Functions;
using System.Security.Policy;
using System.Text;

namespace NRCWebApi.Common
{
    /// <summary>
    /// http请求
    /// </summary>
    public class HttpHelper
    {

        /// <summary>
        /// get请求 
        /// </summary>
        /// <param name="requestUrl">地址</param>
        /// <param name="parameters">参数</param>
        /// <returns></returns>
        public static async Task<string> DoGet(string requestUrl, Dictionary<string, string> parameters = null)
        {
            using (HttpClient client = new HttpClient())
            {
                string address = BuildQuery(requestUrl, parameters);
                var result = client.GetAsync(address);
                //if (!result.IsSuccessStatusCode) return null;
                return await result.Result.Content.ReadAsStringAsync();
            }
        }

        /// <summary>
        /// Get请求发送
        /// </summary>
        /// <param name="requestUrl">url地址</param>
        /// <param name="parameters">参数</param>
        /// <returns></returns>
        public static async Task<T> DoGet<T>(string requestUrl, Dictionary<string, string> parameters = null) where T : new()
        {
            var res = new T();

            using (HttpClient httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Add("Method", "Get");
                string address = BuildQuery(requestUrl, parameters);
                HttpResponseMessage response = await httpClient.GetAsync(address);
                response.EnsureSuccessStatusCode();
                string responseBody = await response.Content.ReadAsStringAsync();
                res = JsonConvert.DeserializeObject<T>(responseBody);
            }
            return res;
        }

        /// <summary>
        /// 返回字节
        /// </summary>
        /// <param name="requestUrl">url地址</param>
        /// <param name="parameters">参数</param>
        /// <returns></returns>
        public static async Task<byte[]> GetHttpToByte(string requestUrl, Dictionary<string, string> parameters = null)
        {
            using (HttpClient client = new HttpClient())
            {
                string address = BuildQuery(requestUrl, parameters);
                var result = client.GetAsync(address);
                return await result.Result.Content.ReadAsByteArrayAsync();
            }
        }

        /// <summary>
        /// post请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static async Task<string> DoPost(string url, object message)
        {
            using (HttpClient client = new HttpClient())
            {
                string jsonContent = JsonConvert.SerializeObject(message);
                var result = client.PostAsync(url, new StringContent(jsonContent, Encoding.UTF8, "application/json"));
                //post传过去的字符串最好编码
                //if (!result.IsSuccessStatusCode) return null;
                return await result.Result.Content.ReadAsStringAsync();
            }
        }

        /// <summary>
        /// post请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static async Task<T> DoPost<T>(string url, object message) where T : new()
        {
            var res = new T();
            string jsonContent = JsonConvert.SerializeObject(message);
            string responseBody = string.Empty;
            using (HttpClient httpClient = new HttpClient())
            {
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                httpClient.DefaultRequestHeaders.Add("Method", "Post");
                HttpResponseMessage response = await httpClient.PostAsync(url, content);
                response.EnsureSuccessStatusCode();
                responseBody = await response.Content.ReadAsStringAsync();
                res = JsonConvert.DeserializeObject<T>(responseBody);
            }
            return res;
        }


        /// <summary>
        /// 组装普通文本请求参数。
        /// </summary>
        /// <param name="url">地址</param>
        /// <param name="parameters">Key-Value形式请求参数字典</param>
        /// <returns>URL编码后的请求数据</returns>
        private static string BuildQuery(string url, IDictionary<string, string> parameters)
        {
            if (parameters == null || !parameters.Any())
            {
                return url;
            }

            var postData = new StringBuilder();
            var hasParam = false;
            using var dem = parameters.GetEnumerator();
            while (dem.MoveNext())
            {
                var name = dem.Current.Key;
                var value = dem.Current.Value;
                // 忽略参数名或参数值为空的参数
                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(value)) continue;
                if (hasParam)
                {
                    postData.Append("&");
                }

                postData.Append(name);
                postData.Append("=");
                postData.Append(value);
                hasParam = true;
            }
            string urlnew = $"{url}?{postData.ToString()}";

            return urlnew;
        }
    }
}
