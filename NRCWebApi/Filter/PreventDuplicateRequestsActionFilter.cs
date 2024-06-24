using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.AspNetCore.Mvc.ModelBinding;
using Newtonsoft.Json;
using StackExchange.Redis;

namespace NRCWebApi.Filter
{
    /// <summary>
    /// 防重组件
    /// 防重组件的思路很简单，将第一次请求的某些参数作为标识符存入redis中，并设置过期时间，下次请求过来，先检查redis相同的请求是否已被处理
    /// </summary>
    public class PreventDuplicateRequestsActionFilter : IAsyncActionFilter, IAsyncResultFilter
    {
        /// <summary>
        /// 失效时间：秒
        /// </summary>
        public TimeSpan? AbsoluteExpirationRelativeToNow { get; set; }

        private bool _isIdempotencyCache = false;
        private string _idempotentKey;
        private readonly IDatabase _database;
        private readonly ILogger<PreventDuplicateRequestsActionFilter> _logger;

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="redis"></param>
        /// <param name="logger"></param>
        public PreventDuplicateRequestsActionFilter(IConnectionMultiplexer redis, ILogger<PreventDuplicateRequestsActionFilter> logger)
        {
            _database = redis.GetDatabase();
            _logger = logger;
        }

        /// <summary>
        /// action执行前调用
        /// </summary>
        /// <param name="context"></param>
        /// <param name="next"></param>
        /// <returns></returns>
        public async Task OnActionExecutionAsync(ActionExecutingContext context, ActionExecutionDelegate next)
        {
            _idempotentKey = GetDistributedCacheKey(context);
            RedisValue cacheData = await _database.StringGetAsync(_idempotentKey);
            if (cacheData.HasValue)
            {
                context.Result = JsonConvert.DeserializeObject<ObjectResult>(cacheData.ToString());
                _isIdempotencyCache = true;
            }
            else
            {
                await next();
            }
        }

        /// <summary>
        /// action执行后调用
        /// </summary>
        /// <param name="context"></param>
        /// <param name="next"></param>
        /// <returns></returns>
        public async Task OnResultExecutionAsync(ResultExecutingContext context, ResultExecutionDelegate next)
        {
            //有缓存，直接返回 缓存的数据
            if (_isIdempotencyCache)
            {
                await next();
                return;
            }

            string contextResult = JsonConvert.SerializeObject(context.Result);
            bool success = await _database.StringSetAsync(_idempotentKey, contextResult, AbsoluteExpirationRelativeToNow, When.NotExists);
            if (success)
            {
                await next();
            }
            else
            {
                _logger.LogWarning($"Received duplicate request({_idempotentKey}), short-circuiting...");
                context.Result = new AcceptedResult(); //返回202
            }

        }


        /// <summary>
        /// 获取redis key
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        private string GetDistributedCacheKey(ActionExecutingContext context)
        {

            string getStringParams = "";
            string methodType = context.HttpContext.Request.Method;
            if (methodType.ToLower().Equals("get"))
            {
                // 获取GET请求的所有查询字符串参数
                Dictionary<string, string> getPostparams = context.HttpContext.Request.Query.ToDictionary(
                     pair => pair.Key,
                     pair => pair.Value.ToString(),
                     StringComparer.OrdinalIgnoreCase);

                getStringParams = JsonConvert.SerializeObject(getPostparams, Formatting.Indented);
            }
            else if (methodType.ToLower().Equals("post"))
            {
                getStringParams = JsonConvert.SerializeObject(context.ActionArguments.FirstOrDefault().Value);
            }


            var idempotentKey = $"{context.HttpContext.Request.Path.Value}@{methodType}@{getStringParams}";

            return idempotentKey;

        }


      
    }
}
