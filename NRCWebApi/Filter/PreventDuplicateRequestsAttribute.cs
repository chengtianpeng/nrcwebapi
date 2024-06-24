using Microsoft.AspNetCore.Mvc.Filters;

namespace NRCWebApi.Filter
{
    /// <summary>
    /// PreventDuplicateRequestsAttribute
    /// </summary>
    [AttributeUsage(AttributeTargets.Method)]
    public class PreventDuplicateRequestsAttribute : Attribute, IFilterFactory
    {
        private readonly int _expiredSeconds;

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="expiredSeconds"></param>
        public PreventDuplicateRequestsAttribute(int expiredSeconds)
        {
            _expiredSeconds = expiredSeconds;
        }

        /// <summary>
        /// 实例化
        /// </summary>
        /// <param name="serviceProvider"></param>
        /// <returns></returns>
        public IFilterMetadata CreateInstance(IServiceProvider serviceProvider)
        {
            var filter = serviceProvider.GetService<PreventDuplicateRequestsActionFilter>();
            filter.AbsoluteExpirationRelativeToNow = TimeSpan.FromSeconds(_expiredSeconds);
            return filter;
        }
        public bool IsReusable => false;
    }
}
