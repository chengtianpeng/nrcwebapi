using Newtonsoft.Json;
using NRCWebApi.Dto;
using NRCWebApi.Dto.appsettings;
using StackExchange.Redis;
using System.Collections.Concurrent;

namespace NRCWebApi.Common.Redis
{
    /// <summary>
    /// Redis帮助类
    /// </summary>
    public class RedisHelper : IDisposable
    {

        //连接字符串
        private string _connectionString;
        //实例名称
        private string _instanceName;
        //默认数据库
        private int _defaultDB = 0;
        private ConcurrentDictionary<string, ConnectionMultiplexer> _connections;

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="redisServer"></param>
        public RedisHelper(RedisServer redisServer)
        {
            _connectionString = redisServer.Connection;
            _instanceName = redisServer.InstanceName;
            _defaultDB = redisServer.DefaultDB;
            _connections = new ConcurrentDictionary<string, ConnectionMultiplexer>();
        }

        /// <summary>
        /// 获取ConnectionMultiplexer
        /// </summary>
        /// <returns></returns>
        private ConnectionMultiplexer GetConnect()
        {
            return _connections.GetOrAdd(_instanceName, p => ConnectionMultiplexer.Connect(_connectionString));
        }

        /// <summary>
        /// 获取数据库
        /// </summary>
        /// <returns></returns>
        public IDatabase GetDatabase()
        {
            return GetConnect().GetDatabase(_defaultDB);
        }

        /// <summary>
        /// 获取数据库
        /// </summary>
        /// <param name="DBNum">默认为0：优先代码的db配置，其次config中的配置</param>
        /// <returns></returns>
        public IDatabase GetDatabase(int DBNum)
        {
            return GetConnect().GetDatabase(DBNum);
        }

       
        #region Subscribe
        /// <summary>
        /// 订阅
        /// </summary>
        /// <param name="channel">频道</param>
        /// <param name="handle">事件</param>
        public void Subscribe(RedisChannel channel, Action<RedisChannel, RedisValue> handle)
        {
            //getSubscriber() 获取到指定服务器的发布者订阅者的连接
            var sub = GetConnect().GetSubscriber();
            //订阅执行某些操作时改变了 优先/主动 节点广播
            sub.Subscribe(channel, handle);
        }
        /// <summary>
        /// 发布
        /// </summary>
        /// <param name="channel"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public long Publish(RedisChannel channel, RedisValue message)
        {
            var sub = GetConnect().GetSubscriber();
            return sub.Publish(channel, message);
        }
        /// <summary>
        /// 发布（使用序列化）
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="channel"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public long Publish<T>(RedisChannel channel, T message)
        {
            var sub = GetConnect().GetSubscriber();
            return sub.Publish(channel, JsonConvert.SerializeObject(message));
        }

        /// <summary>
        /// 订阅
        /// </summary>
        /// <param name="redisChannel"></param>
        /// <param name="handle"></param>
        /// <returns></returns>
        public async Task SubscribeAsync(RedisChannel redisChannel, Action<RedisChannel, RedisValue> handle)
        {
            var sub = GetConnect().GetSubscriber();
            await sub.SubscribeAsync(redisChannel, handle);
        }
        /// <summary>
        /// 取消订阅
        /// </summary>
        /// <param name="redisChannel"></param>
        /// <param name="handle"></param>
        /// <returns></returns>
        public async Task UnsubscribeAsync(RedisChannel redisChannel)
        {
            var sub = GetConnect().GetSubscriber();
            await sub.UnsubscribeAsync(redisChannel);
        }
        /// <summary>
        /// 取消订阅
        /// </summary>
        /// <param name="redisChannel"></param>
        /// <param name="handle"></param>
        /// <returns></returns>
        public void Unsubscribe(RedisChannel redisChannel)
        {
            var sub = GetConnect().GetSubscriber();
            sub.Unsubscribe(redisChannel);
        }


        /// <summary>
        /// 发布
        /// </summary>
        /// <param name="redisChannel"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public async Task<long> PublishAsync(RedisChannel redisChannel, RedisValue message)
        {
            var sub = GetConnect().GetSubscriber();
            return await sub.PublishAsync(redisChannel, message);
        }
        /// <summary>
        /// 发布（使用序列化）
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="redisChannel"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public async Task<long> PublishAsync<T>(RedisChannel redisChannel, T message)
        {
            var sub = GetConnect().GetSubscriber();
            return await sub.PublishAsync(redisChannel, JsonConvert.SerializeObject(message));
        }
        #endregion

        #region --KEY/VALUE存取--
        /// <summary>
        /// 单条存值
        /// </summary>
        /// <param name="key">key</param>
        /// <param name="value">The value.</param>
        /// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        public bool StringSet(string key, string value)
        {
            return GetDatabase().StringSet(key, value);
        }
        /// <summary>
        /// 保存单个key value
        /// </summary>
        /// <param name="key">Redis Key</param>
        /// <param name="value">保存的值</param>
        /// <param name="expiry">过期时间</param>
        /// <returns></returns>
        public bool StringSet(string key, string value, TimeSpan? expiry = default(TimeSpan?))
        {
            return GetDatabase().StringSet(key, value, expiry);
        }
        /// <summary>
        /// 保存多个key value
        /// </summary>
        /// <param name="arr">key</param>
        /// <returns></returns>
        public bool StringSet(KeyValuePair<RedisKey, RedisValue>[] arr)
        {
            return GetDatabase().StringSet(arr);
        }
        /// <summary>
        /// 批量存值
        /// </summary>
        /// <param name="keysStr">key</param>
        /// <param name="valuesStr">The value.</param>
        /// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        public bool StringSetMany(string[] keysStr, string[] valuesStr)
        {
            var count = keysStr.Length;
            var keyValuePair = new KeyValuePair<RedisKey, RedisValue>[count];
            for (int i = 0; i < count; i++)
            {
                keyValuePair[i] = new KeyValuePair<RedisKey, RedisValue>(keysStr[i], valuesStr[i]);
            }

            return GetDatabase().StringSet(keyValuePair);
        }

        /// <summary>
        /// 保存一个对象
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="key"></param>
        /// <param name="obj"></param>
        /// <returns></returns>
        public bool SetStringKey<T>(string key, T obj, TimeSpan? expiry = default(TimeSpan?))
        {
            string json = JsonConvert.SerializeObject(obj);
            return GetDatabase().StringSet(key, json, expiry);
        }
        /// <summary>
        /// 追加值
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        public void StringAppend(string key, string value)
        {
            //追加值，返回追加后长度
            long appendlong = GetDatabase().StringAppend(key, value);
        }

        /// <summary>
        /// 获取单个key的值
        /// </summary>
        /// <param name="key">Redis Key</param>
        /// <returns></returns>
        public RedisValue GetStringKey(string key)
        {
            return GetDatabase().StringGet(key);
        }
        /// <summary>
        /// 根据Key获取值
        /// </summary>
        /// <param name="key">键值</param>
        /// <returns>System.String.</returns>
        public string StringGet(string key)
        {
            try
            {
                return GetDatabase().StringGet(key);
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// 获取多个Key
        /// </summary>
        /// <param name="listKey">Redis Key集合</param>
        /// <returns></returns>
        public RedisValue[] GetStringKey(List<RedisKey> listKey)
        {
            return GetDatabase().StringGet(listKey.ToArray());
        }
        /// <summary>
        /// 批量获取值
        /// </summary>
        public string[] StringGetMany(string[] keyStrs)
        {
            var count = keyStrs.Length;
            var keys = new RedisKey[count];
            var addrs = new string[count];

            for (var i = 0; i < count; i++)
            {
                keys[i] = keyStrs[i];
            }
            try
            {

                var values = GetDatabase().StringGet(keys);
                for (var i = 0; i < values.Length; i++)
                {
                    addrs[i] = values[i];
                }
                return addrs;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        /// <summary>
        /// 获取一个key的对象
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="key"></param>
        /// <returns></returns>
        public T GetStringKey<T>(string key)
        {
            return JsonConvert.DeserializeObject<T>(GetDatabase().StringGet(key));
        }

        #endregion

        #region --删除设置过期--
        /// <summary>
        /// 删除单个key
        /// </summary>
        /// <param name="key">redis key</param>
        /// <returns>是否删除成功</returns>
        public bool KeyDelete(string key)
        {
            return GetDatabase().KeyDelete(key);
        }
        /// <summary>
        /// 删除多个key
        /// </summary>
        /// <param name="keys">rediskey</param>
        /// <returns>成功删除的个数</returns>
        public long KeyDelete(RedisKey[] keys)
        {
            return GetDatabase().KeyDelete(keys);
        }
        /// <summary>
        /// 判断key是否存储
        /// </summary>
        /// <param name="key">redis key</param>
        /// <returns></returns>
        public bool KeyExists(string key)
        {
            return GetDatabase().KeyExists(key);
        }
        /// <summary>
        /// 重新命名key
        /// </summary>
        /// <param name="key">就的redis key</param>
        /// <param name="newKey">新的redis key</param>
        /// <returns></returns>
        public bool KeyRename(string key, string newKey)
        {
            return GetDatabase().KeyRename(key, newKey);
        }
        /// <summary>
        /// 删除hasekey
        /// </summary>
        /// <param name="key"></param>
        /// <param name="hashField"></param>
        /// <returns></returns>
        public bool HaseDelete(RedisKey key, RedisValue hashField)
        {
            return GetDatabase().HashDelete(key, hashField);
        }
        /// <summary>
        /// 移除hash中的某值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="key"></param>
        /// <param name="dataKey"></param>
        /// <returns></returns>
        public bool HashRemove(string key, string dataKey)
        {
            return GetDatabase().HashDelete(key, dataKey);
        }
        /// <summary>
        /// 设置缓存过期
        /// </summary>
        /// <param name="key"></param>
        /// <param name="datetime"></param>
        public void SetExpire(string key, DateTime datetime)
        {
            GetDatabase().KeyExpire(key, datetime);
        }
        /// <summary>
        /// 设置永久有效
        /// </summary>
        /// <param name="key"></param>
        /// <param name="datetime"></param>
        public void SetPersist(string key)
        {
            GetDatabase().KeyPersist(key);
        }
        #endregion


        /// <summary>
        /// 关闭
        /// </summary>
        /// <exception cref="NotImplementedException"></exception>
        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}
