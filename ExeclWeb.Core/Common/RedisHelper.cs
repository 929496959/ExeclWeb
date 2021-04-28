using System;
using System.Diagnostics;
using Newtonsoft.Json;
using FreeRedis;

namespace ExeclWeb.Core.Common
{
    public static class Redis
    {
        static readonly Lazy<RedisClient> CliLazy = new Lazy<RedisClient>(() =>
        {
            var redisConnection = new ConfigHelper().GetValue<string>("RedisConnection");
            var r = new RedisClient(redisConnection)
            {
                Serialize = JsonConvert.SerializeObject,
                Deserialize = JsonConvert.DeserializeObject
            };
            r.Notice += (s, e) => Trace.WriteLine(e.Log);
            return r;
        });
        public static RedisClient Cli => CliLazy.Value;
    }
    public static class TimeOut
    {
        /// <summary>
        /// 超时时间-1秒
        /// </summary>
        public const int TimeOutSecond = 1;
        /// <summary>
        /// 超时时间-1分钟
        /// </summary>
        public const int TimeOutMinute = TimeOutSecond * 60;
        /// <summary>
        /// 超时时间-1小时
        /// </summary>
        public const int TimeOutHour = TimeOutMinute * 60;
        /// <summary>
        /// 超时时间-1天
        /// </summary>
        public const int TimeOutDay = TimeOutHour * 24;
    }
}
