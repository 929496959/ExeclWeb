using System.IO;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;

namespace ExeclWeb.Core.Common
{
    /// <summary>
    /// Config帮助类
    /// </summary>
    public class ConfigHelper
    {
        private readonly string _fileName;
        public ConfigHelper(string fileName = "appsettings.json")
        {
            _fileName = fileName;
        }

        public IConfiguration BuildConfig()
        {
            var builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile(_fileName).Build();
            return builder;
        }

        /// <summary>
        /// 根据Key获取T
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="key"></param>
        /// <returns></returns>
        public T GetSettings<T>(string key) where T : class, new()
        {
            var builder = BuildConfig();
            var config = new ServiceCollection().AddOptions().Configure<T>(builder.GetSection(key)).BuildServiceProvider().GetService<IOptions<T>>().Value;
            return config;
        }

        /// <summary>
        /// 根据Ket获取Value
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public T GetValue<T>(string key)
        {
            var builder = BuildConfig();
            return builder.GetValue<T>(key);
        }
    }
}
