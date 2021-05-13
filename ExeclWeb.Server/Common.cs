using System;
using System.Text;
using System.Web;
using ExeclWeb.Core.Common;

namespace ExeclWeb.Server
{
    public class Common
    {
        /// <summary>
        /// 客户端Gzip压缩数据解压
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string GzipEncoding(string text)
        {
            try
            {
                var isoBytes = Encoding.GetEncoding("ISO-8859-1").GetBytes(text);
                var gZipBytes = SharpZipHelper.GZipDeCompress(isoBytes);
                text = HttpUtility.UrlDecode(Encoding.Default.GetString(gZipBytes));
                return text;
            }
            catch (Exception e)
            {
                NLogger.Error(e);
            }
            return null;
        }

        /// <summary>
        /// 获取参数
        /// </summary>
        /// <param name="query">请求Path</param>
        /// <param name="key">参数key</param>
        /// <returns></returns>
        public static string GetQueryParam(string query, string key)
        {
            if (string.IsNullOrWhiteSpace(query)) return null;

            query = query.Split("/execlws?")[1];
            var array = query.Split("&");
            for (int i = 0; i < array.Length; i++)
            {
                var item = array[i].Split("=");
                if (item[0].Trim().ToLower() == key.Trim().ToLower())
                {
                    return item[1];
                }
            }
            return null;
        }

        /// <summary>
        /// 获取时间戳
        /// </summary>
        /// <returns></returns>
        public static string TimeStamp()
        {
            return new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds().ToString();
        }

    }
}
