using System;
using System.Text;
using System.Web;
using Common;

namespace OnLine.Web.Common
{
    public static class Util
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
        /// 获取时间戳
        /// </summary>
        /// <returns></returns>
        public static string TimeStamp()
        {
            return new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds().ToString();
        }
    }
}
