using System.Collections.Generic;
using Fleck;

namespace ExeclWeb.Server
{
    /// <summary>
    /// 会话组
    /// </summary>
    public class SessionGroup
    {
        public string Group { get; set; }
        public List<WebSocketConnectionPool> Pools { get; set; }
    }
    /// <summary>
    /// 连接池
    /// </summary>
    public class WebSocketConnectionPool
    {
        /// <summary>
        /// 会话Id
        /// </summary>
        public string SessionId { get; set; }
        /// <summary>
        /// 会话连接池
        /// </summary>
        public IWebSocketConnection WebSocketConnection { get; set; }
    }
}
