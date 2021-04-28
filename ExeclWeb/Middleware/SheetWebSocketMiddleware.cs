using System;
using System.Collections.Concurrent;
using System.IO;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Common;
using Data.ViewModel;
using ExeclWeb.Common;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json.Linq;
using OnLine.Web.Common;

namespace ExeclWeb.Middleware
{
    public class SheetWebSocketMiddleware
    {
        private static readonly SheetWebSocket SheetWebSocket = new SheetWebSocket();
        private static readonly ConcurrentDictionary<string, WebSocket> Sockets = new ConcurrentDictionary<string, WebSocket>();
        private readonly RequestDelegate _next;

        public SheetWebSocketMiddleware(RequestDelegate next)
        {
            _next = next;
        }

        public async Task Invoke(HttpContext context)
        {
            if (!context.WebSockets.IsWebSocketRequest)
            {
                await _next.Invoke(context);
                return;
            }
            //WebSocket dummy;

            CancellationToken ct = context.RequestAborted;
            var currentSocket = await context.WebSockets.AcceptWebSocketAsync();

            // 参数获取

            //string socketId = Guid.NewGuid().ToString();
            string socketId = context.Request.Query["userid"].ToString();
            string gridKey = context.Request.Query["gridkey"].ToString();

            if (!Sockets.ContainsKey(socketId))
            {
                Sockets.TryAdd(socketId, currentSocket);
            }

            //_sockets.TryRemove(socketId, out dummy);
            //_sockets.TryAdd(socketId, currentSocket);

            while (true)
            {
                if (ct.IsCancellationRequested)
                {
                    break;
                }
                // 接收并解压数据
                string receive = await ReceiveStringAsync(currentSocket, ct);

                // 如果是客户端发来的探针，跳过
                if (receive.Equals("rub"))
                {
                    continue;
                }

                // 解压数据
                var decompress = Util.GzipEncoding(receive);
                var jObject = decompress.ToObject<JObject>();
                if (jObject == null || jObject.Value<string>("t").Equals("mv"))
                {
                    continue;
                }
                NLogger.Info(jObject.ToJson());
                //await SheetWebSocket.Process(jObject, gridKey);

                if (string.IsNullOrEmpty(receive))
                {
                    if (currentSocket.State != WebSocketState.Open)
                    {
                        break;
                    }
                    continue;
                }

                foreach (var socket in Sockets)
                {
                    if (socket.Value.State != WebSocketState.Open)
                    {
                        continue;
                    }

                    // 返回数据
                    var repMsg = new CellResponseMsg()
                    {
                        createTime = Util.TimeStamp(),
                        data = decompress,
                        id = "7a",
                        returnMessage = "success",
                        status = 0,
                        type = 1,
                        username = "Aaron"
                    };
                    await SendStringAsync(socket.Value, repMsg.ToJson(), ct);

                    //if (socket.Key == msg.ReceiverID || socket.Key == socketId)
                    //{
                    //    await SendStringAsync(socket.Value, JsonConvert.SerializeObject(msg), ct);
                    //}
                }
            }

            //_sockets.TryRemove(socketId, out dummy);

            await currentSocket.CloseAsync(WebSocketCloseStatus.NormalClosure, "Closing", ct);
            currentSocket.Dispose();
        }

        /// <summary>
        /// 发送
        /// </summary>
        /// <param name="socket"></param>
        /// <param name="data"></param>
        /// <param name="ct"></param>
        /// <returns></returns>
        private static Task SendStringAsync(WebSocket socket, string data, CancellationToken ct = default(CancellationToken))
        {
            var buffer = Encoding.UTF8.GetBytes(data);
            var segment = new ArraySegment<byte>(buffer);
            return socket.SendAsync(segment, WebSocketMessageType.Text, true, ct);
        }

        /// <summary>
        /// 接收
        /// </summary>
        /// <param name="socket"></param>
        /// <param name="ct"></param>
        /// <returns></returns>
        private static async Task<string> ReceiveStringAsync(WebSocket socket, CancellationToken ct = default(CancellationToken))
        {
            var buffer = new ArraySegment<byte>(new byte[8192]);
            using (var ms = new MemoryStream())
            {
                WebSocketReceiveResult result;
                do
                {
                    ct.ThrowIfCancellationRequested();

                    result = await socket.ReceiveAsync(buffer, ct);
                    //ms.Write(buffer.Array, buffer.Offset, result.Count);
                    await ms.WriteAsync(buffer.Array, buffer.Offset, result.Count, ct);
                }
                while (!result.EndOfMessage);

                ms.Seek(0, SeekOrigin.Begin);
                if (result.MessageType != WebSocketMessageType.Text)
                {
                    return null;
                }

                using (var reader = new StreamReader(ms, Encoding.UTF8))
                {
                    return await reader.ReadToEndAsync();
                }
            }
        }
    }
}
