using System;
using System.Linq;
using System.Collections.Generic;
using ExeclWeb.Core.ViewModel;
using ExeclWeb.Core.Common;
using Fleck;
using Newtonsoft.Json.Linq;

namespace ExeclWeb.Server
{
    class Program
    {
        private static readonly List<SessionGroup> SessionGroup = new List<SessionGroup>();
        private static readonly SheetProcess SheetProcess = new SheetProcess();

        static void Main(string[] args)
        {
            try
            {
                FleckLog.Level = LogLevel.Debug;
                var server = new WebSocketServer("ws://0.0.0.0:26780");
                server.Start(socket =>
                {
                    // 会话Id
                    var sessionId = socket.ConnectionInfo.Id.ToString();
                    var path = socket.ConnectionInfo.Path;
                    var gridKey = Common.GetParam(path, "g");
                    var userId = Common.GetParam(path, "q");

                    socket.OnOpen = () =>
                    {
                        var group = SessionGroup.FirstOrDefault(p => p.Group == gridKey);
                        if (group != null)
                        {
                            // 添加会话
                            if (!group.Pools.Exists(p => p.SessionId == sessionId))
                            {
                                var pool = new WebSocketConnectionPool()
                                {
                                    SessionId = sessionId,
                                    WebSocketConnection = socket
                                };
                                group.Pools.Add(pool);
                            }
                        }
                        else
                        {
                            // 添加会话组
                            var groupModel = new SessionGroup()
                            {
                                Group = gridKey,
                                Pools = new List<WebSocketConnectionPool>()
                                {
                                    new WebSocketConnectionPool()
                                    {
                                        SessionId = sessionId,
                                        WebSocketConnection=socket
                                    }
                                }
                            };
                            SessionGroup.Add(groupModel);
                        }
                        Console.WriteLine($"{sessionId} connection success...");
                    };
                    socket.OnClose = () =>
                    {
                        var group = SessionGroup.FirstOrDefault(p => p.Group == gridKey);
                        if (group != null)
                        {
                            var pool = group.Pools.FirstOrDefault(p => p.SessionId == sessionId);
                            group.Pools.Remove(pool);
                            // 如果会话全部关闭，则移除会话组
                            if (!group.Pools.Any())
                            {
                                SessionGroup.Remove(group);
                            }
                            Console.WriteLine($"{sessionId} close connection...");
                        }
                    };
                    socket.OnMessage = message =>
                    {
                        var requestData = Common.GzipEncoding(message);
                        NLogger.Info(requestData);
                        // 单元格操作
                        var requestJObject = requestData.ToObject<JObject>();
                        SheetProcess.Process(requestJObject, gridKey).Wait();
                        Console.WriteLine(requestData);
                        var group = SessionGroup.FirstOrDefault(p => p.Group == gridKey);
                        if (group != null)
                        {
                            foreach (var item in group.Pools)
                            {
                                var rep = new CellResponseMsg()
                                {
                                    createTime = Common.TimeStamp(),
                                    data = requestData,
                                    id = sessionId,
                                    returnMessage = "success",
                                    status = 0,
                                    type = requestJObject.Value<string>("t") == "mv" ? 3 : 2,
                                    username = sessionId
                                };
                                item.WebSocketConnection.Send(rep.ToJson());
                            }
                        }
                    };
                });
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e);
            }

            Console.ReadKey();
        }
    }
}

