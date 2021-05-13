using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Collections.Generic;
using ExeclWeb.Core.ViewModel;
using ExeclWeb.Core.Common;
using Fleck;

namespace ExeclWeb.Server
{
    class Program
    {
        private static readonly List<SessionGroup> SessionGroup = new List<SessionGroup>();
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
                    var gridKey = Common.GetQueryParam(path, "t");
                    //var userid = Common.GetQueryParam(path, "userid");

                    socket.OnOpen = () =>
                    {
                        var group = SessionGroup.FirstOrDefault(p => p.Group == gridKey);
                        if (group != null)
                        {
                            // 添加会话
                            var pools = group.Pools;
                            if (!pools.Exists(p => p.SessionId == sessionId))
                            {
                                var pool = new WebSocketConnectionPool()
                                {
                                    SessionId = sessionId,
                                    WebSocketConnection = socket
                                };
                                pools.Add(pool);
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
                            Console.WriteLine($"{sessionId} close connection...");
                        }
                    };
                    socket.OnMessage = message =>
                    {
                        //var msg = Common.GzipEncoding(message);
                        var rep = new CellResponseMsg()
                        {
                            createTime = Common.TimeStamp(),
                            data = message,
                            id = "7a",
                            returnMessage = "success",
                            status = 0,
                            type = 0,
                            username = "aaron"
                        };
                        //allSockets.ToList().ForEach(s => s.Send(rep.ToJson()));
                    };
                });

                //var input = Console.ReadLine();
                //while (input != "exit")
                //{
                //    foreach (var socket in allSockets.ToList())
                //    {
                //        socket.Send(input);
                //    }
                //    input = Console.ReadLine();
                //}
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

