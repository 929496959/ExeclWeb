﻿using System;
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
                    var gridKey = Common.GetParam(path, "t");
                    //var userid = Common.GetParam(path, "userid");

                    // 会话组信息
                    var group = SessionGroup.FirstOrDefault(p => p.Group == gridKey);
                    socket.OnOpen = () =>
                    {
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
                        //var msg = Common.GzipEncoding(message);
                        if (group != null)
                        {
                            foreach (var item in group.Pools)
                            {
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

