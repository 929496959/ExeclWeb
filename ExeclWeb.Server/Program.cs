using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExeclWeb.Core.Common;
using ExeclWeb.Core.ViewModel;
using Fleck;

namespace OnLine.WebSocket
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                FleckLog.Level = LogLevel.Debug;
                var allSockets = new List<IWebSocketConnection>();
                var server = new WebSocketServer("ws://0.0.0.0:26780");
                server.Start(socket =>
                {
                    socket.OnOpen = () =>
                    {
                        Console.WriteLine("Open!");
                        allSockets.Add(socket);
                    };
                    socket.OnClose = () =>
                    {
                        Console.WriteLine("Close!");
                        allSockets.Remove(socket);
                    };
                    socket.OnMessage = message =>
                    {
                        var msg = Util.GzipEncoding(message);
                        var rep = new CellResponseMsg()
                        {
                            createTime = Util.TimeStamp(),
                            data = msg,
                            id = "7a",
                            returnMessage = "success",
                            status = 0,
                            type = 0,
                            username = "aaron"
                        };
                        allSockets.ToList().ForEach(s => s.Send(rep.ToJson()));
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
                Console.WriteLine(e);
            }

            Console.ReadKey();
        }
    }
}

