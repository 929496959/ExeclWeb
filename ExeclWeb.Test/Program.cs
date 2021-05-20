using System;
using System.IO;
using System.Threading.Tasks;

namespace ExeclWeb.Test
{
    class Program
    {
        static async Task Main(string[] args)
        {
            JObjectTest.JObject();
            //await ProcessTest.Process();

            Console.WriteLine("end...");
            Console.ReadKey();
        }
    }
}
