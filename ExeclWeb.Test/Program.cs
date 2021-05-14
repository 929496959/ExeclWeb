using System;
using System.IO;
using System.Threading.Tasks;

namespace ExeclWeb.Test
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                //JObjectTest.JObject();
                await ProcessTest.Process();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e);
            }
            Console.WriteLine("end...");
            Console.ReadKey();
        }
    }
}
