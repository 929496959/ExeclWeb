using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using Common;

namespace ConsoleTest
{
    public class JObjectTest
    {
        public static void JObject()
        {
            try
            {
                var JObject = "[{\"name\":\"Sheet1\",\"index\":\"1\",\"order\":0,\"status\":1,\"row\":54,\"column\":60,\"celldata\":[]},{\"name\":\"Sheet2\",\"index\":\"2\",\"order\":1,\"status\":0,\"row\":54,\"column\":60,\"celldata\":[]},{\"name\":\"Sheet3\",\"index\":\"3\",\"order\":2,\"status\":0,\"row\":54,\"column\":60,\"celldata\":[]}]";
                var jsonObject = JObject.ToObject<JArray>();

                // 查询
                var sheet1 = jsonObject.Value<JObject>(0);
                var sheet2 = jsonObject.FirstOrDefault(p => p.Value<string>("name") == "Sheet2");
                var index = jsonObject.IndexOf(sheet2); // 获取下标
                var sheet3 = jsonObject.Where(p => p.Value<string>("name").Contains("Sheet"));
                Console.WriteLine(sheet1.ToJson());
                Console.WriteLine(sheet2.ToJson());
                Console.WriteLine(sheet3.ToJson());

                // 判断key是否存在
                //var jk = sheet1.Property("jk");
                bool jk = sheet1.ContainsKey("jk");
                if (!jk)
                {
                    //sheet1.Add(new JProperty("jk",100));
                    //sheet1.Property("name");
                }

                // 添加
                sheet1.Add("jo", "jo-ok");
                var obj = new
                {
                    t1 = "t1",
                    t2 = new List<int>() { 1, 2, 3 }
                };
                sheet1.Add("jo2", JToken.FromObject(obj));
                Console.WriteLine(sheet1.ToJson());

                // 修改
                sheet1["jo"] = "jo123";
                Console.WriteLine(sheet1.ToJson());

                // 删除
                sheet1.Remove("jo2");
                Console.WriteLine(sheet1.ToJson());
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e);
            }
        }
    }
}
