using System;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using ExeclWeb.Core.Application;
using ExeclWeb.Core.Common;

namespace ExeclWeb.Test
{
    public class ProcessTest
    {
        private static readonly SheetService SheetService = new SheetService();

        public static async Task Process()
        {
            var gridKey = "9BC4E24A-D545-4EEF-4680-050ED4C3BF54";
            try
            {
                // 客户端数据
                //var v = "{\"t\":\"v\",\"i\":\"1\",\"v\":{\"v\":99,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"99\"},\"r\":0,\"c\":3}";
                //var rv = "{\"t\":\"rv\",\"i\":\"1\",\"v\":[[{\"v\":1,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"1\"}],[{\"v\":2,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"2\"}],[{\"v\":3,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"3\"}],[{\"v\":4,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"4\"}],[{\"v\":5,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"5\"}]],\"range\":{\"row\":[0,4],\"column\":[0,0]}}";
                //var all = "{\"t\":\"all\",\"i\":\"1\",\"v\":{},\"k\":\"config\"}";
                //var all = "{\"t\":\"all\",\"i\":\"1\",\"v\":[],\"k\":\"luckysheet_conditionformat_save\"}";
                //var cg = "{\"t\":\"cg\",\"i\":\"1\",\"v\":[{\"rangeType\":\"range\",\"borderType\":\"border-outside\",\"color\":\"#000\",\"style\":\"1\",\"range\":[{\"row\":[0,1],\"column\":[0,0],\"row_focus\":0,\"column_focus\":0,\"left\":0,\"width\":73,\"top\":0,\"height\":19,\"left_move\":0,\"width_move\":73,\"top_move\":0,\"height_move\":39}]}],\"k\":\"borderInfo\"}";
                var fc = "{\"t\":\"fc\",\"i\":\"1\",\"v\":\"{\\\"r\\\":1,\\\"c\\\":1,\\\"index\\\":\\\"0\\\",\\\"func\\\":[true,3,\\\"=sum(A1:B1)\\\"]}\",\"op\":\"add\",\"pos\":0}";
                var requestMsg = fc.ToObject<JObject>();

                // 工作博数据
                //var sheetData = await OnLineSheetService.LoadSheetData(gridKey);
                //var sheetJsonData = "[{\"name\":\"Sheet1\",\"index\":\"1\",\"order\":0,\"status\":1,\"row\":54,\"column\":60,\"celldata\":[]},{\"name\":\"Sheet2\",\"index\":\"2\",\"order\":1,\"status\":0,\"row\":54,\"column\":60,\"celldata\":[]},{\"name\":\"Sheet3\",\"index\":\"3\",\"order\":2,\"status\":0,\"row\":54,\"column\":60,\"celldata\":[]}]";
                var sheetJsonData = "[{\"name\":\"Sheet1\",\"index\":\"1\",\"order\":0,\"status\":1,\"row\":84,\"column\":60,\"config\":{},\"celldata\":[{\"v\":1,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"1\",\"r\":0,\"c\":0},{\"v\":2,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"2\",\"r\":1,\"c\":0},{\"v\":3,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"3\",\"r\":2,\"c\":0},{\"v\":4,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"4\",\"r\":3,\"c\":0},{\"v\":5,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"5\",\"r\":4,\"c\":0}],\"calcChain\":[\"{\\\"r\\\":1,\\\"c\\\":1,\\\"index\\\":\\\"0\\\",\\\"func\\\":[true,3,\\\"=sum(A1:B1)\\\"]}\"]},{\"name\":\"Sheet2\",\"index\":\"2\",\"order\":1,\"status\":0,\"row\":84,\"column\":60,\"config\":{},\"celldata\":[]},{\"name\":\"Sheet3\",\"index\":\"3\",\"order\":2,\"status\":0,\"row\":84,\"column\":60,\"config\":{},\"celldata\":[]}]";

                string data;
                string type = requestMsg.Value<string>("t");
                switch (type)
                {
                    case "v":
                        // 单个单元格刷新
                        data = Operation_V(requestMsg, sheetJsonData);
                        break;
                    case "rv":
                        // 范围单元格刷新
                        data = Operation_Rv(requestMsg, sheetJsonData);
                        break;
                    case "cg":
                        // config操作
                        data = Operation_Cg(requestMsg, sheetJsonData);
                        break;
                    case "all":
                        // 通用保存
                        data = Operation_All(requestMsg, sheetJsonData);
                        break;
                    case "fc":
                        // 函数链操作
                        data = Operation_Fc(requestMsg, sheetJsonData);
                        break;
                    case "drc":
                        // 删除行或列
                        data = Operation_Drc(requestMsg, sheetJsonData);
                        break;
                    case "arc":
                        // 增加行或列
                        data = Operation_Arc(requestMsg, sheetJsonData);
                        break;
                    case "fsc":
                        // 清除筛选
                        data = Operation_Fsc(requestMsg, sheetJsonData);
                        break;
                    case "fsr":
                        // 恢复筛选
                        data = Operation_Fsr(requestMsg, sheetJsonData);
                        break;
                    case "sha":
                        // 新建sheet
                        break;
                    case "shs":
                        // 切换到指定sheet
                        break;
                    case "shc":
                        // 复制sheet
                        break;
                    case "na":
                        // 修改工作簿名称
                        break;
                    case "shd":
                        // 删除sheet
                        break;
                    case "shre":
                        // 删除sheet后恢复操作
                        break;
                    case "shr":
                        // 调整sheet位置
                        break;
                    case "sh":
                        // sheet属性(隐藏或显示)
                        break;
                    default:
                        break;
                }
                // 处理最新sheet json_data数据

            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e);
            }
        }

        /// <summary>
        /// 单个单元格刷新
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="sheetJsonData">工作薄json_data</param>
        /// <returns></returns>
        public static string Operation_V(JObject requestMsg, string sheetJsonData)
        {
            // 获取请求信息
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");//当前sheet的index值
            var v = jObject.Value<JObject>("v");//单元格的值
            var r = jObject.Value<int>("r");//单元格的行号
            var c = jObject.Value<int>("c");//单元格的列号

            // 获取工作薄
            var jArray = sheetJsonData.ToObject<JArray>();
            var sheet = (JObject)jArray.FirstOrDefault(p => p.Value<string>("index") == i);
            var cellData = sheet.Value<JArray>("celldata");
            if (cellData.Count > 0)
            {
                // 有单元格处理
                var item = (JObject)cellData.FirstOrDefault(p => p.Value<int>("r") == r && p.Value<int>("c") == c);
                if (item != null)
                {
                    var index = cellData.IndexOf(item);
                    var cell = new JObject()
                    {
                        {"r",r},
                        {"v",v},
                        {"v",v}
                    };
                    cellData[index] = cell;
                    // 如果该单元格的 v 是null，删除该单元格
                    if (item.Value<JObject>("v") == null)
                    {
                        cellData.Remove(index);
                    }
                }
            }
            else
            {
                // 如果 celldata 是空，则添加
                var cell = new JObject()
                {
                    {"r",r},
                    {"v",v},
                    {"v",v}
                };
                cellData.Add(cell);
            }
            return jArray.ToJson();
        }

        /// <summary>
        /// 范围单元格刷新
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="sheetJsonData">工作薄json_data</param>
        /// <returns></returns>
        public static string Operation_Rv(JObject requestMsg, string sheetJsonData)
        {
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");//当前sheet的index值
            var vArray = jObject.Value<JArray>("v");//二维数组
            var row = jObject.Value<JObject>("range").Value<JArray>("row");
            var column = jObject.Value<JObject>("range").Value<JArray>("column");

            var jArray = sheetJsonData.ToObject<JArray>();
            var sheet = (JObject)jArray.FirstOrDefault(p => p.Value<string>("index") == i);
            var cellData = sheet.Value<JArray>("celldata");
            //遍历行列，对符合行列的内容进行更新
            for (int ri = row.Value<int>(0); ri <= row.Value<int>(1); ri++)
            {
                for (int ci = column.Value<int>(0); ci <= column.Value<int>(1); ci++)
                {
                    var v = (JObject)vArray[ri][ci];
                    v.Add("r", ri);
                    v.Add("c", ci);
                    var cell = cellData.FirstOrDefault(p => p.Value<int>("r") == ri && p.Value<int>("c") == ci);
                    if (cell == null)
                    {
                        cellData.Add(v);
                    }
                    else
                    {
                        cellData[cellData.IndexOf(cell)] = v;
                    }
                }
            }
            return jArray.ToJson();
        }

        /// <summary>
        /// config操作
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="sheetJsonData">工作薄json_data</param>
        /// <returns></returns>
        public static string Operation_Cg(JObject requestMsg, string sheetJsonData)
        {
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var k = jObject.Value<string>("k");
            var v = jObject.Value<JToken>("v");

            var jArray = sheetJsonData.ToObject<JArray>();
            var sheet = (JObject)jArray.FirstOrDefault(p => p.Value<string>("index") == i);
            var config = sheet.Value<JObject>("config");
            bool isExist = config.ContainsKey(k);
            if (isExist)
            {
                config[k] = v;
            }
            else
            {
                if (v.GetType() == typeof(JObject))
                {
                    config.Add(k, new JObject());
                }
                if (v.GetType() == typeof(JObject))
                {
                    config.Add(k, new JArray());
                }
            }

            return jArray.ToJson();
        }

        /// <summary>
        /// 通用保存
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="sheetJsonData">工作薄json_data</param>
        /// <returns></returns>
        public static string Operation_All(JObject requestMsg, string sheetJsonData)
        {
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var k = jObject.Value<string>("k");
            var v = jObject.Value<JToken>("v");

            var jArray = sheetJsonData.ToObject<JArray>();
            var sheet = (JObject)jArray.FirstOrDefault(p => p.Value<string>("index") == i);
            bool isExist = sheet.ContainsKey(k);
            if (isExist)
            {
                sheet[k] = v;
            }
            else
            {
                sheet.Add(k, v);
            }

            return jArray.ToJson();
        }

        /// <summary>
        /// 函数链操作
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="sheetJsonData">工作薄json_data</param>
        /// <returns></returns>
        public static string Operation_Fc(JObject requestMsg, string sheetJsonData)
        {
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var v = jObject.Value<JToken>("v");
            var op = jObject.Value<string>("op");
            var pos = jObject.Value<int>("pos");

            var jArray = sheetJsonData.ToObject<JArray>();
            var sheet = (JObject)jArray.FirstOrDefault(p => p.Value<string>("index") == i);

            // 判断公式链节点是否存在
            if (sheet != null && !sheet.ContainsKey("calcChain"))
            {
                sheet.Add("calcChain", new JArray());
            }
            // op处理
            if (op == "add")
            {
                sheet.Value<JArray>("calcChain").Add(v);
            }
            if (op == "update")
            {
                sheet["calcChain"][pos] = v;
            }
            if (op == "del")
            {
                sheet.Value<JArray>("calcChain").RemoveAt(pos);
            }

            return jArray.ToJson();
        }

        /// <summary>
        /// 删除行或列
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="sheetJsonData">工作薄json_data</param>
        /// <returns></returns>
        public static string Operation_Drc(JObject requestMsg, string sheetJsonData)
        {
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var rc = jObject.Value<string>("r");// r 行，c 列
            var v = jObject.Value<JObject>("v");
            int vIndex = v.Value<int>("index");
            int vLen = v.Value<int>("len");

            var jArray = sheetJsonData.ToObject<JArray>();
            var sheet = (JObject)jArray.FirstOrDefault(p => p.Value<string>("index") == i);
            var cellData = sheet.Value<JArray>("celldata");
            foreach (var item in cellData)
            {
                var cell = (JObject)item;
                if (rc == "r")
                {
                    //删除行所在区域的内容
                    if (cell.Value<int>("r") >= vIndex && cell.Value<int>("r") < vIndex + vLen)
                    {
                        cellData.Remove(cell);
                    }
                    //增加大于,最大删除行的的行号
                    if (cell.Value<int>("r") >= vIndex + vLen)
                    {
                        cellData.Remove(cell);
                        cell["r"] = cell.Value<int>("r") - vLen;
                        cellData.Add(cell);
                    }
                }
                else
                {
                    //删除列所在区域的内容
                    if (cell.Value<int>("c") >= vIndex && cell.Value<int>("c") < vIndex + vLen)
                    {
                        cellData.Remove(cell);
                    }
                    //增加大于,最大删除列的的列号
                    if (cell.Value<int>("c") >= vIndex + vLen)
                    {
                        cellData.Remove(cell);
                        cell["c"] = cell.Value<int>("c") - vLen;
                        cellData.Add(cell);
                    }
                }
            }
            if (rc == "r")
            {
                sheet["row"] = sheet.Value<int>("row") - vLen;
            }
            else
            {
                sheet["column"] = sheet.Value<int>("column") - vLen;
            }
            return jArray.ToJson();
        }

        /// <summary>
        /// 增加行或列
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="sheetJsonData">工作薄json_data</param>
        /// <returns></returns>
        public static string Operation_Arc(JObject requestMsg, string sheetJsonData)
        {
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var rc = jObject.Value<string>("r");// r 行，c 列
            var v = jObject.Value<JObject>("v");
            int vIndex = v.Value<int>("index");
            int vLen = v.Value<int>("len");
            string vDirection = v.Value<string>("direction");
            JArray vData = v.Value<JArray>("data");

            var jArray = sheetJsonData.ToObject<JArray>();
            var sheet = (JObject)jArray.FirstOrDefault(p => p.Value<string>("index") == i);
            var cellData = sheet.Value<JArray>("celldata");
            foreach (var item in cellData)
            {
                var cell = (JObject)item;
                if (rc == "r")
                {
                    //如果是增加行，且是向左增加
                    if (cell.Value<int>("r") >= vIndex && vDirection == "lefttop")
                    {
                        cellData.Remove(cell);
                        cell["r"] = cell.Value<int>("r") + vLen;
                        cellData.Add(cell);
                    }
                    //如果是增加行，且是向右增加
                    if (cell.Value<int>("r") > vIndex && vDirection == "rightbottom")
                    {
                        cellData.Remove(cell);
                        cell["r"] = cell.Value<int>("r") + vLen;
                        cellData.Add(cell);
                    }
                }
                else
                {
                    //如果是增加列，且是向上增加
                    if (cell.Value<int>("c") >= vIndex && vDirection == "lefttop")
                    {
                        cellData.Remove(cell);
                        cell["c"] = cell.Value<int>("c") + vLen;
                        cellData.Add(cell);
                    }
                    //如果是增加列，且是向下增加
                    if (cell.Value<int>("c") > vIndex && vDirection == "rightbottom")
                    {
                        cellData.Remove(cell);
                        cell["c"] = cell.Value<int>("c") + vLen;
                        cellData.Add(cell);
                    }
                }
            }
            if (rc == "r")
            {
                sheet["row"] = sheet.Value<int>("row") + vLen;
                for (int r = 0; r < vData.Count; r++)
                {
                    for (int c = 0; c < vData[0].Count(); c++)
                    {
                        var newV = vData[r][c];
                        if (newV == null)
                        {
                            continue;
                        }
                        var newCell = new JObject { { "v", newV }, { "r", r + vIndex }, { "c", c } };
                        cellData.Add(newCell);
                    }
                }
            }
            else
            {
                sheet["column"] = sheet.Value<int>("column") + vLen;
                for (int r = 0; r < vData.Count; r++)
                {
                    for (int c = 0; c < vData[0].Count(); c++)
                    {
                        var newV = vData[r][c];
                        if (newV == null)
                        {
                            continue;
                        }
                        var newCell = new JObject { { "v", newV }, { "r", r }, { "c", c + vIndex } };
                        cellData.Add(newCell);
                    }
                }
            }
            return jArray.ToJson();
        }

        /// <summary>
        /// 清除筛选
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="sheetJsonData">工作薄json_data</param>
        /// <returns></returns>
        public static string Operation_Fsc(JObject requestMsg, string sheetJsonData)
        {
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var v = jObject.Value<JObject>("v");

            var jArray = sheetJsonData.ToObject<JArray>();
            var sheet = (JObject)jArray.FirstOrDefault(p => p.Value<string>("index") == i);
            if (v == null)
            {
                sheet.Remove("filter");
                sheet.Remove("filter_select");
            }

            return jArray.ToJson();
        }

        /// <summary>
        /// 恢复筛选
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="sheetJsonData">工作薄json_data</param>
        /// <returns></returns>
        public static string Operation_Fsr(JObject requestMsg, string sheetJsonData)
        {
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var v = jObject.Value<JObject>("v");

            var jArray = sheetJsonData.ToObject<JArray>();
            var sheet = (JObject)jArray.FirstOrDefault(p => p.Value<string>("index") == i);
            if (v != null)
            {
                sheet["filter"] = v["filter"];
                sheet["filter_select"] = v["filter_select"];
            }

            return jArray.ToJson();
        }
    }
}
