﻿using System;
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
                var v = "{\"t\":\"v\",\"i\":\"1\",\"v\":{\"v\":99,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"99\"},\"r\":0,\"c\":3}";
                //var rv = "{\"t\":\"rv\",\"i\":\"1\",\"v\":[[{\"v\":1,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"1\"}],[{\"v\":2,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"2\"}],[{\"v\":3,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"3\"}],[{\"v\":4,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"4\"}],[{\"v\":5,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"5\"}]],\"range\":{\"row\":[0,4],\"column\":[0,0]}}";
                //var all = "{\"t\":\"all\",\"i\":\"1\",\"v\":{},\"k\":\"config\"}";
                //var all = "{\"t\":\"all\",\"i\":\"1\",\"v\":[],\"k\":\"luckysheet_conditionformat_save\"}";
                //var cg = "{\"t\":\"cg\",\"i\":\"1\",\"v\":[{\"rangeType\":\"range\",\"borderType\":\"border-outside\",\"color\":\"#000\",\"style\":\"1\",\"range\":[{\"row\":[0,1],\"column\":[0,0],\"row_focus\":0,\"column_focus\":0,\"left\":0,\"width\":73,\"top\":0,\"height\":19,\"left_move\":0,\"width_move\":73,\"top_move\":0,\"height_move\":39}]}],\"k\":\"borderInfo\"}";
                //var fc = "{\"t\":\"fc\",\"i\":\"1\",\"v\":\"{\\\"r\\\":1,\\\"c\\\":1,\\\"index\\\":\\\"0\\\",\\\"func\\\":[true,3,\\\"=sum(A1:B1)\\\"]}\",\"op\":\"add\",\"pos\":0}";
                //var shr = "{\"t\":\"shr\",\"i\":null,\"v\":{\"0\":1,\"1\":0,\"2\":2,\"3\":3,\"4\":4,\"5\":5,\"6\":6,\"7\":7,\"Sheet_6az6nei65t1i_1596209937084\":8}}";
                var requestMsg = v.ToObject<JObject>();

                string type = requestMsg.Value<string>("t");
                switch (type)
                {
                    case "v":
                        // 单个单元格刷新
                        await Operation_V(requestMsg, gridKey);
                        break;
                    case "rv":
                        // 范围单元格刷新
                        await Operation_Rv(requestMsg, gridKey);
                        break;
                    case "cg":
                        // config操作
                        await Operation_Cg(requestMsg, gridKey);
                        break;
                    case "all":
                        // 通用保存
                        await Operation_All(requestMsg, gridKey);
                        break;
                    case "fc":
                        // 函数链操作
                        await Operation_Fc(requestMsg, gridKey);
                        break;
                    case "drc":
                        // 删除行或列
                        await Operation_Drc(requestMsg, gridKey);
                        break;
                    case "arc":
                        // 增加行或列
                        await Operation_Arc(requestMsg, gridKey);
                        break;
                    case "fsc":
                        // 清除筛选
                        await Operation_Fsc(requestMsg, gridKey);
                        break;
                    case "fsr":
                        // 恢复筛选
                        await Operation_Fsr(requestMsg, gridKey);
                        break;
                    case "sha":
                        // 新建sheet
                        await Operation_Sha(requestMsg, gridKey);
                        break;
                    case "shc":
                        // 复制sheet
                        await Operation_Shc(requestMsg, gridKey);
                        break;
                    case "shd":
                        // 删除sheet
                        await Operation_Shd(requestMsg, gridKey);
                        break;
                    case "shre":
                        // 删除sheet后恢复操作
                        await Operation_Shre(requestMsg, gridKey);
                        break;
                    case "shr":
                        // 调整sheet位置
                        await Operation_Shr(requestMsg, gridKey);
                        break;
                    case "shs":
                        // 切换到指定sheet
                        await Operation_Shs(requestMsg, gridKey);
                        break;
                    case "sh":
                        // sheet属性(隐藏或显示)
                        await Operation_Sh(requestMsg, gridKey);
                        break;
                    case "na":
                        // 修改工作簿名称
                        Operation_Na(requestMsg, gridKey);
                        break;
                }
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
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_V(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var v = jObject.Value<JObject>("v");
            var r = jObject.Value<int>("r");
            var c = jObject.Value<int>("c");

            // sheet页
            var sheetModel = await SheetService.GetSheet(gridKey, i);
            var sheet = (JObject)sheetModel.json_data;
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
                // 如果celldata 是空，则添加
                var cell = new JObject()
                {
                    {"r",r},
                    {"v",v},
                    {"v",v}
                };
                cellData.Add(cell);
            }

            // 提交修改
            await SheetService.UpdateSheet(sheetModel.id, gridKey, sheetModel.index, sheet.ToJson(), sheetModel.status, sheetModel.order, sheetModel.is_delete);
        }

        /// <summary>
        /// 范围单元格刷新
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Rv(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var vArray = jObject.Value<JArray>("v");
            var row = jObject.Value<JObject>("range").Value<JArray>("row");
            var column = jObject.Value<JObject>("range").Value<JArray>("column");

            // sheet页
            var sheetModel = await SheetService.GetSheet(gridKey, i);
            var sheet = (JObject)sheetModel.json_data;
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

            // 提交修改
            await SheetService.UpdateSheet(sheetModel.id, gridKey, sheetModel.index, sheet.ToJson(), sheetModel.status, sheetModel.order, sheetModel.is_delete);
        }

        /// <summary>
        /// config操作
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Cg(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var k = jObject.Value<string>("k");
            var v = jObject.Value<JToken>("v");

            // sheet页
            var sheetModel = await SheetService.GetSheet(gridKey, i);
            var sheet = (JObject)sheetModel.json_data;
            //如果不存在，则创建 'config' 节点
            if (!sheet.ContainsKey("config"))
            {
                sheet.Add("config", null);
            }
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

            // 提交修改
            await SheetService.UpdateSheet(sheetModel.id, gridKey, sheetModel.index, sheet.ToJson(), sheetModel.status, sheetModel.order, sheetModel.is_delete);
        }

        /// <summary>
        /// 通用保存
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_All(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var k = jObject.Value<string>("k");
            var v = jObject.Value<JToken>("v");

            // sheet页
            var sheetModel = await SheetService.GetSheet(gridKey, i);
            var sheet = (JObject)sheetModel.json_data;
            bool isExist = sheet.ContainsKey(k);
            if (isExist)
            {
                sheet[k] = v;
            }
            else
            {
                sheet.Add(k, v);
            }

            // 提交修改
            await SheetService.UpdateSheet(sheetModel.id, gridKey, sheetModel.index, sheet.ToJson(), sheetModel.status, sheetModel.order, sheetModel.is_delete);
        }

        /// <summary>
        /// 函数链操作
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Fc(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var v = jObject.Value<JToken>("v");
            var op = jObject.Value<string>("op");
            var pos = jObject.Value<int>("pos");

            // sheet页
            var sheetModel = await SheetService.GetSheet(gridKey, i);
            var sheet = (JObject)sheetModel.json_data;

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

            // 提交修改
            await SheetService.UpdateSheet(sheetModel.id, gridKey, sheetModel.index, sheet.ToJson(), sheetModel.status, sheetModel.order, sheetModel.is_delete);
        }

        /// <summary>
        /// 删除行或列
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Drc(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var rc = jObject.Value<string>("r");// r 行，c 列
            var v = jObject.Value<JObject>("v");
            int vIndex = v.Value<int>("index");
            int vLen = v.Value<int>("len");

            // sheet页
            var sheetModel = await SheetService.GetSheet(gridKey, i);
            var sheet = (JObject)sheetModel.json_data;
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

            // 提交修改
            await SheetService.UpdateSheet(sheetModel.id, gridKey, sheetModel.index, sheet.ToJson(), sheetModel.status, sheetModel.order, sheetModel.is_delete);
        }

        /// <summary>
        /// 增加行或列
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Arc(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var rc = jObject.Value<string>("r");// r 行，c 列
            var v = jObject.Value<JObject>("v");
            int vIndex = v.Value<int>("index");
            int vLen = v.Value<int>("len");
            string vDirection = v.Value<string>("direction");
            JArray vData = v.Value<JArray>("data");

            // sheet页
            var sheetModel = await SheetService.GetSheet(gridKey, i);
            var sheet = (JObject)sheetModel.json_data;
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

            // 提交修改
            await SheetService.UpdateSheet(sheetModel.id, gridKey, sheetModel.index, sheet.ToJson(), sheetModel.status, sheetModel.order, sheetModel.is_delete);
        }

        /// <summary>
        /// 清除筛选
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Fsc(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var v = jObject.Value<JObject>("v");

            // sheet页
            var sheetModel = await SheetService.GetSheet(gridKey, i);
            var sheet = (JObject)sheetModel.json_data;
            if (v == null)
            {
                sheet.Remove("filter");
                sheet.Remove("filter_select");
            }

            // 提交修改
            await SheetService.UpdateSheet(sheetModel.id, gridKey, sheetModel.index, sheet.ToJson(), sheetModel.status, sheetModel.order, sheetModel.is_delete);
        }

        /// <summary>
        /// 恢复筛选
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Fsr(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var v = jObject.Value<JObject>("v");

            // sheet页
            var sheetModel = await SheetService.GetSheet(gridKey, i);
            var sheet = (JObject)sheetModel.json_data;
            if (v != null)
            {
                sheet["filter"] = v["filter"];
                sheet["filter_select"] = v["filter_select"];
            }

            // 提交修改
            await SheetService.UpdateSheet(sheetModel.id, gridKey, sheetModel.index, sheet.ToJson(), sheetModel.status, sheetModel.order, sheetModel.is_delete);
        }

        /// <summary>
        /// 新建sheet
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Sha(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var v = jObject.Value<JObject>("v");

            string index = v.Value<string>("index");
            int status = v.Value<int>("status");
            int order = v.Value<int>("order");

            // 提交修改
            await SheetService.AddSheet(gridKey, index, v.ToJson(), status, order, 0);
        }

        /// <summary>
        /// 复制sheet
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Shc(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var v = jObject.Value<JObject>("v");
            var vCopyIndex = v.Value<int>("copyindex");
            var vName = v.Value<string>("name");

            // sheet页
            var sheetList = await SheetService.GetSheets(gridKey);
            var sheetModel = sheetList.ToList()[vCopyIndex];
            var sheet = (JObject)sheetModel.json_data;
            sheet["index"] = i;
            sheet["name"] = vName;

            // 提交修改
            await SheetService.AddSheet(gridKey, i, sheet.ToJson(), sheetModel.status, sheetModel.order, sheetModel.is_delete);
        }

        /// <summary>
        /// 删除sheet
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Shd(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var v = jObject.Value<JObject>("v");
            var vDeleIndex = v.Value<int>("deleIndex");

            // sheet页
            var sheetList = await SheetService.GetSheets(gridKey);
            var sheetModel = sheetList.ToList()[vDeleIndex];

            // 提交修改
            await SheetService.UpdateIsDelete(sheetModel.id, 1);
        }

        /// <summary>
        /// 删除sheet后恢复
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Shre(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var v = jObject.Value<JObject>("v");
            var vReIndex = v.Value<int>("reIndex");

            // sheet页
            var sheetList = await SheetService.GetSheets(gridKey);
            var sheetModel = sheetList.ToList()[vReIndex];

            // 提交修改
            await SheetService.UpdateIsDelete(sheetModel.id, 0);
        }

        /// <summary>
        /// 调整sheet位置
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Shr(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var v = jObject.Value<JObject>("v");

            // sheet页
            var sheetList = await SheetService.GetSheets(gridKey);
            foreach (var item in v)
            {
                var sheetModel = sheetList.FirstOrDefault(p => p.index == item.Key);
                if (sheetModel != null)
                {
                    var sheet = (JObject)sheetModel.json_data;
                    sheet["order"] = item.Value;
                    await SheetService.UpdateSheet(sheetModel.id, gridKey, sheetModel.index, sheet.ToJson(), sheetModel.status, Convert.ToInt32(item.Value), sheetModel.is_delete);
                }
            }
        }

        /// <summary>
        /// 切换到指定sheet
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Shs(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var v = jObject.Value<int>("v");

            // sheet页
            var sheetList = await SheetService.GetSheets(gridKey);
            foreach (var item in sheetList)
            {
                var vSheet = sheetList.ToList()[v];
                if (vSheet == null) break;
                await SheetService.UpdateStatus(vSheet.id, item.index == vSheet.index ? 1 : 0);
            }
        }

        /// <summary>
        /// sheet属性(隐藏或显示)
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static async Task Operation_Sh(JObject requestMsg, string gridKey)
        {
            // 请求信息
            var jObject = requestMsg.Value<JObject>();
            var i = jObject.Value<string>("i");
            var v = jObject.Value<int>("v");
            var op = jObject.Value<string>("op");
            var cur = jObject.Value<string>("cur");

            // sheet页
            var sheetModel = await SheetService.GetSheet(gridKey, i);
            var sheet = (JObject)sheetModel.json_data;

            var curSheetModel = await SheetService.GetSheet(gridKey, cur);
            var curSheet = (JObject)curSheetModel.json_data;
            if (op == "hide")
            {
                // 隐藏
                sheet["hide "] = 1;
                sheet["status"] = 0;
                curSheet["status"] = 1;
            }
            else
            {
                // 显示
                sheet["hide "] = 0;
                sheet["status"] = 1;
                curSheet["status"] = 0;
            }
            // 提交修改
            await SheetService.UpdateSheet(sheetModel.id, gridKey, sheetModel.index, sheet.ToJson(), sheetModel.status, sheetModel.order, sheetModel.is_delete);
            await SheetService.UpdateSheet(curSheetModel.id, gridKey, curSheetModel.index, curSheet.ToJson(), curSheetModel.status, curSheetModel.order, curSheetModel.is_delete);
        }

        /// <summary>
        /// sheet属性(隐藏或显示)
        /// </summary>
        /// <param name="requestMsg">请求信息</param>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        private static void Operation_Na(JObject requestMsg, string gridKey)
        {

        }
    }
}
