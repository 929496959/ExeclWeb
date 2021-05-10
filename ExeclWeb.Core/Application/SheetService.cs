using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using ExeclWeb.Core.Repository;
using ExeclWeb.Core.Entitys;
using ExeclWeb.Core.Common;

namespace ExeclWeb.Core.Application
{
    public class SheetService
    {
        private readonly SheetRepository _sheetRepository;
        public SheetService()
        {
            _sheetRepository = new SheetRepository();
        }

        /// <summary>
        /// 判断execl文档是否存在
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <returns></returns>
        public async Task<bool> IsExist(string gridKey)
        {
            return await _sheetRepository.IsExist(gridKey);
        }

        /// <summary>
        /// 初始化execl文档sheet数据
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        public async Task InitSheet(string gridKey)
        {
            string defaultSheet = "[{\"name\":\"Sheet1\",\"chart\":[],\"color\":\"\",\"index\":\"1\",\"order\":0,\"row\":84,\"column\":60,\"config\":{},\"status\":1,\"celldata\":[],\"ch_width\":4748,\"rowsplit\":[],\"rh_height\":1790,\"scrollTop\":0,\"scrollLeft\":0,\"visibledatarow\":[],\"visibledatacolumn\":[],\"jfgird_select_save\":[],\"jfgrid_selection_range\":{}},{\"name\":\"Sheet2\",\"chart\":[],\"color\":\"\",\"index\":\"2\",\"order\":1,\"row\":84,\"column\":60,\"config\":{},\"status\":0,\"celldata\":[],\"ch_width\":4748,\"rowsplit\":[],\"rh_height\":1790,\"scrollTop\":0,\"scrollLeft\":0,\"visibledatarow\":[],\"visibledatacolumn\":[],\"jfgird_select_save\":[],\"jfgrid_selection_range\":{}},{\"name\":\"Sheet3\",\"chart\":[],\"color\":\"\",\"index\":\"3\",\"order\":2,\"row\":84,\"column\":60,\"config\":{},\"status\":0,\"celldata\":[],\"ch_width\":4748,\"rowsplit\":[],\"rh_height\":1790,\"scrollTop\":0,\"scrollLeft\":0,\"visibledatarow\":[],\"visibledatacolumn\":[],\"jfgird_select_save\":[],\"jfgrid_selection_range\":{}}]";
            var jArray = defaultSheet.ToObject<JArray>();
            foreach (var item in jArray)
            {
                string index = item.Value<string>("index");
                string itemJson = item.ToJson();
                int status = item.Value<int>("status");
                int order = item.Value<int>("order");
                await _sheetRepository.Add(gridKey, index, itemJson, status, order, 0);
            }
        }

        /// <summary>
        /// 获取execl文档所有的sheet页
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <returns></returns>
        public async Task<IEnumerable<Sheet>> GetSheets(string gridKey)
        {
            return await _sheetRepository.GetSheets(gridKey);
        }

        /// <summary>
        /// 根据sheet下标获取sheet
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <param name="index">sheet下标</param>
        /// <returns></returns>
        public async Task<Sheet> GetSheet(string gridKey, string index)
        {
            return await _sheetRepository.GetSheet(gridKey, index);
        }

        /// <summary>
        /// 加载execl文档sheet页
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <returns></returns>
        public async Task<JArray> LoadSheet(string gridKey)
        {
            var sheets = (await _sheetRepository.GetSheets(gridKey)).Where(p => p.is_delete == 0).OrderByDescending(p => p.status).ThenBy(p => p.order);
            var jArray = new JArray();
            foreach (var item in sheets)
            {
                var itemJObject = item.json_data.ToObject<JObject>();
                if (item.status == 1)
                {
                    jArray.Add(itemJObject);
                }
                else
                {
                    var newItem = new JObject
                    {
                        {"name", itemJObject["name"]},
                        {"index", itemJObject["index"]},
                        {"order", itemJObject["order"]},
                        {"status", itemJObject["status"]}
                    };
                    jArray.Add(newItem);
                }
            }
            return jArray;
        }

        /// <summary>
        /// 加载其它sheet页
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <param name="index">sheet下标</param>
        /// <returns></returns>
        public async Task<JObject> LoadOtherSheet(string gridKey, string index)
        {
            var sheet = await _sheetRepository.GetSheet(gridKey, index);
            var jsonData = sheet.json_data.ToObject<JObject>();
            var jObject = new JObject
            {
                {jsonData.Value<string>("index"), jsonData.Value<JArray>("celldata")}
            };
            return jObject;
        }

        /// <summary>
        /// 提交execl文档
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <param name="jsonData">execl数据</param>
        /// <returns></returns>
        public async Task SubmitSheet(string gridKey, string jsonData)
        {
            if (await _sheetRepository.Delete(gridKey))
            {
                var jArray = jsonData.ToObject<JArray>();
                foreach (var item in jArray)
                {
                    string index = item.Value<string>("index");
                    string itemJson = item.ToJson();
                    int status = item.Value<int>("status");
                    int order = item.Value<int>("order");
                    await _sheetRepository.Add(gridKey, index, itemJson, status, order, 0);
                }
            }
        }

        /// <summary>
        /// 添加sheet
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <param name="index">sheet下标</param>
        /// <param name="jsonData">sheet json数据</param>
        /// <param name="status">sheet状态</param>
        /// <param name="order">sheet序号</param>
        /// <param name="isDelete">是否删除</param>
        /// <returns></returns>
        public async Task<bool> AddSheet(string gridKey, string index, string jsonData, int status, int order, int isDelete)
        {
            return await _sheetRepository.Add(gridKey, index, jsonData, status, order, isDelete);
        }

        /// <summary>
        /// 修改sheet页数据
        /// </summary>
        /// <param name="id">主键</param>
        /// <param name="gridKey">execl文档主键</param>
        /// <param name="index">sheet下标</param>
        /// <param name="jsonData">sheet json数据</param>
        /// <param name="status">sheet状态</param>
        /// <param name="order">sheet序号</param>
        /// <param name="isDelete">是否删除</param>
        /// <returns></returns>
        public async Task<bool> UpdateSheet(int id, string gridKey, string index, string jsonData, int status, int order, int isDelete)
        {
            return await _sheetRepository.Update(id, gridKey, index, jsonData, status, order, isDelete);
        }

        /// <summary>
        /// 修改sheet页的是否删除状态
        /// </summary>
        /// <param name="id">主键id</param>
        /// <param name="isDelete">是否删除</param>
        /// <returns></returns>
        public async Task<bool> UpdateIsDelete(int id, int isDelete)
        {
            return await _sheetRepository.UpdateIsDelete(id, isDelete);
        }

        /// <summary>
        /// 修改sheet页的是否删除状态
        /// </summary>
        /// <param name="id">主键id</param>
        /// <param name="status">激活状态</param>
        /// <returns></returns>
        public async Task<bool> UpdateStatus(int id, int status)
        {
            return await _sheetRepository.UpdateStatus(id, status);
        }

        /// <summary>
        /// 删除execl
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <returns></returns>
        public async Task<bool> DeleteSheet(string gridKey)
        {
            return await _sheetRepository.Delete(gridKey);
        }
    }
}
