using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using ExeclWeb.Core.Common;
using ExeclWeb.Core.Repository;

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
        /// 判断是否存在execl文件
        /// </summary>
        /// <param name="gridKey">execl主键</param>
        /// <returns></returns>
        public async Task<bool> IsExist(string gridKey)
        {
            return await _sheetRepository.IsExist(gridKey);
        }

        /// <summary>
        /// 初始化execl sheet数据
        /// </summary>
        /// <param name="gridKey">execl主键</param>
        public async Task InitSheet(string gridKey)
        {
            string defaultSheet = "[{\"name\":\"Sheet1\",\"chart\":[],\"color\":\"\",\"index\":\"1\",\"order\":0,\"row\":84,\"column\":60,\"config\":{},\"status\":1,\"celldata\":[],\"ch_width\":4748,\"rowsplit\":[],\"rh_height\":1790,\"scrollTop\":0,\"scrollLeft\":0,\"visibledatarow\":[],\"visibledatacolumn\":[],\"jfgird_select_save\":[],\"jfgrid_selection_range\":{}},{\"name\":\"Sheet2\",\"chart\":[],\"color\":\"\",\"index\":\"2\",\"order\":0,\"row\":84,\"column\":60,\"config\":{},\"status\":0,\"celldata\":[],\"ch_width\":4748,\"rowsplit\":[],\"rh_height\":1790,\"scrollTop\":0,\"scrollLeft\":0,\"visibledatarow\":[],\"visibledatacolumn\":[],\"jfgird_select_save\":[],\"jfgrid_selection_range\":{}},{\"name\":\"Sheet3\",\"chart\":[],\"color\":\"\",\"index\":\"3\",\"order\":0,\"row\":84,\"column\":60,\"config\":{},\"status\":0,\"celldata\":[],\"ch_width\":4748,\"rowsplit\":[],\"rh_height\":1790,\"scrollTop\":0,\"scrollLeft\":0,\"visibledatarow\":[],\"visibledatacolumn\":[],\"jfgird_select_save\":[],\"jfgrid_selection_range\":{}}]";
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
        /// 获取Sheet工作簿
        /// </summary>
        /// <param name="gridKey">工作簿key</param>
        /// <returns></returns>
        public async Task<JArray> LoadSheet(string gridKey)
        {
            var sheets = await _sheetRepository.GetSheets(gridKey);
            sheets = sheets.Where(p => p.is_delete == 0).ToList();
            var jArray = new JArray();
            foreach (var item in sheets)
            {
                jArray.Add(item.json_data.ToObject<JObject>());
            }
            return jArray;
        }

        /// <summary>
        /// 提交Sheet文档
        /// </summary>
        /// <param name="gridKey">工作簿key</param>
        /// <param name="jsonData">工作簿数据</param>
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
        /// 删除execl
        /// </summary>
        /// <param name="gridKey">工作簿key</param>
        /// <returns></returns>
        public async Task<bool> DeleteSheet(string gridKey)
        {
            return await _sheetRepository.Delete(gridKey);
        }
    }
}
