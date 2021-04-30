using System.Threading.Tasks;
using System.Transactions;
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
        /// 提交Sheet文档
        /// </summary>
        /// <param name="gridKey">工作簿key</param>
        /// <param name="jsonData">工作簿数据</param>
        /// <returns></returns>
        public async Task<string> SubmitSheet(string gridKey, string jsonData)
        {
            bool status = await _sheetRepository.SubmitSheetAsync(gridKey, jsonData);
            if (status)
            {
                // 更新缓存
                Redis.Cli.Set(gridKey, jsonData, TimeOut.TimeOutMinute * 20);
            }
            return jsonData;
        }

        /// <summary>
        /// 获取Sheet工作簿
        /// </summary>
        /// <param name="gridKey">工作簿key</param>
        /// <returns></returns>
        public async Task<string> LoadSheetData(string gridKey)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                string data = null;
                // 从缓存加载数据
                if (Redis.Cli.Exists(gridKey))
                {
                    data = Redis.Cli.Get<string>(gridKey);
                }
                // 从数据库加载数据
                if (data == null)
                {
                    var sheet = await _sheetRepository.GetSheetAsync(gridKey);
                    if (sheet == null)
                    {
                        // 如果数据库不存在，初始一条数据
                        string _defaultSheet = "[{\"name\":\"Sheet1\",\"chart\":[],\"color\":\"\",\"index\":\"1\",\"order\":0,\"row\":84,\"column\":60,\"config\":{},\"status\":1,\"celldata\":[],\"ch_width\":4748,\"rowsplit\":[],\"rh_height\":1790,\"scrollTop\":0,\"scrollLeft\":0,\"visibledatarow\":[],\"visibledatacolumn\":[],\"jfgird_select_save\":[],\"jfgrid_selection_range\":{}},{\"name\":\"Sheet2\",\"chart\":[],\"color\":\"\",\"index\":\"2\",\"order\":0,\"row\":84,\"column\":60,\"config\":{},\"status\":0,\"celldata\":[],\"ch_width\":4748,\"rowsplit\":[],\"rh_height\":1790,\"scrollTop\":0,\"scrollLeft\":0,\"visibledatarow\":[],\"visibledatacolumn\":[],\"jfgird_select_save\":[],\"jfgrid_selection_range\":{}},{\"name\":\"Sheet3\",\"chart\":[],\"color\":\"\",\"index\":\"3\",\"order\":0,\"row\":84,\"column\":60,\"config\":{},\"status\":0,\"celldata\":[],\"ch_width\":4748,\"rowsplit\":[],\"rh_height\":1790,\"scrollTop\":0,\"scrollLeft\":0,\"visibledatarow\":[],\"visibledatacolumn\":[],\"jfgird_select_save\":[],\"jfgrid_selection_range\":{}}]";
                        var addStatus = await _sheetRepository.AddSheetAsync(gridKey, _defaultSheet);
                        if (addStatus)
                        {
                            // 将数据写入缓存
                            Redis.Cli.Set(gridKey, _defaultSheet, TimeOut.TimeOutMinute * 20);
                            data = _defaultSheet;
                        }
                    }
                    else
                    {
                        // 将数据写入缓存
                        Redis.Cli.Set(gridKey, sheet.json_data, TimeOut.TimeOutMinute * 20);
                        data = sheet.json_data;
                    }
                }
                scope.Complete();
                return data;
            }
        }

        /// <summary>
        /// 修改Sheet工作薄数据
        /// </summary>
        /// <param name="gridKey"></param>
        /// <param name="jsonData"></param>
        /// <returns></returns>
        public async Task<bool> UpdateSheet(string gridKey, string jsonData)
        {
            return await _sheetRepository.UpdateSheetAsync(gridKey, jsonData);
        }

        /// <summary>
        /// 删除Sheet工作薄
        /// </summary>
        /// <param name="gridKey">工作簿key</param>
        /// <returns></returns>
        public async Task<bool> DeleteSheet(string gridKey)
        {
            using (TransactionScope scope = new TransactionScope(TransactionScopeAsyncFlowOption.Enabled))
            {
                // 删除缓存
                if (Redis.Cli.Exists(gridKey))
                {
                    Redis.Cli.Del(gridKey);
                }

                // 删除数据
                var status = await _sheetRepository.DeleteSheetAsync(gridKey);

                scope.Complete();
                return status;
            }
        }
    }
}
