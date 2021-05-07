using System;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using ExeclWeb.Core.Application;
using ExeclWeb.Core.Common;

namespace ExeclWeb.Common
{
    public class SheetWebSocket
    {
        private readonly SheetService _sheetService;
        public SheetWebSocket()
        {
            _sheetService = new SheetService();
        }

        public async Task<bool> Process(JObject requestMsg, string gridKey)
        {
            try
            {
                if (requestMsg == null) return false;

                var sheetJsonData = await _sheetService.LoadSheet(gridKey);

                string type = requestMsg.Value<string>("t");
                switch (type)
                {
                    case "v":
                        // 单元格刷新
                        break;
                    case "rv":
                        // 范围单元格刷新
                        break;
                    case "cg":
                        // config操作
                        break;
                    case "all":
                        // 通用保存
                        break;
                    case "fc":
                        // 函数链操作
                        break;
                    case "drc":
                        // 删除行或列
                        break;
                    case "arc":
                        // 增加行或列
                        break;
                    case "fsc":
                        // 清除筛选
                        break;
                    case "fsr":
                        // 恢复筛选
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

                return true;
            }
            catch (Exception e)
            {
                NLogger.Error(e.Message);
            }

            return false;
        }
    }
}
