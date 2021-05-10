using System;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using ExeclWeb.Core.Application;
using ExeclWeb.Core.Common;

namespace ExeclWeb.Controllers
{
    public class SheetController : Controller
    {
        private readonly SheetService _sheetService;
        public SheetController()
        {
            _sheetService = new SheetService();
        }

        /// <summary>
        /// 普通办公
        /// </summary>
        /// <returns></returns>
        public IActionResult SheetIndex()
        {
            return View();
        }

        /// <summary>
        /// 加载execl
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> LoadSheet(string gridKey)
        {
            if (!await _sheetService.IsExist(gridKey))
            {
                await _sheetService.InitSheet(gridKey);
            }
            var jArray = await _sheetService.LoadSheet(gridKey);
            return Json(jArray.ToJson());
        }

        /// <summary>
        /// 加载其它sheet页
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <param name="index">sheet下标</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> LoadOtherSheet(string gridKey, string index)
        {
            var jObject = await _sheetService.LoadOtherSheet(gridKey, index);
            return Json(jObject.ToJson());
        }

        /// <summary>
        /// 提交execl文档
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <param name="data">execl数据</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> SubmitSheet(string gridKey, string data)
        {
            await _sheetService.SubmitSheet(gridKey, data);
            return Json(data);
        }

        /// <summary>
        /// 删除execl
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> DeleteSheet(string gridKey)
        {
            var status = await _sheetService.DeleteSheet(gridKey);
            return Json(status);
        }

        /// <summary>
        /// 协同办公
        /// </summary>
        /// <returns></returns>
        public IActionResult SynergySheetIndex()
        {
            return View();
        }

        /// <summary>
        /// 加载execl文档
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> LoadSynergySheet(string gridKey)
        {
            if (!await _sheetService.IsExist(gridKey))
            {
                await _sheetService.InitSheet(gridKey);
            }
            var jArray = await _sheetService.LoadSheet(gridKey);
            return Json(jArray.ToJson());
        }

        /// <summary>
        /// 加载其它sheet页
        /// </summary>
        /// <param name="gridKey">execl文档主键</param>
        /// <param name="index">sheet下标</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> LoadOtherSynergySheet(string gridKey, string index)
        {
            var jObject = await _sheetService.LoadOtherSheet(gridKey, index);
            return Json(jObject.ToJson());
        }
    }
}
