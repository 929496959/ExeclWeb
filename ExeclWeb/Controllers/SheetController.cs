﻿using System;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Data.Application;

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
        /// 加载工作簿
        /// </summary>
        /// <param name="gridKey">工作key</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> LoadSheet(string gridKey)
        {
            var data = await _sheetService.LoadSheetData(gridKey);
            return Json(data);
        }

        /// <summary>
        /// 加载其它页 celldata
        /// </summary>
        /// <param name="gridKey">唯一key</param>
        /// <param name="index">sheet主键合集，格式为["sheet_01","sheet_02","sheet_03"]</param>
        /// <returns></returns>
        [Obsolete]
        [HttpPost]
        public IActionResult LoadOtherSheet(string gridKey, string index)
        {
            return Json("{}");
        }

        /// <summary>
        /// 提交sheet
        /// </summary>
        /// <param name="gridKey">工作薄key</param>
        /// <param name="data">工作薄数据</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> SubmitSheet(string gridKey, string data)
        {
            await _sheetService.SubmitSheet(gridKey, data);
            return Json(data);
        }

        /// <summary>
        /// 删除工作薄
        /// </summary>
        /// <param name="gridKey"></param>
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
        /// 加载工作簿
        /// </summary>
        /// <param name="gridKey">唯一key</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> LoadSynergySheet(string gridKey)
        {
            //var data = await _sheetService.LoadSheetData(gridKey);
            //return Json(data);

            var json = "[{\"name\":\"Sheet1\",\"index\":\"1\",\"order\":0,\"status\":1,\"config\":{},\"celldata\":[{\"v\":1,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"1\",\"r\":0,\"c\":0},{\"v\":2,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"2\",\"r\":1,\"c\":0},{\"v\":3,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"3\",\"r\":2,\"c\":0},{\"v\":4,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"4\",\"r\":3,\"c\":0},{\"v\":5,\"ct\":{\"fa\":\"General\",\"t\":\"n\"},\"m\":\"5\",\"r\":4,\"c\":0}]},{\"name\":\"Sheet2\",\"index\":\"2\",\"order\":1,\"status\":0,\"config\":{\"borderInfo\":null},\"celldata\":[]},{\"name\":\"Sheet3\",\"index\":\"3\",\"order\":2,\"status\":0,\"config\":{},\"celldata\":[]}]";
            return Json(json);
        }

        /// <summary>
        /// 加载其它页 celldata
        /// </summary>
        /// <param name="gridKey"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        [Obsolete]
        [HttpPost]
        public IActionResult LoadOtherSynergySheet(string gridKey, string index)
        {
            return Json("{}");
        }
    }
}