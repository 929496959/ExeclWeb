using System;
using System.Linq;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Pipes;
using System.Threading.Tasks;
using System.Web;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Handy.DotNETCoreCompatibility.ColourTranslations;
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
        /// <param name="gridKey">execl文档key</param>
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
        /// <param name="gridKey">execl文档key</param>
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
        /// <param name="gridKey">execl文档key</param>
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
        /// <param name="gridKey">execl文档key</param>
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
        /// <param name="gridKey">execl文档key</param>
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
        /// <param name="gridKey">execl文档key</param>
        /// <param name="index">sheet下标</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<IActionResult> LoadOtherSynergySheet(string gridKey, string index)
        {
            var jObject = await _sheetService.LoadOtherSheet(gridKey, index);
            return Json(jObject.ToJson());
        }

        /// <summary>
        /// 导出execl
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        [HttpPost]
        public IActionResult DownExcel(string data)
        {
            List<JObject> dowExcels = JsonConvert.DeserializeObject<List<JObject>>(data);
            IWorkbook workbook;
            workbook = new HSSFWorkbook();
            //赋予表格默认样式，因为样式文件存储只能有4000个
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.VerticalAlignment = VerticalAlignment.Center; //垂直居中
            cellStyle.WrapText = true; //自动换行
            cellStyle.Alignment = HorizontalAlignment.Center; //水平居中
            cellStyle.FillPattern = FillPattern.NoFill; //背景色是必须要这数据
            cellStyle.FillForegroundColor = IndexedColors.White.Index;//默认背景色
            foreach (var item in dowExcels)
            {
                var columnlen = item["config"] != null ? item["config"]["columnlen"] : null;
                var rowlen = item["config"] != null ? item["config"]["rowlen"] : null;
                var borderInfo = item["config"] != null ? item["config"]["borderInfo"] : null;
                var merge = item["config"] != null ? item["config"]["merge"] : null;
                var Cjarray = item["celldata"];

                //读取了模板内所有sheet内容
                ISheet sheet = workbook.CreateSheet(item["name"].ToString());
                //判断是否有值，并且赋予样式
                if (Cjarray != null)
                {
                    for (int i = 0; i < Cjarray.Count(); i++)
                    {
                        //判断行，存不存在，不存在创建
                        IRow row = sheet.GetRow(int.Parse(Cjarray[i]["r"].ToString()));
                        if (row == null)
                        {
                            row = sheet.CreateRow(int.Parse(Cjarray[i]["r"].ToString()));
                        }
                        var ct = Cjarray[i]["v"]["ct"];
                        var cct = Cjarray[i]["v"];
                        if (ct != null && ct["s"] != null)
                        {
                            //合并单元格的走这边
                            string celldatas = "";
                            //合并单元格时，会导致文字丢失，提前处理文字信息
                            for (int j = 0; j < ct["s"].Count(); j++)
                            {
                                var stv = ct["s"][j];
                                celldatas += stv["v"] != null ? stv["v"].ToString() : "";
                            }
                            //判断列，不存在创建
                            ICell Cell = row.GetCell(int.Parse(Cjarray[i]["c"].ToString()));
                            if (Cell == null)
                            {
                                HSSFRichTextString richtext = new HSSFRichTextString(celldatas);
                                Cell = row.CreateCell(int.Parse(Cjarray[i]["c"].ToString()));
                                for (int k = 0; k < ct["s"].Count(); k++)
                                {
                                    IFont font = workbook.CreateFont();
                                    var stv = ct["s"][k];
                                    //文字颜色
                                    if (stv["fc"] != null)
                                    {
                                        var rGB = HTMLColorTranslator.Translate(stv["fc"].ToString());
                                        var color = Color.FromArgb(rGB.R, rGB.G, rGB.B);
                                        font.Color = ((HSSFWorkbook)workbook).GetCustomPalette().FindSimilarColor(color.R, color.G, color.B).Indexed;
                                    }
                                    else
                                        font.Color = HSSFColor.Black.Index;
                                    //是否加粗
                                    if (stv["bl"] != null)
                                    {
                                        font.IsBold = !string.IsNullOrEmpty(stv["bl"].ToString()) && (stv["bl"].ToString() == "1" ? true : false);
                                        font.Boldweight = stv["bl"].ToString() == "1" ? (short)FontBoldWeight.Bold : (short)FontBoldWeight.None;
                                    }
                                    else
                                    {
                                        font.IsBold = false;
                                        font.Boldweight = (short)FontBoldWeight.None;
                                    }
                                    //是否斜体
                                    if (stv["it"] != null)
                                        font.IsItalic = !string.IsNullOrEmpty(stv["it"].ToString()) && (stv["it"].ToString() == "1" ? true : false);
                                    else
                                        font.IsItalic = false;
                                    //下划线
                                    if (stv["un"] != null)
                                    {
                                        font.Underline = stv["un"].ToString() == "1" ? FontUnderlineType.Single : FontUnderlineType.None;
                                    }
                                    else
                                        font.Underline = FontUnderlineType.None;
                                    //字体
                                    if (stv["ff"] != null)
                                        font.FontName = stv["ff"].ToString();
                                    //文字大小
                                    if (stv["fs"] != null)
                                        font.FontHeightInPoints = double.Parse(stv["fs"].ToString());
                                    Cell.CellStyle.SetFont(font);
                                    richtext.ApplyFont(celldatas.IndexOf(stv["v"].ToString()), celldatas.IndexOf(stv["v"].ToString()) + stv["v"].ToString().Length, font);
                                    Cell.SetCellValue(richtext);
                                }
                                //背景颜色
                                if (cct["bg"] != null)
                                {
                                    ICellStyle cellStyle1 = workbook.CreateCellStyle();
                                    cellStyle1.CloneStyleFrom(cellStyle);
                                    if (cct["bg"] != null)
                                    {
                                        var rGB = HTMLColorTranslator.Translate(cct["bg"].ToString());
                                        var color = Color.FromArgb(rGB.R, rGB.G, rGB.B);
                                        cellStyle1.FillPattern = FillPattern.SolidForeground;
                                        cellStyle1.FillForegroundColor = ((HSSFWorkbook)workbook).GetCustomPalette().FindSimilarColor(color.R, color.G, color.B).Indexed;
                                    }
                                    Cell.CellStyle = cellStyle1;
                                }
                                else
                                {
                                    Cell.CellStyle = cellStyle;
                                }
                            }
                        }
                        else
                        {
                            //没有合并单元格的走这边
                            //判断列，不存在创建
                            ICell Cell = row.GetCell(int.Parse(Cjarray[i]["c"].ToString()));
                            if (Cell == null)
                            {
                                Cell = row.CreateCell(int.Parse(Cjarray[i]["c"].ToString()));
                                IFont font = workbook.CreateFont();
                                ct = Cjarray[i]["v"];
                                //字体颜色
                                if (ct["fc"] != null)
                                {
                                    var rGB = HTMLColorTranslator.Translate(ct["fc"].ToString());
                                    var color = Color.FromArgb(rGB.R, rGB.G, rGB.B);
                                    font.Color = ((HSSFWorkbook)workbook).GetCustomPalette().FindSimilarColor(color.R, color.G, color.B).Indexed;
                                }
                                else
                                    font.Color = HSSFColor.Black.Index;
                                //是否加粗
                                if (ct["bl"] != null)
                                {
                                    font.IsBold = !string.IsNullOrEmpty(ct["bl"].ToString()) && (ct["bl"].ToString() == "1" ? true : false);
                                    font.Boldweight = ct["bl"].ToString() == "1" ? (short)FontBoldWeight.Bold : (short)FontBoldWeight.None;
                                }
                                else
                                {
                                    font.IsBold = false;
                                    font.Boldweight = (short)FontBoldWeight.None;
                                }
                                //斜体
                                if (ct["it"] != null)
                                    font.IsItalic = !string.IsNullOrEmpty(ct["it"].ToString()) && (ct["it"].ToString() == "1" ? true : false);
                                else
                                    font.IsItalic = false;
                                //下划线
                                if (ct["un"] != null)
                                    font.Underline = ct["un"].ToString() == "1" ? FontUnderlineType.Single : FontUnderlineType.None;
                                else
                                    font.Underline = FontUnderlineType.None;
                                //字体
                                if (ct["ff"] != null)
                                    font.FontName = ct["ff"].ToString();
                                //文字大小
                                if (ct["fs"] != null)
                                    font.FontHeightInPoints = double.Parse(ct["fs"].ToString());
                                Cell.CellStyle.SetFont(font);
                                //判断背景色
                                if (ct["bg"] != null)
                                {
                                    ICellStyle cellStyle1 = workbook.CreateCellStyle();
                                    cellStyle1.CloneStyleFrom(cellStyle);
                                    if (ct["bg"] != null)
                                    {
                                        var rGB = HTMLColorTranslator.Translate(ct["bg"].ToString());
                                        var color = Color.FromArgb(rGB.R, rGB.G, rGB.B);
                                        cellStyle1.FillPattern = FillPattern.SolidForeground;
                                        cellStyle1.FillForegroundColor = ((HSSFWorkbook)workbook).GetCustomPalette().FindSimilarColor(color.R, color.G, color.B).Indexed;
                                    }
                                    Cell.CellStyle = cellStyle1;
                                }
                                else
                                {
                                    Cell.CellStyle = cellStyle;
                                }
                                Cell.SetCellValue(ct["v"] != null ? ct["v"].ToString() : "");
                            }
                        }
                    }
                    sheet.ForceFormulaRecalculation = true;
                }
                //判断是否要设置列宽度
                if (columnlen != null)
                {
                    foreach (var cols in columnlen)
                    {
                        var p = cols as JProperty;
                        sheet.SetColumnWidth(int.Parse(p.Name), int.Parse(p.Value.ToString()) * 38);
                    }
                }
                //判断是否要设置行高度
                if (rowlen != null)
                {
                    foreach (var rows in rowlen)
                    {
                        var p = rows as JProperty;
                        sheet.GetRow(int.Parse(p.Name)).HeightInPoints = float.Parse(p.Value.ToString());
                    }
                }
                //判断是否要加边框
                if (borderInfo != null)
                {
                    for (int i = 0; i < borderInfo.Count(); i++)
                    {
                        var bordervalue = borderInfo[i]["value"];
                        if (bordervalue != null)
                        {
                            var rowindex = bordervalue["row_index"];
                            var colindex = bordervalue["col_index"];
                            var l = bordervalue["l"];
                            var r = bordervalue["r"];
                            var t = bordervalue["t"];
                            var b = bordervalue["b"];
                            if (rowindex != null)
                            {
                                IRow rows = sheet.GetRow(int.Parse(bordervalue["row_index"].ToString()));
                                if (colindex != null)
                                {
                                    ICell cell = rows.GetCell(int.Parse(bordervalue["col_index"].ToString()));
                                    if (b != null)
                                    {
                                        cell.CellStyle.BorderBottom = ExcelHepler.GetBorderStyle(int.Parse(b["style"].ToString()));
                                        var rGB = HTMLColorTranslator.Translate(b["color"].ToString());
                                        var bcolor = Color.FromArgb(rGB.R, rGB.G, rGB.B);
                                        cell.CellStyle.BottomBorderColor = ((HSSFWorkbook)workbook).GetCustomPalette().FindSimilarColor(bcolor.R, bcolor.G, bcolor.B).Indexed;
                                    }
                                    else
                                    {
                                        cell.CellStyle.BorderBottom = BorderStyle.None;
                                        cell.CellStyle.BottomBorderColor = HSSFColor.COLOR_NORMAL;
                                    }
                                    if (t != null)
                                    {
                                        cell.CellStyle.BorderTop = ExcelHepler.GetBorderStyle(int.Parse(t["style"].ToString()));
                                        var rGB = HTMLColorTranslator.Translate(t["color"].ToString());
                                        var tcolor = Color.FromArgb(rGB.R, rGB.G, rGB.B);
                                        cell.CellStyle.TopBorderColor = ((HSSFWorkbook)workbook).GetCustomPalette().FindSimilarColor(tcolor.R, tcolor.G, tcolor.B).Indexed;
                                    }
                                    else
                                    {
                                        cell.CellStyle.BorderBottom = BorderStyle.None;
                                        cell.CellStyle.BottomBorderColor = HSSFColor.COLOR_NORMAL;
                                    }
                                    if (l != null)
                                    {
                                        cell.CellStyle.BorderLeft = ExcelHepler.GetBorderStyle(int.Parse(l["style"].ToString()));
                                        var rGB = HTMLColorTranslator.Translate(l["color"].ToString());
                                        var lcolor = Color.FromArgb(rGB.R, rGB.G, rGB.B);
                                        cell.CellStyle.LeftBorderColor = ((HSSFWorkbook)workbook).GetCustomPalette().FindSimilarColor(lcolor.R, lcolor.G, lcolor.B).Indexed;
                                    }
                                    else
                                    {
                                        cell.CellStyle.BorderBottom = BorderStyle.None;
                                        cell.CellStyle.BottomBorderColor = HSSFColor.COLOR_NORMAL;
                                    }
                                    if (r != null)
                                    {
                                        cell.CellStyle.BorderRight = ExcelHepler.GetBorderStyle(int.Parse(r["style"].ToString()));
                                        var rGB = HTMLColorTranslator.Translate(r["color"].ToString());
                                        var rcolor = Color.FromArgb(rGB.R, rGB.G, rGB.B);
                                        cell.CellStyle.RightBorderColor = ((HSSFWorkbook)workbook).GetCustomPalette().FindSimilarColor(rcolor.R, rcolor.G, rcolor.B).Indexed;
                                    }
                                    else
                                    {
                                        cell.CellStyle.BorderBottom = BorderStyle.None;
                                        cell.CellStyle.BottomBorderColor = HSSFColor.COLOR_NORMAL;
                                    }
                                }
                            }
                        }
                    }
                }
                //判断是否要合并单元格
                if (merge != null)
                {
                    foreach (var imerge in merge)
                    {
                        var firstmer = imerge.First();
                        int r = int.Parse(firstmer["r"].ToString());//主单元格的行号,开始行号
                        int rs = int.Parse(firstmer["rs"].ToString());//合并单元格占的行数,合并多少行
                        int c = int.Parse(firstmer["c"].ToString());//主单元格的列号,开始列号
                        int cs = int.Parse(firstmer["cs"].ToString());//合并单元格占的列数,合并多少列
                        CellRangeAddress region = new CellRangeAddress(r, r + rs - 1, c, c + cs - 1);
                        sheet.AddMergedRegion(region);
                    }
                }
            }

            var dir = Path.Combine(Directory.GetCurrentDirectory(), @"wwwroot\Files");
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            var fileId = Guid.NewGuid();
            var fileName = $"{fileId}.xls";
            FileStream stream = new FileStream(Path.Combine(dir, fileName), FileMode.OpenOrCreate);
            workbook.Write(stream);
            stream.Seek(0, SeekOrigin.Begin);
            workbook.Close();
            stream.Close();
            return Ok(fileId);
        }
    }
}
