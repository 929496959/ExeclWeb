using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;

namespace ExeclWeb.Core.Common
{
    public static class ExcelHepler
    {
        /// <summary>
        /// 获取单元格的值
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static object GetCellValue(ICell item)
        {
            if (item == null)
            {
                return string.Empty;
            }
            switch (item.CellType)
            {
                case CellType.Boolean:
                    return item.BooleanCellValue;
                case CellType.Error:
                    return ErrorEval.GetText(item.ErrorCellValue);
                case CellType.Formula:
                    switch (item.CachedFormulaResultType)
                    {
                        case CellType.Boolean:
                            return item.BooleanCellValue;

                        case CellType.Error:
                            return ErrorEval.GetText(item.ErrorCellValue);

                        case CellType.Numeric:
                            if (DateUtil.IsCellDateFormatted(item))
                            {
                                return item.DateCellValue.ToString("yyyy-MM-dd");
                            }
                            else
                            {
                                return item.NumericCellValue;
                            }
                        case CellType.String:
                            string str = item.StringCellValue;
                            if (!string.IsNullOrEmpty(str))
                            {
                                return str.ToString();
                            }
                            else
                            {
                                return string.Empty;
                            }
                        case CellType.Unknown:
                        case CellType.Blank:
                        default:
                            return string.Empty;
                    }
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(item))
                    {
                        return item.DateCellValue.ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        return item.NumericCellValue;
                    }
                case CellType.String:
                    string strValue = item.StringCellValue;
                    return strValue.ToString();

                case CellType.Unknown:
                case CellType.Blank:
                default:
                    return string.Empty;
            }
        }

        /// <summary>
        /// 获取边框样式
        /// </summary>
        /// <param name="styleId"></param>
        /// <returns></returns>
        public static BorderStyle GetBorderStyle(int styleId)
        {
            switch (styleId)
            {
                case 0:
                    return BorderStyle.None;
                case 1:
                    return BorderStyle.Thin;
                case 2:
                    return BorderStyle.Medium;
                case 3:
                    return BorderStyle.Dashed;
                case 4:
                    return BorderStyle.Dotted;
                case 5:
                    return BorderStyle.Thick;
                case 6:
                    return BorderStyle.Double;
                case 7:
                    return BorderStyle.Hair;
                case 8:
                    return BorderStyle.MediumDashed;
                case 9:
                    return BorderStyle.DashDot;
                case 10:
                    return BorderStyle.MediumDashDot;
                case 11:
                    return BorderStyle.DashDotDot;
                case 12:
                    return BorderStyle.MediumDashDotDot;
                case 13:
                    return BorderStyle.SlantedDashDot;
                default:
                    return BorderStyle.None;
            }
        }
    }
}
