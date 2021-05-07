using System;

namespace ExeclWeb.Core.Entitys
{
    public class Sheet
    {
        /// <summary>
        /// 主键
        /// </summary>
        public int id { get; set; }
        /// <summary>
        /// sheetId
        /// </summary>
        public string sheet_id { get; set; }
        /// <summary>
        /// sheet下标
        /// </summary>
        public string index { get; set; }
        /// <summary>
        /// sheet json数据
        /// </summary>
        public string json_data { get; set; }
        /// <summary>
        /// 状态，是否当前sheet页
        /// </summary>
        public int status { get; set; }
        /// <summary>
        /// sheet序号
        /// </summary>
        public int order { get; set; }
        /// <summary>
        /// 是否删除
        /// </summary>
        public int is_delete { get; set; }
        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime create_time { get; set; }
        /// <summary>
        /// 修改时间
        /// </summary>
        public DateTime? update_time { get; set; }
    }
}
