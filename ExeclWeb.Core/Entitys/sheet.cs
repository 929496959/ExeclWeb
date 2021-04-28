using System;

namespace ExeclWeb.Core.Entitys
{
    public class sheet
    {
        /// <summary>
        /// 主键
        /// </summary>
        public int id { get; set; }
        /// <summary>
        /// sheet文档id
        /// </summary>
        public string sheet_id { get; set; }
        /// <summary>
        /// sheet数据
        /// </summary>
        public string json_data { get; set; }
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
