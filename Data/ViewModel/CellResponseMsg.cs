namespace Data.ViewModel
{
    /// <summary>
    /// 单元格响应Model
    /// </summary>
    public class CellResponseMsg
    {
        /// <summary>
        /// 命令发送时间
        /// </summary>
        public string createTime { get; set; }
        /// <summary>
        /// 修改的命令
        /// </summary>
        public object data { get; set; }
        /// <summary>
        /// websocket的id
        /// </summary>
        public string id { get; set; }
        /// <summary>
        /// 返回状态
        /// </summary>
        public string returnMessage { get; set; }
        /// <summary>
        /// 0告诉前端需要根据data的命令修改  1无意义
        /// </summary>
        public int status { get; set; }
        /// <summary>
        /// 0：连接成功，1：发送给当前连接的用户，2：发送信息给其他用户，3：发送选区位置信息，999：用户连接断开
        /// </summary>
        public int type { get; set; }
        /// <summary>
        /// 用户名称
        /// </summary>
        public string username { get; set; }
    }
}
