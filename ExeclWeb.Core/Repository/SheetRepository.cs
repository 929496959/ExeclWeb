using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Dapper;
using ExeclWeb.Core.Data;
using ExeclWeb.Core.Entitys;

namespace ExeclWeb.Core.Repository
{
    public class SheetRepository
    {
        private readonly DapperHelper _dapperHelper;
        public SheetRepository()
        {
            _dapperHelper = new DapperHelper();
        }

        /// <summary>
        /// 获取execl所有sheet页
        /// </summary>
        /// <param name="gridKey">execl主键</param>
        /// <returns></returns>
        public async Task<IEnumerable<Sheet>> GetSheets(string gridKey)
        {
            var sql = "SELECT * FROM sheet WHERE grid_key=@grid_key;";
            return await _dapperHelper.QueryAsync<Sheet>(sql, new { grid_key = gridKey });
        }

        /// <summary>
        /// 获取sheet页
        /// </summary>
        /// <param name="gridKey">execl文档key</param>
        /// <param name="index">sheet下标</param>
        /// <returns></returns>
        public async Task<Sheet> GetSheet(string gridKey, string index)
        {
            var sql = "SELECT * FROM sheet WHERE grid_key=@grid_key AND `index`=@index;";
            return await _dapperHelper.QueryFirstAsync<Sheet>(sql, new { grid_key = gridKey, index });
        }

        /// <summary>
        /// 添加sheet
        /// </summary>
        /// <param name="gridKey">execl文档key</param>
        /// <param name="index">sheet下标</param>
        /// <param name="jsonData">sheet json数据</param>
        /// <param name="status">sheet状态</param>
        /// <param name="order">sheet序号</param>
        /// <param name="isDelete">是否删除</param>
        /// <returns></returns>
        public async Task<bool> Add(string gridKey, string index, string jsonData, int status, int order, int isDelete)
        {
            var param = new DynamicParameters();
            param.Add("@grid_key", gridKey);
            param.Add("@index", index);
            param.Add("@json_data", jsonData);
            param.Add("@status", status);
            param.Add("@order", order);
            param.Add("@is_delete", isDelete);
            param.Add("@create_time", DateTime.Now);

            var sql = "INSERT INTO sheet(`grid_key`,`index`,`json_data`,`status`,`order`,`is_delete`,`create_time`) VALUE(@grid_key,@index,@json_data,@status,@order,@is_delete,@create_time);";
            return await _dapperHelper.ExecuteAsync(sql, param);
        }

        /// <summary>
        /// 添加sheet
        /// </summary>
        /// <param name="id">主键</param>
        /// <param name="gridKey">execl文档key</param>
        /// <param name="index">sheet下标</param>
        /// <param name="jsonData">sheet json数据</param>
        /// <param name="status">sheet状态</param>
        /// <param name="order">sheet序号</param>
        /// <param name="isDelete">是否删除</param>
        /// <returns></returns>
        public async Task<bool> Update(int id, string gridKey, string index, string jsonData, int status, int order, int isDelete)
        {
            var param = new DynamicParameters();
            param.Add("@id", id);
            param.Add("@grid_key", gridKey);
            param.Add("@index", index);
            param.Add("@json_data", jsonData);
            param.Add("@status", status);
            param.Add("@order", order);
            param.Add("@is_delete", isDelete);
            param.Add("@update_time", DateTime.Now);

            var sql = "UPDATE sheet SET `grid_key`=@grid_key,`index`=@index,`json_data`=@json_data,`status`=@status,`order`=@order,`is_delete`=@is_delete WHERE id=@id;";
            return await _dapperHelper.ExecuteAsync(sql, param);
        }

        /// <summary>
        /// 修改sheet删除状态
        /// </summary>
        /// <param name="id">主键id</param>
        /// <param name="isDelete">是否删除</param>
        /// <returns></returns>
        public async Task<bool> UpdateIsDelete(int id, int isDelete)
        {
            var param = new DynamicParameters();
            param.Add("@id", id);
            param.Add("@is_delete", isDelete);

            var sql = "UPDATE sheet SET `is_delete`=@is_delete WHERE id=@id;";
            return await _dapperHelper.ExecuteAsync(sql, param);
        }

        /// <summary>
        /// 修改sheet激活状态
        /// </summary>
        /// <param name="id">主键id</param>
        /// <param name="status">激活状态</param>
        /// <returns></returns>
        public async Task<bool> UpdateStatus(int id, int status)
        {
            var param = new DynamicParameters();
            param.Add("@id", id);
            param.Add("@status", status);

            var sql = "UPDATE sheet SET `status`=@status WHERE id=@id;";
            return await _dapperHelper.ExecuteAsync(sql, param);
        }

        /// <summary>
        /// 判断execl是否存在
        /// </summary>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        public async Task<bool> IsExist(string gridKey)
        {
            var sql = "SELECT COUNT(*) as rows FROM sheet WHERE grid_key=@grid_key;";
            return await _dapperHelper.ExecuteScalarAsync<int>(sql, new { grid_key = gridKey }) > 0;
        }

        /// <summary>
        /// 删除execl
        /// </summary>
        /// <param name="gridKey">execl文档key</param>
        /// <returns></returns>
        public async Task<bool> Delete(string gridKey)
        {
            var sql = "DELETE FROM sheet WHERE grid_key=@grid_key;";
            return await _dapperHelper.ExecuteAsync(sql, new { grid_key = gridKey });
        }
    }
}
