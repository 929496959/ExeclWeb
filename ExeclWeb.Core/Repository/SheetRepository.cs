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
        /// 根据execl所有sheet
        /// </summary>
        /// <param name="gridKey">execl主键</param>
        /// <returns></returns>
        public async Task<IEnumerable<Sheet1>> GetSheets(string gridKey)
        {
            var sql = "SELECT * FROM sheet WHERE grid_key=@grid_key;";
            return await _dapperHelper.QueryAsync<Sheet1>(sql, new { grid_key = gridKey });
        }

        /// <summary>
        /// 添加sheet
        /// </summary>
        /// <param name="gridKey">execl主键</param>
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

            var insertSql = "INSERT INTO sheet(`grid_key`,`index`,`json_data`,`status`,`order`,`is_delete`,`create_time`) VALUE(@grid_key,@index,@json_data,@status,@order,@is_delete,@create_time);";
            return await _dapperHelper.ExecuteAsync(insertSql, param);
        }

        /// <summary>
        /// 添加sheet
        /// </summary>
        /// <param name="id">主键</param>
        /// <param name="gridKey">execl主键</param>
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

            var updateSql = "UPDATE sheet SET `grid_key`=@grid_key,`index`=@index,`json_data`=@json_data,`status`=@status,`order`=@order,`is_delete`=@is_delete WHERE id=@id;";
            return await _dapperHelper.ExecuteAsync(updateSql, param);
        }

        /// <summary>
        /// 判断是否存在
        /// </summary>
        /// <param name="gridKey">execl主键</param>
        /// <returns></returns>
        public async Task<bool> IsExist(string gridKey)
        {
            var sql = "SELECT COUNT(*) as rows FROM sheet WHERE grid_key=@grid_key;";
            return await _dapperHelper.ExecuteScalarAsync<int>(sql, new { grid_key = gridKey }) > 0;
        }

        /// <summary>
        /// 删除
        /// </summary>
        /// <param name="gridKey">execl主键</param>
        /// <returns></returns>
        public async Task<bool> Delete(string gridKey)
        {
            var sql = "DELETE FROM sheet WHERE grid_key=@grid_key;";
            return await _dapperHelper.ExecuteAsync(sql, new { grid_key = gridKey });
        }
    }
}
