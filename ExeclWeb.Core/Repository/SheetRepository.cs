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
        /// 获取全部
        /// </summary>
        /// <returns></returns>
        public async Task<List<sheet>> GetSheetsAsync()
        {
            var sql = "SELECT t1.id,t1.sheet_id,t1.json_data,t1.create_time,t1.update_time FROM sheet t1;";
            var list = await _dapperHelper.QueryFirstAsync<List<sheet>>(sql);
            return list;
        }

        /// <summary>
        /// 获取
        /// </summary>
        /// <param name="sheetId">工作簿id</param>
        /// <returns></returns>
        public async Task<sheet> GetSheetAsync(string sheetId)
        {
            var sql = "SELECT t1.id,t1.sheet_id,t1.json_data,t1.create_time,t1.update_time FROM sheet t1 WHERE t1.sheet_id=@sheet_id;";
            var entity = await _dapperHelper.QueryFirstAsync<sheet>(sql, new { sheet_id = sheetId });
            return entity;
        }

        /// <summary>
        /// 添加
        /// </summary>
        /// <param name="sheetId">工作簿id</param>
        /// <param name="jsonData">json_data数据</param>
        /// <returns></returns>
        public async Task<bool> AddSheetAsync(string sheetId, string jsonData)
        {
            var param = new DynamicParameters();
            param.Add("@sheet_id", sheetId);
            param.Add("@json_data", jsonData);
            param.Add("@create_time", DateTime.Now);

            var insertSql = "INSERT INTO sheet(sheet_id,json_data,create_time) VALUE(@sheet_id,@json_data,@create_time);";
            return await _dapperHelper.ExecuteAsync(insertSql, param);
        }

        /// <summary>
        /// 修改
        /// </summary>
        /// <param name="sheetId">工作簿id</param>
        /// <param name="jsonData">json_data数据</param>
        /// <returns></returns>
        public async Task<bool> UpdateSheetAsync(string sheetId, string jsonData)
        {
            var param = new DynamicParameters();
            param.Add("@sheet_id", sheetId);
            param.Add("@json_data", jsonData);
            param.Add("@update_time", DateTime.Now);

            var updateSql = "UPDATE sheet SET json_data=@json_data,update_time=@update_time WHERE sheet_id=@sheet_id;";
            return await _dapperHelper.ExecuteAsync(updateSql, param);
        }

        /// <summary>
        /// 删除
        /// </summary>
        /// <param name="sheetId">工作簿id</param>
        /// <returns></returns>
        public async Task<bool> DeleteSheetAsync(string sheetId)
        {
            var sql = "DELETE FROM sheet WHERE sheet_id=@sheet_id;";
            return await _dapperHelper.ExecuteAsync(sql, new { sheet_id = sheetId });
        }

        /// <summary>
        /// 判断是否存在
        /// </summary>
        /// <param name="sheetId">工作簿id</param>
        /// <returns></returns>
        public async Task<bool> IsExist(string sheetId)
        {
            var isExistSql = "SELECT COUNT(*) as rows FROM sheet WHERE sheet_id=@sheet_id;";
            return await _dapperHelper.ExecuteScalarAsync<int>(isExistSql, new { sheet_id = sheetId }) > 0;
        }

        /// <summary>
        /// 提交
        /// </summary>
        /// <param name="sheetId">工作簿id</param>
        /// <param name="jsonData">工作簿数据</param>
        /// <returns></returns>
        public async Task<bool> SubmitSheetAsync(string sheetId, string jsonData)
        {
            using (var dbConnection = _dapperHelper.Connection())
            {
                dbConnection.Open();
                using (var transaction = dbConnection.BeginTransaction())
                {
                    var param = new DynamicParameters();
                    param.Add("@sheet_id", sheetId);
                    param.Add("@json_data", jsonData);
                    try
                    {
                        // 1.先判断当前文档是否存在
                        var isExistSql = "SELECT COUNT(*) as rows FROM sheet WHERE sheet_id=@sheet_id;";
                        var isExist = await dbConnection.ExecuteScalarAsync<int>(isExistSql, new { sheet_id = sheetId }) > 0;
                        if (isExist)
                        {
                            param.Add("@update_time", DateTime.Now);
                            var updateSql = "UPDATE sheet SET json_data=@json_data,update_time=@update_time WHERE sheet_id=@sheet_id;";
                            await dbConnection.ExecuteAsync(updateSql, param);
                        }
                        else
                        {
                            param.Add("@create_time", DateTime.Now);
                            var insertSql = "INSERT INTO sheet(sheet_id,json_data,create_time) VALUE(@sheet_id,@json_data,@create_time);";
                            await dbConnection.ExecuteAsync(insertSql, param);
                        }
                        transaction.Commit();
                        return true;
                    }
                    catch (Exception)
                    {
                        transaction.Rollback();
                        return false;
                    }
                }
            }
        }
    }
}
