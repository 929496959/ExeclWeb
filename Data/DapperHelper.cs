using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using Dapper;
using Common;

namespace Data
{
    /// <summary>
    /// Dapper帮助类
    /// </summary>
    public class DapperHelper
    {
        private readonly ConfigHelper _configHelper = new ConfigHelper();
        private readonly string _connectionString;
        public DapperHelper()
        {
            _connectionString = _configHelper.GetValue<string>("MySQLConnection");
        }

        public IDbConnection Connection()
        {
            return new MySqlConnection(_connectionString);
        }

        #region Execute

        public virtual bool Execute(string sql, object param = null)
        {
            using (var conn = Connection())
            {
                return conn.Execute(sql, param) > 0;
            }
        }
        public virtual async Task<bool> ExecuteAsync(string sql, object param = null)
        {
            using (var conn = Connection())
            {
                return await conn.ExecuteAsync(sql, param) > 0;
            }
        }

        #endregion

        #region ExecuteScalar

        public virtual T ExecuteScalar<T>(string sql, object param = null)
        {
            using (var conn = Connection())
            {
                return conn.ExecuteScalar<T>(sql, param);
            }
        }
        public virtual async Task<T> ExecuteScalarAsync<T>(string sql, object param = null)
        {
            using (var conn = Connection())
            {
                return await conn.ExecuteScalarAsync<T>(sql, param);
            }
        }

        #endregion

        #region Query

        public virtual IEnumerable<T> Query<T>(string sql, object param = null)
        {
            using (var conn = Connection())
            {
                return conn.Query<T>(sql, param);
            }
        }
        public virtual async Task<IEnumerable<T>> QueryAsync<T>(string sql, object param = null)
        {
            using (var conn = Connection())
            {
                return await conn.QueryAsync<T>(sql, param);
            }
        }

        #endregion

        #region QueryFirst

        public virtual T QueryFirst<T>(string sql, object param = null)
        {
            using (var conn = Connection())
            {
                return conn.QueryFirstOrDefault<T>(sql, param);
            }
        }
        public virtual async Task<T> QueryFirstAsync<T>(string sql, object param = null)
        {
            using (var conn = Connection())
            {
                return await conn.QueryFirstOrDefaultAsync<T>(sql, param);
            }
        }

        #endregion
    }
}
