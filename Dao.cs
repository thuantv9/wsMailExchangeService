using System;
using System.Data;
using System.Data.SqlClient;

namespace wsEmailExchange
{
    public class Dao
    {
        public static bool GetConnect(ref SqlConnection conn, string connStr, ref string errMsg)
        {
            try
            {
                conn = new SqlConnection(connStr);
                conn.Open();
                return true;
            }
            catch (Exception ex)
            {
                Common.GhiLog("GetConnect", ex, "Connection String : " + connStr);
                errMsg = ex.Message;
                return false;
            }
        }

        public static DataTable GetTable(string TSQL, SqlConnection conn)
        {
            try
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (DataTable dataTable = new DataTable())
                {
                    using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(TSQL, conn))
                    {
                        sqlDataAdapter.SelectCommand.CommandTimeout = 60000;
                        sqlDataAdapter.Fill(dataTable);
                    }
                    return dataTable;
                }
            }
            catch (Exception ex)
            {
                Common.GhiLog("GetTable", ex, TSQL);
                throw ex;
            }
            finally
            {
                conn.Close();
            }
        }

        public static int ExecSQL(string TSQL, SqlConnection conn)
        {
            try
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand sqlCommand = new SqlCommand(TSQL, conn))
                {
                    sqlCommand.CommandTimeout = 60000;
                    return sqlCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Common.GhiLog("ExecSQL", ex, TSQL);
                throw ex;
            }
            finally
            {
                conn.Close();
            }
        }



        public static int ExecSQLGetID(string TSQL, SqlConnection conn)
        {
            try
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand sqlCommand = new SqlCommand(TSQL, conn))
                {
                    sqlCommand.CommandTimeout = 60000;
                    return (int)sqlCommand.ExecuteScalar();
                }
            }
            catch (Exception ex)
            {
                Common.GhiLog("Exec_SQL_GetID", ex, TSQL);
                throw ex;
            }
            finally
            {
                conn.Close();
            }
        }

        public static void UpdateDataTable(DataTable dt, string TABLE_NAME, SqlConnection conn)
        {
            try
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM " + TABLE_NAME + " WHERE 1<>1", conn))
                {
                    using (new SqlCommandBuilder(adapter))
                        adapter.Update(dt);
                }
            }
            catch (Exception ex)
            {
                Common.GhiLog("UpdateDataTable", ex, "");
                throw ex;
            }
            finally
            {
                conn.Close();
            }
        }
    }
}
