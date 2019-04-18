using System;
using System.Data.Common;
using System.Data.SqlClient;
using System.Windows.Forms;
using Tutorial.SqlConn;

namespace OutputTableProject
{
    public class SqlCommand
    {
        //Выполнение Команды SQL
        public static void Query(String sqlquery)
        {
            SqlConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            try
            {
                string sql = sqlquery;

                System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand
                {
                    Connection = conn,
                    CommandText = sql
                };
                cmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e);
                MessageBox.Show(e.StackTrace);
            }
            finally
            {
                conn.Dispose();
                conn = null;
            }
            Console.Read();
        }
        
        //Команда SELECT c типом данных DECIMAL 
        public decimal Select(String col, String tablename, string val)
        {
            SqlConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            decimal ret = 0;
            try
            {
                string sql = "Select " + col + " from " + tablename + " where " + val;

                System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = sql;
                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            int colNameIndex = reader.GetOrdinal(col);
                            decimal colName = reader.GetDecimal(colNameIndex);
                            ret = colName;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e);
                MessageBox.Show(e.StackTrace);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
            Console.Read();
            return ret;
        }
        
        //Команда SELECT с типом данных INT
        public int SelectInt(String col, String tablename, string val)
        {
            SqlConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            int ret = 0;
            try
            {
                string sql = "Select " + col + " from " + tablename + " where " + val;

                System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = sql;
                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            int colNameIndex = reader.GetOrdinal(col);
                            int colName = reader.GetInt32(colNameIndex);
                            ret = colName;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e);
                MessageBox.Show(e.StackTrace);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
            Console.Read();
            return ret;
        }
        
        //Команда SELECT с типом данных STRING
        public string SelectStr(String col, String tablename, string val)
        {
            SqlConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            string ret = "";
            try
            {
                string sql = "Select " + col + " from " + tablename + " where " + val;

                System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = sql;
                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            int colNameIndex = reader.GetOrdinal(col);
                            string colName = reader.GetString(colNameIndex);
                            ret = colName;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e);
                MessageBox.Show(e.StackTrace);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
            Console.Read();
            return ret;
        }

        //Команда SELECT для суммирования ячеек
        public decimal SelectSum(String col, String tablename)
        {
            SqlConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            decimal ret = 0;
            try
            {
                string sql = "Select sum(" + col + ") from " + tablename;

                System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = sql;
                ret = ((decimal)cmd.ExecuteScalar());
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e);
                MessageBox.Show(e.StackTrace);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
            Console.Read();
            return ret;
        }

        //Команда для добавления всех отделов в выпадающий список
        public static void Spisok(ComboBox comboBox1, String tablename, String col)
        {
            SqlConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            try
            {
                string sql = "Select " + col + " from " + tablename;

                System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = sql;
                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            int colNameIndex = reader.GetOrdinal(col);
                            string colName = reader.GetString(colNameIndex);
                            comboBox1.Items.Add(colName);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e);
                MessageBox.Show(e.StackTrace);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }
            Console.Read();
        }
    }
}

