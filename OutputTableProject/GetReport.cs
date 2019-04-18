using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;
using Tutorial.SqlConn;

namespace OutputTableProject
{
    //Класс для формирования ведомости в базе данных
    class GetReport
    {
        private static ComboBox comboBox2;

        public static void GetOutputTable(TextBox textBox20)
        {
            string date = "";
            date = textBox20.Text;
            SqlCommand sqlcom = new SqlCommand();
            List<int> numbers = new List<int> { };

            //Делаем запрос всех Id с указанной датой
            SqlConnection conn = DBUtils.GetDBConnection();
            conn.Open();
            try
            {
                string sql = "Select Id from InputTable where date = " + date;

                System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = sql;
                
                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        
                        while (reader.Read())
                        {                         
                            int colIdIndex = reader.GetOrdinal("Id");
                            int colId = reader.GetInt32(colIdIndex);
                            numbers.Add(colId);
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
            //Заполнение выходной таблицы данными


            decimal dec1 = 0, dec2 = 0, dec3 = 0;
            int n,j = 0;
            string s = "";
            int[] id = numbers.ToArray<int>();

            SqlCommand.Query("Delete from OutputTable where Id != 0");
            SqlCommand.Query("DBCC CHECKIDENT('OutputTable', RESEED, " + 0 + ")");
           
            for (int i = 0; i < id.Length; i++) {
                LoadReportToDB.LoadReport(textBox20, textBox20, textBox20, textBox20, textBox20, textBox20, textBox20, textBox20, textBox20, textBox20, textBox20,
             textBox20, textBox20, textBox20, textBox20, textBox20, textBox20, textBox20, comboBox2, true, id[i]);

                SqlCommand.Query("Insert into OutputTable(OrderNum) values (" + (i+1) + ")");

                n = sqlcom.SelectInt("OrderNum", "InputTable", "Id = " + id[i]);
                SqlCommand.Query("Update OutputTable set OrderNum = " + n + " where Id = " + (i+1));

                n = sqlcom.SelectInt("OtdelId", "InputTable", "Id = " + id[i]);
                s = sqlcom.SelectStr("Name", "Otdel", "Id = " + n);
                SqlCommand.Query("Update OutputTable set OtdelId = '" + s + "' where Id = " + (i+1));

                s = sqlcom.SelectStr("WorkName", "InputTable", "Id = " + id[i]);
                SqlCommand.Query("Update OutputTable set WorkName = '" + s + "' where Id = " + (i+1));

                n = sqlcom.SelectInt("ObFact", "InputTable", "Id = " + id[i]);
                dec1 = (decimal)n;
                SqlCommand.Query("Update OutputTable set Vol = " + dec1 + " where Id = " + (i+1));

                n = sqlcom.SelectInt("Tiraj", "InputTable", "Id = " + id[i]);
                dec1 = (decimal)n;
                SqlCommand.Query("Update OutputTable set Tiraj = " + dec1 + " where Id = " + (i+1));

                dec1 = sqlcom.SelectSum("Sum", "PrintOnOfset");
                dec2 = sqlcom.SelectSum("Cost", "TirajOnColPrint");
                dec3 = dec1 + dec2;
                dec1 = sqlcom.SelectSum("Cost", "TirajOnKseroks");
                dec2 = sqlcom.SelectSum("Cost", "TirajOnRizograph");
                dec3 = dec3 + dec1 + dec2;
                s = dec3.ToString();
                s = s.Replace(",", ".");
                SqlCommand.Query("Update OutputTable set CostofDoneWork = " + s + " where Id = " + (i+1));

                dec1 = sqlcom.Select("Sum", "PaperExpense", "Id = 1");
                s = dec1.ToString();
                s = s.Replace(",", ".");
                SqlCommand.Query("Update OutputTable set PaperOfset65 = " + s + " where Id = " + (i + 1));

                dec1 = sqlcom.Select("Sum", "PaperExpense", "Id = 2");
                s = dec1.ToString();
                s = s.Replace(",", ".");
                SqlCommand.Query("Update OutputTable set PaperOfset80 = " + s + " where Id = " + (i + 1));

                dec1 = sqlcom.Select("Sum", "PaperExpense", "Id = 3");
                s = dec1.ToString();
                s = s.Replace(",", ".");
                SqlCommand.Query("Update OutputTable set PaperOfset120 = " + s + " where Id = " + (i + 1));

                dec1 = sqlcom.Select("Sum", "PaperExpense", "Id = 4");
                s = dec1.ToString();
                s = s.Replace(",", ".");
                SqlCommand.Query("Update OutputTable set PaperMag48 = " + s + " where Id = " + (i + 1));

                dec1 = sqlcom.Select("AmountPages", "PaperExpense", "Id = 5");
                s = dec1.ToString();
                s = s.Replace(",", ".");
                SqlCommand.Query("Update OutputTable set PaperMel200 = " + s + " where Id = " + (i + 1));

                dec1 = sqlcom.Select("AmountPages", "PaperExpense", "Id = 6");
                s = dec1.ToString();
                s = s.Replace(",", ".");
                SqlCommand.Query("Update OutputTable set PaperMel250 = " + s + " where Id = " + (i + 1));

                dec1 = sqlcom.Select("AmountPages", "PaperExpense", "Id = 7");
                s = dec1.ToString();
                s = s.Replace(",", ".");
                SqlCommand.Query("Update OutputTable set PaperMel115 = " + s + " where Id = " + (i + 1));

                dec1 = sqlcom.Select("AmountPages", "PaperExpense", "Id = 8");
                s = dec1.ToString();
                s = s.Replace(",", ".");
                SqlCommand.Query("Update OutputTable set PaperMelKart = " + s + " where Id = " + (i + 1));

                dec1 = sqlcom.Select("AmountPages", "PaperExpense", "Id = 9");
                s = dec1.ToString();
                s = s.Replace(",", ".");
                SqlCommand.Query("Update OutputTable set ColorPaper = " + s + " where Id = " + (i + 1));

                j = i;
            }
            //Суммирование расходов и вывод в ИТОГ
            SqlCommand.Query("Insert into OutputTable(WorkName) values ('Итого:')");

            dec1 = sqlcom.SelectSum("CostOfDoneWork", "OutputTable");
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update OutputTable set CostofDoneWork = " + s + " where Id = " + (j + 2));

            dec1 = sqlcom.SelectSum("PaperOfset65", "OutputTable");
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update OutputTable set PaperOfset65 = " + s + " where Id = " + (j + 2));

            dec1 = sqlcom.SelectSum("PaperOfset80", "OutputTable");
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update OutputTable set PaperOfset80 = " + s + " where Id = " + (j + 2));

            dec1 = sqlcom.SelectSum("PaperOfset120", "OutputTable");
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update OutputTable set PaperOfset120 = " + s + " where Id = " + (j + 2));

            dec1 = sqlcom.SelectSum("PaperMag48", "OutputTable");
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update OutputTable set PaperMag48 = " + s + " where Id = " + (j + 2));

            dec1 = sqlcom.SelectSum("PaperMel200", "OutputTable");
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update OutputTable set PaperMel200 = " + s + " where Id = " + (j + 2));

            dec1 = sqlcom.SelectSum("PaperMel250", "OutputTable");
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update OutputTable set PaperMel250 = " + s + " where Id = " + (j + 2));

            dec1 = sqlcom.SelectSum("PaperMel115", "OutputTable");
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update OutputTable set PaperMel115 = " + s + " where Id = " + (j + 2));

            dec1 = sqlcom.SelectSum("PaperMelKart", "OutputTable");
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update OutputTable set PaperMelKart = " + s + " where Id = " + (j + 2));

            dec1 = sqlcom.SelectSum("ColorPaper", "OutputTable");
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update OutputTable set ColorPaper = " + s + " where Id = " + (j + 2));

        }
    }
}
