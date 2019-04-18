using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tutorial.SqlConn;


namespace OutputTableProject
{
    //Класс для добавления данных в базу данных
    public class LoadReportToDB
    {
        public static void LoadReport(TextBox textbox1, TextBox textbox2, TextBox textbox3, TextBox textbox4, TextBox textbox5, TextBox textbox6, TextBox textbox7, 
                        TextBox textbox8, TextBox textbox9, TextBox textbox10, TextBox textbox11,TextBox textbox12, TextBox textbox13, TextBox textbox14, 
                        TextBox textbox15, TextBox textbox16, TextBox textbox17, TextBox textbox18, ComboBox comboBox1, Boolean fromDB, int id)
        {
            string otdel = "", ordnum = "", tir = "", fact = "", plist = "", colpap4 = "", colpap3 = "", pap65 = "", pap80 = "", pap120 = "", gazpap = "", mel200 = "", mel220 = "", mel115 = "",
                    melkar = "", melcol = "", wname = "", form = "", date = "", str = "";
            int n = 0;
            if (fromDB == true)
            {
                SqlCommand sqlque = new SqlCommand();
                n = sqlque.SelectInt("OrderNum", "InputTable", "Id = " + id);
                ordnum = n.ToString();

                n = sqlque.SelectInt("OtdelId", "InputTable", "Id = " + id);
                str = sqlque.SelectStr("Name", "Otdel", "Id = " + n);
                otdel = str;

                str = sqlque.SelectStr("WorkName", "InputTable", "Id = " + id);
                wname = str;

                str = sqlque.SelectStr("Format", "InputTable", "Id = " + id);
                form = str;

                n = sqlque.SelectInt("Tiraj", "InputTable", "Id = " + id);
                tir = n.ToString();

                n = sqlque.SelectInt("ObFact", "InputTable", "Id = " + id);
                fact = n.ToString();

                n = sqlque.SelectInt("ObPrintList", "InputTable", "Id = " + id);
                plist = n.ToString();

                n = sqlque.SelectInt("ColPageA4", "InputTable", "Id = " + id);
                colpap4 = n.ToString();

                n = sqlque.SelectInt("ColPageA3", "InputTable", "Id = " + id);
                colpap3 = n.ToString();

                n = sqlque.SelectInt("Paper65A3", "InputTable", "Id = " + id);
                pap65 = n.ToString();

                n = sqlque.SelectInt("Paper80A3", "InputTable", "Id = " + id);
                pap80 = n.ToString();

                n = sqlque.SelectInt("Paper120A3", "InputTable", "Id = " + id);
                pap120 = n.ToString();

                n = sqlque.SelectInt("PaperMagA3", "InputTable", "Id = " + id);
                gazpap = n.ToString();

                n = sqlque.SelectInt("PaperMel200", "InputTable", "Id = " + id);
                mel200 = n.ToString();

                n = sqlque.SelectInt("PaperMel220", "InputTable", "Id = " + id);
                mel220 = n.ToString();

                n = sqlque.SelectInt("PaperMel115", "InputTable", "Id = " + id);
                mel115 = n.ToString();

                n = sqlque.SelectInt("PaperMelKart", "InputTable", "Id = " + id);
                melkar = n.ToString();

                n = sqlque.SelectInt("ColPage", "InputTable", "Id = " + id);
                melcol = n.ToString();
            }
            else
            {
                ordnum = textbox16.Text;
                otdel = comboBox1.Text;
                wname = textbox18.Text;
                form = textbox1.Text;
                date = textbox17.Text;
                tir = textbox2.Text;
                fact = textbox3.Text;
                plist = textbox4.Text;
                colpap4 = textbox5.Text;
                colpap3 = textbox6.Text;
                pap65 = textbox7.Text;
                pap80 = textbox8.Text;
                pap120 = textbox9.Text;
                gazpap = textbox10.Text;
                mel200 = textbox11.Text;
                mel220 = textbox12.Text;
                mel115 = textbox13.Text;
                melkar = textbox14.Text;
                melcol = textbox15.Text;

                //Добавление id отдела
                SqlConnection conn = DBUtils.GetDBConnection();
                conn.Open();
                try
                {
                    string sql = "Select Id from Otdel where Name = '" + otdel + "'";
                    System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                    cmd.Connection = conn;
                    cmd.CommandText = sql;
                    using (DbDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                int empIdIndex = reader.GetOrdinal("Id");
                                long empId = Convert.ToInt64(reader.GetValue(0));
                                string Id = empId.ToString();
                                otdel = Id;
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

                SqlConnection connection = DBUtils.GetDBConnection();
                connection.Open();
                try
                {
                    string sql = "Insert into InputTable (OrderNum, OtdelId, WorkName, Format, Tiraj, ObFact, ObPrintList, ColPageA4, ColPageA3, Paper65A3, Paper80A3," +
                        "Paper120A3, PaperMagA3, PaperMel200, PaperMel220, PaperMel115, PaperMelKart, ColPage, Date) "
                                                         + " values ( " + ordnum + "," + otdel + "," + "'" + wname + "'" + "," + "'" + form + "'" + "," + tir + "," + fact + "," + plist + "," + colpap4 + "," + colpap3 + "," + pap65 + "," + pap80 + "," + pap120 + "," + gazpap + "," + mel200 + "," +
                                                           mel220 + "," + mel115 + "," + melkar + "," + melcol + "," + "'" + date + "'" + ") ";
                    System.Data.SqlClient.SqlCommand cmd = connection.CreateCommand();
                    cmd.CommandText = sql;
                    cmd.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Error: " + e);
                    MessageBox.Show(e.StackTrace);
                }
                finally
                {
                    connection.Close();
                    connection.Dispose();
                    connection = null;
                }
                Console.Read();



            }                                 
            //Добавление данных в расчётные таблицы
            decimal dec1 = 0 , dec2 = 0, dec3 = 0;
            int n1 = 0, n2 = 0;
            string s = "";
            SqlCommand sqlsel = new SqlCommand();

            //----------------------------------------------------------------------------------------------
            //Печать на офсетных
            SqlCommand.Query("Update PrintOnOfset set Amount = " + plist + " where Id = 1");

            Int32.TryParse(plist, out n1);
            Int32.TryParse(tir, out n2);       
            SqlCommand.Query("Update PrintOnOfset set Amount = " + n1*n2 + " where Id = 2");
            
            dec1 = sqlsel.Select("Price", "PrintOnOfset", "Id = 3");
            dec1 = dec1 * n1;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PrintOnOfset set Amount = " + s  + " where Id = 3");

            Int32.TryParse(fact, out n1);
            SqlCommand.Query("Update PrintOnOfset set Amount = " + ((n1 * 1000)/16) + " where Id = 4");
            
            //Подсчет суммы в расчётной таблице печати на офсетных
            dec1 = sqlsel.Select("Price", "PrintOnOfset", "Id = 1");
            dec2 = sqlsel.Select("Amount", "PrintOnOfset", "Id = 1");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PrintOnOfset set Sum = " + s + " where Id = 1");

            dec1 = sqlsel.Select("Price", "PrintOnOfset", "Id = 2");
            dec2 = sqlsel.Select("Amount", "PrintOnOfset", "Id = 2");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PrintOnOfset set Sum = " + s + " where Id = 2");

            dec1 = sqlsel.Select("Price", "PrintOnOfset", "Id = 3");
            dec2 = sqlsel.Select("Amount", "PrintOnOfset", "Id = 3");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PrintOnOfset set Sum = " + s + " where Id = 3");

            dec1 = sqlsel.Select("Price", "PrintOnOfset", "Id = 4");
            dec2 = sqlsel.Select("Amount", "PrintOnOfset", "Id = 4");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PrintOnOfset set Sum = " + s + " where Id = 4");

            //----------------------------------------------------------------------------------------------
            //Тиражирование на ризографе
            Int32.TryParse(plist, out n1);
            Int32.TryParse(tir, out n2);
            SqlCommand.Query("Update TirajOnRizograph set Amount = " + n1 + " where Id = 1");
            SqlCommand.Query("Update TirajOnRizograph set Amount = " + n1*n2 + " where Id = 2");
            
            //Подсчет суммы в расчётной таблице тиражирования на ризографе
            dec1 = sqlsel.Select("Price", "TirajOnRizograph", "Id = 1");
            dec2 = sqlsel.Select("Amount", "TirajOnRizograph", "Id = 1");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update TirajOnRizograph set Cost = " + s + " where Id = 1");

            dec1 = sqlsel.Select("Price", "TirajOnRizograph", "Id = 2");
            dec2 = sqlsel.Select("Amount", "TirajOnRizograph", "Id = 2");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update TirajOnRizograph set Cost = " + s + " where Id = 2");

            //----------------------------------------------------------------------------------------------
            //Тиражирование на цветном принтере
            SqlCommand.Query("Update TirajOnColPrint set Vol = " + colpap3 + " where Id = 1");
            SqlCommand.Query("Update TirajOnColPrint set Tiraj = " + tir + " where Id = 1");

            SqlCommand.Query("Update TirajOnColPrint set Vol = " + colpap4 + " where Id = 2");
            SqlCommand.Query("Update TirajOnColPrint set Tiraj = " + tir + " where Id = 2");

            SqlCommand.Query("Update TirajOnColPrint set Vol = " + colpap3 + " where Id = 3");
            SqlCommand.Query("Update TirajOnColPrint set Tiraj = " + tir + " where Id = 3");

            SqlCommand.Query("Update TirajOnColPrint set Vol = " + colpap4 + " where Id = 4");
            SqlCommand.Query("Update TirajOnColPrint set Tiraj = " + tir + " where Id = 4");
            
            //Подсчет суммы в расчётной таблице тиражирования на цветном принтере
            dec1 = sqlsel.Select("Vol", "TirajOnColPrint", "Id = 1");
            dec2 = sqlsel.Select("Tiraj", "TirajOnColPrint", "Id = 1");
            dec3 = sqlsel.Select("Price", "TirajOnColPrint", "Id = 1");
            dec1 = dec1 * dec2 * dec3;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update TirajOnColPrint set Cost = " + s + " where Id = 1");

            dec1 = sqlsel.Select("Vol", "TirajOnColPrint", "Id = 2");
            dec2 = sqlsel.Select("Tiraj", "TirajOnColPrint", "Id = 2");
            dec3 = sqlsel.Select("Price", "TirajOnColPrint", "Id = 2");
            dec1 = dec1 * dec2 * dec3;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update TirajOnColPrint set Cost = " + s + " where Id = 2");

            dec1 = sqlsel.Select("Vol", "TirajOnColPrint", "Id = 3");
            dec2 = sqlsel.Select("Tiraj", "TirajOnColPrint", "Id = 3");
            dec3 = sqlsel.Select("Price", "TirajOnColPrint", "Id = 3");
            dec1 = dec1 * dec2 * dec3;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update TirajOnColPrint set Cost = " + s + " where Id = 3");

            dec1 = sqlsel.Select("Vol", "TirajOnColPrint", "Id = 4");
            dec2 = sqlsel.Select("Tiraj", "TirajOnColPrint", "Id = 4");
            dec3 = sqlsel.Select("Price", "TirajOnColPrint", "Id = 4");
            dec1 = dec1 * dec2 * dec3;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update TirajOnColPrint set Cost = " + s + " where Id = 4");

            //----------------------------------------------------------------------------------------------
            //Тиражирование на ксероксе
            Int32.TryParse(plist, out n1);
            n2 = n1 / 2;
            SqlCommand.Query("Update TirajOnKseroks set Vol = " + n2 + " where Id = 1");
            SqlCommand.Query("Update TirajOnKseroks set Tiraj = " + tir + " where Id = 1");

            SqlCommand.Query("Update TirajOnKseroks set Vol = " + n1 + " where Id = 2");
            SqlCommand.Query("Update TirajOnKseroks set Tiraj = " + tir + " where Id = 2");

            SqlCommand.Query("Update TirajOnKseroks set Tiraj = " + tir + " where Id = 3");

            SqlCommand.Query("Update TirajOnKseroks set Tiraj = " + tir + " where Id = 4");

            //Подсчет суммы в расчётной таблице тиражирования ксероксе
            dec1 = sqlsel.Select("Vol", "TirajOnKseroks", "Id = 1");
            dec2 = sqlsel.Select("Tiraj", "TirajOnKseroks", "Id = 1");
            dec3 = sqlsel.Select("Price", "TirajOnKseroks", "Id = 1");
            dec1 = dec1 * dec2 * dec3;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update TirajOnKseroks set Cost = " + s + " where Id = 1");

            dec1 = sqlsel.Select("Vol", "TirajOnKseroks", "Id = 2");
            dec2 = sqlsel.Select("Tiraj", "TirajOnKseroks", "Id = 2");
            dec3 = sqlsel.Select("Price", "TirajOnKseroks", "Id = 2");
            dec1 = dec1 * dec2 * dec3;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update TirajOnKseroks set Cost = " + s + " where Id = 2");

            dec1 = sqlsel.Select("Tiraj", "TirajOnKseroks", "Id = 3");
            dec2 = sqlsel.Select("Price", "TirajOnKseroks", "Id = 3");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update TirajOnKseroks set Cost = " + s + " where Id = 3");

            dec1 = sqlsel.Select("Tiraj", "TirajOnKseroks", "Id = 4");
            dec2 = sqlsel.Select("Price", "TirajOnKseroks", "Id = 4");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update TirajOnKseroks set Cost = " + s + " where Id = 4");

            //----------------------------------------------------------------------------------------------
            //Расход бумаги
            Int32.TryParse(pap65, out n1);
            Int32.TryParse(tir, out n2);
            n1 = n1 * n2;
            SqlCommand.Query("Update PaperExpense set ToPrint = " + n1 + " where Id = 1");
            n2 = n1 * 14 / 100;
            SqlCommand.Query("Update PaperExpense set ToPrilad = " + n2 + " where Id = 1");
            n1 = n1 + n2;
            SqlCommand.Query("Update PaperExpense set AmountPages = " + n1 + " where Id = 1");
            n2 = n1 * 10 / 1000;
            SqlCommand.Query("Update PaperExpense set Sum = " + n2 + " where Id = 1");

            Int32.TryParse(pap80, out n1);
            Int32.TryParse(tir, out n2);
            n1 = n1 * n2;
            SqlCommand.Query("Update PaperExpense set ToPrint = " + n1 + " where Id = 2");
            n2 = n1 * 14 / 100;
            SqlCommand.Query("Update PaperExpense set ToPrilad = " + n2 + " where Id = 2");
            n1 = n1 + n2;
            SqlCommand.Query("Update PaperExpense set AmountPages = " + n1 + " where Id = 2");
            n2 = n1 * 10 / 1000;
            SqlCommand.Query("Update PaperExpense set Sum = " + n2 + " where Id = 2");

            Int32.TryParse(pap120, out n1);
            Int32.TryParse(tir, out n2);
            n1 = n1 * n2;
            SqlCommand.Query("Update PaperExpense set ToPrint = " + n1 + " where Id = 3");
            n2 = n1 * 14 / 100;
            SqlCommand.Query("Update PaperExpense set ToPrilad = " + n2 + " where Id = 3");
            n1 = n1 + n2;
            SqlCommand.Query("Update PaperExpense set AmountPages = " + n1 + " where Id = 3");
            n2 = n1 * 14 / 1000;
            SqlCommand.Query("Update PaperExpense set Sum = " + n2 + " where Id = 3");

            Int32.TryParse(gazpap, out n1);
            Int32.TryParse(tir, out n2);
            n1 = n1 * n2;
            SqlCommand.Query("Update PaperExpense set ToPrint = " + n1 + " where Id = 4");
            n2 = n1 * 14 / 100;
            SqlCommand.Query("Update PaperExpense set ToPrilad = " + n2 + " where Id = 4");
            n1 = n1 + n2;
            SqlCommand.Query("Update PaperExpense set AmountPages = " + n1 + " where Id = 4");
            n2 = n1 * 7 / 1000;
            SqlCommand.Query("Update PaperExpense set Sum = " + n2 + " where Id = 4");

            Int32.TryParse(mel200, out n1);
            Int32.TryParse(tir, out n2);
            n1 = n1 * n2;
            SqlCommand.Query("Update PaperExpense set ToPrint = " + n1 + " where Id = 5");
            n2 = n1 * 4 / 100;
            SqlCommand.Query("Update PaperExpense set ToPrilad = " + n2 + " where Id = 5");
            n1 = n1 + n2;
            SqlCommand.Query("Update PaperExpense set AmountPages = " + n1 + " where Id = 5");

            Int32.TryParse(mel220, out n1);
            Int32.TryParse(tir, out n2);
            n1 = n1 * n2;
            SqlCommand.Query("Update PaperExpense set ToPrint = " + n1 + " where Id = 6");
            n2 = n1 * 4 / 100;
            SqlCommand.Query("Update PaperExpense set ToPrilad = " + n2 + " where Id = 6");
            n1 = n1 + n2;
            SqlCommand.Query("Update PaperExpense set AmountPages = " + n1 + " where Id = 6");

            Int32.TryParse(mel115, out n1);
            Int32.TryParse(tir, out n2);
            n1 = n1 * n2;
            SqlCommand.Query("Update PaperExpense set ToPrint = " + n1 + " where Id = 7");
            n2 = n1 * 4 / 100;
            SqlCommand.Query("Update PaperExpense set ToPrilad = " + n2 + " where Id = 7");
            n1 = n1 + n2;
            SqlCommand.Query("Update PaperExpense set AmountPages = " + n1 + " where Id = 7");

            Int32.TryParse(melkar, out n1);
            Int32.TryParse(tir, out n2);
            n1 = n1 * n2;
            SqlCommand.Query("Update PaperExpense set ToPrint = " + n1 + " where Id = 8");
            n2 = n1 * 4 / 100;
            SqlCommand.Query("Update PaperExpense set ToPrilad = " + n2 + " where Id = 8");
            n1 = n1 + n2;
            SqlCommand.Query("Update PaperExpense set AmountPages = " + n1 + " where Id = 8");

            Int32.TryParse(melcol, out n1);
            Int32.TryParse(tir, out n2);
            n1 = n1 * n2;
            SqlCommand.Query("Update PaperExpense set ToPrint = " + n1 + " where Id = 9");
            n2 = n1 * 4 / 100;
            SqlCommand.Query("Update PaperExpense set ToPrilad = " + n2 + " where Id = 9");
            n1 = n1 + n2;
            SqlCommand.Query("Update PaperExpense set AmountPages = " + n1 + " where Id = 9");

            //Подсчет суммы в расчётной таблице по расходу бумаги
            dec1 = sqlsel.Select("Sum", "PaperExpense", "Id = 1");
            dec2 = sqlsel.Select("Price", "PaperExpense", "Id = 1");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PaperExpense set Cost = " + s + " where Id = 1");

            dec1 = sqlsel.Select("Sum", "PaperExpense", "Id = 2");
            dec2 = sqlsel.Select("Price", "PaperExpense", "Id = 2");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PaperExpense set Cost = " + s + " where Id = 2");

            dec1 = sqlsel.Select("Sum", "PaperExpense", "Id = 3");
            dec2 = sqlsel.Select("Price", "PaperExpense", "Id = 3");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PaperExpense set Cost = " + s + " where Id = 3");

            dec1 = sqlsel.Select("AmountPages", "PaperExpense", "Id = 4");
            dec2 = sqlsel.Select("Price", "PaperExpense", "Id = 4");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PaperExpense set Cost = " + s + " where Id = 4");

            dec1 = sqlsel.Select("AmountPages", "PaperExpense", "Id = 5");
            dec2 = sqlsel.Select("Price", "PaperExpense", "Id = 5");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PaperExpense set Cost = " + s + " where Id = 5");

            dec1 = sqlsel.Select("AmountPages", "PaperExpense", "Id = 6");
            dec2 = sqlsel.Select("Price", "PaperExpense", "Id = 6");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PaperExpense set Cost = " + s + " where Id = 6");

            dec1 = sqlsel.Select("AmountPages", "PaperExpense", "Id = 7");
            dec2 = sqlsel.Select("Price", "PaperExpense", "Id = 7");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PaperExpense set Cost = " + s + " where Id = 7");

            dec1 = sqlsel.Select("AmountPages", "PaperExpense", "Id = 8");
            dec2 = sqlsel.Select("Price", "PaperExpense", "Id = 8");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PaperExpense set Cost = " + s + " where Id = 8");

            dec1 = sqlsel.Select("AmountPages", "PaperExpense", "Id = 9");
            dec2 = sqlsel.Select("Price", "PaperExpense", "Id = 9");
            dec1 = dec1 * dec2;
            s = dec1.ToString();
            s = s.Replace(",", ".");
            SqlCommand.Query("Update PaperExpense set Cost = " + s + " where Id = 9");
        }

    }
}
