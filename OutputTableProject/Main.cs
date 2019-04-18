using OutputTableCroject;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tutorial.SqlConn;

namespace OutputTableProject
{
    public partial class Main : Form
    {
        //Главный класс управления программой для вызова нужных методов
        public Main()
        {
            InitializeComponent();
            SqlCommand.Spisok(comboBox1, "Otdel", "Name");
        }

        private void load_button_Click(object sender, EventArgs e)
        {
            LoadReportToDB.LoadReport(textBox1, textBox2, textBox3, textBox4, textBox5, textBox6,  textBox7,  textBox8, textBox9, textBox10, textBox11,
             textBox12, textBox13, textBox14, textBox15, textBox16,  textBox17, textBox18, comboBox1, false , 0);
        }

        private void get_calc_tables_button_Click(object sender, EventArgs e)
        {
            GetCalcTables.ConvertCalcTables();

        }

        private void get_output_table_button_Click(object sender, EventArgs e)
        {
            GetReport.GetOutputTable(textBox20);
            ConvertToExcel.ConvertToXLS();
        }

        private void Main_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void оПрограммеToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Messages.MessageShow("info");
        }

        private void загрузкаОтчётаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Messages.MessageShow("loadreport");
        }

        private void получениеРасчётныхТаблицToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Messages.MessageShow("getcalctables");
        }

        private void получениеВедомостиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Messages.MessageShow("getreport");
        }
    }
}

