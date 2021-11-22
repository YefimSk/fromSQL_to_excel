using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace SqlToExcel
{
    public partial class Form1 : Form
    {
        DataSet datset;
        SqlDataAdapter adapter;
        SqlCommandBuilder commandBuilder;
        public string CmdatTabext = "SELECT * FROM Products";
        public string conct = @"Data Source=USER-PC;" +
                "Initial Catalog=Northwind;" +
                "Integrated Security=SSPI;";
        public Form1()
        {

        InitializeComponent();
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.AllowUserToAddRows = false;

            using (SqlConnection connection = new SqlConnection(conct))
            {
                connection.Open();
                adapter = new SqlDataAdapter(CmdatTabext, connection);

                datset = new DataSet();
                adapter.Fill(datset);
                dataGridView1.DataSource = datset.Tables[0];
            }


            }

        //Екземпляр приложения Excel
        Excel.Application appl;
        //Лист
        Excel.Worksheet lisst;
        //Выделеная область
        Excel.Range lisstR;

        private void button1_Click(object sender, EventArgs e)
        {
            appl = new Excel.Application();

            try
            {
                //добавляем книгу
                appl.Workbooks.Add(Type.Missing);

                //делаем временно неактивным документ
                appl.Interactive = true;
                appl.EnableEvents = true;
                //appl.Visible = true;
                //appl.ScreenUpdating = true;

                //выбираем лист на котором будем работать (Лист 1)
                lisst = (Excel.Worksheet)appl.Sheets[1];
                //Название листа
                lisst.Name = "Таблица Products";

                //Выгрузка данных
                DataTable datTab = GetData();

                int cln = 0;
                int rln = 0;
                string data = "";

                //называем колонки
                for (int i = 0; i < datTab.Columns.Count; i++)
                {
                    data = datTab.Columns[i].ColumnName.ToString();
                    lisst.Cells[1, i + 1] = data;

                    //выделяем первую строку
                    lisstR = lisst.get_Range("A1:Z1", Type.Missing);

                    //делаем полужирный текст и перенос слов
                    lisstR.WrapText = true;
                    lisstR.Font.Bold = true;
                }

                //заполняем строки
                for (rln = 0; rln < datTab.Rows.Count; rln++)
                {
                    for (cln = 0; cln < datTab.Columns.Count; cln++)
                    {
                        data = datTab.Rows[rln].ItemArray[cln].ToString();
                        lisst.Cells[rln + 2, cln + 1] = data;
                    }
                }

                //выбираем всю область данных
                lisstR = lisst.UsedRange;

                //выравниваем строки и колонки по их содержимому
                lisstR.Columns.AutoFit();
                lisstR.Rows.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                //Показываем ексель
                appl.Visible = true;

                appl.Interactive = true;
                appl.ScreenUpdating = true;
                appl.UserControl = true;

                //Отсоединяемся от Excel
                releaseObject(lisstR);
                releaseObject(lisst);
                releaseObject(appl);
            }
        }

        private DataTable GetData()
        {
            //строка соединения
            string conct = "Data Source= USER-PC;" +
                "Initial Catalog=Northwind;" +
                "Integrated Security=SSPI; " +
                "Connection Timeout=260";

            //соединение
            SqlConnection con = new SqlConnection(conct);

            DataTable datTab = new DataTable();
            try
            {
                string shtbl = @"SELECT * FROM Products";
                SqlCommand comm = new SqlCommand(shtbl, con);

                con.Open();
                SqlDataAdapter datadap = new SqlDataAdapter(comm);
                DataSet datset = new DataSet();
                datadap.Fill(datset);
                datTab = datset.Tables[0];
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                con.Close();
                con.Dispose();
            }
            return datTab;
        }

        //Освобождаем ресуры (закрываем Excel)
        void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
