using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace randomProduct
{
    public partial class Form1 : Form
    {
        public static string customerName = "";
        public static double probability = 0.0;

        public Form1()
        {
            InitializeComponent();
        }

        // 添加按钮
        private void button2_Click(object sender, EventArgs e)
        {
            string name=this.textBox1.Text.Trim();
            if (!string.IsNullOrEmpty(name))
            {
                this.listBox1.Items.Add(name);
                this.textBox1.Text = "";
                this.textBox1.Focus();
            }
        }

        //回车
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                string name = this.textBox1.Text.Trim();
                if (!string.IsNullOrEmpty(name))
                {
                    this.listBox1.Items.Add(name);
                    this.textBox1.Text = "";
                    this.textBox1.Focus();
                }
            }
        }

        // 删除名单 
        private void button4_Click(object sender, EventArgs e)
        {
            string name = this.listBox1.SelectedItem.ToString();
            // MessageBox.Show(name);
            if (!string.IsNullOrEmpty(name))
            {
                this.listBox1.Items.RemoveAt(this.listBox1.SelectedIndex);
            }
        }

        // 选择文件并写入到listbox控件中
        #region btnClick 点击"选择文件"按钮, 打开选择文件对话框
        private void button1_Click(object sender, EventArgs e)
        {
            //DataTable excelDataTable = new DataTable();
            DataSet excelDataTable = new DataSet();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Files|*.xls;*.xlsx";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string filePath = "";
            if (openFileDialog.ShowDialog()==DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
                excelDataTable = ReadExcelToTable(filePath);
                for (int i=1; i<excelDataTable.Tables[0].Rows.Count;i++)
                {
                    this.listBox1.Items.Add(excelDataTable.Tables[0].Rows[i].ItemArray[0].ToString().Trim());
                }

            }
        }
        #endregion

        // 读取excel文件
        public DataSet ReadExcelToTable(string filePath)
        {
            string strConn = "";
            if (IsExcelXmlFileFormat(filePath))
            {
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=Excel 12.0;";
            }
            else
            {
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
            }
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string strExcel = "select * from [sheet1$]";
            OleDbDataAdapter da = new OleDbDataAdapter(strExcel, strConn);
            DataSet ds = new DataSet();
            try
            {
                da.Fill(ds);
            }
            catch (Exception ex)
            {
                throw new Exception("读取Excel失败：" + ex.Message);
            }
            return ds;
        }

        // 判断是否为xlsx格式文件
        private static bool IsExcelXmlFileFormat(string fileName)
        {
            return fileName.EndsWith("xlsx", StringComparison.OrdinalIgnoreCase);
        }

        // 随机抽签
        private void button3_Click(object sender, EventArgs e)
        {
            this.textBox2.Text = "";
            int list = this.listBox1.Items.Count;
            if (list>0)
            {
                if (probability==0)
                {
                    Random random = new Random();
                    int id = random.Next(0, list);
                    this.textBox2.Text = this.listBox1.Items[id].ToString();
                }
                else
                {
                    Random random = new Random();
                    int id = random.Next(0, 100);
                    int j = 0;
                    float jilu = (float)(1- probability) / (list-1) * 100;
                    for (int i =0;i<list;i++)
                    {
                        if (customerName==this.listBox1.Items[i].ToString().Trim())
                        {
                            j += (int)(probability*100);
                        }
                        else
                        {
                            j += (int)jilu;
                        }
                        if (id<j)
                        {
                            this.textBox2.Text = this.listBox1.Items[i].ToString().Trim();
                            return;
                        }
                        if (i == list - 1)
                        {
                            this.textBox2.Text = this.listBox1.Items[i].ToString().Trim();
                            return;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("列表中没有单位名单，请先输入后再点抽签");
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            //f9呼出配置项
            if (e.KeyValue==120)
            {
                Form2 form2 = new Form2();
                form2.ShowDialog();
            }

        }




        /*
        private DataTable ReadExcelToTable(string path)
        {
            OleDbConnectionStringBuilder connStringBuilder = new OleDbConnectionStringBuilder();
            connStringBuilder.DataSource = path;
            if (IsExcelXmlFileFormat(path))
            {
                connStringBuilder.Provider = "Microsoft.ACE.OLEDB.12.0";
                connStringBuilder.Add("Extended Properties", "Excel 8.0;HDR=NO;");
            }
            else
            {
                connStringBuilder.Provider = "Microsoft.Jet.OLEDB.4.0";
                connStringBuilder.Add("Extended Properties", "Excel 8.0;");
            }
            DataSet data = new DataSet();
            try
            {
                using (OleDbConnection dbConn = new OleDbConnection(connStringBuilder.ConnectionString))
                {
                    dbConn.Open();

                    DataTable sheets = dbConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    using (OleDbCommand selectCmd = new OleDbCommand(
                      String.Format("SELECT * FROM [{0}]", sheets.Rows[0]["TABLE_NAME"]), dbConn))
                    {
                        using (OleDbDataAdapter dbAdapter = new OleDbDataAdapter())
                        {
                            dbAdapter.SelectCommand = selectCmd;
                            dbAdapter.Fill(data, "mytable");
                        }
                    }
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message.ToString());
                return null;
            }
            
            this.listBox1.ValueMember = data.Tables["mytable"].Columns[0].ColumnName;
            this.listBox1.DisplayMember = data.Tables["mytable"].Columns[0].ColumnName;
            this.listBox1.DataSource = data.Tables["mytable"];

            return null;
        }

        
        */


        /*
        #region private 根据excle的路径把第一个sheet中的内容放入datatable
        /// <summary>
        /// 根据excle的路径把第一个sheet中的内容放入datatable
        /// </summary>
        /// <param name="path">excel文件存放的路径</param>
        /// <returns>DataTable</returns>

        private static DataTable ReadExcelToTable(string path)
        {
            try
            {
                // 连接字符串(Office 07及以上版本 不能出现多余的空格 而且分号注意)
                //string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
                // 连接字符串(Office 07以下版本, 基本上上面的连接字符串就可以了) 
                string connstring = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
                using (OleDbConnection conn = new OleDbConnection(connstring))
                {
                    conn.Open();
                    // 取得所有sheet的名字
                    DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
                    // 取得第一个sheet的名字
                    string firstSheetName = sheetsName.Rows[0][2].ToString();

                    // 查询字符串 
                    string sql = string.Format("SELECT * FROM [{0}]", firstSheetName);

                    // OleDbDataAdapter是充当 DataSet 和数据源之间的桥梁，用于检索和保存数据
                    OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);

                    // DataSet是不依赖于数据库的独立数据集合
                    DataSet set = new DataSet();

                    // 使用 Fill 将数据从数据源加载到 DataSet 中
                    ada.Fill(set);

                    return set.Tables[0];
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message.ToString());
                return null;
            }
        }
        #endregion
         */


    }
}
