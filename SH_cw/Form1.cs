using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;
namespace SH_cw
{
    public partial class Form1 : Form
    {
        DataTable BaseTable = new DataTable();
        DataTable Table = new DataTable();
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 读取Excel数据返回DataTable
        /// </summary>
        /// <param name="Excel_File">Excel的路径</param>
        /// <returns></returns>
        private DataTable GetExcel(string Excel_File)
        {
            string connStr = "";
            string FileType = System.IO.Path.GetExtension(Excel_File);
            if (string.IsNullOrEmpty(FileType)) return null;
            if (FileType == ".xls")
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Excel_File + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            else
                connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + Excel_File + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            string sql_F = "select * from [{0}]";
            OleDbConnection conn = null;
            OleDbDataAdapter da = null;
            DataTable dtSheetName = null;
            try
            {
                conn = new OleDbConnection(connStr);
                conn.Open();
                string SheetName = "";
                dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                da = new OleDbDataAdapter();
                for (int i = 0; i < dtSheetName.Rows.Count; i++)
                {
                    SheetName = dtSheetName.Rows[i]["TABLE_NAME"].ToString();
                    if (SheetName.Contains("$") && !SheetName.Replace("'", "").EndsWith("$"))
                    {
                        continue;
                    }
                    da.SelectCommand = new OleDbCommand(string.Format(sql_F, SheetName), conn);

                    DataSet ds = new DataSet();
                    da.Fill(ds, SheetName);
                    return ds.Tables[0];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {

                conn.Close();
                da.Dispose();
                conn.Dispose();
            }
            return null;
        }
        /// <summary>
        /// 把DataTable表格保存城Csv
        /// </summary>
        /// <param name="dt">需要保存的表格</param>
        /// <param name="FilePath">保存路径</param>
        private void DataTableToCsv(DataTable dt, string FilePath)
        {
            FileStream fs = null;
            StreamWriter sw = null;
            try
            {
                StringBuilder sb = new StringBuilder();
                fs = new FileStream(FilePath, FileMode.OpenOrCreate);
                sw = new StreamWriter(new BufferedStream(fs), System.Text.Encoding.Default);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sb.Append(dt.Columns[i].ColumnName + "\t");//"\t"切换到同行的下一个单元格
                }
                sb = sb.Remove(sb.Length - 1, 1);
                sb.Append("\n");//"\n"切换到下一行
                sw.Write(sb);
                sb = new StringBuilder();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        sb.Append(dt.Rows[i][j].ToString() + "\t");
                    }
                    sb.Remove(sb.Length - 1, 1);
                    sb.Append("\n");
                }
                sw.Write(sb);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                sw.Close();
                fs.Close();
            }
        }
        /// <summary>
        /// Form清空
        /// </summary>
        private void Clear()
        {
            dataGridView1.DataSource = null;
            BaseTable = new DataTable();
            Table = new DataTable();
        }
        private void JXCJ(DataTable dt)
        {
            DataTable CJ_table = new DataTable();
            CJ_table.Columns.Add("名称");
            CJ_table.Columns.Add("金额");
            foreach(DataRow dr in dt.Rows)
            {
                DataRow temp_dr = CJ_table.NewRow();
                temp_dr[0] = dr[0];
                string MC = dr[0].ToString();
                decimal Amount = 0;
                foreach(Match match in Regex.Matches(MC, @"[、“”【】-（）《》·\[\]\w]*\s?[、“”【】-（）《》·\[\]\w]*\s\*\s\d*"))
                {
                    string[] split = match.ToString().Split('*');
                    string name = split[0].Trim();
                    int count = int.Parse(split[1]);
                    foreach(DataRow Bdr in BaseTable.Rows)
                    {
                        if (Bdr["商品名称"].ToString()==name)
                            Amount += decimal.Parse(Bdr["市本级分佣"].ToString())*count;
                    }
                }
                temp_dr[1] = Amount;
                CJ_table.Rows.Add(temp_dr);
            }
            string Csv_name = DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + " " + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + "-" + DateTime.Now.Second.ToString();
            DataTableToCsv(CJ_table, @"F:\SH_cw\out\"+Csv_name+".xls");
            MessageBox.Show(@"已保存在  F:\SH_cw\out\");
        }
        private void dataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Link;
            else e.Effect = DragDropEffects.None;
        }
        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            string dataFile = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            string[] arrystring = dataFile.Split('.');
            if (arrystring[arrystring.Length - 1].ToLower() == "xls" || arrystring[arrystring.Length - 1].ToLower() == "xlsx")
            {
                Table = GetExcel(@dataFile);
                dataGridView1.DataSource = Table;
                dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (!Directory.Exists(@"F:\SH_cw\out\"))
            {
                Directory.CreateDirectory(@"F:\SH_cw\out\");
            }
            string file = @"F:\SH_cw\酬金.xls";
            if (System.IO.File.Exists(file))
            {
                BaseTable = GetExcel(file);
            }
            else
            {
                MessageBox.Show("缺少  “酬金”表格");
                this.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            JXCJ(Table);
        }
    }
}
