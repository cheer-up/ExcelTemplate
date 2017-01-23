using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Xml.Serialization;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading;

namespace ExcelTempl
{
    
    public partial class Form1 : Form
    {public static bool SelectFirstCell=false;
     public static bool SelectEndCell=false;
        public static DataTable dt = null;
        public static Point SelectedTemplCell;
       static Dictionary<Point, TemplateRange> templAdress =
           new Dictionary<Point, TemplateRange>();
        private delegate void SetDGVValueDelegate();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           

        }
        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DataTable tbl = new DataTable();
                // foreach (var firstRowCell in ws.Cells[1, 1, ws.Dimension.End.Row, ws.Dimension.End.Column])
                var letter = 'A';
                for (int i = 0; i < ws.Dimension.End.Column; i++)
                {
                    //tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                    tbl.Columns.Add(letter.ToString());
                    letter++;
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        /*    dataGridView2.DataSource = GetDataTableFromExcel(@"C:\Report.xlsx",false);
            dataGridView2.AutoResizeColumns(
         DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader);*/
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            dataGridView2.DataSource = GetDataTableFromExcel(openFileDialog1.FileName, false);
            dataGridView2.AutoResizeColumns(
         DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader);
        }

        private void открытьФайлExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void открытьФайлШаблонаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog2.ShowDialog();
        }

        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            dt = GetDataTableFromExcel(openFileDialog2.FileName, false);
            dataGridView1.DataSource = dt;
            for (int i=0; i<dataGridView1.ColumnCount;i++)

            dataGridView1.AutoResizeColumns(
         DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader);
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView2.ClearSelection();
            label4.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            //SelectedTemplCell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            SelectedTemplCell = new Point(e.RowIndex, e.ColumnIndex);
            if (!templAdress.ContainsKey(SelectedTemplCell))
            {
                templAdress.Add(SelectedTemplCell, new TemplateRange(new Point(0, 0), new Point(0, dataGridView2.RowCount - 1)));
            }
            else {
                for (int i = templAdress[SelectedTemplCell].FirstCell.X; i <= templAdress[SelectedTemplCell].EndCell.X; i++)
                {
                    try
                    { dataGridView2.Rows[i].Cells[templAdress[SelectedTemplCell].FirstCell.Y].Selected = true; }
                    catch { }
                }
            }
            label5.Text = templAdress[SelectedTemplCell].FirstCellString();
            label6.Text = templAdress[SelectedTemplCell].EndCellString();



        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            
            dataGridView2.DefaultCellStyle.SelectionBackColor = Color.Red;
            SelectFirstCell = true;
            SelectEndCell = false;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            dataGridView2.DefaultCellStyle.SelectionBackColor = Color.Red;
            SelectEndCell = true;
            SelectFirstCell = false;
        }
       

        private void dataGridView2_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (SelectFirstCell)
            {
                templAdress[SelectedTemplCell].FirstCell1 = new Point(e.RowIndex,e.ColumnIndex);
                if ((templAdress[SelectedTemplCell].EndCell.X < templAdress[SelectedTemplCell].FirstCell.X) ||
                    (templAdress[SelectedTemplCell].EndCell.Y != templAdress[SelectedTemplCell].FirstCell.Y)) {
                    templAdress[SelectedTemplCell].EndCell = new Point(dataGridView2.RowCount-1, templAdress[SelectedTemplCell].FirstCell.Y);
                }
                label5.Text = templAdress[SelectedTemplCell].FirstCellString();
                label6.Text = templAdress[SelectedTemplCell].EndCellString();
                for (int i = templAdress[SelectedTemplCell].FirstCell.X; i<= templAdress[SelectedTemplCell].EndCell.X; i++)
                {
                    dataGridView2.Rows[i].Cells[templAdress[SelectedTemplCell].FirstCell.Y].Selected = true;
                }
            }
            if (SelectEndCell)
            {
                Point EndCellPosibly = new Point(e.RowIndex, e.ColumnIndex);
                if ((EndCellPosibly.X < templAdress[SelectedTemplCell].FirstCell.X) ||
                    (EndCellPosibly.Y != templAdress[SelectedTemplCell].FirstCell.Y))
                {
                    templAdress[SelectedTemplCell].EndCell = new Point(dataGridView2.RowCount - 1, templAdress[SelectedTemplCell].FirstCell.Y);
                }
                else { templAdress[SelectedTemplCell].EndCell = EndCellPosibly; }
                label5.Text = templAdress[SelectedTemplCell].FirstCellString();
                label6.Text = templAdress[SelectedTemplCell].EndCellString();
                for (int i = templAdress[SelectedTemplCell].FirstCell.X; i <= templAdress[SelectedTemplCell].EndCell.X; i++)
                {
                    dataGridView2.Rows[i].Cells[templAdress[SelectedTemplCell].FirstCell.Y].Selected = true;
                }
            }
        }
        public void PrintToGrid()
        {
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                // DataTable dt=(DataTable) dataGridView1.DataSource;
                if (templAdress.ContainsKey(new Point(0, i)))
                {
                    
                    TemplateRange tmp = templAdress[new Point(0, i)];
                    int l = 1;
                    if (dataGridView2.DataSource != null)
                    {

                        for (int j = tmp.FirstCell.X; j < tmp.EndCell.X; j++, l++)
                        {

                            try { dt.Rows[l][i] = dataGridView2.Rows[j].Cells[tmp.EndCell.Y].Value; } // dataGridView1.Rows[l].Cells[i].Value = dataGridView2.Rows[j].Cells[tmp.EndCell.Y].Value; }
                            catch (Exception)
                            {
                                DataRow dr = dt.NewRow();
                                dt.Rows.Add(dr);
                                dt.Rows[l][i] = dataGridView2.Rows[j].Cells[tmp.EndCell.Y].Value;
                                // dataGridView1.Rows[l].Cells[i].Value = dataGridView2.Rows[j].Cells[tmp.EndCell.Y].Value;
                            }
                        }
                        }
                        else
                        {
                            try { dt.Rows[l][i] = tmp.FirstCellString() + ":" + tmp.EndCellString(); } // dataGridView1.Rows[l].Cells[i].Value = dataGridView2.Rows[j].Cells[tmp.EndCell.Y].Value; }
                            catch (Exception)
                            {
                                DataRow dr = dt.NewRow();
                                dt.Rows.Add(dr);
                                dt.Rows[l][i] = tmp.FirstCellString() + ":" + tmp.EndCellString();
                                // dataGridView1.Rows[l].Cells[i].Value = dataGridView2.Rows[j].Cells[tmp.EndCell.Y].Value;
                            }
                            break;
                        }
                    
                }
            }
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
 dataGridView1.ColumnHeadersVisible = false;
 dataGridView1.DataSource = dt;
            dataGridView1.ColumnHeadersVisible = true;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            //dataGridView1.Invoke((Action)(() => dataGridView1.DataSource = dt));
            //            dataGridView1.Invalidate();

        }
        private void button3_Click(object sender, EventArgs e)
        {
            PrintToGrid();
        }
     
        private void button4_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
        }
        public static void WriteToExcelM(string path)
        {
            using (ExcelPackage pck = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("List1");
                ws.Cells["A1"].LoadFromDataTable(dt, true);
                pck.Save();
            }
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            WriteToExcelM(saveFileDialog1.FileName);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            saveFileDialog2.ShowDialog();
            
       
        }
        public static void SerializeDic(Dictionary<Point, TemplateRange> dic,string path)
        {
            /*TextWriter writer = new StreamWriter(path);
            XmlSerializer serializer = new XmlSerializer(typeof(DicXML[]),
                                 new XmlRootAttribute() { ElementName = "items" });

            serializer.Serialize(writer,
                          dic.Select(kv => new DicXML() { Cell = kv.Key, Range = kv.Value }).ToArray());*/
            Config cfg = new Config(dic, dt.Rows[0].ItemArray);
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
            formatter.Serialize(stream, cfg);
            stream.Close();
            MessageBox.Show("Конфигурация сохранена");
        }
        public  void DeSerializeDic(string path)
        {
           dt= new DataTable();
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            Config obj = (Config)formatter.Deserialize(stream);
            stream.Close();
            templAdress = obj.dic;
            var letter = 'A';
           
            for (int i = 0; i < obj.header.Length; i++)
            {
                //tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                dt.Columns.Add(letter.ToString());
                letter++;
            }
            dt.Rows.Add(obj.header);
            PrintToGrid();
            


        }

        private void saveFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            SerializeDic(templAdress, saveFileDialog2.FileName);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            openFileDialog3.ShowDialog();
        }

        private void openFileDialog3_FileOk(object sender, CancelEventArgs e)
        {
            DeSerializeDic(openFileDialog3.FileName);
        }
    }
}
