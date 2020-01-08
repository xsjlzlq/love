using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace 小工具
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView2.AutoGenerateColumns = false;
            this.comboBox1.SelectedIndex = 1;
        }
       private List<Images> img_List;
       private List<Qlrs> qlr_List;
        public void InitShowImages()
        {
            img_List = this.textBox2.Text.GetImageInfo(is_check);
            if (img_List == null) return;
            this.dataGridView1.DataSource = null;
            this.dataGridView1.DataSource = img_List;
            this.toolStripStatusLabel2.Text = string.Format("{0}数量{1}", this.dataGridView1.Columns[0].HeaderText.Substring(0, 2), img_List.Count);
        }
        public void InitShowQlrs()
        {
            qlr_List = this.textBox1.Text.GetQlrInfo();
            if (qlr_List == null) return;
            this.dataGridView2.DataSource = null;
            this.dataGridView2.DataSource = qlr_List;
            this.toolStripStatusLabel1.Text = "权利人数量" + qlr_List.Count;
        }
        //权利人列表
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "打开权利人列表";
            ofd.Filter = "(*.xlsx)|*.xlsx|(*.txt)|*.txt|(*.csv)|*.csv";
            ofd.ShowDialog();
            if (ofd.FileName == string.Empty) return;
            this.textBox1.Text = ofd.FileName;
            this.InitShowQlrs();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowDialog();
            if (fbd.SelectedPath == string.Empty) return;
            this.textBox2.Text = fbd.SelectedPath;
            this.InitShowImages();
        }
        /// <summary>
        /// is_check
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        bool is_check = true;
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                is_check = true;
                this.dataGridView1.Columns[0].HeaderText = "照片名称";
                
            }
            else
            {
                is_check = false;
                this.dataGridView1.Columns[0].HeaderText = "文件名称";
            }
            this.InitShowImages();
            }
         //点击排序
        bool state = true;
        private void dataGridView1_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (img_List == null || img_List.Count == 0) return;
            DataGridViewColumn column = dataGridView1.Columns[e.ColumnIndex];
            OrderBy<Images>(dataGridView1, column, img_List);
            this.toolStripStatusLabel2.Text = string.Format("{0}数量{1}", this.dataGridView1.Columns[0].HeaderText.Substring(0, 2), img_List.Count);
        }
        private void dataGridView2_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (qlr_List == null || qlr_List.Count == 0) return;
            DataGridViewColumn column = dataGridView2.Columns[e.ColumnIndex];
            OrderBy<Qlrs>(dataGridView2, column, qlr_List);
            this.toolStripStatusLabel1.Text = "权利人数量" + qlr_List.Count;
        }
        private void button3_Click(object sender, EventArgs e)
        {
             int k=Int32.Parse(comboBox1.SelectedItem.ToString());
             if (img_List == null || qlr_List == null) return;
             if (img_List.Count != k * (qlr_List.Count))
             {
                 if (MessageBox.Show(string.Format("{0}数量{1}不等于{2}倍的权利人数量，是否继续？",
                 dataGridView1.Columns[0].HeaderText.Substring(0, 2), img_List.Count, k), "提示",
                 MessageBoxButtons.YesNo) == DialogResult.Yes)
                 {
                     FolderBrowserDialog fbd = new FolderBrowserDialog();
                     fbd.ShowDialog();
                     if (fbd.SelectedPath == string.Empty) return;

                     try
                     {
                         fbd.SelectedPath.Move(qlr_List, img_List, k);
                     }
                     catch (Exception ex) { MessageBox.Show(ex.Message); }
                     finally { MessageBox.Show("匹配成功！", "提示"); }
                 }
             }
             else
             {
                 FolderBrowserDialog fbd = new FolderBrowserDialog();
                 fbd.ShowDialog();
                 if (fbd.SelectedPath == string.Empty) return;

                 try
                 {
                     fbd.SelectedPath.Move(qlr_List, img_List, k);
                 }
                 catch (Exception ex) { MessageBox.Show(ex.Message); }
                 finally { MessageBox.Show("匹配成功！", "提示"); }
             }

                 
                 
        }
        ////右键菜单1
        //private void listBox1_MouseDown(object sender, MouseEventArgs e)
        //{
        //    if (e.Button == MouseButtons.Right)
        //        contextMenuStrip1.Show(MousePosition.X, MousePosition.Y);
            
        //}
        ////右键删除功能1
        //private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    if (this.listBox1.Items.Count == 0) return;
        //    foreach (var item in listBox1.SelectedItems)
        //    {
        //        qlr_List.Remove(item.ToString());
        //    }
            
        //    listBox1.Items.Clear();
        //    this.listBox1.Items.AddRange(qlr_List.ToArray());
        //    this.toolStripStatusLabel1.Text = "权利人数量" + qlr_List.Count;
        //}
        ////右键刷新1
        //private void 刷新ToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    if (this.textBox1.Text == "") return;
        //    qlr_List = this.textBox1.Text.GetQlrInfo();
        //    if (this.listBox1.Items.Count > 0)
        //        listBox1.Items.Clear();
        //    this.listBox1.Items.AddRange(qlr_List.ToArray());
        //    this.toolStripStatusLabel1.Text = "权利人数量" + qlr_List.Count;
        //}
        //右键菜单2
        private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
                contextMenuStrip2.Show(MousePosition.X, MousePosition.Y);
        }
        private void dataGridView2_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
                contextMenuStrip1.Show(MousePosition.X, MousePosition.Y);
        }
        /// <summary>
        /// 右键删除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 删除DToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataBingdings<Images>(this.dataGridView1,img_List);
            this.toolStripStatusLabel2.Text = string.Format("{0}数量{1}", this.dataGridView1.Columns[0].HeaderText.Substring(0, 2), img_List.Count);
          
        }
        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (qlr_List == null || qlr_List.Count == 0) return;
            DataBingdings<Qlrs>(this.dataGridView2, qlr_List);
            this.toolStripStatusLabel1.Text = "权利人数量" + qlr_List.Count;

        }
        private void 刷新FToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.InitShowImages();
        }
        private void 刷新ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.InitShowQlrs();
        }
        /// <summary>
        /// 排序
        /// </summary>
        public void OrderBy<T>(DataGridView gridView,DataGridViewColumn column,List<T> list)
        {
           
            Type type = typeof(T);
            if (state == true)
            {
                list = list.OrderBy(m => type.GetProperty(column.Name).GetValue(m, null)).ToList();
                state = false;
            }
            else
            {
                list = list.OrderByDescending(m => type.GetProperty(column.Name).GetValue(m, null)).ToList();
                state = true;
            }
            gridView.DataSource = null;
            gridView.DataSource = list;
        }
        //右键删除
        public void DataBingdings<T>(DataGridView gridView, List<T> list)
        {
            if (list == null||list.Count==0) return;
            foreach (DataGridViewRow row in gridView.SelectedRows)
            {
                string Name = row.Cells[0].Value.ToString();
                T obj = (from b in list where typeof(T).GetProperties()[0].GetValue(b, null).Equals(Name) select b).FirstOrDefault();
                list.Remove(obj);
            }
            gridView.DataSource = null;
            gridView.DataSource = list;
           
        }

    }
}
