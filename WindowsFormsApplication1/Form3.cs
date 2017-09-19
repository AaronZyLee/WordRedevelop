using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordRedevelop
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = @"C:\Users\DELL\Desktop";
            openFileDialog1.Filter = "Word文档(*.docx)|*.docx";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
                textBox1.Text = System.IO.Path.GetFileName(openFileDialog1.FileName);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("未选择要更改的文件！");
                return;
            }
            if (textBox2.Text == "")
            {
                MessageBox.Show("未选择保存路径！");
                return;
            }
            Form1.fileName = openFileDialog1.FileName;
            Form1.saveDir = folderBrowserDialog1.SelectedPath;
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = @"C:\Users\DELL\Desktop\";
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                textBox2.Text = System.IO.Path.GetFileName(folderBrowserDialog1.SelectedPath);
        }
    }
}
