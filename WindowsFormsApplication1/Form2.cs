using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordRedevelop
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("未输入要替换的文字");
                return;
            }
            ExchangeText();
            this.Close();

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 替换文本
        /// </summary>
        private void ExchangeText()
        {
            Microsoft.Office.Interop.Word.Application app = null;
            Document doc = null;

            try
            {
                object missing = System.Reflection.Missing.Value;
                //object fileName = @"C:\Users\DELL\Desktop\松轩可研说明-输变电工程.docx";
                //Form1.saveFile = @"C:\Users\DELL\Desktop\松轩可研说明-输变电工程（修改）.docx";
                Form1.saveFile = (string)Form1.saveDir + "\\松轩可研说明-输变电工程（修改）.docx";
                app = new Microsoft.Office.Interop.Word.Application();
                doc = app.Documents.Open(ref Form1.fileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                Dictionary<string, string> datas = new Dictionary<string, string>();
                datas.Add("上海金山松轩110kV输变电工程", textBox1.Text);

                object replace = WdReplace.wdReplaceAll;
                foreach (var item in datas)
                {
                    app.Selection.Find.Replacement.ClearFormatting();
                    app.Selection.Find.ClearFormatting();
                    app.Selection.Find.Text = item.Key;
                    app.Selection.Find.Replacement.Text = item.Value;

                    app.Selection.Find.Execute(
                    ref missing, ref missing,
                    ref missing, ref missing,
                    ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref replace,
                    ref missing, ref missing,
                    ref missing, ref missing);
                }

                if (System.IO.File.Exists((string)Form1.saveFile))
                    System.IO.File.Delete((string)Form1.saveFile);
                doc.SaveAs2(Form1.saveFile, missing, missing, missing, missing, missing, missing, missing,
                    missing, missing, missing, missing, missing, missing, missing, missing, missing);
                MessageBox.Show("替换成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (doc != null)
                    doc.Close();
                if (app != null)
                    app.Quit();

            }

        }
    }
}
