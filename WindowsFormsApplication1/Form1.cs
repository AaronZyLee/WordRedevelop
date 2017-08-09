using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordRedevelop
{
    public partial class Form1 : Form
    {

        object saveFile;

        public Form1()
        {
            InitializeComponent();
        }
        ApplicationClass word = new Microsoft.Office.Interop.Word.ApplicationClass();


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExchangeText();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            CovertWord();
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
                object fileName = @"C:\Users\DELL\Desktop\松轩可研说明-输变电工程.docx";
                saveFile = @"C:\Users\DELL\Desktop\松轩可研说明-输变电工程（修改）.docx";
                app = new Microsoft.Office.Interop.Word.Application();
                doc = app.Documents.Open(ref fileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                Dictionary<string, string> datas = new Dictionary<string, string>();
                datas.Add("上海金山松轩110kV输变电工程", "哈哈哈哈哈");

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

                doc.SaveAs2(saveFile, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                if (doc != null)
                    doc.Close();
                if (app != null)
                    app.Quit();
                MessageBox.Show("替换成功！");
            }

        }

        /// <summary>
        /// 拆分word
        /// </summary>
        private void CovertWord()
        {
            object missing = System.Reflection.Missing.Value;

            try
            {
                string path = @"C:\Users\DELL\Desktop\split";
                //Directory.CreateDirectory(path);

                //Document doc = ReadDocument(@"C:\Users\DELL\Desktop\示例word.doc");
                Document doc = ReadDocument((string)saveFile);
                      

                int[,] positions = GetPosition(doc);


                object oStart = doc.Content.Start;
                object oEnd = 0;
                for (int i = 0; i < doc.Bookmarks.Count; i++)
                {
                    if (i != doc.Bookmarks.Count-1)
                    {
                        oEnd = positions[i, 0];
                    }
                    else
                    {
                        oEnd = doc.Content.End;
                    }

                    Range tocopy = doc.Range(ref oStart, ref oEnd);
                    tocopy.Copy();

                    Document docto = CreateDocument();
                    docto.Content.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);

                    object filename = path + @"\" + "test" + i.ToString() + ".docx";
                    object format = WdSaveFormat.wdFormatDocumentDefault; 
                    docto.SaveAs(ref filename, ref format, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing);
                    docto.Close(ref missing, ref missing, ref missing);

                    oStart = oEnd;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                word.Quit(ref missing, ref missing, ref missing);
                MessageBox.Show("拆分成功!");
            }


        }

        /// <summary>
        /// 合并word
        /// </summary>
        private void MergeWord()
        {

        }

        #region Helper Function
        /// <summary>
        /// 创建word文档
        /// </summary>
        /// <returns></returns>
        private Document CreateDocument()
        {
            object missing = System.Reflection.Missing.Value;
            Document newdoc = word.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            return newdoc;
        }

        /// <summary>
        /// 读取word文档
        /// </summary>
        /// <param name="path">word文档路径</param>
        /// <returns></returns>
        private Document ReadDocument(string path)
        {
            object missing = System.Reflection.Missing.Value;

            Type wordType = word.GetType();

            Documents docs = word.Documents;
            Type docsType = docs.GetType();
            object objDocName = path;
            Document doc = (Document)docsType.InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, docs, new Object[] { objDocName, true, true });

            return doc;
        }

        /// <summary>
        /// 获取书签的位置
        /// </summary>
        /// <param name="word">文档对象</param>
        /// <returns></returns>
        private int[,] GetPosition(Document word)
        {
            
            int bmcount = word.Bookmarks.Count;
            int[,] result = new int[bmcount, 2];
            for (int i = 1; i <= bmcount; i++)
            {
                object index = i;
                Bookmark bm = word.Bookmarks.get_Item(ref index);
                result[i-1, 0] = bm.Start;
                result[i-1, 1] = bm.End;
                
            }
            return result;
        }
        #endregion


    }
}
