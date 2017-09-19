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

        public static object saveDir;
        public static object fileName;
        public static object saveFile;
        public static bool isCoverted = false;

        public Form1()
        {
            InitializeComponent();
        }
        ApplicationClass word = new Microsoft.Office.Interop.Word.ApplicationClass();


        private void Form1_Load(object sender, EventArgs e)
        {
            Form3 form = new Form3();
            form.ShowDialog();
            form.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form2 form = new Form2();
            form.ShowDialog();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            CovertWord();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            MergeWord();
        }


        /// <summary>
        /// 依据书签位置拆分word
        /// </summary>
        public void CovertWord()
        {
            object missing = System.Reflection.Missing.Value;

            try
            {
                if(saveFile == null){
                    MessageBox.Show("请先替换文本！");
                    return;
                }
                //string path = @"C:\Users\DELL\Desktop\split";
                string path = (string)saveDir;

                /**
                if (System.IO.File.Exists(path))
                    System.IO.File.Delete(path);
                System.IO.Directory.CreateDirectory(path);
                */

                Document doc = ReadDocument((string)saveFile);
                      

                int[,] positions = GetPosition(doc);


                object oStart = doc.Content.Start;
                object oEnd = 0;
                for (int i = 0; i <= doc.Bookmarks.Count; i++)
                {
                    if (i != doc.Bookmarks.Count)
                    {
                        oEnd = positions[i, 0];
                    }
                    else
                    {
                        oEnd = doc.Content.End;
                    }

                    object filename = null;
                    if (i != 0)
                    {
                        string newfile = doc.Bookmarks.get_Item(i).Name;
                        filename = path + @"\" + newfile + ".docx";
                    }
                    else
                        filename = path + @"\" + "b0" + ".docx";
                    

                    Range tocopy = doc.Range(ref oStart, ref oEnd);
                    tocopy.Copy();

                    Document docto = CreateDocument();
                    docto.Content.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);

                    object format = WdSaveFormat.wdFormatDocumentDefault; 
                    docto.SaveAs(ref filename, ref format, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing);
                    docto.Close(ref missing, ref missing, ref missing);

                    oStart = oEnd;
                }
                MessageBox.Show("拆分成功!");
                isCoverted = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            finally
            {
                word.Quit(ref missing, ref missing, ref missing);
            }


        }

        /// <summary>
        /// 合并word
        /// </summary>
        private void MergeWord()
        {
            ApplicationClass appClass = null;
            Document doc = null;
            try
            {
                if (isCoverted == false)
                {
                    MessageBox.Show("未分割文件！");
                    return;
                }
                object missing = System.Reflection.Missing.Value;
                appClass = new ApplicationClass();
                object fileName = (string)saveDir + "\\output.docx";
                doc = appClass.Documents.Add(ref missing,ref missing,ref missing,ref missing);
                doc.Activate();

                string addFolder = (string)saveDir;
                string[] files = System.IO.Directory.GetFiles(addFolder,"*.docx");

                object confirmConversion = false;
                object attachment = false;
                object link = false;
                object pBreak = (int)Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
                foreach (string s in files)
                {
                    if (s != saveFile)
                    {
                        appClass.Selection.InsertFile(s, ref missing, confirmConversion, link, attachment);
                        appClass.Selection.InsertBreak(ref pBreak);
                    }
                }
                if(System.IO.File.Exists((string)fileName))
                    System.IO.File.Delete((string)fileName);
                doc.SaveAs2(ref fileName,missing,missing,missing,missing,missing,missing,missing,missing,missing,missing,missing,missing,missing
                    ,missing,missing,missing);
                MessageBox.Show("合并成功！");
                isCoverted = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (doc != null)
                    doc.Close();
                if (appClass != null)
                    appClass.Quit();
            }

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
