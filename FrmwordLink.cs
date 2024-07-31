using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;

using Microsoft.Office.Interop.Word;
namespace WindowsFormsApp1
{
    public partial class FrmwordLink : Form
    {
        /// <summary>
        /// 处理pdf 的超链接 将超链接移除掉 
        /// </summary>
        public FrmwordLink()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string path = textBox1.Text;
            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("请选择转换前文件夹路径");
                return;

            }

            DirectoryInfo dir = new DirectoryInfo(path);
            // 遍历文件夹中的文件
            foreach (FileInfo file in dir.GetFiles())
            {

                if (file.Extension.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) || file.Extension.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    GetLink(file.FullName);
                } 
            }


            MessageBox.Show("成功");
            

        }

        public void GetLink(string path)
        {

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            try
            {
                Document doc = wordApp.Documents.Open(path);

                // 获取文档中的所有超链接  
                foreach (Hyperlink link in doc.Hyperlinks)
                {

                    // 移除超链接  
                    //   link.Range.Text = link.TextToDisplay; // 将超链接文本替换为显示文本  
                    link.Delete(); // 删除超链接   
                }
                // 保存并关闭  
                doc.Save();
                doc.Close();
                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);

            }
            catch (Exception)
            {

                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog op = new FolderBrowserDialog();

            if (op.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = op.SelectedPath;
            }

        }
    }
}
