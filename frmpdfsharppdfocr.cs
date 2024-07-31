//using PdfSharp.Pdf.IO;
//using PdfSharp.Pdf;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class frmpdfsharppdfocr : Form
    {
        /// <summary>
        /// 区分pdf文件是不是带图片的 带图片的需要ocr
        /// </summary>
        public frmpdfsharppdfocr()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog op = new FolderBrowserDialog();

            if (op.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = op.SelectedPath;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog op = new FolderBrowserDialog();

            if (op.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = op.SelectedPath;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = textBox1.Text;

            string path2 = textBox2.Text;
            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("请选择转换前文件夹路径");
                return;

            }
            if (string.IsNullOrEmpty(path2))
            {

                //在转换前文件夹下新建一个转换后文件夹
                path2 = path + "\\NewOCRFile";

            }
            if (!Directory.Exists(path))
            {

                MessageBox.Show("转换前文件夹不存在");
                return;

            }
            DirectoryInfo dir = new DirectoryInfo(path);
            DeleteFolderRecursively(path2);//删除转换后的旧文件
            TraverseDirectory(dir, path, path2);//遍历文件夹及子文件夹的文件

            MessageBox.Show("转换成功");
        }
        public void TraverseDirectory(DirectoryInfo directory, string path, string path2)
        {
            // 遍历文件夹中的文件
            foreach (FileInfo file in directory.GetFiles())
            {
                if (file.Extension.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    //判断是否解压过 就是同名称文件夹

                    string filename = System.IO.Path.GetFileNameWithoutExtension(file.FullName);

                    string wordFilePath = string.Format(file.DirectoryName + "\\" + filename);
                    if (!Directory.Exists(wordFilePath))
                    {
                        //没有同名称文件夹
                        // 解压缩文件到同一文件夹下    

                        ExtractZipFile(file.FullName, path2 + "\\" + filename);
                    }
                }
                if (file.Extension.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    //处理文件夹下的文件
                    string filename = System.IO.Path.GetFileNameWithoutExtension(file.FullName);
                    string filenameAll = System.IO.Path.GetFileName(file.FullName);
                    string wordFilePath = string.Format(file.DirectoryName + "\\" + filename);

                    if (HasImagesInPdfSharp(file.FullName))
                    {
                        wordFilePath = string.Format(file.DirectoryName + "\\O");
                    }
                    else
                    {
                        wordFilePath = string.Format(file.DirectoryName + "\\C");
                    }

                    wordFilePath = wordFilePath.Replace(path, path2);
                    if (!Directory.Exists(wordFilePath))
                    {
                        Directory.CreateDirectory(wordFilePath);
                    }
                    wordFilePath = string.Format(wordFilePath + "\\" + filenameAll);
                    if (File.Exists(wordFilePath))
                    {
                        File.Delete(wordFilePath);
                    }
                    File.Copy(file.FullName, wordFilePath);
                }

            }

            // 遍历子文件夹
            foreach (DirectoryInfo subDirectory in directory.GetDirectories())
            {
                if (!subDirectory.FullName.EndsWith("NewOCRFile", StringComparison.OrdinalIgnoreCase))
                {
                    TraverseDirectory(subDirectory, path, path2); // 递归调用
                }
            }
        }



        public static void DeleteFolderRecursively(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                //  Console.WriteLine($"The folder {folderPath} does not exist.");
                return;
            }

            try
            {
                foreach (var file in Directory.GetFiles(folderPath))
                {
                    File.Delete(file);
                }

                foreach (var subFolder in Directory.GetDirectories(folderPath))
                {
                    Directory.Delete(subFolder, true); // 第二个参数设置为true表示删除子目录及其内容
                }

                Directory.Delete(folderPath, true);
                //  Console.WriteLine($"Folder {folderPath} and its contents have been successfully deleted.");
            }
            catch (Exception ex)
            {
                //  Console.WriteLine($"An error occurred while deleting the folder: {ex.Message}");
            }
        }


        /// <summary>
        /// 解压文件进行文件的处理
        /// </summary>
        /// <param name="zipFilePath"></param>
        /// <param name="extractPath"></param>
        public void ExtractZipFile(string zipFilePath, string extractPath)
        {

            // 使用ZipFile.ExtractToDirectory方法解压ZIP文件
            if (!Directory.Exists(extractPath))
            {
                Directory.CreateDirectory(extractPath);
            }
            ZipFile.ExtractToDirectory(zipFilePath, extractPath);
            foreach (var file in Directory.GetFiles(extractPath, "*.zip"))
            {
                // 检查文件是否是zip文件
                if (file.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    // 解压缩文件到同一文件夹下   
                    string filename = System.IO.Path.GetFileNameWithoutExtension(file);

                    ExtractZipFile(file, extractPath + "\\" + filename);

                    System.IO.File.Delete(file);

                    foreach (var Newfile in Directory.GetFiles(extractPath + "\\" + filename))
                    {
                        if (Newfile.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                        {
                            string filenameAll = System.IO.Path.GetFileName(Newfile);
                            string wordFilePath = string.Format(extractPath + "\\" + filename);

                            if (HasImagesInPdfSharp(Newfile))
                            {
                                wordFilePath = string.Format(extractPath + "\\" + filename + "\\O");
                            }
                            else
                            {
                                wordFilePath = string.Format(extractPath + "\\" + filename + "\\C");
                            }
                            if (!Directory.Exists(wordFilePath))
                            {
                                Directory.CreateDirectory(wordFilePath);
                            }
                            wordFilePath = string.Format(wordFilePath + "\\" + filenameAll);
                            if (File.Exists(wordFilePath))
                            {
                                File.Delete(wordFilePath);
                            }

                            File.Copy(Newfile, wordFilePath);
                            if (File.Exists(Newfile))
                            {
                                File.Delete(Newfile);
                            }
                        }
                    }

                }
            }
        }

        /// <summary>
        /// PdfSharp 判断pdf文件第二页是否有图片
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool HasImagesInPdfSharp(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return false;
            }

            try
            {
                bool isimg = false;
                using (PdfSharp.Pdf.PdfDocument document = PdfSharp.Pdf.IO.PdfReader.Open(filePath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.ReadOnly))
                {
                    int maxcount = document.PageCount > 3 ? 3 : document.PageCount;
                    if (document.PageCount > 1)
                    {
                        for (int pageIndex = 1; pageIndex < maxcount; pageIndex++)
                        {
                            PdfSharp.Pdf.PdfPage page = document.Pages[pageIndex];

                            PdfSharp.Pdf.PdfDictionary res = page.Elements.GetDictionary("/Resources");
                            if (res != null)
                            {

                                PdfSharp.Pdf.PdfDictionary pd = res.Elements.GetDictionary("/XObject");
                                if (pd != null)
                                {
                                    System.Collections.Generic.ICollection<PdfSharp.Pdf.PdfItem> pdfit = pd.Elements.Values;
                                    foreach (PdfSharp.Pdf.PdfItem item in pdfit)
                                    {
                                        PdfSharp.Pdf.Advanced.PdfReference pdfr = item as PdfSharp.Pdf.Advanced.PdfReference;
                                        if (pdfr != null)
                                        {
                                            PdfSharp.Pdf.PdfDictionary keyValues = pdfr.Value as PdfSharp.Pdf.PdfDictionary;
                                            if (keyValues != null && keyValues.Elements.GetString("/Subtype") == "/Image")
                                            {
                                                isimg = true;
                                                break;
                                            }
                                        }
                                    }
                                }

                            }
                        }

                    }

                }
                return isimg;

            }
            catch (Exception ex)
            {


                //   MessageBox.Show(filePath + "转换错误：" + ex.Message);
                //pdfsharp读取失败用itextsharp
                return HasImagesInPdf(filePath);
            }
        }

        public bool HasImagesInPdf(string filePath)
        { 
            bool isimg = false;
            // 使用  ItextSharp 加载 PDF 文档  
            using (PdfReader reader = new PdfReader(filePath))
            {
                int numberOfPages = reader.NumberOfPages;
                int maxcount = numberOfPages >=3 ? 3 : numberOfPages;
                // 遍历每一页  
                for (int pageNumber = 2; pageNumber <= maxcount; pageNumber++)
                {
                   
                Console.WriteLine($"Checking page {pageNumber}...");

                // 获取页面的字典  
                PdfDictionary pageDict = reader.GetPageN(pageNumber);

                // 获取页面的资源字典（如果有的话）  
                PdfDictionary resources = (PdfDictionary)PdfReader.GetPdfObject(pageDict.Get(PdfName.RESOURCES));

                // 检查资源字典中是否有 XObject 字典（通常包含图片等资源）  
                if (resources != null && resources.Contains(PdfName.XOBJECT))
                {
                    PdfDictionary xobject = (PdfDictionary)PdfReader.GetPdfObject(resources.Get(PdfName.XOBJECT));

                    // 遍历 XObject 字典中的条目  
                    foreach (PdfName name in xobject.Keys)
                    {
                        PdfObject obj = xobject.Get(name);

                        // 检查对象是否是间接引用，并尝试获取其类型  
                        if (obj.IsIndirect())
                        {
                            PdfDictionary xobj = (PdfDictionary)PdfReader.GetPdfObject(obj);

                            // 检查对象的类型，这里以图片（/Image）为例  
                            if (PdfName.IMAGE.Equals(xobj.Get(PdfName.SUBTYPE)))
                            {
                                isimg = true;
                                Console.WriteLine($"Page {pageNumber} contains an image with name {name}.");
                                break;
                            }
                        }
                    }
                }
                 }
            }
            return isimg;

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
