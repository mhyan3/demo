using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Reflection;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using Org.BouncyCastle.Asn1.Pkcs;
using System.Reflection.Emit;
using System.Drawing.Drawing2D;
using DocumentFormat.OpenXml.Vml;
using System.Windows.Interop;
using System.Text;
using System.Xml;
using iTextSharp.text;

namespace WindowsFormsApp1
{
    public partial class FrmWordToTxt : Form
    {
        /// <summary>
        /// 读取word内容 转成txt简易标签
        /// </summary>
        public FrmWordToTxt()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = textBox1.Text;

            string path2 = "";
            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("请选择转换前文件夹路径");
                return;

            }
            //if (string.IsNullOrEmpty(path2))
            //{

            //    //在转换前文件夹下新建一个转换后文件夹
            //    path2 = path + "\\NewTextFile";

            //}

            if (path.EndsWith(".txt", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                if (path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    if (File.Exists(path))
                    {
                        FileInfo file = new FileInfo(path);
                        string meg = "";
                        meg = TranFile(file, meg); ;
                        if (!string.IsNullOrEmpty(meg))
                        {

                            MessageBox.Show("转换成功;" + meg);
                        }
                        else
                        {

                            MessageBox.Show("转换成功");
                        }
                    }
                }
            }
            else
            {
                if (!Directory.Exists(path))
                {

                    MessageBox.Show("转换前文件夹不存在");
                    return;

                }
                DirectoryInfo dir = new DirectoryInfo(path);
                //DeleteFolderRecursively(path2);//删除转换后的旧文件
                string meg = "";
                meg = TraverseDirectory(dir, path, path2, meg);//遍历文件夹及子文件夹的文件
                if (!string.IsNullOrEmpty(meg))
                {

                    MessageBox.Show("转换成功;" + meg);
                }
                else
                {

                    MessageBox.Show("转换成功");
                }
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
        public static void DeleteFolderRecursively(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                return;
            }

            try
            {
                foreach (var file in Directory.GetFiles(folderPath))
                {
                    if (file.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                    {
                        File.Delete(file);
                    }
                }


            }
            catch (Exception ex)
            {
                //  Console.WriteLine($"An error occurred while deleting the folder: {ex.Message}");
            }
        }
        public string TraverseDirectory(DirectoryInfo directory, string path, string path2, string msg)
        {
            // 遍历文件夹中的文件
            foreach (FileInfo file in directory.GetFiles())
            {

                msg = TranFile(file, msg);

            }

            // 遍历子文件夹
            foreach (DirectoryInfo subDirectory in directory.GetDirectories())
            {
                if (!subDirectory.FullName.EndsWith("NewTextFile", StringComparison.OrdinalIgnoreCase))
                {
                    msg = TraverseDirectory(subDirectory, path, path2, msg); // 递归调用
                }
            }
            return msg;
        }
        public string TranFile(FileInfo file, string msg)
        {

            if (file.Extension.EndsWith(".doc", StringComparison.OrdinalIgnoreCase))
            {
                //处理文件夹下的文件
                string filename = System.IO.Path.GetFileNameWithoutExtension(file.FullName);


                string filenameAll = System.IO.Path.GetFileName(file.FullName);
                string TextFilePath = string.Format(file.DirectoryName + "\\" + filename + ".txt");
                if (File.Exists(TextFilePath))
                {
                    File.Delete(TextFilePath);
                }
                //读取word内容 存入text
                string[] textval = { };

                string Jsonpath = System.Windows.Forms.Application.StartupPath;
                Jsonpath = Jsonpath + "\\fileTemp";
                if (!Directory.Exists(Jsonpath))
                {
                    Directory.CreateDirectory(Jsonpath);
                }

                string wordPath = string.Format(Jsonpath + "\\" + filename + ".doc");
                if (File.Exists(wordPath))
                {
                    File.Delete(wordPath);
                }
                File.Copy(file.FullName, wordPath);

                string wordPath1 = string.Format(Jsonpath + "\\" + filename + ".docx");
                if (File.Exists(wordPath1))
                {
                    File.Delete(wordPath1);
                }
                //doc转docx 
                wordPath = DocToDocx.ToDocxSaveAs(wordPath, false);


                List<string> TextList = textval.ToList();
                TextList = GetWordText(wordPath);
                if (TextList != null && TextList.Count() > 0)
                {

                    textval = TextList.ToArray();
                    File.WriteAllLines(TextFilePath, textval);
                }
                else
                {
                    msg += filename + "转化失败，请查询该文件是否存在或是否被打开过；";
                }

            }

            if (file.Extension.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                //处理文件夹下的文件
                string filename = System.IO.Path.GetFileNameWithoutExtension(file.FullName);
                string filenameAll = System.IO.Path.GetFileName(file.FullName);
                string TextFilePath = string.Format(file.DirectoryName + "\\" + filename + ".txt");
                if (File.Exists(TextFilePath))
                {
                    File.Delete(TextFilePath);
                }
                //读取word内容 存入text
                string[] textval = { };

                List<string> TextList = textval.ToList();
                TextList = GetWordText(file.FullName);
                if (TextList != null && TextList.Count() > 0)
                {

                    textval = TextList.ToArray();
                    File.WriteAllLines(TextFilePath, textval);
                }
                else
                {
                    msg += filename + "转化失败，请查询该文件是否存在或是否被打开过；";
                }

            }
            return msg;
        }

        //记录编号列表
        public class ListLevelModel
        {
            public int Level { get; set; } = 0;
            public int ParNumID { get; set; }
            public string IndentLeft { get; set; }//左缩进
            public string FirstLine { get; set; }//首行缩进
            public bool IslistString { get; set; } = false;//是否是编号
            public bool IsListLevelEnd { get; set; } = false;//是否结束
            public int EndParNumID { get; set; } = 0;// 结束的段落id  这样在结束段落前打结束符
            public bool IsListLevelStart { get; set; } = false;//是否开始
            public string Paragraptext { get; set; }
        }
        public List<string> GetWordText(string path)
        {
            List<string> TextList = new List<string>();
            try
            {

                List<InteropWordParagraphText> wordParagraphTexts = GetListString(path);


                List<INodeContent> nodeContents = new List<INodeContent>();
                using (WordprocessingDocument wordDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(path, true))

                {
                    int TableIndex = 1;
                    int ParagrapIndex = 1;
                    foreach (var item in wordDoc.MainDocumentPart.Document.Body.Elements())
                    {
                        if (item is DocumentFormat.OpenXml.Wordprocessing.Paragraph)
                        {
                            DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph = (DocumentFormat.OpenXml.Wordprocessing.Paragraph)item;

                            ParagrapData data = GetParagrapData(paragraph, wordParagraphTexts, true);
                            InteropWordParagraphText InteropWordParagraphTexts = wordParagraphTexts.Where(a => a.IsTable == false && a.TableIndex == ParagrapIndex).FirstOrDefault();
                            if (InteropWordParagraphTexts != null)
                            {
                                data.pageNumber = InteropWordParagraphTexts.pageNumber;
                            }
                            nodeContents.Add(data);
                            ParagrapIndex++;

                        }

                        if (item is DocumentFormat.OpenXml.Wordprocessing.Table)
                        {

                            TableData tableData = new TableData();
                            List<TableRowData> tableRows = new List<TableRowData>();

                            List<TableRowSpanData> tableRowSpans = new List<TableRowSpanData>();
                            DocumentFormat.OpenXml.Wordprocessing.Table table = (DocumentFormat.OpenXml.Wordprocessing.Table)item;

                            Dictionary<int, int> widths = new Dictionary<int, int>();
                            if (table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableGrid>() != null && table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableGrid>().Count() > 0)
                            {
                                int ColumnIndex = 1;
                                DocumentFormat.OpenXml.Wordprocessing.TableGrid tableGrid = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableGrid>().FirstOrDefault();
                                foreach (var gridColumn in tableGrid.Elements<GridColumn>())
                                {
                                    widths.Add(ColumnIndex, Convert.ToInt32(gridColumn.Width.Value));
                                    ColumnIndex++;

                                }
                            }

                            tableData.Widths = widths;


                            int RowIndex = 1;
                            foreach (var TableRow in table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>())
                            {
                                int ColumnIndex = 1;
                                foreach (var TableCell in TableRow.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>())
                                {
                                    int ColSpan = 0;
                                    if (TableCell.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCellProperties>() != null && TableCell.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCellProperties>().Count() > 0)
                                    {
                                        DocumentFormat.OpenXml.Wordprocessing.TableCellProperties tableCellProperties = TableCell.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCellProperties>().FirstOrDefault();

                                        if (tableCellProperties.Elements<GridSpan>() != null && tableCellProperties.Elements<GridSpan>().Count() > 0)
                                        {
                                            GridSpan gridSpan = tableCellProperties.Elements<GridSpan>().FirstOrDefault();
                                            ColSpan = gridSpan.Val;
                                        }


                                        if (tableCellProperties.Elements<VerticalMerge>() != null && tableCellProperties.Elements<VerticalMerge>().Count() > 0)
                                        {
                                            VerticalMerge verticalMerge = tableCellProperties.Elements<VerticalMerge>().FirstOrDefault();
                                            if (verticalMerge.Val != null)
                                            {
                                                if (verticalMerge.Val == MergedCellValues.Restart)
                                                {
                                                    //开始合并行
                                                    tableRowSpans.Add(new TableRowSpanData() { RowIndex = RowIndex, ColumnIndex = ColumnIndex, StartRowSpanIndex = RowIndex, EndRowSpanIndex = 0 });
                                                }
                                                if (verticalMerge.Val == MergedCellValues.Continue)
                                                {
                                                    //合并行 后面的行
                                                    //找到上一个合并行的 在替换结束行
                                                    if (tableRowSpans != null && tableRowSpans.Count() > 0)
                                                    {
                                                        TableRowSpanData rowSpanData = tableRowSpans.Where(a => a.ColumnIndex == ColumnIndex).OrderBy(a => a.RowIndex).LastOrDefault();
                                                        if (rowSpanData != null)
                                                        {
                                                            rowSpanData.EndRowSpanIndex = RowIndex;

                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                //合并行 后面的行
                                                //找到上一个合并行的 在替换结束行
                                                if (tableRowSpans != null && tableRowSpans.Count() > 0)
                                                {
                                                    TableRowSpanData rowSpanData = tableRowSpans.Where(a => a.ColumnIndex == ColumnIndex).OrderBy(a => a.RowIndex).LastOrDefault();
                                                    if (rowSpanData != null)
                                                    {
                                                        rowSpanData.EndRowSpanIndex = RowIndex;

                                                    }
                                                }
                                            }

                                        }

                                    }
                                    List<ParagrapData> ParagrapDatas = new List<ParagrapData>();
                                    if (TableCell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>() != null && TableCell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Count() > 0)
                                    {
                                        foreach (var paragraph in TableCell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                                        {

                                            ParagrapDatas.Add(GetParagrapData(paragraph, wordParagraphTexts, false));
                                        }


                                    }
                                    tableRows.Add(new TableRowData() { CellText = TableCell.InnerText, RowIndex = RowIndex, ColumnIndex = ColumnIndex, ColumnSpan = ColSpan, EndColSpanIndex = (ColSpan > 1 ? ColumnIndex + ColSpan - 1 : 0), ParagrapDatas = ParagrapDatas, ISColumnSpan = (ColSpan > 1 ? true : false) });
                                    if (ColSpan > 1)
                                    {
                                        ColumnIndex = ColumnIndex + ColSpan - 1;
                                    }
                                    ColumnIndex++;
                                }
                                RowIndex++;
                            }

                            //处理合并行
                            foreach (var row in tableRowSpans)
                            {

                                foreach (var cell in tableRows.Where(a => a.ColumnIndex == row.ColumnIndex && (a.RowIndex >= row.StartRowSpanIndex && a.RowIndex <= row.EndRowSpanIndex)))
                                {

                                    cell.RowSpan = row.EndRowSpanIndex - row.StartRowSpanIndex + 1;
                                    cell.StartRowSpanIndex = row.StartRowSpanIndex;
                                    cell.EndRowSpanIndex = row.EndRowSpanIndex;
                                    cell.ISRowSpan = true;
                                }
                            }
                            tableData.TableIndex = TableIndex;
                            InteropWordParagraphText InteropWordParagraphTexts = wordParagraphTexts.Where(a => a.IsTable == true && a.TableIndex == TableIndex).FirstOrDefault();
                            if (InteropWordParagraphTexts != null)
                            {
                                tableData.pageNumber = InteropWordParagraphTexts.pageNumber;
                            }
                            tableData.tableRowDatas = tableRows;
                            nodeContents.Add(tableData);
                            TableIndex++;


                        }

                    }

                    wordDoc.Close();
                }

                List<ListLevelModel> listleveldata = new List<ListLevelModel>();

                int ParNumID = 1;
                foreach (var content in nodeContents)
                {
                    if (content is ParagrapData)
                    {
                        ParagrapData paragrap = content as ParagrapData;
                        ListLevelModel listLevelModel = new ListLevelModel() { ParNumID = ParNumID };
                        if (paragrap.IslistString && paragrap.ListLevel > 0)
                        {
                            //小标题
                            //判断上一个标题和该等级是否一致  一致代表是同级   

                            listLevelModel.Level = paragrap.ListLevel;

                            listLevelModel.IndentLeft = paragrap.IndentLeft;
                            listLevelModel.FirstLine = paragrap.FirstLine;
                            listLevelModel.IslistString = true;
                            if (listleveldata != null && listleveldata.Where(a => a.IslistString == true).Count() > 0)
                            {
                                //得到上一个段落是什么等级序号  同等级的不考虑 比上一段落等级高 说明上一段落的编号区间已结束   需要找到同一等级段落  比上一段落等级低则当前等级算开始

                                var TopList = listleveldata.Where(a => a.IslistString == true && a.IsListLevelEnd == false).LastOrDefault();
                                if (TopList != null)
                                {
                                    // 
                                    if (TopList.Level != paragrap.ListLevel)
                                    {

                                        if (TopList.Level < paragrap.ListLevel)
                                        {
                                            //比上一段落等级低则当前等级算开始
                                            //第一个序号
                                            listLevelModel.IsListLevelEnd = false;
                                            listLevelModel.IsListLevelStart = true;
                                        }
                                        if (TopList.Level > 1)
                                        {

                                            TopList.IsListLevelEnd = true;
                                            TopList.EndParNumID = ParNumID;
                                        }



                                    }
                                }
                                else
                                {
                                    //以上所有序号都已结束则当前算开始

                                    //第一个序号
                                    listLevelModel.IsListLevelEnd = false;
                                    listLevelModel.IsListLevelStart = true;
                                }


                            }
                            else
                            {
                                //第一个序号
                                listLevelModel.IsListLevelEnd = false;
                                listLevelModel.IsListLevelStart = true;
                            }


                        }
                        listLevelModel.Paragraptext = paragrap.Paragraptext;
                        listleveldata.Add(listLevelModel);
                        ParNumID++;
                    }
                }
                ParNumID = 1;

                foreach (var content in nodeContents)
                {
                    if (content is ParagrapData)
                    {
                        string HtmlValue = "P<>";
                        ParagrapData paragrap = content as ParagrapData;
                        //段落
                        //对齐方式 
                        if (paragrap.Alignment == "Center")
                        {
                            HtmlValue = "PC<>";
                        }
                        if (paragrap.Alignment == "Right")
                        {
                            HtmlValue = "PR<>";

                        }
                        //判断是否可以加<list>
                        if (listleveldata.Where(a => a.ParNumID == ParNumID && a.IslistString == true && a.IsListLevelStart == true && a.Level > 1).Count() > 0)
                        {
                            HtmlValue = "<list>\n" + HtmlValue;
                        }

                        string runhtml = "";
                        var numrun = paragrap.runDatas.Where(a => a.IsNumer == true).FirstOrDefault();

                        string numberhtml = "";
                        if (numrun != null)
                        {
                            numberhtml = "<label>" + numrun.Runtext.TrimEnd() + "</label> ";
                        }

                        int runIndex = 1;
                        //将序号部分连在一起

                        foreach (var item in paragrap.runDatas.Where(a => a.IsNumer == false))
                        {


                            if (!string.IsNullOrEmpty(item.Runtext))
                            {
                                if (item.IsItalic && !item.IsTab)
                                {
                                    runhtml += "<i>";
                                }
                                if (item.IsBold && !item.IsTab)
                                {
                                    runhtml += "<b>";
                                }
                                if (item.Underline == "Single" && !item.IsTab)
                                {
                                    runhtml += "<u>";
                                }
                                if (item.IsvertAlign && item.IsSubscript)
                                {
                                    runhtml += "<sub>";
                                }
                                if (item.IsvertAlign && item.IsSuperscript)
                                {
                                    runhtml += "<sup>";
                                }
                                if (runIndex == 1 && !string.IsNullOrEmpty(numberhtml))
                                {
                                    runhtml += numberhtml;


                                }
                                runhtml += item.Runtext;

                                if (item.IsvertAlign && item.IsSuperscript)
                                {
                                    runhtml = runhtml.TrimEnd() + "</sup>";
                                }
                                if (item.IsvertAlign && item.IsSubscript)
                                {
                                    runhtml = runhtml.TrimEnd() + "</sub>";
                                }
                                if (item.Underline == "Single" && !item.IsTab)
                                {
                                    runhtml = runhtml.TrimEnd() + "</u>";
                                }
                                if (item.IsBold && !item.IsTab)
                                {
                                    runhtml = runhtml.TrimEnd() + "</b>";
                                }
                                if (item.IsItalic && !item.IsTab)
                                {
                                    runhtml = runhtml.TrimEnd() + "</i>";
                                }
                            }

                            runIndex++;
                        }

                        if (!string.IsNullOrEmpty(runhtml))
                        {

                            HtmlValue += runhtml;

                            //判断是否可以加<list>
                            if (listleveldata.Where(a => a.ParNumID == ParNumID && a.IsListLevelEnd == true).Count() > 0)
                            {
                                HtmlValue = HtmlValue + "\n</list>";
                            }
                            TextList.Add(HtmlValue);
                        }
                        ParNumID++;
                    }
                    if (content is TableData)
                    {
                        //表格

                        TableData table = content as TableData;

                        string tablehtml = "";
                        string TableStyle = "T<>";

                        //得到行号
                        List<int> RowIndexs = table.tableRowDatas.Select(a => a.RowIndex).Distinct().ToList();
                        int rowindexs = table.tableRowDatas.Where(a => a.ISColumnSpan == false).Select(a => a.RowIndex).FirstOrDefault();

                        foreach (var item in table.tableRowDatas.Where(a => a.RowIndex == rowindexs).ToList())
                        {

                            string Ali = "L";
                            var pa = item.ParagrapDatas.FirstOrDefault();
                            if (pa != null)
                            {

                                if (pa.Alignment == "Center")
                                {
                                    Ali = "C";
                                }
                                if (pa.Alignment == "Right")
                                {
                                    Ali = "R";

                                }
                            }

                            string width = Convert.ToInt32(((float)(table.Widths[item.ColumnIndex]) / (float)(table.Widths.Sum(a => a.Value))) * 100).ToString() + "%";
                            TableStyle += Ali + width + ";";

                        }
                        tablehtml += TableStyle;
                        foreach (var item in RowIndexs)
                        {

                            tablehtml += "\nB<>";
                            foreach (var row in table.tableRowDatas.Where(a => a.RowIndex == item).ToList())
                            {

                                if (row.ColumnSpan > 1)
                                {
                                    tablehtml += "<C" + row.ColumnIndex + "-" + row.ColumnSpan + ">";
                                }
                                if (row.RowSpan > 1)
                                {
                                    if (row.StartRowSpanIndex == row.RowIndex)
                                    {
                                        tablehtml += "<R" + row.StartRowSpanIndex + "-" + row.EndRowSpanIndex + ">";
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                tablehtml += "^";

                                foreach (var tableParagrap in row.ParagrapDatas)
                                {
                                    if (!string.IsNullOrEmpty(tableParagrap.Paragraptext))
                                    {



                                        foreach (var items in tableParagrap.runDatas)
                                        {

                                            string runhtml = "";

                                            if (items.IsItalic && !items.IsTab)
                                            {
                                                runhtml += "<i>";
                                            }
                                            if (items.IsBold && !items.IsTab)
                                            {
                                                runhtml += "<b>";
                                            }
                                            if (items.Underline == "Single" && !items.IsTab)
                                            {
                                                runhtml += "<u>";
                                            }
                                            runhtml += items.Runtext;
                                            if (items.Underline == "Single" && !items.IsTab)
                                            {
                                                runhtml = runhtml.TrimEnd() + "</u>";
                                            }
                                            if (items.IsBold && !items.IsTab)
                                            {
                                                runhtml = runhtml.TrimEnd() + "</b>";
                                            }
                                            if (items.IsItalic && !items.IsTab)
                                            {
                                                runhtml = runhtml.TrimEnd() + "</i>";
                                            }

                                            tablehtml = runhtml.TrimEnd() + runhtml;

                                        }

                                    }
                                    else
                                    {
                                        tablehtml += "&nbsp;";

                                    }

                                }

                            }





                        }
                        TextList.Add(tablehtml);

                    }
                }

            }
            catch (Exception)
            {

                TextList = new List<string>();
            }

            return TextList;
        }
        public ParagrapData GetParagrapData(DocumentFormat.OpenXml.Wordprocessing.Paragraph paragraph, List<InteropWordParagraphText> wordParagraphTexts, bool IsNumber)
        {
            bool IsBold = false;//是否加粗  默认不加粗
            bool IsItalic = false;//是否斜体  默认不斜体
            float Size = 12;//字体大小 
            string FontName = "";//字体 
            string Alignment = "";//对齐方式
            string IndentLeft = "";//左缩进
            string IndentRight = "";//右缩进
            string IndentHanging = "";//悬挂缩进
            string FirstLine = "";//首行缩进
            string Underline = "";
            bool IsNumberList = false;//是否有编号列表
            List<RunData> runDatas = new List<RunData>();
            if (paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties>() != null)
            {
                DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties ParagraphProperties = paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties>().FirstOrDefault();
                if (ParagraphProperties != null)
                {

                    Indentation Indentation = ParagraphProperties.Elements<DocumentFormat.OpenXml.Wordprocessing.Indentation>().FirstOrDefault();
                    if (Indentation != null)
                    {
                        if (Indentation.Left != null)
                        {

                            IndentLeft = Indentation.Left;
                        }
                        if (Indentation.Right != null)
                        {


                            IndentRight = Indentation.Right;
                        }
                        if (Indentation.Hanging != null)
                        {

                            IndentHanging = Indentation.Hanging;
                        }
                        if (Indentation.FirstLine != null)
                        {

                            FirstLine = Indentation.FirstLine;
                        }
                    }
                    DocumentFormat.OpenXml.Wordprocessing.Justification Justification = ParagraphProperties.Elements<DocumentFormat.OpenXml.Wordprocessing.Justification>().FirstOrDefault();
                    if (Justification != null)
                    {
                        if (Justification.Val != null)
                        {
                            if (Justification.Val == DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Left)

                            {
                                Alignment = "Left";
                            }
                            if (Justification.Val == JustificationValues.Center)

                            {
                                Alignment = "Center";
                            }
                            if (Justification.Val == JustificationValues.Right)

                            {
                                Alignment = "Right";
                            }
                            if (Justification.Val == JustificationValues.Both)

                            {
                                Alignment = "Both";
                            }
                        }

                    }
                    FontSize FontSize = ParagraphProperties.Elements<FontSize>().FirstOrDefault();
                    if (FontSize != null)
                    {
                        Size = Convert.ToInt32(FontSize.Val) / 2;

                    }

                    Bold Bold = ParagraphProperties.Elements<Bold>().FirstOrDefault();
                    if (Bold != null)
                    {
                        if (Bold.Val == new DocumentFormat.OpenXml.OnOffValue(true) || Bold.Val == null)
                        {
                            IsBold = true;
                        }

                    }

                    Italic Italic = ParagraphProperties.Elements<Italic>().FirstOrDefault();
                    if (Italic != null)
                    {
                        if (Italic.Val == new DocumentFormat.OpenXml.OnOffValue(true) || Italic.Val == null)
                        {
                            IsItalic = true;

                        }

                    }

                    NumberingProperties NumberingProperties = ParagraphProperties.Elements<NumberingProperties>().FirstOrDefault();
                    if (NumberingProperties != null && NumberingProperties.Elements<NumberingLevelReference>() != null)
                    {
                        IsNumberList = true;
                    }
                }
            }
            ParagrapData paragrapData = new ParagrapData();
            paragrapData.IslistString = false;
            string Paragraptext = paragraph.InnerText;
            if (IsNumberList && IsNumber)
            {
                //找编号列表前的编号序号

                InteropWordParagraphText interoptext = wordParagraphTexts.Where(a => a.StringValue == Paragraptext && a.IsTable == false && a.IsSelect == false).FirstOrDefault();
                if (interoptext != null)
                {
                    Paragraptext = interoptext.listString + Paragraptext;
                    interoptext.IsSelect = true;


                    paragrapData.listString = interoptext.listString;
                    paragrapData.IslistString = true;
                    paragrapData.ListLevel = interoptext.ListLevel;
                    string listUnderline = "";
                    if (interoptext.ListWdUnderline != null)
                    {
                        if (interoptext.ListWdUnderline == WdUnderline.wdUnderlineSingle)

                        {
                            listUnderline = "Single";
                        }
                        if (interoptext.ListWdUnderline == WdUnderline.wdUnderlineDouble)

                        {
                            listUnderline = "Double ";
                        }
                        if (interoptext.ListWdUnderline == WdUnderline.wdUnderlineNone)

                        {
                            listUnderline = "None ";
                        }

                    }

                    runDatas.Add(new RunData() { Runtext = interoptext.listString, IsBold = interoptext.ListIsBold, IsItalic = interoptext.ListIsItalic, FontName = interoptext.ListFontName, Size = interoptext.ListSize, Underline = listUnderline, IsNumer = true, IsSubscript = false, IsSuperscript = false, IsvertAlign = false });
                }
            }

            foreach (DocumentFormat.OpenXml.Wordprocessing.Run run in paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>())
            {
                //是否是tab键 
                bool IsEndTab = false;



                bool IsvertAlign = false;//是否是上下标
                bool IsSuperscript = false; //是否是上标 superscript
                bool IsSubscript = false;//是否是下标 subscript
                if (run.Elements<DocumentFormat.OpenXml.Wordprocessing.RunProperties>() != null)
                {
                    IsBold = false;//是否加粗  默认不加粗
                    IsItalic = false;//是否斜体  默认不斜体
                    Underline = "";

                    DocumentFormat.OpenXml.Wordprocessing.RunProperties runProperties = run.Elements<DocumentFormat.OpenXml.Wordprocessing.RunProperties>().FirstOrDefault();
                    if (runProperties != null)
                    {


                        Justification Justification = runProperties.Elements<Justification>().FirstOrDefault();
                        if (Justification != null)
                        {
                            if (Justification.Val != null)
                            {
                                if (Justification.Val == JustificationValues.Left)

                                {
                                    Alignment = "Left";
                                }
                                if (Justification.Val == JustificationValues.Center)

                                {
                                    Alignment = "Center";
                                }
                                if (Justification.Val == JustificationValues.Right)

                                {
                                    Alignment = "Right";
                                }
                                if (Justification.Val == JustificationValues.Both)

                                {
                                    Alignment = "Both";
                                }
                            }

                        }

                        DocumentFormat.OpenXml.Wordprocessing.Underline underlines = runProperties.Elements<DocumentFormat.OpenXml.Wordprocessing.Underline>().FirstOrDefault();
                        if (underlines != null)
                        {
                            if (underlines.Val != null)
                            {
                                if (underlines.Val == UnderlineValues.Single)

                                {
                                    Underline = "Single";
                                }
                                if (underlines.Val == UnderlineValues.Double)

                                {
                                    Underline = "Double ";
                                }
                                if (underlines.Val == UnderlineValues.None)

                                {
                                    Underline = "None ";
                                }

                            }

                        }

                        FontSize FontSize = runProperties.Elements<FontSize>().FirstOrDefault();
                        if (FontSize != null)
                        {
                            Size = Convert.ToInt32(FontSize.Val) / 2;

                        }
                        RunFonts runFonts = runProperties.Elements<RunFonts>().FirstOrDefault();
                        if (runFonts != null)
                        {
                            FontName = runFonts.EastAsia;

                        }
                        VerticalTextAlignment verticalTextAlignment = runProperties.Elements<VerticalTextAlignment>().FirstOrDefault();
                        if (verticalTextAlignment != null)
                        {
                            if (verticalTextAlignment.Val != null && verticalTextAlignment.Val == VerticalPositionValues.Superscript)
                            {
                                IsvertAlign = true;
                                IsSuperscript = true;
                                IsSubscript = false;
                            }
                            if (verticalTextAlignment.Val != null && verticalTextAlignment.Val == VerticalPositionValues.Subscript)
                            {
                                IsvertAlign = true;
                                IsSuperscript = false;
                                IsSubscript = true;
                            }

                        }
                        Bold Bold = runProperties.Elements<Bold>().FirstOrDefault();
                        if (Bold != null)
                        {
                            if (Bold.Val == new DocumentFormat.OpenXml.OnOffValue(true) || Bold.Val == null)
                            {
                                IsBold = true;
                            }

                        }

                        Italic Italic = runProperties.Elements<Italic>().FirstOrDefault();
                        if (Italic != null)
                        {
                            if (Italic.Val == new DocumentFormat.OpenXml.OnOffValue(true) || Italic.Val == null)
                            {
                                IsItalic = true;

                            }

                        }



                    }
                    if (IsBold)
                    {
                        var runs = runDatas.Where(a => a.IsNumer == true).FirstOrDefault();
                        if (runs != null)
                        {
                            runs.IsBold = IsBold;
                        }
                    }
                    if (IsItalic)
                    {
                        var runs = runDatas.Where(a => a.IsNumer == true).FirstOrDefault();
                        if (runs != null)
                        {
                            runs.IsItalic = IsItalic;
                        }
                    }
                    if (!string.IsNullOrEmpty(Underline))
                    {
                        var runs = runDatas.Where(a => a.IsNumer == true).FirstOrDefault();
                        if (runs != null)
                        {
                            runs.Underline = Underline;
                        }
                    }

                }

                string Runtext = "";

                if (run.Elements<DocumentFormat.OpenXml.Wordprocessing.Break>().Count() > 0)
                {
                    Runtext = run.InnerText + "\n";

                }
                else
                {
                    Runtext = run.InnerText;
                }

                bool IsAddRun = false;
                if (run.Elements<DocumentFormat.OpenXml.Wordprocessing.TabChar>().Count() > 0)
                {
                    string EndRuntext = "";
                    bool IsEndRuntext = false;
                    Runtext = "";
                    foreach (var item in run.Elements())
                    {
                        if (item is DocumentFormat.OpenXml.Wordprocessing.Text)
                        {
                            if (IsEndRuntext)
                            {
                                EndRuntext += item.InnerText;
                            }
                            else
                            {
                                Runtext += item.InnerText;

                            }

                        }
                        if (item is DocumentFormat.OpenXml.Wordprocessing.Break)
                        {
                            if (IsEndRuntext)
                            {
                                EndRuntext = EndRuntext + "\n";
                            }
                            else
                            {
                                Runtext = Runtext + "\n";
                            }
                        }
                        if (item is DocumentFormat.OpenXml.Wordprocessing.TabChar)
                        {
                            IsEndRuntext = true;

                        }

                    }

                    if (IsEndRuntext)
                    {
                        RunData lastRun = runDatas.LastOrDefault();
                        if (runDatas == null || runDatas.Count() <= 0)
                        {
                            //直接加数据
                            IsAddRun = true;
                        }
                        else
                        {
                            if (lastRun.IsBold == IsBold && lastRun.IsItalic == IsItalic && lastRun.Underline == Underline /*&& lastRun.FontName == FontName && lastRun.Size == Size*/ && lastRun.IsSubscript == IsSubscript && lastRun.IsSuperscript == IsSuperscript && lastRun.IsvertAlign == IsvertAlign)
                            {
                                lastRun.Runtext = lastRun.Runtext + Runtext;
                            }
                            else
                            {
                                //加数据
                                IsAddRun = true;
                            }
                        }
                        if (IsAddRun)
                        {

                            runDatas.Add(new RunData() { Runtext = Runtext, IsBold = IsBold, IsItalic = IsItalic, FontName = FontName, Size = Size, Underline = Underline, IsSubscript = IsSubscript, IsSuperscript = IsSuperscript, IsvertAlign = IsvertAlign });
                        }
                        if (!string.IsNullOrEmpty(EndRuntext))
                        {
                            runDatas.Add(new RunData() { Runtext = "", IsBold = IsBold, IsItalic = IsItalic, FontName = FontName, Size = Size, IsTab = true, Underline = Underline, IsSubscript = IsSubscript, IsSuperscript = IsSuperscript, IsvertAlign = IsvertAlign });

                        }
                        else
                        {
                            IsEndTab = true;
                        }
                        Runtext = EndRuntext;
                        IsAddRun = true;
                    }

                }
                if (IsAddRun != true)
                {
                    if (runDatas == null || runDatas.Count() <= 0)
                    {
                        //直接加数据
                        IsAddRun = true;
                    }
                    else
                    {
                        RunData lastRun = runDatas.LastOrDefault();

                        if (lastRun.IsBold == IsBold && lastRun.IsItalic == IsItalic && lastRun.Underline == Underline /*&& lastRun.FontName == FontName && lastRun.Size == Size*/ && lastRun.IsTab == false && lastRun.IsSubscript == IsSubscript && lastRun.IsSuperscript == IsSuperscript && lastRun.IsvertAlign == IsvertAlign)
                        {
                            lastRun.Runtext = lastRun.Runtext + Runtext;
                        }
                        else
                        {
                            //加数据
                            IsAddRun = true;
                        }
                    }
                }


                if (IsAddRun)
                {
                    if (!string.IsNullOrEmpty(Runtext))
                    {

                        runDatas.Add(new RunData() { Runtext = Runtext, IsBold = IsBold, IsItalic = IsItalic, FontName = FontName, Size = Size, IsTab = IsEndTab, Underline = Underline, IsSubscript = IsSubscript, IsSuperscript = IsSuperscript, IsvertAlign = IsvertAlign });
                    }
                    else if (IsEndTab && string.IsNullOrEmpty(Runtext))
                    {

                        runDatas.Add(new RunData() { Runtext = "", IsBold = IsBold, IsItalic = IsItalic, FontName = FontName, Size = Size, IsTab = IsEndTab, Underline = Underline, IsSubscript = IsSubscript, IsSuperscript = IsSuperscript, IsvertAlign = IsvertAlign });
                    }
                }

            }
            paragrapData.Paragraptext = Paragraptext;
            paragrapData.Alignment = Alignment;
            paragrapData.IndentHanging = IndentHanging;
            paragrapData.IndentLeft = IndentLeft;
            paragrapData.IndentRight = IndentRight;
            paragrapData.FirstLine = FirstLine;
            paragrapData.runDatas = runDatas;


            return paragrapData;
        }
        public class InteropWordParagraphText
        {
            public string listString { get; set; }
            public bool IsTable { get; set; }
            public int TableIndex { get; set; }
            public string TableValue { get; set; }
            public string StringValue { get; set; }
            public bool IsSelect { get; set; } = false;
            public int ListLevel { get; set; }
            public int pageNumber { get; set; }
            /// <summary>
            /// wdListBullet	2	    项目符号列表。 
            ///wdListListNumOnly	1	 可在段落正文中使用的 ListNum 域。 
            ///wdListMixedNumbering	5	 混合数字列表。 
            ///wdListNoNumbering	0	 不带项目符号、编号或分级显示的列表。 
            ///wdListOutlineNumbering	4	 分级显示的列表。 
            ///wdListPictureBullet	6	 图片项目符号列表。 
            ///wdListSimpleNumbering	3	 简单数字列表。
            /// </summary>
            public Microsoft.Office.Interop.Word.WdListType ListType { get; set; }//编号包含的列表的类型 
            public bool ListIsBold { get; set; } = false;//编号列表 是否加粗  默认不加粗
            public bool ListIsItalic { get; set; } = false;//编号列表 是否斜体  默认不斜体
            public float ListSize { get; set; } = 0;//编号列表 字体大小 
            public string ListFontName { get; set; }//编号列表 字体
            /// <summary>
            /// wdUnderlineDash	7	   划线。 
            ///wdUnderlineDashHeavy	23	 粗划线。 
            ///wdUnderlineDashLong	39	 长划线。 
            ///wdUnderlineDashLongHeavy	55	 长粗划线。 
            ///wdUnderlineDotDash	9	 点划相间线。 
            ///wdUnderlineDotDashHeavy	25	 粗点划相间线。 
            ///wdUnderlineDotDotDash	10	 点-点-划线相间模式。 
            ///wdUnderlineDotDotDashHeavy	26	 粗点-点-划线相间模式。 
            ///wdUnderlineDotted	4	 点。 
            ///wdUnderlineDottedHeavy	20	 粗点。 
            ///wdUnderlineDouble	3	 双线。 
            ///wdUnderlineNone	0	 无下划线。 
            ///wdUnderlineSingle	1	 单线。 默认值。 
            ///wdUnderlineThick	6	 单粗线。 
            ///wdUnderlineWavy	11	 单波浪线。 
            ///wdUnderlineWavyDouble	43	 双波浪线。 
            ///wdUnderlineWavyHeavy	27	 粗波浪线。 
            ///wdUnderlineWords	2	 仅为单个字加下划线。
            /// </summary>
            public WdUnderline ListWdUnderline { get; set; } = WdUnderline.wdUnderlineNone;//编号列表下划线的类型
        }
        public class ComListData
        {

            public int TableIndex { get; set; }
            public string TableValue { get; set; }
        }

        public List<InteropWordParagraphText> GetListString(string path)
        {

            // 创建Word应用程序实例  
            //Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            //wordApp.Visible = false; // 根据需要设置是否显示Word界面  
            List<InteropWordParagraphText> wordParagraphTexts = new List<InteropWordParagraphText>();
            // 打开文档
            object fileName = path;
            Dictionary<int, float> Level = new Dictionary<int, float>();

            using (var context = new ReportContext(path))

            {
                try
                {
                    int TableIndex = 1;
                    int ParagrapIndex = 1;
                    while (context.Vernier != null)
                    {
                        string text = context.Vernier.Range.Text;
                        int pageNumber = context.Vernier.Range.Information[WdInformation.wdActiveEndPageNumber];
                        if (context.Vernier.Range.Tables.Count > 0)

                        {

                            var table = context.Vernier.Range.Tables[1];


                            wordParagraphTexts.Add(new InteropWordParagraphText() { listString = "", StringValue = "", IsTable = true, TableIndex = TableIndex, IsSelect = false, pageNumber = pageNumber });
                            TableIndex++;
                            //表格 
                            var paragraphsCount = table.Range.Cells[table.Range.Cells.Count].Range.Paragraphs.Count;

                            context.Vernier = table.Range.Cells[table.Range.Cells.Count].Range.Paragraphs[paragraphsCount].Next();
                        }
                        else
                        {
                            ParagrapData paragrapData = new ParagrapData();



                            string listString = "";
                            int ListLevel = 0;
                            Range Range = context.Vernier.Range;
                            if (Range != null)
                            {



                                WdParagraphAlignment wdParagraph = Range.ParagraphFormat.Alignment;

                                InteropWordParagraphText wordParagraphText = new InteropWordParagraphText();
                                // 检查段落是否有列表编号   
                                if (Range.ListFormat.List != null)
                                {

                                    float Left = context.Vernier.Range.ParagraphFormat.LeftIndent;
                                    float firstLineIndent = context.Vernier.Range.ParagraphFormat.FirstLineIndent;
                                    listString = Range.ListFormat.ListString;

                                    float indent = Left + firstLineIndent;
                                    // 读取编号级别  
                                    ListLevel = Range.ListFormat.ListLevelNumber;
                                    if (Level != null && Level.Count() > 0)
                                    {
                                        if (Level.Where(a => a.Value == indent).Count() > 0)
                                        {
                                            ListLevel = Level.Where(a => a.Value == indent).FirstOrDefault().Key;
                                        }
                                        else
                                        {

                                            ListLevel = Level.Where(a => a.Value < indent).Max(a => a.Key) + 1;

                                            Level.Add(ListLevel, indent);
                                        }
                                    }
                                    else
                                    {
                                        Level.Add(ListLevel, indent);

                                    }

                                    wordParagraphText.ListType = Range.ListFormat.ListType;
                                    ListTemplate style = Range.ListFormat.ListTemplate;
                                    if (style != null)
                                    {
                                        int IsItalic = style.ListLevels[Range.ListFormat.ListLevelNumber].Font.Italic;
                                        int IsBold = style.ListLevels[Range.ListFormat.ListLevelNumber].Font.Bold;
                                        if (IsBold != 0)
                                        {
                                            wordParagraphText.ListIsBold = true;
                                        }
                                        if (IsItalic != 0)
                                        {
                                            wordParagraphText.ListIsItalic = true;
                                        }

                                        wordParagraphText.ListSize = style.ListLevels[Range.ListFormat.ListLevelNumber].Font.Size;
                                        wordParagraphText.ListFontName = style.ListLevels[Range.ListFormat.ListLevelNumber].Font.Name;
                                        wordParagraphText.ListWdUnderline = style.ListLevels[Range.ListFormat.ListLevelNumber].Font.Underline;

                                    }

                                }

                                wordParagraphText.listString = listString; wordParagraphText.ListLevel = ListLevel; wordParagraphText.StringValue = text.Replace("\v", "").Replace("\r", ""); wordParagraphText.IsTable = false; wordParagraphText.TableIndex = ParagrapIndex; wordParagraphText.IsSelect = false; wordParagraphText.pageNumber = pageNumber;

                                wordParagraphTexts.Add(wordParagraphText);
                            }

                            ParagrapIndex++;



                        }
                        context.Vernier.Range.Select();
                        context.MoveNext();
                    }
                    // context.Dispose();
                }
                catch (Exception ex)
                {
                    context.Dispose();
                }

            }



            return wordParagraphTexts;
        }
        public interface INodeContent
        {
        }
        public class WordReportContext
        {

            public List<INodeContent> nodeContents { get; set; }



        }

        public class ParagrapData : INodeContent
        {
            //存段落数据
            public string Paragraptext { get; set; }
            public string Alignment { get; set; }//对齐方式
            public string IndentLeft { get; set; }//左缩进
            public string FirstLine { get; set; }//首行缩进
            public string IndentRight { get; set; }//右缩进
            public string IndentHanging { get; set; }//悬挂缩进
            public int pageNumber { get; set; } = 0;// 页码
            public string listString { get; set; }//序号内容
            public bool IslistString { get; set; } //是否是序号
            public int ListLevel { get; set; } //序号等级



            public List<RunData> runDatas { get; set; }

        }

        public class TableData : INodeContent
        {
            //存表格数据

            public int TableIndex { get; set; }//表格号
            public int pageNumber { get; set; } = 0;// 页码

            public Dictionary<int, int> Widths { get; set; }//列宽
            public List<TableRowData> tableRowDatas { get; set; }

        }
        public class RunData
        {
            public string Runtext { get; set; }
            public bool IsNumer { get; set; } = false;//是否序号部分
            public bool IsTab { get; set; } = false;//是否tab
            public bool IsBold { get; set; } = false;//是否加粗  默认不加粗
            public bool IsItalic { get; set; }//是否斜体  默认不斜体
            public float Size { get; set; }//字体大小 
            public string FontName { get; set; }//字体 

            public string Underline { get; set; } //下划线的类型
            public bool IsvertAlign { get; set; } //是否是上下标
            public bool IsSuperscript { get; set; } //是否是上标 superscript
            public bool IsSubscript { get; set; } //是否是下标 subscript

        }
        public class TableRowData
        {
            public string CellText { get; set; }//列内容
            public int RowIndex { get; set; }//行号
            public int ColumnIndex { get; set; }//列号
            public int RowSpan { get; set; }//合并行数
            public int ColumnSpan { get; set; }//合并的列数
            public bool ISRowSpan { get; set; }//是否合并行
            public bool ISColumnSpan { get; set; }//是否合并列
            public float ColumnWidth { get; set; }//列宽
            public int StartRowSpanIndex { get; set; }//开始合并的行号
            public int EndRowSpanIndex { get; set; }//结束合并的行号
            public int EndColSpanIndex { get; set; }//结束合并的列数

            public List<ParagrapData> ParagrapDatas { get; set; }

        }

        //单元格合并列情况
        public class TableRowSpanData
        {

            public int RowIndex { get; set; }
            public int ColumnIndex { get; set; }

            public int StartRowSpanIndex { get; set; }
            public int EndRowSpanIndex { get; set; }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        public string[] TxtVal = null;
        private void button3_Click(object sender, EventArgs e)
        {
            string path = textBox1.Text;
            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("请选择转换前文件夹路径");
                return;

            }
            if (path.EndsWith(".txt", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                if (path.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                {
                    if (File.Exists(path))
                    {


                        FileInfo file = new FileInfo(path);
                        string meg = "";
                        //处理文件夹下的文件
                        string filename = System.IO.Path.GetFileNameWithoutExtension(file.FullName);
                        string filenameAll = System.IO.Path.GetFileName(file.FullName);
                        string xmlFilePath = string.Format(file.DirectoryName + "\\" + filename + ".xml");
                        if (File.Exists(xmlFilePath))
                        {
                            File.Delete(xmlFilePath);
                        }
                        string Jsonpath = System.Windows.Forms.Application.StartupPath;
                        Jsonpath = Jsonpath + "\\fileTemp";
                        if (!Directory.Exists(Jsonpath))
                        {
                            Directory.CreateDirectory(Jsonpath);
                        }
                        string xmlTempFilePath = string.Format(Jsonpath + "\\Template.xml");
                        TxtVal = File.ReadAllLines(file.FullName);


                        if (File.Exists(xmlTempFilePath))
                        {
                            File.Copy(xmlTempFilePath, xmlFilePath);

                            string xmlContent = File.ReadAllText(xmlFilePath);
                             
                            XDocument document = XDocument.Load(xmlFilePath);

                            //// 创建新的 XDocument 并设置新的 DOCTYPE
                            //XDocument newDoc = new XDocument(
                            //    new XDocumentType(document.DocumentType.Name, "H:\\DTD\\caselaw_sma.dtd", "[]", null),
                            //    document.Root
                            //);

                            var xmlnodes = document.Element("case").Elements().ToList();
                            XmlReplace(document.Element("case"));

                            //ap1是否有
                            GetXelment(xmlnodes); 
                       document.Save(xmlFilePath);

                        string  TxtVals   = File.ReadAllText(xmlFilePath);
                            TxtVals = TxtVals.Replace("\"H:\\DTD\\caselaw_sma.dtd\"[]", "\"H:\\DTD\\caselaw_sma.dtd\"").Replace("utf-8", "UTF-8");
                            File.WriteAllText(xmlFilePath, TxtVals);

                        }
                        if (!string.IsNullOrEmpty(meg))
                        {

                            MessageBox.Show("转换成功;" + meg);
                        }
                        else
                        {

                            MessageBox.Show("转换成功");
                        }
                    }
                }
            }
            else
            {
                if (!Directory.Exists(path))
                {

                    MessageBox.Show("转换前文件夹不存在");
                    return;

                }
                DirectoryInfo directory = new DirectoryInfo(path);
                string msg = "";
                foreach (FileInfo file in directory.GetFiles())
                {

                    if (file.Extension.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                    {
                        //处理文件夹下的文件
                        string filename = System.IO.Path.GetFileNameWithoutExtension(file.FullName);
                        string filenameAll = System.IO.Path.GetFileName(file.FullName);
                        string xmlFilePath = string.Format(file.DirectoryName + "\\" + filename + ".xml");
                        if (File.Exists(xmlFilePath))
                        {
                            File.Delete(xmlFilePath);
                        }
                        string Jsonpath = System.Windows.Forms.Application.StartupPath;
                        Jsonpath = Jsonpath + "\\fileTemp";
                        if (!Directory.Exists(Jsonpath))
                        {
                            Directory.CreateDirectory(Jsonpath);
                        }
                        string xmlTempFilePath = string.Format(Jsonpath + "\\Template.xml");
                        TxtVal = File.ReadAllLines(file.FullName);


                        if (File.Exists(xmlTempFilePath))
                        {
                            File.Copy(xmlTempFilePath, xmlFilePath);

                            XDocument document = XDocument.Load(xmlFilePath);


                            var xmlnodes = document.Element("case").Elements().ToList();
                            XmlReplace(document.Element("case"));

                            //ap1是否有
                            GetXelment(xmlnodes);

                            document.Save(xmlFilePath);
                        }

                    }

                }

                MessageBox.Show("转换成功。" + msg);
            }

        }

        public void GetXelment(List<XElement> elements)
        {

            foreach (XElement element in elements)
            {
                if (element.Elements().Count() > 0)
                {
                    GetXelment(element.Elements().ToList());

                    XmlReplace(element);
                }
                else
                {
                    XmlReplace(element);
                }
            }
        }
        static bool IsTwoUppercaseLetters(string input)
        {
            // 检查输入字符串长度是否为2  
            if (input.Length != 2)
            {
                return false;
            }

            // 使用LINQ的All方法来检查字符串中的每个字符是否都是大写英文字母  
            return input.All(char.IsUpper);
        }

        public string[] XmlReplace(XElement element)
        {

            if (element.Name == "case.ref.no.group")
            {
                //判断是否有ap1
                if (TxtVal.Where(a => a.StartsWith("AP1<>")).Count() <= 0)
                {
                    //移除
                    foreach (var item in element.Elements("history"))
                    {
                        item.Remove();
                    }
                }
            }


            if (element.Name == "judge.line" && element.Value.Contains("[DJ]"))
            {
                element.Value = element.Value.Replace("[DJ]", "");
                //插入多个节点
                List<string> DJ = GetTxtList("DJ<>", TxtVal);
                List<string> DJ1 = GetTxtList("DJ1<>", TxtVal);
                if (DJ != null && DJ.Count() > 0 && DJ1 != null && DJ1.Count() > 0 && DJ.Count() == DJ1.Count())
                {
                    int i = 0;
                    foreach (var item in DJ)
                    {
                        var addnode = new XElement("judge");
                        addnode.Add(Environment.NewLine);
                        addnode.Add(new XElement("name") { Value = item });
                        var addtitle = new XElement("job.title") { Value = DJ1[i] };
                        addtitle.Add(new XAttribute("prefix", IsTwoUppercaseLetters(DJ1[i]) ? "no" : "yes"));

                        addnode.Add(Environment.NewLine);
                        addnode.Add(addtitle);

                        element.Add(Environment.NewLine);
                        element.Add(addnode);
                        i++;
                    }

                    TxtVal = TxtVal.Where(a => a.StartsWith("DJ<>") == false).ToArray();
                    TxtVal = TxtVal.Where(a => a.StartsWith("DJ1<>") == false).ToArray();
                }
            }

            if (element.Name == "date.group" && element.Value.Contains("[DL]"))
            {
                element.Value = element.Value.Replace("[DL]", "");
                //插入多个节点
                List<string> DL = GetTxtList("DL<>", TxtVal);

                if (DL != null && DL.Count() > 0)
                {
                    int DLINdex = 0;
                    foreach (var item in DL)
                    {
                        List<string> datalist = GetDatetime(item);

                        var addnode = new XElement("date.line") ;

                        addnode.Add(Environment.NewLine);
                        GetParXelment(addnode, datalist[0]);
                        addnode.Add(Environment.NewLine);
                        var datanode = new XElement("date") { Value = datalist[2] };



                        datanode.Add(new XAttribute("yyyymmdd", datalist[1]));
                        datanode.Add(new XAttribute("type", "unreported"));
                        datanode.Add(new XAttribute("significance", DLINdex > 0 ? "judgment" : "hearing"));
                        addnode.Add(datanode);
                        element.Add(Environment.NewLine);
                        element.Add(addnode);
                        DLINdex++;
                    }

                }
                TxtVal = TxtVal.Where(a => a.StartsWith("DL<>") == false).ToArray();
            }

            if (element.Name == "counsel.group" && element.Value.Contains("[AT]"))
            {
                element.Value = element.Value.Replace("[AT]", "");
                //插入多个节点
                List<string> AT = GetTxtList("AT<>", TxtVal);

                if (AT != null && AT.Count() > 0)
                {

                    foreach (var item in AT)
                    {
                        var addnode = new XElement("counsel.line");
                        GetParXelment(addnode, item);

                        element.Add(Environment.NewLine);
                        element.Add(addnode);
                    }

                }
                TxtVal = TxtVal.Where(a => a.StartsWith("AT<>") == false).ToArray();
            }

            if (element.Name == "judgment" && element.Value.Contains("[Content]"))
            {
                element.Value = element.Value.Replace("[Content]", "");

                GetContontNode(element);
            }
            foreach (var item in element.Attributes())
            {
                item.Value = GetXMlContent(item.Value);
            }
            if (element.Elements().Count() <= 0)
            { 
                if (element.HasElements == false && string.IsNullOrEmpty(element.Value))
                {
                    element = new XElement(element.Name, new XAttribute(XNamespace.Xml + "space", "preserve"), null, true);
                } 
                element.Value = GetXMlContent(element.Value);
            }
            return TxtVal;
            // element.rep
        }

        public string GetXMlContent(string OldValue)
        {
            string Newoldvalue = OldValue;
            if (TxtVal.Contains("CH<>"))
            {

                Newoldvalue = Newoldvalue.Replace("[CN]", "Chinese");
                Newoldvalue = Newoldvalue.Replace("[CNE]", "zh");
            }
            else if (TxtVal.Contains("EN<>"))
            {

                Newoldvalue = Newoldvalue.Replace("[CN]", "English");
                Newoldvalue = Newoldvalue.Replace("[CNE]", "en");
            }
            else
            {

                Newoldvalue = Newoldvalue.Replace("[CN]", "English");
                Newoldvalue = Newoldvalue.Replace("[CNE]", "en");
            }

            Newoldvalue = Newoldvalue.Replace("[DataNow]", DateTime.Now.ToString("yyyyMMdd"));

            if (OldValue.Contains("[TI]") || OldValue.Contains("[TI-Type]"))
            {
                string tI = GetTxtLable("TI<>", TxtVal, 0);
                if (!string.IsNullOrEmpty(tI) && tI.IndexOf("^") != -1)
                {

                    Newoldvalue = Newoldvalue.Replace("[TI]", tI.Substring(0, tI.IndexOf("^")));
                    Newoldvalue = Newoldvalue.Replace("[TI-Type]", tI.Substring(tI.IndexOf("^") + 1, tI.Length - tI.IndexOf("^") - 1));
                }
                else
                {
                    Newoldvalue = Newoldvalue.Replace("[TI]", "");
                    Newoldvalue = Newoldvalue.Replace("[TI-Type]", "");

                }
            }
            if (OldValue.Contains("[TI1]") || OldValue.Contains("[TI1-Type]"))
            {
                string tI1 = GetTxtLable("TI<>", TxtVal, 1);
                if (!string.IsNullOrEmpty(tI1) && tI1.IndexOf("^") != -1)
                {

                    Newoldvalue = Newoldvalue.Replace("[TI1]", tI1.Substring(0, tI1.IndexOf("^")));
                    Newoldvalue = Newoldvalue.Replace("[TI1-Type]", tI1.Substring(tI1.IndexOf("^") + 1, tI1.Length - tI1.IndexOf("^") - 1));
                }
                else
                {
                    Newoldvalue = Newoldvalue.Replace("[TI1]", "");
                    Newoldvalue = Newoldvalue.Replace("[TI1-Type]", "");

                }
            }
            string CITE = GetTxtLable("CITE<>", TxtVal, -1);
            Newoldvalue = Newoldvalue.Replace("[CITE]", CITE);
            string CO = "";
            if (!string.IsNullOrEmpty(CITE))
            {
                string[] parts = CITE.Split(new char[] { ' ', '/' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts != null && parts.Count() >= 3)
                {
                    CO = parts[1].Replace("HK", "");
                }

            }

            Newoldvalue = Newoldvalue.Replace("[CITE-CO]", CO);

            string DN = GetTxtLable("DN<>", TxtVal, -1);
            if (!string.IsNullOrEmpty(DN))
            {
                string[] parts = DN.Split(new char[] { ' ', '/' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts != null && parts.Count() >= 3)
                {
                    Newoldvalue = Newoldvalue.Replace("[DN-1]", parts[0]);
                    Newoldvalue = Newoldvalue.Replace("[DN-2]", parts[1]);
                    Newoldvalue = Newoldvalue.Replace("[DN-3]", parts[2]);
                }
                else
                {
                    Newoldvalue = Newoldvalue.Replace("[DN-1]", "");
                    Newoldvalue = Newoldvalue.Replace("[DN-2]", "");
                    Newoldvalue = Newoldvalue.Replace("[DN-3]", "");

                }
            }
            else
            {
                Newoldvalue = Newoldvalue.Replace("[DN-1]", "");
                Newoldvalue = Newoldvalue.Replace("[DN-2]", "");
                Newoldvalue = Newoldvalue.Replace("[DN-3]", "");

            }
            //if (OldValue.Contains("[DN-3]"))
            //{

            //    TxtVal = TxtVal.Where(a => a.StartsWith("DN<>") == false).ToArray();
            //}

            string APN = GetTxtLable("APN<>", TxtVal, -1);

            Newoldvalue = Newoldvalue.Replace("[AP1]", GetTxtLable("AP1<>", TxtVal, -1));
            //if (OldValue.Contains("[AP1]"))
            //{

            //    TxtVal = TxtVal.Where(a => a.StartsWith("AP1<>") == false).ToArray();
            //}
            if (!string.IsNullOrEmpty(APN))
            {
                string[] parts = APN.Split(new char[] { ' ', '/' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts != null && parts.Count() >= 3)
                {
                    Newoldvalue = Newoldvalue.Replace("[APN-1]", parts[0]);
                    Newoldvalue = Newoldvalue.Replace("[APN-2]", parts[1]);
                    Newoldvalue = Newoldvalue.Replace("[APN-3]", parts[2]);
                }
                else
                {
                    Newoldvalue = Newoldvalue.Replace("[APN-1]", "");
                    Newoldvalue = Newoldvalue.Replace("[APN-2]", "");
                    Newoldvalue = Newoldvalue.Replace("[APN-3]", "");
                }
            }
            else
            {
                Newoldvalue = Newoldvalue.Replace("[APN-1]", "");
                Newoldvalue = Newoldvalue.Replace("[APN-2]", "");
                Newoldvalue = Newoldvalue.Replace("[APN-3]", "");

            }
            //if (OldValue.Contains("[APN-3]"))
            //{

            //    TxtVal = TxtVal.Where(a => a.StartsWith("APN<>") == false).ToArray();
            //}

            return Newoldvalue;
        }


        public class XelementTdata
        {
            public XElement Telement { get; set; } //1层级的全部插入element     
            public int Level { get; set; }
            public bool IsEnd { get; set; } //当找到比上一个层级大的时候需要上一层级结束

            public int Index { get; set; }

        }

        public class XelementQLdata
        {
            public bool IsQStart { get; set; } = false;
            public bool IslistStart = false;
            public XElement TopParQelement { get; set; } = null;
            public XElement TopParLelement { get; set; } = null;
            public XElement ParQelement { get; set; } = null;
            public XElement ParLelement { get; set; } = null;
            public int Index { get; set; }
            public bool IsEnd { get; set; }
            public int Type { get; set; }
        }
        public XElement GetContontNode(XElement element)
        {

            bool isStart = false;


            int TIndex = 0;
            XElement TJelement = null;


            List<XelementTdata> xelementTdatas = new List<XelementTdata>();
            int index = 1;


            // bool IsQStart = false;
            //   bool IslistStart = false;
            XElement TopParQelement = null;
            XElement TopParLelement = null;
            //  XElement ParQelement = null;
            //  XElement ParLelement = null;

            List<XelementQLdata> xelementQLdatas = new List<XelementQLdata>();
            foreach (var item in TxtVal)
            {
                if (item.StartsWith("T<>"))
                {
                    isStart = true;
                }
                if (isStart)
                {
                    if (item.StartsWith("T<>"))
                    {
                        var addnode = new XElement("heading") { Value = item.Replace("T<>", "").Replace("<i>", "").Replace("</i>", "").Replace("<u>", "").Replace("</u>", "").Replace("<b>", "").Replace("</b>", "") };
                        addnode.Add(new XAttribute("align", "left"));
                        if (TIndex == 0)
                        { 
                            element.Add(addnode);
                        }
                        else
                        {
                            if (xelementTdatas != null && xelementTdatas.Count() > 0)
                            {
                                XElement Telement = xelementTdatas.Where(a => a.IsEnd == false).LastOrDefault().Telement;

                                Telement.Add(Environment.NewLine);
                                Telement.Add(addnode);
                            }
                            else
                            {
                                if (TJelement == null)
                                {

                                    TJelement = new XElement("judge.block");

                                }
                                else
                                {
                                    TJelement.Add(Environment.NewLine); 

                                }
                                TJelement.Add(addnode);

                            }

                        }
                        TIndex++;
                    }

                    if (item.StartsWith("TJ<>"))
                    {
                        if (TJelement != null)
                        {
                            ChkXml(xelementTdatas);
                            foreach (var itemdata in xelementTdatas.Where(a => a.Level == 1).OrderBy(a => a.Index).ToList())
                            {
                                TJelement.Add(Environment.NewLine);
                                TJelement.Add(itemdata.Telement);

                            }
                            element.Add(Environment.NewLine);
                            element.Add(TJelement);

                            xelementTdatas = new List<XelementTdata>();
                            TJelement = new XElement("judge.block");
                        }
                        else
                        {
                            TJelement = new XElement("judge.block");

                        }

                        var addnode = new XElement("heading") { Value = item.Replace("TJ<>", "").Replace("<i>", "").Replace("</i>", "").Replace("<u>", "").Replace("</u>", "").Replace("<b>", "").Replace("</b>", "") };
                        addnode.Add(new XAttribute("align", "left"));

                        TJelement.Add(Environment.NewLine);
                        TJelement.Add(addnode);
                    }
                    if (item.StartsWith("P<>"))
                    {
                        var addnode = new XElement("para");

                        GetParXelment(addnode, item);

                        if (xelementQLdatas != null && xelementQLdatas.Where(a => a.Type == 0 && a.IsEnd == false).Count() > 0)
                        {
                            if (xelementQLdatas != null && xelementQLdatas.Where(a => a.Type == 1 && a.IsEnd == false).Count() > 0)
                            {
                                addnode = new XElement("list.item");

                                addnode.Add(Environment.NewLine);
                                GetParXelment(addnode, item);
                                var xel = xelementQLdatas.Where(a => a.Type == 1 && a.IsEnd == false).LastOrDefault();

                                xel.ParLelement.Add(Environment.NewLine);
                                xel.ParLelement.Add(addnode);
                            }
                            else
                            {
                                TopParLelement = addnode;
                                var xel = xelementQLdatas.Where(a => a.Type == 0 && a.IsEnd == false).LastOrDefault();

                                xel.ParQelement.Add(Environment.NewLine);
                                xel.ParQelement.Add(addnode);
                            }
                        }
                        else
                        {
                            if (xelementQLdatas != null && xelementQLdatas.Where(a => a.Type == 1 && a.IsEnd == false).Count() > 0)
                            {

                                addnode = new XElement("list.item");
                                addnode.Add(Environment.NewLine);
                                GetParXelment(addnode, item);
                                var xel = xelementQLdatas.Where(a => a.Type == 1 && a.IsEnd == false).LastOrDefault();

                                xel.ParLelement.Add(Environment.NewLine);
                                xel.ParLelement.Add(addnode);
                            }
                            else
                            {
                                TopParLelement = addnode;
                                if (xelementTdatas != null && xelementTdatas.Count() > 0)
                                {
                                    XElement Telement = xelementTdatas.Where(a => a.IsEnd == false).LastOrDefault().Telement;

                                    Telement.Add(Environment.NewLine);
                                    Telement.Add(addnode);
                                }
                                else
                                {
                                    if (TJelement == null)
                                    {

                                        TJelement = new XElement("judge.block");

                                    }
                                    else {

                                        TJelement.Add(Environment.NewLine);
                                    }
                                    TJelement.Add(addnode);

                                }


                                TopParQelement = addnode;
                            }





                        }


                    }
                    if (item.StartsWith("<q>"))
                    {
                        //IsQStart = true;

                        //ParQelement = new XElement("block.quote");
                        int Xelindex = 0;
                        if (xelementQLdatas != null && xelementQLdatas.Where(a => a.Type == 0 && a.IsEnd == false).Count() > 0)
                        {
                            Xelindex = xelementQLdatas.Where(a => a.Type == 0 && a.IsEnd == false).Max(a => a.Index) + 1;
                        }

                        xelementQLdatas.Add(new XelementQLdata() { ParQelement = new XElement("block.quote"), IsEnd = false, Index = Xelindex, Type = 0 });


                        //  TopParQelement.Add(ParQelement);

                    }
                    if (item.StartsWith("</q>"))
                    {
                        if (xelementQLdatas != null && xelementQLdatas.Where(a => a.Type == 0 && a.IsEnd == false).Count() > 0)
                        {
                            var xel = xelementQLdatas.Where(a => a.Type == 0 && a.IsEnd == false).LastOrDefault();


                            TopParQelement.Add(Environment.NewLine);
                            TopParQelement.Add(xel.ParQelement);
                            xel.IsEnd = true;
                        }

                        //  IsQStart = false;
                        //  TopParQelement.Add(ParQelement);
                        //  ParQelement = null;
                        //   TopParQelement = null;
                    }
                    if (item.StartsWith("<list>"))
                    {
                        //  IslistStart = true;
                        int Xelindex = 0;
                        if (xelementQLdatas != null && xelementQLdatas.Where(a => a.Type == 1 && a.IsEnd == false).Count() > 0)
                        {
                            Xelindex = xelementQLdatas.Where(a => a.Type == 1 && a.IsEnd == false).Max(a => a.Index) + 1;
                        }

                        xelementQLdatas.Add(new XelementQLdata() { ParLelement = new XElement("list"), IsEnd = false, Index = Xelindex, Type = 1 });
                        //  ParLelement = new XElement("list"); 

                    }
                    if (item.StartsWith("</list>"))
                    {
                        //   IslistStart = false;
                        if (xelementQLdatas != null && xelementQLdatas.Where(a => a.Type == 1 && a.IsEnd == false).Count() > 0)
                        {
                            var xel = xelementQLdatas.Where(a => a.Type == 1 && a.IsEnd == false).LastOrDefault();

                            TopParLelement.Add(Environment.NewLine);
                            TopParLelement.Add(xel.ParLelement);
                            xel.IsEnd = true;
                        }
                        //if (IsQStart)
                        //  {

                        //      ParQelement.Add(TopParLelement);
                        //  }
                        //  TopParLelement = null;
                        //   ParLelement = null;
                    }
                    if (item.StartsWith("T1<>") || item.StartsWith("T2<>") || item.StartsWith("T3<>") || item.StartsWith("T4<>"))
                    {
                        int level = Convert.ToInt32(item.Substring(1, 1));
                        if (xelementTdatas != null && xelementTdatas.Count() > 0)
                        {

                            foreach (var xel in xelementTdatas.Where(a => a.Level >= level && a.IsEnd == false))
                            {
                                xel.IsEnd = true;
                            }
                        }

                        var addnode = new XElement("heading") { Value = item.Replace("T1<>", "").Replace("T2<>", "").Replace("T3<>", "").Replace("T4<>", "").Replace("<i>", "").Replace("</i>", "").Replace("<u>", "").Replace("</u>", "").Replace("<b>", "").Replace("</b>", "") };
                        addnode.Add(new XAttribute("align", "left"));

                        var tpar = new XElement("para.group");

                        tpar.Add(Environment.NewLine);
                        tpar.Add(addnode);
                        xelementTdatas.Add(new XelementTdata() { Level = level, Telement = tpar, IsEnd = false, Index = index });
                        index++;

                    }


                }

            }



            //往上累加
            if (TJelement == null)
            {

                TJelement = new XElement("judge.block");

            }
            ChkXml(xelementTdatas);
            foreach (var item in xelementTdatas.Where(a => a.Level == 1).OrderBy(a => a.Index).ToList())
            {
                TJelement.Add(Environment.NewLine);
                TJelement.Add(item.Telement);

            }

            if (TJelement != null)
            {
                element.Add(Environment.NewLine);
                element.Add(TJelement);
            }



            return element;
        }

        public void GetParXelment(XElement xElement, string ParTextVal)
        {
            ParTextVal = ParTextVal.Replace("P<>", "");

            string[] parts = ParTextVal.Split(new string[] { "<" }, StringSplitOptions.None);
            int divindex = 0;
            List<string> strings = new List<string>();

            string TopName = "";
            Dictionary<int, List<string>> F = new Dictionary<int, List<string>>();
            //处理标签
            if (ParTextVal.Contains("<f"))
            {
                F = GetTxtFNList("F<>", TxtVal);

            }

            foreach (string part in parts)
            {
                if (!string.IsNullOrEmpty(part))
                {


                    string newpart = part;
                    if (newpart.StartsWith("/") && newpart.IndexOf(">") != -1 && !newpart.EndsWith(">") && !TopName.StartsWith("<c"))
                    {
                        TopName = "";
                        strings.Add(newpart.Substring(newpart.IndexOf(">") + 1, newpart.Length - newpart.IndexOf(">") - 1));
                    }
                    else
                    {
                        if (divindex != 0)
                        {
                            newpart = "<" + part;


                        }
                        if (newpart.StartsWith("</") && newpart.EndsWith(">"))
                        {
                            TopName = "";
                            strings.Add(newpart.Substring(newpart.IndexOf(">") + 1, newpart.Length - newpart.IndexOf(">") - 1));
                        }
                        else if (newpart.StartsWith("<") && newpart.EndsWith(">"))
                        {
                            if (!newpart.StartsWith("<f"))
                            {

                                TopName += newpart;
                            }
                            else
                            {

                                TopName = "";
                                strings.Add(newpart);
                            }


                        }
                        else
                        {
                            string newtxt = TopName + newpart;
                            if (newtxt.Contains("<f"))
                            {
                                string[] parts1 = newtxt.Split(new string[] { "<f" }, StringSplitOptions.None);
                                foreach (string partf in parts1)
                                {
                                    if (!string.IsNullOrEmpty(partf))
                                    {
                                        string newpartf = "<f" + partf;

                                        strings.Add(newpartf.Substring(0, newpartf.IndexOf(">") + 1));

                                        strings.Add(newpartf.Substring(newpartf.IndexOf(">") + 1, newpartf.Length - newpartf.IndexOf(">") - 1));

                                    }
                                }

                            }
                            else
                            {
                                if (newtxt.Contains("<c"))
                                {
                                    if (newtxt.Contains("</c>"))
                                    {

                                        strings.Add(newtxt.Substring(0, newtxt.IndexOf("</c>")));

                                        if (!newpart.EndsWith("</c>"))
                                        {
                                            strings.Add(newtxt.Substring(newtxt.IndexOf("</c>") + 4, newtxt.Length - newtxt.IndexOf("</c>") - 4));
                                        }
                                        TopName = "";
                                    }
                                    else
                                    {
                                        TopName = newtxt;
                                    }

                                }
                                else
                                {
                                    strings.Add(newtxt);


                                    TopName = "";

                                }

                            }

                        }

                    }
                }
                divindex++;
            }

            foreach (var item in strings)
            {
                if (item.StartsWith("<label>"))
                {
                  //  xElement.Add(Environment.NewLine);
                    xElement.Add(new XElement("label", item.Replace("<label>", "")));
                }
                if ((item.Contains("<i>") || item.Contains("<b>") || item.Contains("<u>") || item.Contains("<sub>") || item.Contains("<sup>")) && !item.StartsWith("<c>"))
                {
                    string type = "";
                    if (item.StartsWith("<i><b><u>"))
                    {
                        type = "bold_italic_underline";
                    }
                    else if (item.StartsWith("<i><u>"))
                    {
                        type = "italic_underline";
                    }
                    else if (item.StartsWith("<i><b>"))
                    {
                        type = "bold_italic";
                    }
                    else if (item.StartsWith("<b><u>"))
                    {
                        type = "bold_underline";
                    }
                    else if (item.StartsWith("<sup>"))
                    {
                        type = "sup";
                    }
                    else if (item.StartsWith("<sub>"))
                    {
                        type = "sub";
                    }
                    else if (item.StartsWith("<u>"))
                    {
                        type = "underline";
                    }
                    else if (item.StartsWith("<b>"))
                    {
                        type = "bold";
                    }
                    else if (item.StartsWith("<i>"))
                    {
                        type = "italic";
                    }

                    XElement xElement1 = new XElement("emphasis", item.Replace("<i>", "").Replace("<sub>", "").Replace("<u>", "").Replace("<sup>", "").Replace("<b>", ""));
                    xElement1.Add(new XAttribute("type", type));

                  //  xElement.Add(Environment.NewLine);
                    xElement.Add(xElement1);
                }
                if (item.StartsWith("<f"))
                {
                    //得到f数量
                    string fval = item.Substring(item.IndexOf("<f") + 2, item.IndexOf(">") - item.IndexOf("<f") - 2);
                    if (fval != null)
                    {
                        int FIndex = Convert.ToInt32(fval);

                        if (F.ContainsKey(FIndex))
                        {
                            XElement FontXel = new XElement("footnote");

                            FontXel.Add(new XAttribute("label", FIndex));
                            foreach (string FVal in F[FIndex])
                            {

                                XElement parXel = new XElement("para");
                                GetParXelment(parXel, FVal.Replace("F<>", "").Replace("FN<>", "").Trim());

                                FontXel.Add(Environment.NewLine);
                                FontXel.Add(parXel);

                            }
                            xElement.Add(Environment.NewLine);
                            xElement.Add(FontXel);
                        }

                    }

                }
                if (item.StartsWith("<c"))
                {
                    XElement xElement1 = new XElement("case.considered");


                    string titile = item.Substring(item.IndexOf("<i>") + 3, item.IndexOf("</i>") - item.IndexOf("<i>") - 3).Trim();
                    string citecitation = item.Substring(item.IndexOf("</i>") + 4, item.Length - item.IndexOf("</i>") - 4).Trim();
                    if (citecitation.StartsWith(",") || citecitation.StartsWith("，"))
                    {
                        citecitation = citecitation.Substring(1, citecitation.Length - 1).Trim();


                    }
                    //string html = "<case.ref BVtable=\"yes\"><citetitle type=\"case\" full =\""+titile+ "\" legtype =\"ord\">"+titile+"</citetitle> <citecitation full=\""+ citecitation + "\">"+ citecitation + "</citecitation></case.ref>";
                    //xElement1.Value = html;
                    XElement xElementc = new XElement("case.ref");
                    xElementc.Add(new XAttribute("BVtable", "yes"));


                    XElement xElementt = new XElement("citetitle") { Value = titile +" " };

                    xElementt.Add(new XAttribute("type", "case"));
                    xElementt.Add(new XAttribute("full", titile));
                    xElementt.Add(new XAttribute("legtype", "ord"));

                    //  xElementc.Add(Environment.NewLine);
                    xElementc.Add(xElementt); 


                    XElement xElementci = new XElement("citecitation") { Value = citecitation };
                    xElementci.Add(new XAttribute("full", citecitation));

                    //  xElementc.Add(Environment.NewLine);
                    xElementc.Add(xElementci);
                    //  xElement1.Add(Environment.NewLine);
                    xElement1.Add(xElementc);

                    xElement.Add(Environment.NewLine);
                    xElement.Add(xElement1);
                }
                if (item.StartsWith("<") == false)
                {
                   // xElement.Add(Environment.NewLine);

                    xElement.Add(item.Trim());
                }
            }

            //    xElement.Value = ParTextVal.Replace("P<>", "");


        }
        public void ChkXml(List<XelementTdata> xelementTdatas)
        {
            //往上累加
            for (int i = 4; i > 1; i--)
            {
                foreach (var item in xelementTdatas.Where(a => a.Level == i).OrderBy(a => a.Index).ToList())
                {

                    //能找到上一层级最后一个
                    var xeldata = xelementTdatas.Where(a => a.Index < item.Index && a.Level < item.Level).OrderBy(a => a.Index).LastOrDefault();
                    //当前index
                    XElement xElement = xeldata.Telement;

                    xElement.Add(Environment.NewLine);
                    xElement.Add(item.Telement);

                }

            }

        }

        public List<string> GetDatetime(string text)
        {

            string dataTime = "";
            DateTime date = DateTime.Now;
            string datatimeval = "";
            List<string> list = new List<string>();
            string CNdata = @"\d{4}年(0?[1-9]|1[012])月(0?[1-9]|[12][0-9]|3[01])日";

            string endata = @"(January|February|March|April|May|June|July|August|September|October|November|December) \d{1,2} \d{4}";
            string endata1 = @"\d{1,2} (January|February|March|April|May|June|July|August|September|October|November|December) \d{4}";

            Regex regex = new Regex(CNdata);
            foreach (Match match in regex.Matches(text))
            {
                datatimeval = match.Value;

                int year = Convert.ToInt32(datatimeval.Substring(0, datatimeval.IndexOf("年")));
                int month = Convert.ToInt32(datatimeval.Substring(datatimeval.IndexOf("年") + 1, datatimeval.IndexOf("月") - datatimeval.IndexOf("年") - 1));
                int day = Convert.ToInt32(datatimeval.Substring(datatimeval.IndexOf("月") + 1, datatimeval.IndexOf("日") - datatimeval.IndexOf("月") - 1));
                date = new DateTime(year, month, day);

            }
            Dictionary<int, string> months = new Dictionary<int, string>();
            months.Add(1, "January");
            months.Add(2, "February");
            months.Add(3, "March");
            months.Add(4, "April");
            months.Add(5, "May");
            months.Add(6, "June");
            months.Add(7, "July");
            months.Add(8, "August");
            months.Add(9, "September");
            months.Add(10, "October");
            months.Add(11, "November");
            months.Add(12, "December");
            regex = new Regex(endata);
            foreach (Match match in regex.Matches(text))
            {
                datatimeval = match.Value;

                int year = Convert.ToInt32(datatimeval.Substring(datatimeval.Length - 4, 4));
                int day = Convert.ToInt32(datatimeval.Substring(datatimeval.Length - 7, 2).Trim());
                string monthvaltext = datatimeval.Substring(0, datatimeval.Length - 8).Trim();
                int month = months.Where(a => a.Value == monthvaltext).Select(a => a.Key).FirstOrDefault();



                date = new DateTime(year, month, day);
            }
            regex = new Regex(endata1);
            foreach (Match match in regex.Matches(text))
            {
                datatimeval = match.Value;

                int year = Convert.ToInt32(datatimeval.Substring(datatimeval.Length - 4, 4));
                int day = Convert.ToInt32(datatimeval.Substring(0, 2).Trim());

                string monthvaltext = datatimeval.Substring(2, datatimeval.Length - 7).Trim();
                int month = months.Where(a => a.Value == monthvaltext).Select(a => a.Key).FirstOrDefault();

                date = new DateTime(year, month, day);
            }


            dataTime = text.Replace(datatimeval, "");

            list.Add(dataTime); list.Add(date.ToString("yyyyMMdd")); list.Add(datatimeval);
            return list;
        }
        public string GetTxtLable(string lableKey, string[] TxtVal, int Orderindex)
        {
            string TxtValue = "";
            List<string> txtlist = TxtVal.Where(a => a.StartsWith(lableKey)).ToList();
            if (Orderindex >= 0)
            {
                if (txtlist != null && txtlist.Count() > 0 && txtlist.Count() >= (Orderindex + 1))
                {
                    TxtValue = txtlist[Orderindex];

                }
            }
            else
            {
                TxtValue = txtlist.FirstOrDefault();
            }

            if (!string.IsNullOrEmpty(TxtValue))
            {
                TxtValue = TxtValue.Replace(lableKey, "");


            }
            return TxtValue;
        }


        public List<string> GetTxtList(string lableKey, string[] TxtVal)
        {
            List<string> txtlist = TxtVal.Where(a => a.StartsWith(lableKey)).Select(a => a.Replace(lableKey, "").Replace(lableKey, "&")).ToList();

            return txtlist;
        }
        public Dictionary<int, List<string>> GetTxtFNList(string lableKey, string[] TxtVal)
        {
            List<string> txtlist = TxtVal.Where(a => a.StartsWith(lableKey) || a.StartsWith("FN<>")).ToList();
            Dictionary<int, List<string>> dic = new Dictionary<int, List<string>>();
            int index = 1;
            foreach (string txt in txtlist)
            {
                if (txt.StartsWith("F<>"))
                {
                    string val = txt.Substring(3, txt.IndexOf(".") - 3);
                    if (!string.IsNullOrEmpty(val))
                    {
                        index = Convert.ToInt32(val);
                    }
                    dic.Add(index, new List<string>() { txt.Replace("F<>" + index + ".", "") });
                    index++;
                }
                else
                {
                    var dick = dic.LastOrDefault();
                    dick.Value.Add(txt);
                }
            }
            return dic;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word文档|*.doc;*.docx;*.txt";
            openFileDialog.Multiselect = true;


            if (openFileDialog.ShowDialog() == DialogResult.OK)
            { textBox1.Text = openFileDialog.FileName; }
        }
    }
}
