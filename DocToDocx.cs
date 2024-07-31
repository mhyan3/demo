using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public class DocToDocx
    {
        /// <summary>
        /// docx转docx
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static string  ToDocxSaveAs(string path, bool isInit = true)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();
            Object oMissing = System.Reflection.Missing.Value;
            Object saveto = path.ToLower().EndsWith(".docx") ? path : path.Replace(".doc", ".docx");
           
            int hWnd = 0;
            string tempFilePath = string.Empty;
            try
            {
                app.Visible = false;
                object openType = Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatWebPages;
                object filepath = path;
                object confirmconversion = false;
                object readOnly = false;
                object visible = true;

                object oallowsubstitution = System.Reflection.Missing.Value;
                object fileFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;


                doc = app.Documents.Open(ref filepath, ref confirmconversion, ref readOnly, ref oMissing,
                                              ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                              ref oMissing, ref oMissing, ref oMissing, ref visible,
                                              ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                //初始化页面布局
                if (isInit == true)
                {
                    CreateDocumentAndInit(app, doc);
                }
                try
                {
                    hWnd = app.ActiveWindow.Hwnd;
                }
                catch (Exception ex)
                {
                }
            //    doc.SetCompatibilityMode((int)Microsoft.Office.Interop.Word.WdCompatibilityMode.wdWord2007);
                doc.SaveAs(ref saveto, ref fileFormat, ref oMissing, ref oMissing, ref oMissing,
                               ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                               ref oMissing, ref oMissing, ref oMissing, ref oallowsubstitution, ref oMissing,
                               ref oMissing);
                doc.Close();
                app.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }
            catch (Exception ex)
            {

                doc.Close();
               

                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

            }
            path = path.ToLower().EndsWith(".docx") ? path : path.Replace(".doc", ".docx");
            return path;
        } /// <summary>
          /// 设置当前选中内容所在页的纸张方向为纵向及其页边距
          /// </summary>
          /// <param name="section">当前选中内容</param>
        public static void SetVerticalPageMargin(Microsoft.Office.Interop.Word.Section section)
        {
            SetVerticalPageMargin(section?.PageSetup);
        }
        /// <summary>
        /// 设置纸张方向为纵向及其页边距
        /// </summary>
        /// <param name="pageSetup"></param>
        private static void SetVerticalPageMargin(Microsoft.Office.Interop.Word.PageSetup pageSetup)
        {
            if (pageSetup != null)
            {
                /**
                 * 调整注释
                 * 说明：根据该处的赋值对象判断，是在设置Word方向为纵向
                 * @tangp 2020-11-13
                 */
                pageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait;
                pageSetup.TopMargin = (float)0.6 * 72;
                pageSetup.BottomMargin = (float)0.3 * 72;
                pageSetup.LeftMargin = (float)0.7 * 72;
                pageSetup.RightMargin = (float)0.5 * 72;
                pageSetup.HeaderDistance = (float)0.6 * 72;
                pageSetup.FooterDistance = (float)0.12 * 72;
            }
        } /// <summary>
          /// 初始化页面布局
          /// </summary>
          /// <param name="app"></param>
          /// <param name="doc"></param>
          /// <returns></returns>
        public static Microsoft.Office.Interop.Word.Document CreateDocumentAndInit(Microsoft.Office.Interop.Word.Application app, Microsoft.Office.Interop.Word.Document doc)
        {


            Microsoft.Office.Interop.Word.Options options = app.Options;

            options.SuggestFromMainDictionaryOnly = false;
            options.IgnoreUppercase = false;
            options.IgnoreMixedDigits = false;
            options.IgnoreInternetAndFileAddresses = false;
            options.RepeatWord = false;
            doc.SpellingChecked = false;
            doc.ShowSpellingErrors = false;
            doc.PageSetup.PaperSize = Microsoft.Office.Interop.Word.WdPaperSize.wdPaperA4;
            doc.SelectAllEditableRanges();
            SetVerticalPageMargin(doc.Sections.First);

            app.Selection.WholeStory();
            app.Selection.Font.Name = "宋体";
            app.Selection.Font.Name = "Times New Roman";



            // 单倍行距
            app.Selection.ParagraphFormat.SpaceBefore = 0;
            app.Selection.ParagraphFormat.SpaceAfter = 0;
            app.Selection.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle;

            app.Selection.ParagraphFormat.KeepWithNext = 0;
            app.Selection.ParagraphFormat.KeepTogether = 0;
            doc.Paragraphs.SpaceBeforeAuto = 0;
            doc.Paragraphs.SpaceAfterAuto = 0;

            //控制整体网格属性
            app.Selection.PageSetup.LayoutMode = Microsoft.Office.Interop.Word.WdLayoutMode.wdLayoutModeDefault;

            //设置标题样式 重写Heading 1   
            Microsoft.Office.Interop.Word.Style style = null;
            try { style = doc.Styles["Heading 1"]; } catch { style = doc.Styles["标题 1"]; }
            if (style != null)
            {
                style.Font.Name = "宋体";
                style.Font.Name = "Times New Roman";
                style.Font.Size = 12;
                style.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;


                style.ParagraphFormat.TabStops.Add(app.InchesToPoints(0.5f), Alignment.Left, Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces);

                style.ParagraphFormat.SpaceBefore = 0;
                style.ParagraphFormat.SpaceAfter = 0;
                style.ParagraphFormat.SpaceBeforeAuto = 0;
                style.ParagraphFormat.SpaceAfterAuto = 0;
                style.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle;

                style.ParagraphFormat.KeepWithNext = 0;
                style.ParagraphFormat.KeepTogether = 0;
            }

            //设置normal样式 重写正文
            style = null;
            try { style = doc.Styles["Normal"]; } catch { style = doc.Styles["正文"]; }
            if (style != null)
            {
                style.Font.Name = "宋体";
                style.Font.Name = "Times New Roman";
                style.Font.Size = 12;
                style.Font.Bold = -1;
                style.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;

                style.ParagraphFormat.SpaceBefore = 0;
                style.ParagraphFormat.SpaceAfter = 0;
                style.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle;

                style.ParagraphFormat.KeepWithNext = 0;
                style.ParagraphFormat.KeepTogether = 0;

            }

            ClearPageFooter(doc.Sections.First);
            doc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[1].Format.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
            doc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[1].Format.Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom].Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;
            doc.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Paragraphs[1].Format.Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom].LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone;

            return doc;
        }
        public static void ClearPageFooter(Microsoft.Office.Interop.Word.Section section)
        {
            section.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = string.Empty;
        }
        public enum Alignment
        {
            Left,
            Right,
            Center,
            General
        }
    }
}
