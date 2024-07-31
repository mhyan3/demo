using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{

    public class ReportContext : IDisposable
    {
        public Microsoft.Office.Interop.Word.Paragraph Vernier { get; set; }
        public Microsoft.Office.Interop.Word.Application App { get; set; }
        public Microsoft.Office.Interop.Word.Document Doc { get; set; }
        public ReportContext(string Path)
        {
            if (string.IsNullOrEmpty(Path))
            {
                throw new ArgumentNullException("文档路径无效");
            }
            if (!File.Exists(Path))
            {
                throw new ArgumentNullException("文档路径无效");

            }
            App = new Microsoft.Office.Interop.Word.Application();
            object fileName = Path;
            Doc = App.Documents.Open(ref fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Vernier = Doc.Paragraphs.Count > 0 ? Doc.Paragraphs[1] : null;
        }
        public ReportContext(string Path, Microsoft.Office.Interop.Word.Application _App, Microsoft.Office.Interop.Word.Document _Doc)
        {
            if (string.IsNullOrEmpty(Path))
            {
                throw new ArgumentNullException("文档路径无效");
            }
            if (!File.Exists(Path))
            {
                throw new ArgumentNullException("文档路径无效");

            }
            App = _App;
            object fileName = Path;
            Doc = _Doc;
            Vernier = Doc.Paragraphs.Count > 0 ? Doc.Paragraphs[1] : null;
        }
        public void MoveNext()
        {
            Vernier = Vernier == null ? null : Vernier.Next();
        }
        public void MoveLast()
        {
            Vernier = Vernier == null ? null : Vernier.Previous();
        }
        public void Save()
        {
            Doc.Save();
        }
        public async void Dispose()
        {
            try
            {

                Doc.Close();
                App.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(Doc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(App);
            }
            catch (Exception ex)
            {
                App.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(App);
            }
        }
    }
}
