using CefSharp;
using CefSharp.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms; 

namespace WindowsFormsApp3
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public partial class frmCefSharp加载html : Form
    {
        public ChromiumWebBrowser browser;
        private BoundObject boundObject;
        public void ReceiveMessageFromWeb(string message)
        {
            richTextBox1.Text = message; // 将消息显示在文本框中
        }
        public frmCefSharp加载html()
        {
            InitializeComponent();
            CefSettings settings = new CefSettings();
            // 初始化 CefSharp  
            Cef.Initialize(settings);

       // 创建并添加 CefSharp 浏览器控件  
browser = new ChromiumWebBrowser(System.Windows.Forms.Application.StartupPath + "\\TinyMce\\index.html");
boundObject = new BoundObject();
boundObject.form1 = this;
browser.BrowserSettings.Javascript = CefState.Enabled;
browser.Dock = DockStyle.Fill;

browser.JavascriptObjectRepository.Settings.LegacyBindingEnabled = true;
browser.JavascriptObjectRepository.Register("boundObject", boundObject, false, BindingOptions.DefaultBinder);//获取html上的text的值
this.panel2.Controls.Add(browser);


browser.LoadingStateChanged += Browser_LoadingStateChanged;
        } private async void Browser_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
 {
     if (!e.IsLoading)
     {
       
            browser.ExecuteScriptAsync("LoadHtml(\"" + html + "\");");
     }
 }
        public class BoundObject
        {
            public RichTextBox ParentRichTextBox { get; set; }
             public frmcs form1 { get; set; } 
            public void sendContentFromCKEditor(string content)
            {
                if (ParentRichTextBox.InvokeRequired)
                {
                    //获取html上的编辑器的值
                    ParentRichTextBox.Invoke((MethodInvoker)delegate
                    {
                        ParentRichTextBox.Text = content.Replace("<figure class=\"table\">", "").Replace("</figure>", "").Replace("<", "\n<");
                    });
                }
                else
                {
                    ParentRichTextBox.Text = content.Replace("<figure class=\"table\">", "").Replace("</figure>", "").Replace("<", "\n<");
                }
            }
        }
        //private void Browser_FrameLoadEnd(object sender, FrameLoadEndEventArgs e)
        //{
        //    if (e.Frame.IsMain)
        //    {
        //        // 在页面加载完成后，添加 JavaScript 代码来监听 CKEditor 内容的改变
        //        browser.ExecuteScriptAsync(@"
        //            CKEDITOR.instances.editor.document.on('change', function() {
        //                var content = CKEDITOR.instances.editor.getData();
        //                boundObject.sendContentFromCKEditor(content);
        //            });
        //        ");
        //    }
        //}
        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            Cef.Shutdown();
        }

        private async  void button1_Click(object sender, EventArgs e)
        {
            // 执行 JavaScript 代码获取 div 内容
            CefSharp.JavascriptResponse  response = await browser.EvaluateScriptAsync("document.getElementById('display-content').innerHTML");

            // 检查返回值是否成功
            if (response.Success)
            {
                // 获取返回的 div 内容
                string divContent = response.Result?.ToString();
                richTextBox1.Text = divContent;
                // 在这里处理获取到的 div 内容
                Console.WriteLine(divContent);
            }
            else
            {
                // 处理执行 JavaScript 代码时的错误
                Console.WriteLine($"JavaScript 执行错误: {response.Message}");
            }

        }

        //private void richTextBox1_TextChanged(object sender, EventArgs e)
        //{
        //    browser.ExecuteScriptAsync($"e.setData('{richTextBox1.Text.Replace("\n<", "<")}');");

        //}
    }
}
