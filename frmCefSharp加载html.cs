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
            browser = new ChromiumWebBrowser(Application.StartupPath + "\\TinyMce\\index.html");
            //browser.BrowserSettings.Javascript = CefState.Enabled;
            //browser.JavascriptObjectRepository.Settings.LegacyBindingEnabled = true;
             


            this.panel1.Controls.Add(browser);
            //boundObject = new BoundObject();
            //boundObject.ParentRichTextBox = this.richTextBox1;
            browser.Dock = DockStyle.Fill;
            //browser.JavascriptObjectRepository.Register("boundObject", boundObject, false, BindingOptions.DefaultBinder);//获取html上的编辑器的值
            // 处理浏览器加载完成事件
            //  browser.FrameLoadEnd += Browser_FrameLoadEnd;


            //// 加载网页完成后执行脚本
            //browser.LoadingStateChanged += (sender, args) =>
            //{
            //    if (args.IsLoading == false)
            //    {
            //        // 执行 JavaScript 脚本给 div 传值
            //        browser.ExecuteScriptAsync("document.getElementById('display-content').innerHTML = '<p>23223123234334</p><p>32245234</p><p>234452345</p>';");
            //      //  browser.ExecuteScriptAsync($"editor.setData('<p>23223123234334</p><p>32245234</p><p>234452345</p>');");
            //    }
            //};
        }
        public class BoundObject
        {
            public RichTextBox ParentRichTextBox { get; set; }
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
