using mshtml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace 提取答案
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            isFirstExam = false;
            htmlStr = File.ReadAllText(@"C:\Users\ght\Desktop\a.txt");
            //使用Microsoft Internet Controls取得所有的已经打开的IE(以Tab计算)
            SHDocVw.ShellWindows sws = new SHDocVw.ShellWindows();
            //sws为当前打开的所有IE窗口每个一个Tab都可以操作，每个Tab对应Com Object的SHDocVw.InternetExplorer
            foreach (SHDocVw.InternetExplorer iw in sws)
            {
                if (iw.LocationName.Contains("查看考试结果")) //iw.LocationURL 当前链接路径   考试
                {
                    //全部答案
                    string daan = "";
                    //提取答案块
                    Regex regex = new Regex(@"{""answer"".*}]'\s*/>", RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.CultureInvariant);
                    if (regex.IsMatch(htmlStr))
                    {
                        MatchCollection matchCollection = regex.Matches(htmlStr);
                        daan = matchCollection[0].ToString().Replace("/>", "").Trim().Replace("]'", "");
                    }
                    //把答案分组
                    string strTmp = "},{";
                    string[] lstDaAn = Regex.Split(daan, strTmp, RegexOptions.IgnoreCase);
                    for (int i = 0; i < lstDaAn.Length; i++)
                    {
                        if (lstDaAn[i].Contains("{"))
                        {
                            lstDaAn[i] = lstDaAn[i].Replace("{", "");
                        }
                        //第一条会有个{   最后一条最后有个}
                        if (lstDaAn[i].Contains("}"))
                        {
                            lstDaAn[i] = lstDaAn[i].Replace("}", "");
                        }
                        //"type":"single"   单选  "realAnswer":"D"  
                        if (lstDaAn[i].Contains("\"type\":\"single\""))
                        {
                            //正确答案选项
                            string xuanxiang = lstDaAn[i].Substring(lstDaAn[i].IndexOf("\"realAnswer\":\"") + "\"realAnswer\":\"".Length, 1);
                            //题目编号
                            string id = lstDaAn[i].Substring(lstDaAn[i].IndexOf("\"id\":\"") + "\"id\":\"".Length, lstDaAn[i].IndexOf("\", \"optionOrder\"") - lstDaAn[i].IndexOf("\"id\":\"") - "\"id\":\"".Length);
                            //自动答题 通过设置cookie直接答题
                            MessageBox.Show(id);






                            //取得每个Tab之后，就可以通过InternetExplorer的Document取得每个页面的Dom
                           // HTMLDocument doc = (HTMLDocument)iw.Document;
                            //通过DOM操作IE页面
                            //mshtml.IHTMLElementCollection inputs = (mshtml.IHTMLElementCollection)doc2.all.tags("INPUT");
                            //mshtml.HTMLInputElement input1 = (mshtml.HTMLInputElement)inputs.item("kw1", 0);
                            //input1.value = "test";
                            //mshtml.IHTMLElement element2 = (mshtml.IHTMLElement)inputs.item("su1", 0);
                            //element2.click();

                            ////变换一下，通过答案内容确定
                            //mshtml.IHTMLElementCollection allLi= (mshtml.IHTMLElementCollection)doc.all.tags("li");
                            //IEnumerable<IHTMLElement> EnHEColl = allLi.Cast<IHTMLElement>();
                            //////这个he1就是题干元素
                            ////IHTMLElement he1 = EnHEColl.FirstOrDefault(p => p.innerHTML != null && p.innerHTML.Contains(content));
                            //////这个是答案
                            ////IHTMLElement he2 = EnHEColl.FirstOrDefault(p => p.innerHTML != null && p.innerHTML.Contains("A.民办高校专业技术人员"));
                            ////MessageBox.Show(he2.innerHTML);
                            //////这个he2是li元素，他里面还有个input
                            //////看桌面的txt

                            //foreach (IHTMLElement elt in EnHEColl)
                            //{
                            //    if (elt.innerHTML.Contains(content))
                            //    {
                            //        //这个elt就是题干的li元素
                            //        //elt.innerText
                            //    }
                            //}


                        }
                        //"type":"multiple" 多选
                        else if (lstDaAn[i].Contains("\"type\":\"multiple\""))
                        {
                            //正确答案选项
                            string xuanxiang = lstDaAn[i].Substring(lstDaAn[i].IndexOf("\"realAnswer\":\"") + "\"realAnswer\":\"".Length, lstDaAn[i].IndexOf("\",\"realScore\"") - lstDaAn[i].IndexOf("\"realAnswer\":\"") - "\"realAnswer\":\"".Length);
                            //题目
                            string content = lstDaAn[i].Substring(lstDaAn[i].IndexOf("\"content\":\"") + "\"content\":\"".Length, lstDaAn[i].IndexOf("\",\"filePath\"") - lstDaAn[i].IndexOf("\"content\":\"") - "\"content\":\"".Length);
                            //自动答题
                        }
                        //"type":"judge"判断
                        else if (lstDaAn[i].Contains("\"type\":\"judge\""))
                        {
                            //正确答案选项
                            string xuanxiang = lstDaAn[i].Substring(lstDaAn[i].IndexOf("\"realAnswer\":\"") + "\"realAnswer\":\"".Length, 1);
                            //题目
                            string content = lstDaAn[i].Substring(lstDaAn[i].IndexOf("\"content\":\"") + "\"content\":\"".Length, lstDaAn[i].IndexOf("\",\"filePath\"") - lstDaAn[i].IndexOf("\"content\":\"") - "\"content\":\"".Length);
                            //自动答题
                        }
                    }
                }
            }
        }
        //是否是第一次考试的查看考试结果,默认是
        bool isFirstExam = true;
        //第一次考试后查看考试结果的页面的代码
        string htmlStr = "";
        private void watchExam_Tick(object sender, EventArgs e)
        {
            //第一次考试的结果，目的要取得整个页面源代码
            if (isFirstExam)
            {
                try
                {
                    //使用Microsoft Internet Controls取得所有的已经打开的IE(以Tab计算)
                    SHDocVw.ShellWindows sws = new SHDocVw.ShellWindows();
                    //sws为当前打开的所有IE窗口每个一个Tab都可以操作，每个Tab对应Com Object的SHDocVw.InternetExplorer
                    foreach (SHDocVw.InternetExplorer iw in sws)
                    {
                        //提交后查看考试结果
                        if (iw.LocationName.Contains("答题后")) //iw.LocationURL 当前链接路径    查看考试结果
                        {
                            //取得每个Tab之后，就可以通过InternetExplorer的Document取得每个页面的Dom
                            mshtml.HTMLDocumentClass doc = (mshtml.HTMLDocumentClass)iw.Document;
                            mshtml.HTMLBody body = (mshtml.HTMLBody)doc.body;
                            htmlStr = body.innerHTML.ToString();
                            isFirstExam = false;

                            //取得Dom之后，基本上就已经取得了操作IE的所有权限了，可以继续使用HTML Object Library对页面进行操作
                            //或者通过注册JavaScript，对页面进行操作：
                            //mshtml.IHTMLScriptElement script = dom.createElement("script") as mshtml.IHTMLScriptElement; \\创建script标签
                            //script.text = "$(\"[name='wd']\").val('刘德华');"; \\通过Jquery，对百度进行操作
                            //mshtml.HTMLBody body = dom.body as mshtml.HTMLBody; \\取得body对象
                            //body.appendChild((mshtml.IHTMLDOMNode)script); \\注册JavaScript
                            //关闭
                            //iw.Quit();
                        }
                    }
                }
                catch { }
            }
            else //第二次进入考试，在这操作答案
            {
                //使用Microsoft Internet Controls取得所有的已经打开的IE(以Tab计算)
                SHDocVw.ShellWindows sws = new SHDocVw.ShellWindows();
                //sws为当前打开的所有IE窗口每个一个Tab都可以操作，每个Tab对应Com Object的SHDocVw.InternetExplorer
                foreach (SHDocVw.InternetExplorer iw in sws)
                {
                    if (iw.LocationName.Contains("答题前")) //iw.LocationURL 当前链接路径   考试
                    {
                        //全部答案
                        string daan = "";
                        //提取答案块
                        Regex regex = new Regex(@"{""answer"".*}]'\s*/>", RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.CultureInvariant);
                        if (regex.IsMatch(htmlStr))
                        {
                            MatchCollection matchCollection = regex.Matches(htmlStr);
                            daan = matchCollection[0].ToString().Replace("/>", "").Trim().Replace("]'", "");
                        }
                        //把答案分组
                        string strTmp = "},{";
                        string[] lstDaAn = Regex.Split(daan, strTmp, RegexOptions.IgnoreCase);
                        for (int i = 0; i < lstDaAn.Length; i++)
                        {
                            if (lstDaAn[i].Contains("{"))
                            {
                                lstDaAn[i] = lstDaAn[i].Replace("{", "");
                            }
                            //第一条会有个{   最后一条最后有个}
                            if (lstDaAn[i].Contains("}"))
                            {
                                lstDaAn[i] = lstDaAn[i].Replace("}", "");
                            }
                            //"type":"single"   单选  "realAnswer":"D"  
                            if (lstDaAn[i].Contains("\"type\":\"single\""))
                            {
                                //正确答案选项
                                string xuanxiang = lstDaAn[i].Substring(lstDaAn[i].IndexOf("\"realAnswer\":\"") + "\"realAnswer\":\"".Length, 1);
                                //题目
                                string content = lstDaAn[i].Substring(lstDaAn[i].IndexOf("\"content\":\"") + "\"content\":\"".Length, lstDaAn[i].IndexOf("\",\"filePath\"") - lstDaAn[i].IndexOf("\"content\":\"") - "\"content\":\"".Length);
                                //自动答题

                                //取得每个Tab之后，就可以通过InternetExplorer的Document取得每个页面的Dom
                                HTMLDocument doc = (HTMLDocument)iw.Document;
                                //通过DOM操作IE页面
                                //mshtml.IHTMLElementCollection inputs = (mshtml.IHTMLElementCollection)doc2.all.tags("INPUT");
                                //mshtml.HTMLInputElement input1 = (mshtml.HTMLInputElement)inputs.item("kw1", 0);
                                //input1.value = "test";
                                //mshtml.IHTMLElement element2 = (mshtml.IHTMLElement)inputs.item("su1", 0);
                                //element2.click();

                                mshtml.IHTMLElementCollection inputs = (mshtml.IHTMLElementCollection)doc.all.tags("li");
                                IEnumerable<IHTMLElement> EnHEColl = inputs.Cast<IHTMLElement>();
                                IHTMLElement he1 = EnHEColl.FirstOrDefault(p => p.innerHTML != null && p.innerHTML.Contains(content));
                                MessageBox.Show(he1.innerText);
                                //mshtml.HTMLInputElement input1 = (mshtml.HTMLInputElement)inputs.item("kw1", 0);

                            }
                            //"type":"multiple" 多选
                            else if (lstDaAn[i].Contains("\"type\":\"multiple\""))
                            {
                                //正确答案选项
                                string xuanxiang = lstDaAn[i].Substring(lstDaAn[i].IndexOf("\"realAnswer\":\"") + "\"realAnswer\":\"".Length, lstDaAn[i].IndexOf("\",\"realScore\"") - lstDaAn[i].IndexOf("\"realAnswer\":\"") - "\"realAnswer\":\"".Length);
                                //题目
                                string content = lstDaAn[i].Substring(lstDaAn[i].IndexOf("\"content\":\"") + "\"content\":\"".Length, lstDaAn[i].IndexOf("\",\"filePath\"") - lstDaAn[i].IndexOf("\"content\":\"") - "\"content\":\"".Length);
                                //自动答题
                            }
                            //"type":"judge"判断
                            else if (lstDaAn[i].Contains("\"type\":\"judge\""))
                            {
                                //正确答案选项
                                string xuanxiang = lstDaAn[i].Substring(lstDaAn[i].IndexOf("\"realAnswer\":\"") + "\"realAnswer\":\"".Length, 1);
                                //题目
                                string content = lstDaAn[i].Substring(lstDaAn[i].IndexOf("\"content\":\"") + "\"content\":\"".Length, lstDaAn[i].IndexOf("\",\"filePath\"") - lstDaAn[i].IndexOf("\"content\":\"") - "\"content\":\"".Length);
                                //自动答题
                            }
                        }
                    }
                }
            }
        }
    }
}
