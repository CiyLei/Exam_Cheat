using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using System.Text.RegularExpressions;
using System.Threading;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        string contr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
        OleDbCommand cmd;
        OleDbConnection con;
        OleDbDataReader reader;
        List<String> drawfileList = new List<string>();      //存放是第几套题库 0选择题  1文字题  2Word题  3Excel题  4Ppt题  5Windows题  6Internet题
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 控件透明化 
        /// </summary>
        /// <param name="image">背景图片</param>
        /// <param name="alpha">透明值</param>
        private Image PictureAlpha(Image image, int alpha)
        {
            //Image image = Image.FromFile(Path);
            Bitmap img = new Bitmap(image);
            Bitmap bmp = new Bitmap(img.Width, img.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            Graphics g = Graphics.FromImage(image);
            g.DrawImage(img, 0, 0);
            for (int h = 0; h <= img.Height - 1; h++)
            {
                for (int w = 0; w <= img.Width - 1; w++)
                {
                    Color c = img.GetPixel(w, h);
                    bmp.SetPixel(w, h, Color.FromArgb(alpha, c.R, c.G, c.B));
                }
            }
            return (Image)bmp.Clone();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.BackColor = Color.FromArgb(222, 222, 222);
            label3.BackColor = Color.Transparent;
            label11.BackColor = Color.Transparent;
            label13.BackColor = Color.Transparent;
            button1.FlatStyle = FlatStyle.Popup;    //功能是可以让按钮透明
            button1.BackColor = Color.Transparent;
            button1.Image = PictureAlpha(this.BackgroundImage, 0);
            button3.FlatStyle = FlatStyle.Popup;    //功能是可以让按钮透明
            button3.BackColor = Color.Transparent;
            button3.Image = PictureAlpha(this.BackgroundImage, 0);
            button4.FlatStyle = FlatStyle.Popup;    //功能是可以让按钮透明
            button4.BackColor = Color.Transparent;
            button4.Image = PictureAlpha(this.BackgroundImage, 0);
            button5.FlatStyle = FlatStyle.Popup;    //功能是可以让按钮透明
            button5.BackColor = Color.Transparent;
            button5.Image = PictureAlpha(this.BackgroundImage, 0);
            panel1.BackColor = Color.Transparent;
            panel2.BackColor = Color.Transparent;
            panel3.BackColor = Color.Transparent;
            panel4.BackColor = Color.Transparent;
            getclfilename();
            Control.CheckForIllegalCrossThreadCalls = false;
            MessageBox.Show("记住要先登陆考试系统哦∩ω∩");
        }
        private string getclfilename() //cl为缓存目录(给我操作的目录)
        {
            if (!Directory.Exists(@".\cl"))
            {
                Directory.CreateDirectory(@".\cl");
            }
            return System.Environment.CurrentDirectory + @"\cl\";
        }
        private void OpenNkf(Object o)
        {
            OpenFileDialog ofd=(OpenFileDialog)o;
            //if (label3.Text.Equals("Null"))
                //ofd.FileName = "定死的路径";
            string filename = ofd.FileName;
            label3.Text = filename;
            //先清空缓存目录
            Directory.Delete(getclfilename(), true);
            getclfilename();
            //把数据库文件移动到cl缓存目录下
            File.Copy(filename, getclfilename() + ofd.SafeFileName);
            //得到新的nkf后缀名文件的路径
            filename = getclfilename() + ofd.SafeFileName;
            //改后缀名为accdb
            string dfileName = Path.ChangeExtension(filename, ".accdb");
            File.Move(filename, dfileName);
            //得到新的accdb数据库文件的路径
            dfileName = getclfilename() + ofd.SafeFileName.Replace("nkf", "accdb");
            //连接数据库
            contr += dfileName + "; Jet OLEDB:Database Password=ZHL-JHQ-YJX-YGZ-XJ";
            con = new OleDbConnection(contr);
            con.Open();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "nkf文件 (*.nkf)|*.nkf";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    ParameterizedThreadStart pts = new ParameterizedThreadStart(OpenNkf);  //因为备份数据库要耗时  所以放到线程里
                    Thread t = new Thread(pts);
                    object o = ofd;
                    t.Start(o);
                }
            }
            catch { MessageBox.Show("意外的出错了,请重试！＞０＜"); }
        }
        /// <summary>
        /// 从数据库里读出zip并解压出来的耗时操作
        /// </summary>
        private void AccdbToUnZip_t()
        {
            label2.Text = "正在操作";
            panel1.BackColor = Color.LightSkyBlue;
            AccdbToUnZip("TheoryTopicTypeTable01", "QuestionTextOLE_Field", drawfileList[0], "选择题问题");  //解压选择题并记录 //解压成品Work
            panel1.BackColor = Color.MediumSpringGreen;
            label2.Text = "操作完成";
            label8.Text = "正在操作";
            panel2.BackColor = Color.LightSkyBlue;
            string workPath = AccdbToUnZip("QuestionFileOptAndDocumentEdtTableDoc", "QuestionKeyObjectZipFileOLE_Field", drawfileList[2], "成品Work");
            File.Copy(workPath, label11.Text + "\\Work.docx");
            panel2.BackColor = Color.MediumSpringGreen;
            label8.Text = "操作完成";
            label9.Text = "正在操作";
            panel3.BackColor = Color.LightSkyBlue;
            string pptPath = AccdbToUnZip("QuestionFileOptAndDocumentEdtTablePpt", "QuestionKeyObjectZipFileOLE_Field", drawfileList[4], "成品Ppt");
            File.Copy(pptPath, label11.Text + "\\Ppt.pptx");
            panel3.BackColor = Color.MediumSpringGreen;
            label9.Text = "操作完成";
            label10.Text = "正在操作";
            panel4.BackColor = Color.LightSkyBlue;
            string xlsPath = AccdbToUnZip("QuestionFileOptAndDocumentEdtTableXls", "QuestionKeyObjectZipFileOLE_Field", drawfileList[3], "成品Xls");
            File.Copy(xlsPath, label11.Text + "\\Xls.xlsx");
            panel4.BackColor = Color.MediumSpringGreen;
            label10.Text = "操作完成";
            reader.Close();
            con.Close();
            //Directory.Delete(getclfilename(),true);
            //删除cl文件夹  记得
            //删除cl文件夹  记得
            //删除cl文件夹  记得
        }
        /// <summary>
        /// 从数据库里读出zip并解压出来 返回路径 
        /// </summary>
        /// <param name="table_name">表名</param>
        /// <param name="value">表头名</param>
        /// <param name="id">获取第几个数据</param>
        /// <param name="fileName">解压出来的文件名</param>
        /// <param name="isXZT">是否是选择题</param>
        private string AccdbToUnZip(string table_name, string value ,string id, string fileName)
        {
            string UnZipPath = "";
            string sql = "select * from " + table_name + " where TopicID_Field = '" + id + "'"; //查询答案表
            cmd = new OleDbCommand(sql, con);
            reader = cmd.ExecuteReader();
            int i = 1;
            //StreamWriter sw = new StreamWriter(label11.Text + "\\选择题.txt", false, Encoding.Default); ;
            while (reader.Read())
            {
                byte[] bytes1 = (byte[])reader[value];   //查询答案数据(16进制流)
                Stream stream = new MemoryStream(bytes1);
                string ZipFilePath = getclfilename() + fileName + i.ToString() + ".zip";
                StreamToFile(stream, ZipFilePath);
                UnZipPath = UnZip(ZipFilePath, getclfilename(), "caIeduCnz07B", true, i.ToString());
                if (table_name.Equals("TheoryTopicTypeTable01"))    //如果是在读取选择题表的话,记录下答案
                {
                    string question = GetPathToTxt(UnZipPath);       //获取题目内容
                    File.AppendAllText(label11.Text + "\\选择题.txt", "第" + reader["QuestionID_Field"] + "题的答案是:" + reader["QuestionCriteria_Field"].ToString().Replace("|||", "") + "\r\n");
                    //File.AppendAllText(label11.Text + "\\选择题.txt", question.Replace("\n","\r\n") + "\r\n" + "答案是:" + reader["QuestionCriteria_Field"].ToString().Replace("|||", "") + "\r\n");
                }
                i++;
            }
            reader.Close();
            return UnZipPath;
        }
        /// <summary>
        /// 读取RTB文件
        /// </summary>
        /// <param name="Path">文件路径</param>
        private string GetPathToTxt(string Path)
        {
            StreamReader sr = new StreamReader(Path, Encoding.Default);
            string txt = sr.ReadToEnd();
            sr.Close();
            RichTextBox rtb = new RichTextBox();    //RTF格式文件不能直接读取  RichTextBox可以读取  所以先用RichTextBox读取再取出
            rtb.Rtf = txt;
            return rtb.Text;
        }
        /// <summary>
        /// 把流写入文件
        /// </summary>
        /// <param name="stream">流</param>
        /// <param name="fileName">文件路径</param>
        public void StreamToFile(Stream stream, string fileName)
        {
            // 把 Stream 转换成 byte[]
            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);
            // 设置当前流的位置为流的开始
            stream.Seek(0, SeekOrigin.Begin);
            // 把 byte[] 写入文件
            if (!File.Exists(fileName))
                File.Create(fileName).Close();
            FileStream fs = new FileStream(fileName, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(bytes);
            bw.Close();
            fs.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
        }
        /// <summary>
        /// 解压缩一个 zip 文件。
        /// </summary>
        /// <param name="zipedFile">The ziped file.</param>
        /// <param name="strDirectory">The STR directory.</param>
        /// <param name="password">zip 文件的密码。</param>
        /// <param name="overWrite">是否覆盖已存在的文件。</param>
        /// <param name="i">因为压缩包里面的文件名都一样,解压会重名,所以添一个标识。</param>
        public string UnZip(string zipedFile, string strDirectory, string password, bool overWrite,string i)
        {
            string directoryName = "";
            string fileName = "";
            if (strDirectory == "")
                strDirectory = Directory.GetCurrentDirectory();
            if (!strDirectory.EndsWith("\\"))
                strDirectory = strDirectory + "\\";
            using (ZipInputStream s = new ZipInputStream(File.OpenRead(zipedFile)))
            {
                s.Password = password;
                ZipEntry theEntry;
                bool Str_Key = false;       //存放文件名是否带Key  如果带  直接就跳过其他
                while ((theEntry = s.GetNextEntry()) != null && !Str_Key)
                {
                    directoryName = "";
                    string pathToZip = "";
                    pathToZip = theEntry.Name;
                    if (pathToZip != "")
                        directoryName = Path.GetDirectoryName(pathToZip) + "\\";
                    fileName = Path.GetFileName(pathToZip);
                    Directory.CreateDirectory(strDirectory + directoryName);
                    if (fileName != "")
                    {
                        if ((File.Exists(strDirectory + directoryName + fileName) && overWrite) || (!File.Exists(strDirectory + directoryName + fileName)))
                        {
                            using (FileStream streamWriter = File.Create(strDirectory + directoryName + i + fileName))
                            {
                                int size = 2048;
                                byte[] data = new byte[2048];
                                while (true)
                                {
                                    size = s.Read(data, 0, data.Length);
                                    if (size > 0)
                                        streamWriter.Write(data, 0, size);
                                    else
                                        break;
                                }
                                streamWriter.Close();
                            }
                        }
                    }
                    if (fileName.IndexOf("Key") != -1)
                        Str_Key = true;
                }
                s.Close();
            }
            return strDirectory + directoryName + i + fileName;
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            //注册热键F1，Id号为100。HotKey.KeyModifiers.None表示没添加任何辅助键
            HotKey.RegisterHotKey(Handle, 100, HotKey.KeyModifiers.None, Keys.F1);
            HotKey.RegisterHotKey(Handle, 101, HotKey.KeyModifiers.None, Keys.F2);
        }

        private void Form1_Leave(object sender, EventArgs e)
        {
            //注销Id号为100的热键设定
            HotKey.UnregisterHotKey(Handle, 100);
            //注销Id号为101的热键设定
            HotKey.UnregisterHotKey(Handle, 101);
        }
        protected override void WndProc(ref Message m)
        {
            const int WM_HOTKEY = 0x0312;//按快捷键 
            switch (m.Msg)
            {
                case WM_HOTKEY:
                    switch (m.WParam.ToInt32())
                    {
                        case 100:    //按下的是F1
                            this.Hide();
                            break;
                        case 101:    //按下的是F1
                            this.Show();
                            break;
                    }
                    break;
            }
            base.WndProc(ref m);
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (label11.Text == "Null" || label13.Text == "Null")
            {
                MessageBox.Show("请选择答案保存目录和打开有Ktu文件夹所在的目录和打开nkf文件＞︿＜");
                return;
            }
            Thread t = new Thread(new ThreadStart(AccdbToUnZip_t));
            t.Start();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
                label11.Text = fbd.SelectedPath;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "不是选Tku文件夹哦！是选Tku文件夹所在的目录⊙ω⊙";
            try
            {
                string drawfilePanth = "";               //存放drawfilePanth的路径
                string drawfileTxt = "";               //存放drawfilePanth的内容
                if (fbd.ShowDialog() == DialogResult.OK)      //读取drawfilePanth的内容
                {
                    //if (label13.Text.Equals("Null"))
                        //fbd.SelectedPath = "定死的目录";
                    label13.Text = fbd.SelectedPath;
                    DirectoryInfo[] dif = new DirectoryInfo(label13.Text + "\\MyBackup\\").GetDirectories();   //获取MyBackup目录下的子文件
                    drawfilePanth = label13.Text + "\\MyBackup\\" + dif[0].ToString() + "\\";
                    dif = new DirectoryInfo(drawfilePanth).GetDirectories();   //获取学号目录下的子文件
                    drawfilePanth += dif[0].ToString() + "\\drawfile.dat";
                    StreamReader sr = new StreamReader(drawfilePanth, Encoding.Default);
                    drawfileTxt = sr.ReadToEnd();                                         //读取drawfilePanth的内容
                    Regex reg = new Regex("(?<=TopicID\\w{2}=)\\w{2}");     //读取抽取第几套题库
                    MatchCollection mc = reg.Matches(drawfileTxt);
                    foreach (Match m in mc)
                        drawfileList.Add(m.Value);
                }
            }
            catch { MessageBox.Show("意外的出错了！看看是不是学号填错了！还有,注意不是选Tku文件夹哦！是选Tku文件夹所在的目录哦！！！⊙ω⊙"); label13.Text = ""; }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            string xlsPath = AccdbToUnZip(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
        }
    }
}
