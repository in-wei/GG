using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//--
using System.IO;
using UsuallyCommond.MyEnum;
using UsuallyCommond;
using GlobalCommond;
using GlobalCommond.ViewModel;
using Excel = Microsoft.Office.Interop.Excel;

namespace QR_Code_FindOneDayTicket.UI
{
    public partial class Form1_QR : Form
    {
        private MyCommonds MyCommond = new MyCommonds(true, false, true);
        private string ThisReceive;
        private QR_Code_FindOneDayTicket MainServer = new QR_Code_FindOneDayTicket();

        private string FirstGetFormName = "";
        public Form1_QR()
        {
            InitializeComponent();
            ThisReceive = "Form_" + this.Text;
            MyCommond.WriteLog(ThisReceive, "==程式啟動==");
            string ProgramFullPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName.ToString();
            string ProgramLastWriteTime = File.GetLastWriteTime(ProgramFullPath).ToString();
            MyCommond.WriteLog(ThisReceive, $"This Program Last Write Time： {ProgramLastWriteTime}");
            if ( true) MyCommond.CheckApplicationAndKill("EXCEL");
            RegisterEvents();

        }

        // Form UI掛聽
        private void RegisterEvents()
        {
            //Form
            this.Load += Form_Load;
            this.MinimumSizeChanged += Form_MinimumSizeChanged;
            this.FormClosing += Form_Close;
            this.SizeChanged += Form_SizeChanged;

            //UI
            dateTimePicker1.ValueChanged += DateTimePicker1_ValueChanged;
            dateTimePicker2.ValueChanged += DateTimePicker2_ValueChanged;

            textBox1.TextChanged += TextBox1_TextChanged;
            textBox2.DoubleClick += TextBox2_DoubleClick;

            label3.DoubleClick += Label3_DoubleClick;

            button1.Click += Button1_Click;
            button2.Click += Button2_Click;
            button3.Click += Button3_Click;

            // Menu
            this.MenuItem_Close.Click += MenuItem_Close_Click;

        }

        #region Form Event

        public void Form_Load(object sender, EventArgs e)
        {
            DateTime dateTime = DateTime.Now;
            dateTimePicker1.Value = Convert.ToDateTime(dateTime.AddDays(-1).ToString("yyyy/MM/dd 05:00:00"));
            dateTimePicker2.Value = Convert.ToDateTime(dateTime.AddDays( 0).ToString("yyyy/MM/dd 05:00:00"));
            this.listView1.Items.Clear();
            //button1.PerformClick();
            Timer_Set();
            SetTipUI();
        }

        private void Form_MinimumSizeChanged(object sender, EventArgs e)
        {

        }

        public void Form_Close(object sender, FormClosingEventArgs e)
        {
            if (!true)
            {
                this.ShowInTaskbar = false;
                this.WindowState = FormWindowState.Minimized;
                this.Hide();
                e.Cancel = true;
            }
            else
            {
                GC.Collect();
                MyCommond.WriteLog(ThisReceive, "程式關閉");
                System.Environment.Exit(0);
            }
        }

        private void Form_SizeChanged(object sender, EventArgs e)
        {
            this.listView1.Width = this.Width - listView1.Location.X - 30;
            this.listView1.Height = this.Height - listView1.Location.Y - 50;
            //label3.Width = this.Width - label3.Location.X - 30;
            //progressBar1.Width = this.Width - progressBar1.Location.X - 30;
            liseView1_Set();
        }

        private void liseView1_Set()
        {
            //int LW = listView1.Width - 10;
            //listView1.Columns[0].Width = 70;
            //listView1.Columns[1].Width = 70;
            //listView1.Columns[2].Width = 70;
            //listView1.Columns[3].Width = 50;
            //listView1.Columns[3].Width = 100;
            //for (int i = 0; i < listView1.Columns.Count - 1; i++) LW -= listView1.Columns[i].Width;
            //if (LW > 100) { listView1.Columns[4].Width = LW; } else { listView1.Columns[4].Width = 100; }
        }

        #endregion

        #region Menu

        private void MenuItem_Close_Click(object sender, EventArgs e)
        {
            GC.Collect();
            MyCommond.WriteLog(ThisReceive, "程式關閉");
            System.Environment.Exit(0);
        }

        #endregion

        #region Tip

        private void SetTipUI()
        {
            toolTip1.SetToolTip(textBox2, "點兩下進入選檔案");

        }

        #endregion

        #region Timer

        Timer timer = new Timer();

        private void Timer_Set()
        {
            timer = new Timer();
            timer.Interval = 1000;
            timer.Tick += Timer_ListView_Tick;
        }

        bool isloop = true;
        private void Timer_ListView_Tick(object sender, EventArgs e)
        {
            MyCommond.InvokeIfRequired(this, () =>
            {
                if (FirstGetFormName.Length == 0) { FirstGetFormName = this.Text; }

                this.Text = $"{FirstGetFormName}_{MainServer.RunStep}";

                if (MainServer.IsRun)
                {
                    if (MainServer.qrViewList.Count <= 0) { goto NextTime; }
                    else { isloop = true; }
                }
                else { isloop = false; }

                this.listView1.Items.Clear();

                for (int i = 0; i < MainServer.qrViewList.Count; i++)
                {
                    string[] str = new string[listView1.Columns.Count];
                    str[0] = MainServer.qrViewList[i].OperationDate;
                    str[1] = MainServer.qrViewList[i].OperationLine;
                    str[2] = MainServer.qrViewList[i].Cooperate;
                    str[3] = MainServer.qrViewList[i].EntryCount.ToString();
                    str[4] = MainServer.qrViewList[i].ExitCount.ToString();
                    str[5] = MainServer.qrViewList[i].TicketCount.ToString();
                    ListViewItem Lwi = new ListViewItem(str);
                    this.listView1.Items.Add(Lwi);
                }
            NextTime:
                if (!isloop) { timer.Stop(); ffsfsf(true); }
            });
        }

        #endregion

        #region UI Event - Normal
        bool startloop = false;
        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (!startloop)
            {
                startloop = true;
                dateTimePicker1.Value = Convert.ToDateTime(dateTimePicker1.Value.ToString("yyyy/MM/dd 05:00:00"));
                MyCommond.WriteLog(ThisReceive, $"營運日更動:{dateTimePicker1.Value.ToString("yyyy/MM/dd")}");
                dateTimePicker2.Value = dateTimePicker1.Value.AddDays(1);
                MainServer.OperationDate_Start = dateTimePicker1.Value;
                MainServer.OperationDate_End = dateTimePicker2.Value;
                //MainServer.OperationStartDate = Convert.ToDateTime(dateTimePicker1.Value.ToString("yyyy/MM/dd 05:00:00"));
                startloop = false;
            }
            else
            {

            }
        }

        private void DateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

            MyCommond.WriteLog(ThisReceive, $"清分日更動:{dateTimePicker2.Value.ToString("yyyy/MM/dd")}");
            textBox1.Text = $"MP_{dateTimePicker2.Value.ToString("yyyyMMdd")}.csv";
            MainServer.OperationDate_End = dateTimePicker2.Value;
            
            //MainServer.OperationEndDate = Convert.ToDateTime(dateTimePicker2.Value.ToString("yyyy/MM/dd 05:00:00"));
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            MainServer.TxnDataFileName = textBox1.Text;
        }

        private void Label3_DoubleClick(object sender, EventArgs e) => Clipboard.SetData(DataFormats.Text, MainServer.password);

        private void TextBox2_DoubleClick(object sender, EventArgs e)
        {
            string aa = MyCommond.SelectFileForm(new FileLastName()
            {
                xlsx = true
            });

            if (aa != "")
            {
                MyCommond.InvokeIfRequired(this, () =>
                {
                    textBox2.Text = aa;
                    toolTip1.SetToolTip(textBox2, aa);
                });
            }
        }

        #endregion

        #region UI Event - Button

        private void Button1_Click(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"Button1 Click!");
            ffsfsf(false);
            if (true)
            {
                if (true)
                {
                    timer.Start();
                    //System.Threading.Thread thread = new System.Threading.Thread(() =>
                    //{
                    //    timer.Start();
                    //});
                    //thread.IsBackground = true;
                    //thread.Start();
                }
                if (true)
                {
                    System.Threading.Thread thread = new System.Threading.Thread(() =>
                    {
                        MainServer.Start();
                    });
                    thread.IsBackground = true;
                    thread.Start();
                }
            }
            else
            {
                MyCommond.InvokeIfRequired(this, () => { MainServer.Start(); });
            }
        }

        private void Button2_Click(object sender, EventArgs e) => Clipboard.SetData(DataFormats.Text, this.textBox1.Text.Split('.')[0]);

        private void Button3_Click(object sender, EventArgs e)
        {

            string NewCompare_File = textBox2.Text;
            if (NewCompare_File == "點我兩下選檔案") return;
            string Temp_FileName_Full = MyCommond.Path_Program + MyCommond.Path_Template + CheckCooperate(NewCompare_File);

            if (true)
            {
                string Cooperate_Title = GetExcelTitle(NewCompare_File);    // A - 新檔案
                string Template_Title = GetExcelTitle(Temp_FileName_Full);  // B - 累計檔

                string[] A = Cooperate_Title.Split(',');
                string[] B = Template_Title.Split(',');

                int[] a = new int[A.Length];

                for (int i = 0; i < a.GetLength(0); i++)
                {
                    a[i] = -1;
                }

                int A_C = A.Length;
                int B_C = B.Length;

                int ab = 0;
                int ac = 0;
                bool fa = false;
                foreach (var item in A)
                {
                    ac = 0;
                    foreach (var jtem in B)
                    {
                        fa = CompareString(item, jtem);
                        if (fa) 
                        {
                            bool DataNull = true;
                            for (int i = 0; i < a.GetLength(0); i++)
                            {
                                if (a[i] == ac) { DataNull = false; }
                            }
                            if (DataNull) { break; }
                        }
                        ac++;
                    }
                    a[ab++] = ((fa) ? ac : -1);
                }



            }
        }

        #endregion

        private void ffsfsf(bool TF)
        {
            MyCommond.InvokeIfRequired(this, () =>
            {
                dateTimePicker1.Enabled = TF;
                dateTimePicker2.Enabled = TF;
                textBox1.Enabled = TF;
                button1.Enabled = TF;
                button2.Enabled = TF;
                groupBox1.Enabled = TF;
            });
        }

        /// <summary>
        /// 拆分新檔案的名稱並判斷是哪個合作單位。
        /// </summary>
        /// <param name="FullPath">完整的路徑。</param>
        /// <returns>回傳比對檔的檔名。</returns>
        private string CheckCooperate(string FullPath)
        {
            char[] q1 = { '/', '\\' };
            char[] q2 = { '.', '-', '_', '(', ')', ' ' };

            string[] t1 = { "KKDAY", "kkday", "KKDay", "kkDay", "KKday" };
            string[] t2 = { "新北幣", "NTPay", "NewTaiPay" };

            var ss_0 = FullPath.Split(q1);
            var ss_1 = ss_0[ss_0.Length - 1].Split('.');
            string ss_2 = ss_1[0];

            var sa_0 = ss_2.Split(q2);

            bool Is_KKDay = false;
            bool Is_NTPay = false;

            foreach (var s in sa_0)
            {
                if (t1.Contains(s)) { return "Kkday_Temp.xlsx"; }
                if (t2.Contains(s)) { return "NewTaiPeiCoin_Temp.xlsx"; }
            }
            return "";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="FullPath">完整的路徑。</param>
        /// <returns>回傳檔案的欄位名稱。</returns>
        private string GetExcelTitle(string FullPath)
        {
            string Excel_titleGather = "";

            Excel.Application Excel_App = new Excel.Application();
            Excel_App.Visible = true;
            try
            {
                Excel.Workbook Excel_WorkBook = Excel_App.Workbooks.Open(Filename: FullPath, ReadOnly: false, Password: "69278085");
                Excel.Worksheet Excel_WorkSheet = (Excel.Worksheet)Excel_WorkBook.Worksheets[1];
                Excel.Range Excel_Range = Excel_WorkSheet.UsedRange;

                string CellValue = "First";
                int ColumnOffset = 1;
                int RowOffset = 1;
                int ExitLoop = 0;

                while (CellValue != "")
                {
                    CellValue = Convert.ToString((Excel_Range[1, ColumnOffset] as Excel.Range).Value);
                    if (CellValue != null) { Excel_titleGather += CellValue + ","; }
                    else { ExitLoop++; }
                    if (ExitLoop > 10) break;
                    ColumnOffset++;
                }

                Excel_titleGather = Excel_titleGather.Substring(0, Excel_titleGather.Length - 1);
                Excel_WorkBook.Close();
            }
            catch (Exception ex)
            {

            }
            finally
            {
                Excel_App.Quit();
            }
            return Excel_titleGather;
        }

        /// <summary>
        /// 比對兩字串相似度，至少需要連續2字元相同。
        /// </summary>
        /// <param name="String_A">新檔案的欄位。</param>
        /// <param name="String_B">程式的欄位。</param>
        /// <returns>回傳True 或 False。</returns>
        private bool CompareString(string String_A, string String_B)
        {
            int A_Length = String_A.Length;
            int B_Length = String_B.Length;
            int Min_Length = 1;

            for (int A_I = 0; A_I < String_A.Length - Min_Length; A_I++)
            {
                for (int A_J = 0; A_J < String_A.Length - Min_Length; A_J++)
                {
                    string s_a = "";
                    try { s_a = String_A.Substring(A_I, A_Length - A_J); }
                    catch (Exception ex) { continue; }
                    for (int B_I = 0; B_I < String_B.Length - Min_Length; B_I++)
                    {
                        for (int B_J = 0; B_J < String_B.Length - Min_Length; B_J++)
                        {
                            string s_b = "";
                            try { s_b = String_B.Substring(B_I, B_Length - B_J); }
                            catch (Exception ex) { continue; }
                            if (s_a == s_b) { return true; }
                        }
                    }
                }
            }
            return false;
        }

    }
}
