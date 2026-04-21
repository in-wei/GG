using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//-
//using Report.Mrt.ViewModel;
//using Report.Mrt;
using System.IO;
using UsuallyCommond.MyEnum;
using UsuallyCommond;
using GlobalCommond;
using GlobalCommond.ViewModel;

namespace Report.Mrt.UI
{
    public partial class Form_Mrt : Form
    {
        private MrtMain MainServer;
        private MyCommonds MyCommond;
        private string ThisReceive;
        List<ReportList> reportLists;

        public Form_Mrt()
        {
            //new testCatch_2().GetComputerInfo();
            InitializeComponent();
            ThisReceive = "Form_" + this.Text;
            MainServer = new MrtMain();
            MainServer.globalCommond = new GlobalCommonds(true, true, true);
            //MainServer.globalCommond = new GlobalCommonds("Mrt", true, true, true);
            MainServer.SetMyCommand();
            //MyCommond = MainServer.globalCommond.MyCommond;
            MyCommond = new MyCommonds();
            MyCommond.WriteLog(ThisReceive, "==程式啟動==");
            string ProgramFullPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName.ToString();
            string ProgramLastWriteTime = File.GetLastWriteTime(ProgramFullPath).ToString();
            MyCommond.WriteLog(ThisReceive, $"This Program Last Write Time： {ProgramLastWriteTime}");

            if (MainServer.globalCommond.UsingSystem.Equals("mrt", StringComparison.CurrentCultureIgnoreCase) && MainServer.globalCommond.OperationCode_String != null)
            {
                Form_UI_Set();
                liseView1_Set();
                RegisterEvents();
                Timer_Set();
                Other_Init();
            }
            else
            {
                MessageBox.Show($@"設定檔中未發現捷運設定，請檢查下列檔案{MyCommond.Path_Program}{MyCommond.Path_Config}\{MyUseFile.Config_Setting}"
                        , "錯誤"
                        , MessageBoxButtons.OK
                        , MessageBoxIcon.Question);
                GC.Collect();
                MyCommond.WriteLog(ThisReceive, "程式關閉");
                System.Environment.Exit(0);
            }
        }

        public DateTime setDateTime { get; private set; }

        public Form_Mrt(DateTime mDt)
        {
            setDateTime = mDt;
            
            
        }

        public void Other_Init()
        {

            // 掃描Template資料夾
            if (!MyCommond.ScanFolderAndExport($@"{MyCommond.Path_Program}{MyCommond.Path_Template}", 2))
            {
                System.Threading.Thread thread = new System.Threading.Thread(() =>
                {
                    MyCommond.MsgShow($@"發現路徑：{MyCommond.Path_Program}{MyCommond.Path_Template} 中檔案有缺失，不影響主程式執行，但部分報表將無法產製。"
                        , "提示"
                        , MessageBoxButtons.OK
                        , MessageBoxIcon.Question);
                });
                thread.IsBackground = true;
                thread.Start();
            }

        }

        public void Form_UI_Set()
        {
            dateTimePicker1.Value = Convert.ToDateTime(DateTime.Now.AddDays(-1).ToString("yyyy/MM/dd 05:00:00"));
            dateTimePicker2.Value = Convert.ToDateTime(DateTime.Now.AddDays(0).ToString("yyyy/MM/dd 05:00:00"));
            button2.Enabled = false;
        }

        // Form UI掛聽
        private void RegisterEvents()
        {
            //Form
            this.Load += Form_Load;
            this.MinimumSizeChanged += Form_MinimumSizeChanged;
            this.FormClosing += Form_Close;
            this.SizeChanged += Form_SizeChanged;

            // Menu
            this.MenuItem_Close.Click += MenuItem_Close_Click;

            //UI
            dateTimePicker1.ValueChanged += DateTimePicker1_ValueChanged;
            dateTimePicker2.ValueChanged += DateTimePicker2_ValueChanged;
            checkBox1.CheckedChanged  += CheckBox1_CheckedChanged;
            checkBox2.CheckedChanged  += CheckBox2_CheckedChanged;
            checkBox3.CheckedChanged  += CheckBox3_CheckedChanged;
            checkBox4.CheckedChanged  += CheckBox4_CheckedChanged;
            checkBox5.CheckedChanged  += CheckBox5_CheckedChanged;
            checkBox6.CheckedChanged  += CheckBox6_CheckedChanged;
            checkBox7.CheckedChanged  += CheckBox7_CheckedChanged;
            checkBox8.CheckedChanged  += CheckBox8_CheckedChanged;
            checkBox9.CheckedChanged  += CheckBox9_CheckedChanged;
            checkBox10.CheckedChanged += CheckBox10_CheckedChanged;
            //checkBox11.CheckedChanged += CheckBox11_CheckedChanged;
            //checkBox12.CheckedChanged += CheckBox12_CheckedChanged;
            //checkBox13.CheckedChanged += CheckBox13_CheckedChanged;
            //checkBox14.CheckedChanged += CheckBox14_CheckedChanged;
            //checkBox15.CheckedChanged += CheckBox15_CheckedChanged;
            //checkBox16.CheckedChanged += CheckBox16_CheckedChanged;
            //checkBox17.CheckedChanged += CheckBox17_CheckedChanged;
            //checkBox18.CheckedChanged += CheckBox18_CheckedChanged;
            //checkBox19.CheckedChanged += CheckBox19_CheckedChanged;
            //checkBox20.CheckedChanged += CheckBox20_CheckedChanged;
            button1.Click += Button1_Click;
            button2.Click += Button2_Click;
            button3.Click += Button3_Click;
        }

        #region Timer

        Timer Timer_ListView;
        Timer Timer_ProgerssBar;

        private void Timer_Set()
        {
            Timer_ListView = new Timer();
            Timer_ListView.Interval = 1000;
            Timer_ListView.Tick += Timer_ListView_Tick;
            Timer_ListView.Start();

            Timer_ProgerssBar = new Timer();
            Timer_ProgerssBar.Interval = 1000;
            Timer_ProgerssBar.Tick += Timer_Progerss_Tick;
            //Timer_ProgerssBar.Start();
        }

        private void Timer_ListView_Tick(object sender, EventArgs e)
        {
            if (!Timer_ProgerssBar.Enabled) if (MainServer.CardCount_More > 0) Timer_ProgerssBar.Start();
            UpdateClientsList();
        }

        private void Timer_Progerss_Tick(object sender, EventArgs e)
        {
            UpdateProgress();
        }

        int bcc = 0;
        private void UpdateProgress()
        {
            MyCommond.InvokeIfRequired(this, () =>
            {
                if (true)
                {
                    var BarMax = MainServer.TxnReportData_Max;
                    var BarMin = MainServer.TxnReportData_Min;
                    var BarCount = MainServer.TxnReportData_Count;

                    var CC = BarMax == BarMin;

                    if (CC)
                    {
                        bcc = 0;
                        BarMax = 0;
                        BarMin = 0;
                        BarCount = 0;
                    }

                    if (MainServer.TxnReportData_Run && !CC)
                    {
                        if (bcc == 0) { bcc++; }

                        label3.Text = $"交易資料載入:{BarCount,8} / {BarMax,8}";
                        progressBar1.Maximum = BarMax;
                        progressBar1.Minimum = BarMin;
                        progressBar1.Value = BarCount;
                    }

                    if (bcc == 0)
                    {
                        label3.Text = $"程式執行中..";
                    }
                    else
                    {
                        if (!MainServer.TxnReportData_Run)
                        {
                            label3.Text = $"完成載入，開始分析";
                        }
                        else if (BarMax == BarCount)
                        {
                            label3.Text = $"目前正在:{MainServer.TxnReportData_Str}";
                        }
                    }
                }
                else
                {
                    DateTime dt = DateTime.Now;
                    var oftenTime = dt - MainServer.StartLoop;
                    if (((oftenTime.Hours * 60 + oftenTime.Minutes) * 60 + oftenTime.Seconds) <= 0) return;
                    double finishTime = oftenTime.TotalSeconds / (MainServer.CardCount_More + MainServer.CardCount_Less) * MainServer.TxnAllCard.Count();
                    var finishHour = Math.Floor(finishTime / 60 / 60);
                    var finishMinute = Math.Floor((finishTime - finishHour * 60 * 60) / 60);
                    var finishSecond = finishTime - (finishHour * 60 + finishMinute) * 60;
                    TimeSpan ff = new TimeSpan((int)finishHour, (int)finishMinute, (int)finishSecond);
                    TimeSpan ss = ff - oftenTime;
                    label3.Text = $"{(MainServer.CardCount_More + MainServer.CardCount_Less)}/{MainServer.TxnAllCard.Count()}(預計剩餘時間:{ss.ToString(@"d\.hh\:mm\:ss")})";
                    progressBar1.Maximum = MainServer.TxnAllCard.Count();
                    progressBar1.Value = (MainServer.CardCount_More + MainServer.CardCount_Less);
                }
            });


            /**
             
            TxnReportData_Max = 100;
            TxnReportData_Min = 0;
            TxnReportData_Count = 0;
             */
        }

        private void UpdateClientsList()
        {
            MyCommond.InvokeIfRequired(this, () =>
            {
                // 更新進度條
                if (MainServer.CardCount_More > 0)
                {

                }

                // 更新listView
                try
                {
                    this.Text = $"NTMC(MRT) ~ 執行營運日:{MainServer.OperationDate:yyyy-MM-dd} ~ 啟用邏輯線程數量:{MainServer.MrtClients.Count(),4}";

                    var Sr = MainServer.MrtClients;
                    var Lw = listView1.Items;

                    if (!MainServer.IsRunning)
                    {
                        //MyCommond.WriteLog(ThisReceive, $"計算完成");
                        MainServer.Stop();
                        dateTimePicker1.Enabled = true;
                        dateTimePicker2.Enabled = true;
                        button1.Enabled = true;
                        groupBox1.Enabled = true;
                        Sr.Clear();
                        MainServer = new MrtMain();
                        Timer_ListView.Stop();
                        Timer_ProgerssBar.Stop();
                        GC.Collect();
                    }

                    if (Sr.Count() <= 0)
                    {
                        Lw.Clear();
                        return;
                    }

                    foreach (var item in Sr)
                    {
                        bool haveItem = false;
                        bool IsAdd = false;
                        for (int i = 0; i < Lw.Count; i++)
                        {
                            if (Lw[i].SubItems[0].Text == item.ID.ToString())
                            {
                                haveItem = true;
                                Lw[i].SubItems[0].Text = item.ID.ToString();
                                Lw[i].SubItems[1].Text = item.Status.ToString();
                                Lw[i].SubItems[2].Text = item.MrtMethodeName.ToString();
                                Lw[i].SubItems[3].Text = $"{item.Counts.ToString(),6}/{item.CardList.Count}";
                                Lw[i].SubItems[4].Text = item.CardSN;
                                break;
                            }
                        }
                        if (!haveItem)
                        {
                            string[] str = new string[listView1.Columns.Count];

                            str[0] = item.ID.ToString();
                            str[1] = item.Status.ToString();
                            str[2] = item.MrtMethodeName.ToString();
                            str[3] = item.Counts.ToString();
                            str[4] = item.CardSN;

                            ListViewItem Lwi = new ListViewItem(str);
                            listView1.Items.Add(Lwi);
                            IsAdd = true;
                        }
                    }
                    System.Threading.Thread.Sleep(30);

                }
                catch (Exception ex)
                {

                }
            });
        }

        #endregion

        #region Form Event

        public void Form_Load(object sender, EventArgs e)
        {
            reportLists = new List<ReportList>();
            AutoCheck();
            SetTipUI();
            //button1.PerformClick();
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
                while (CheckFromIsLife)
                {
                    CheckFromClose = true;
                    System.Threading.Thread.Sleep(100);
                }
                GC.Collect();
                MyCommond.WriteLog(ThisReceive, "程式關閉");
                System.Environment.Exit(0);
            }
        }

        private void Form_SizeChanged(object sender, EventArgs e)
        {
            this.listView1.Width = this.Width - listView1.Location.X - 30;
            this.listView1.Height = this.Height - listView1.Location.Y - 50;
            label3.Width = this.Width - label3.Location.X - 30;
            progressBar1.Width = this.Width - progressBar1.Location.X - 30;
            liseView1_Set();
        }

        private void liseView1_Set()
        {
            int LW = listView1.Width - 10;
            listView1.Columns[0].Width = 70;
            listView1.Columns[1].Width = 70;
            listView1.Columns[2].Width = 70;
            listView1.Columns[3].Width = 50;
            //listView1.Columns[3].Width = 100;
            for (int i = 0; i < listView1.Columns.Count - 1; i++) LW -= listView1.Columns[i].Width;
            if (LW > 100) { listView1.Columns[4].Width = LW; } else { listView1.Columns[4].Width = 100; }
        }

        private void AutoCheck()
        {
            checkBox1.Checked = true;
            checkBox2.Checked = true;
            checkBox3.Checked = true;
            //checkBox4.Checked = true;
            checkBox5.Checked = true;
            checkBox6.Checked = true;
            checkBox7.Checked = true;
            checkBox8.Checked = true;
            checkBox9.Checked = true;
            //checkBox10.Checked = true;
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
            toolTip1.SetToolTip(checkBox7, "勾選或點擊後方按鈕進入詳細選單");
        }

        #endregion

        #region UI Event - Normal

        #region DateTimePicker
        private void DateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"清分開始日更動:{dateTimePicker1.Value.ToString("yyyy/MM/dd")}");
            MainServer.OperationStartDate = Convert.ToDateTime(dateTimePicker1.Value.ToString("yyyy/MM/dd 05:00:00"));
        }

        private void DateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"清分結束日更動:{dateTimePicker2.Value.ToString("yyyy/MM/dd")}");
            MainServer.OperationEndDate = Convert.ToDateTime(dateTimePicker2.Value.ToString("yyyy/MM/dd 05:00:00"));
        }
        #endregion

        #region CheckBox
        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.EzsyAnalyze.ToDescription()} - {((checkBox1.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.TxnTypeAnalyze.ToDescription()} - {((checkBox2.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox3_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.IOAnalyze.ToDescription()} - {((checkBox3.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox4_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.CompareAnalyze.ToDescription()} - {((checkBox4.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox5_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.TrnAnalyze.ToDescription()} - {((checkBox5.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox6_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.TPassAnalyze.ToDescription()} - {((checkBox6.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox7_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.VolumnReportAnalyze.ToDescription()} - {((checkBox7.Checked) ? "啟動" : "關閉")}");
            if (!CancalCheck)
            {
                if (checkBox7.Checked)
                {
                    C7Select();
                }
                else
                {
                    reportLists = new List<ReportList>();
                }
            }
            
        }

        private void CheckBox8_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.Day_AddValue.ToDescription()} - {((checkBox8.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox9_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.Day_EveryIssuer.ToDescription()} - {((checkBox9.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox10_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.C10Analyze.ToDescription()} - {((checkBox10.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox11_CheckedChanged(object sender, EventArgs e)
        {
            //MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.C11Analyze.ToDescription()} - {((checkBox11.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox12_CheckedChanged(object sender, EventArgs e)
        {
            //MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.C12Analyze.ToDescription()} - {((checkBox12.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox13_CheckedChanged(object sender, EventArgs e)
        {
            //MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.C13Analyze.ToDescription()} - {((checkBox13.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox14_CheckedChanged(object sender, EventArgs e)
        {
            //MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.C14Analyze.ToDescription()} - {((checkBox14.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox15_CheckedChanged(object sender, EventArgs e)
        {
            //MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.C15Analyze.ToDescription()} - {((checkBox15.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox16_CheckedChanged(object sender, EventArgs e)
        {
            //MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.C16Analyze.ToDescription()} - {((checkBox16.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox17_CheckedChanged(object sender, EventArgs e)
        {
            //MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.C17Analyze.ToDescription()} - {((checkBox17.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox18_CheckedChanged(object sender, EventArgs e)
        {
            //MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.C18Analyze.ToDescription()} - {((checkBox18.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox19_CheckedChanged(object sender, EventArgs e)
        {
            //MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.C19Analyze.ToDescription()} - {((checkBox19.Checked) ? "啟動" : "關閉")}");

        }

        private void CheckBox20_CheckedChanged(object sender, EventArgs e)
        {
            //MyCommond.WriteLog(ThisReceive, $"{MrtTxnClientSwitch.C20Analyze.ToDescription()} - {((checkBox20.Checked) ? "啟動" : "關閉")}");

        }

        #endregion

        #endregion

        #region UI Event - Button

        private void Button1_Click(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"Button1 Click!");
            dateTimePicker1.Enabled = false;
            dateTimePicker2.Enabled = false;
            button1.Enabled = false;
            groupBox1.Enabled = false;
            if (true)
            {
                MyCommond.InvokeIfRequired(this, () =>
                {
                    Timer_ProgerssBar.Start();
                    MainServer.LoadTxn(new CheckBoxStatus()
                    {
                        OperationStart /*------*/ = dateTimePicker1.Value,
                        OperationEnd /*--------*/ = dateTimePicker2.Value,
                        EasyAnalyze /*---------*/ = checkBox1.Checked,
                        TxnTypeAnalyze /*------*/ = checkBox2.Checked,
                        IOAnalyze /*-----------*/ = checkBox3.Checked,
                        CompareAnalyze /*------*/ = checkBox4.Checked,
                        TrnAnalyze /*----------*/ = checkBox5.Checked,
                        TPassAnalyze /*--------*/ = checkBox6.Checked,
                        VolumnReportAnalyze /*-*/ = checkBox7.Checked,
                        AddValueAnalyze /*-----*/ = checkBox8.Checked,
                        EveryIssuerAnalyze /*--*/ = checkBox9.Checked,
                        checkbox10 /*----------*/ = checkBox10.Checked,
                        checkbox11 /*----------*/ = false,
                        checkbox12 /*----------*/ = false,
                        checkbox13 /*----------*/ = false,
                        checkbox14 /*----------*/ = false,
                        checkbox15 /*----------*/ = false,
                        checkbox16 /*----------*/ = false,
                        checkbox17 /*----------*/ = false,
                        checkbox18 /*----------*/ = false,
                        checkbox19 /*----------*/ = false,
                        checkbox20 /*----------*/ = false,
                    }, reportLists);
                });
            }
            else
            {
                MyCommond.ExcelColumnStr(100);
                MyCommond.ExportViewTitle<CheckBoxStatus>();
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            MyCommond.InvokeIfRequired(this, () =>
            {
                //button1.Enabled = true;
                System.Threading.Thread thr = new System.Threading.Thread(() =>
                {
                    //MainMrt.Stop();
                });
                thr.IsBackground = true;
                thr.Start();
                //MainMrt.test();
            });
        }
        
        private void Button3_Click(object sender, EventArgs e)
        {
            C7Select();
        }
        
        #endregion

        bool CancalCheck = false;
        bool C7IsFirstCheck = true;
        bool CheckFromClose = false;
        bool CheckFromIsLife = true;
        private void C7Select()
        {
            CancalCheck = true;
            //this.checkBox7.Checked = true;
            Report.Mrt.Check.ReportCheck1 fff = new Report.Mrt.Check.ReportCheck1(C7IsFirstCheck, reportLists);
            ///if (!C7IsFirstCheck) fff.tempReportList = reportLists.ToList();
            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {

                while (fff.IsLife)
                {
                    if (CheckFromClose) { fff.IsLife = false; break; }
                    CheckFromIsLife = fff.IsLife;
                    System.Threading.Thread.Sleep(100);
                }
                CheckFromIsLife = false;
                reportLists = fff.tempReportList.ToList();
                if (reportLists.Count > 0)
                {
                    MyCommond.InvokeIfRequired(this, () => { this.checkBox7.Checked = true; });
                }
                else
                {
                    MyCommond.InvokeIfRequired(this, () => { this.checkBox7.Checked = false; });
                }
                C7IsFirstCheck = false;
                CancalCheck = false;
                //MyCommond.ConsoleExportArray("mrt_button3", MyCommond.TResultToArray<ReportList>(reportLists));
            });
            if (C7IsFirstCheck) { thread.Start(); }
            else 
            { 
                fff.Show();
                
                Point pp = new Point()
                {
                    X = this.Location.X + (this.Size.Width  - fff.Width ) / 2,
                    Y = this.Location.Y + (this.Size.Height - fff.Height) / 2,
                };
                fff.Location = pp; 
                thread.Start(); 
            }

            
        }
    }
}
