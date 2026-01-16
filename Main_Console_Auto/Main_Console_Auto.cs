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
using System.Threading;
using GlobalCommond.ViewModel;
using UsuallyCommond.MyEnum;
using GlobalCommond;
using System.IO;
using System.IO.Compression;

namespace Main_Console_Auto
{
    public partial class Form1_Main_Auto : Form
    {
        MyCommonds MyCommond;
        string ThisReceive = "MainForm_Auto";
        Report.Mrt.MrtMain MainServer;

        private int Time_1s = 1000;

        private string MouseMoveShowStatueString = "";
        string Path_MrtData_Backup = @"MrtData_Backup\";
        private TimeSpan next_ds = new TimeSpan(0, 0, 0);
        private DateTime ProgramStartDate;

        private int bcc = 0;
        private int tempMinute = -1;

        private bool AfterThread;
        private bool Detect_Start;
        private bool RunFinish;

        private bool Test_;

        public Form1_Main_Auto()
        {
            InitializeComponent();

            ProgramStartDate = DateTime.Now;

            Class_Init();       // 激活所有class
            Variable_Init();    // 初始化變數
            Other_Init();       // 其他執行
            RegisterEvents();   // 掛聽程式

            this.ShowInTaskbar = false;
            this.WindowState = FormWindowState.Minimized;
            this.Hide();
        }

        private void Class_Init()
        {
            if (MyCommond != null) MyCommond = null;
            if (MainServer != null) MainServer = null;


            MyCommond = new MyCommonds();

            MyCommond.WriteLog(ThisReceive, "==程式啟動==");
            string ProgramFullPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName.ToString();
            string ProgramLastWriteTime = File.GetLastWriteTime(ProgramFullPath).ToString();
            MyCommond.WriteLog(ThisReceive, $"This Program Last Write Time： {ProgramLastWriteTime}");

            MainServer = new Report.Mrt.MrtMain();
            MainServer.globalCommond = new GlobalCommonds(true, true, true);
            //MainServer.globalCommond = new GlobalCommonds("Mrt", true, true, true);

        }

        private void Variable_Init()
        {
            Test_ = false;

            AfterThread = true;
            Detect_Start = false;
            RunFinish = false;
        }

        /// <summary>
        /// 其他初始化或是執行程式。
        /// </summary>
        private void Other_Init()
        {
            if (true) MyCommond.CheckApplicationAndKill("EXCEL");

            // 掃描Template資料夾
            if (!MyCommond.ScanFolderAndExport($@"{MyCommond.Path_Program}{MyCommond.Path_Template}", 2))
            {
                Thread thread = new Thread(() =>
                {
                    MyCommond.MsgShow($@"發現路徑：{MyCommond.Path_Program}{MyCommond.Path_Template} 中檔案有缺失，不影響主程式執行，但部分報表將無法產製。"
                        , "提示"
                        , MessageBoxButtons.OK
                        , MessageBoxIcon.Question);
                });
                thread.IsBackground = true;
                thread.Start();
            }

            return;
            if (!true) MyCommond.ScanFolderAndExport($@"C:\");
            if (!true) MyCommond.ScanFolderAndExport($@"D:\");

        }

        /// <summary>
        /// 掛聽 From 中的物件。
        /// </summary>
        private void RegisterEvents()
        {
            //Form
            this.Load += Form_Load;

            // Menu
            this.notifyIcon1.MouseMove += NotifyIcon1_MouseMove;
            this.notifyIcon1.MouseClick += NotifyIcon1_MouseClick;

            this.toolStripMenuItem1.Click += ToolStripMenuItem1_Click;
            this.toolStripMenuItem1.MouseLeave += ToolStripMenuItem1_MouseLeave;
        }

        private void ToolStripMenuItem1_MouseLeave(object sender, EventArgs e)
        {
            contextMenuStrip1.Hide();
        }

        /// <summary>
        /// 浮動menu中的關閉功能。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (true) MyCommond.CheckApplicationAndKill("EXCEL");
            if (true) WhileEnd();
        }

        private bool contextMenuAllowed = false;

        /// <summary>
        /// 最小化的Item活屬點擊偵測。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NotifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Console.WriteLine("左鍵");
            } // 左鍵
            else if (e.Button == MouseButtons.Right)
            {
                Console.WriteLine("右鍵");
                contextMenuStrip1.Visible = true;
                contextMenuStrip1.Show(System.Windows.Forms.Cursor.Position);
                System.Threading.Thread thread = new System.Threading.Thread(() => { Thread.Sleep(10 * Time_1s); MyCommond.InvokeIfRequired(this, () => { contextMenuStrip1.Visible = false; }); });
                thread.IsBackground = true;
                thread.Start();
            } // 右鍵
            else if (e.Button == MouseButtons.Middle)
            {
                Console.WriteLine("中鍵");
            } // 中鍵
            else
            {
                Console.WriteLine("其他" + e.Button.ToString());
            } // 其他  不知道有啥用
        }

        /// <summary>
        /// 載入主畫面程式。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form_Load(object sender, EventArgs e)
        {
            MyCommond.InvokeIfRequired(this, () => notifyIcon1.Visible = true);
            NotifyShowBalloonTip("最小化執行", "程式於後台執行中...");
            SetTimer();
            if (true) { timer_Schedule_1.Start(); } // 開始程式執行
        }

        private System.Windows.Forms.Timer timer_Schedule_1;
        private System.Windows.Forms.Timer timer_MrtRunStatus;

        /// <summary>
        /// 初始化 Timer。
        /// </summary>
        private void SetTimer()
        {
            timer_Schedule_1 = new System.Windows.Forms.Timer();
            timer_Schedule_1.Interval = 1000;
            timer_Schedule_1.Tick += Schedule_1;

            timer_MrtRunStatus = new System.Windows.Forms.Timer();
            timer_MrtRunStatus.Interval = 1000;
            timer_MrtRunStatus.Tick += MrtRunStatus;
        }

        /// <summary>
        /// 監聽 Mrt 計算主程式執行狀況。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MrtRunStatus(object sender, EventArgs e)
        { 
            MyCommond.InvokeIfRequired(this, () =>
            {
                try
                {
                    var BarMax = MainServer.TxnReportData_Min;
                    var BarCount = MainServer.TxnReportData_Count;
                    var i = MainServer.MrtClients.Count;

                    /** 添加線程前 */if ( AfterThread && i != 0)
                    {
                        AfterThread = false;
                        MouseMoveShowStatueString = "正在分析資料...";
                        NotifyShowBalloonTip("開始執行", MouseMoveShowStatueString);
                    }
                    /** 添加線程前 */if ( AfterThread && i == 0 && MainServer.TxnReportData_Run) 
                    {
                        // if (bcc == 0) { bcc++; }
                        MouseMoveShowStatueString = $"交易資料載入:{BarCount,8} / {BarMax,8}";
                        if (BarMax == BarCount) { MouseMoveShowStatueString = $"目前正在:{MainServer.TxnReportData_Str}"; }
                    }
                    /** 添加線程後 */if (!AfterThread && i == 0)
                    {
                        AfterThread = true;
                        RunFinish = true;
                        MouseMoveShowStatueString = "待機中...";
                        NotifyShowBalloonTip("完成執行", MouseMoveShowStatueString);
                        timer_MrtRunStatus.Stop();
                    }
                    /** 添加線程後 */if (!AfterThread && i != 0 && MainServer.TxnReportData_Run) 
                    {
                        MouseMoveShowStatueString = $"完成載入，開始分析";
                        NotifyShowBalloonTip("開始執行", MouseMoveShowStatueString);
                    }
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $@"異常!訊息:{ex.Message}");
                }
            });
        }

        /// <summary>
        /// 監聽 From 程式執行狀況。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Schedule_1(object sender, EventArgs e)
        {
            MyCommond.InvokeIfRequired(this, () =>
            {
                DateTime dt = DateTime.Now;
                TimeSpan ds = new TimeSpan(dt.Hour, dt.Minute, dt.Second);

                try
                {
                    if (Test_)
                    {
                        int mH = 0;
                        int mm = 0;
                        int mS = 0;
                        next_ds = new TimeSpan(mH, mm, mS);
                    }

                    if (ds == next_ds)
                    {
                        Class_Init();
                        Variable_Init();
                        Other_Init();
                    } // 跨日後重新載入class用法及變數

                    if (dt.Minute != tempMinute)
                    {
                        tempMinute = dt.Minute;  // 將值存放到暫存變數
                        MyCommond.WriteLog(ThisReceive, $"Program is life. Memory Using: {MyCommond.GetMemory()}");  // 輸出存活字串
                    } // 暫存分鐘的值跟當前分鐘的值不相等才輸出(每分鐘發送一次存活字串) (不管是不是在執行計算都回報每分鐘都回報一次)
                    else
                    {

                    } // 值都相同就不動作

                    //if (true) return;
                    if (!Detect_Start && !RunFinish && MainServer.globalCommond.Detect_Time_Start < ds && ds < MainServer.globalCommond.Detect_Time_End)
                    {
                        Detect_Start = true;
                        MouseMoveShowStatueString = "執行中...";
                        NotifyShowBalloonTip("開始執行", MouseMoveShowStatueString);
                        if (Test_)
                        {
                            int mY = 2024;
                            int mM = 3;
                            int mD = 5;
                            dt = new DateTime(mY, mM, mD, 5, 0, 0);
                        }
                        StartTxn(Convert.ToDateTime(dt.ToString("yyyy/MM/dd 05:00:00")).AddDays(-1));
                    }  // 啟動程式
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $@"狀態回報錯誤!訊息：{ex.Message}");
                }
            });
        }

        /// <summary>
        /// 排程時間內啟動要執行的程式。
        /// </summary>
        /// <param name="Start_Dt"></param>
        private void StartTxn(DateTime Start_Dt)
        {
            Thread thread = new Thread(() =>
            {
                string filename_1 = $@"history{Start_Dt.ToString("yyyyMMdd")}_nc";
                string history_fullpath = MyCommond.Path_FileZilla_root + MyCommond.Path_FileZilla_History + $@"{filename_1}.7z";
                string analyze_fullpath = MyCommond.Path_Program + MyCommond.Path_Analyze + $@"{filename_1}.csv";
                int find_histroy_count = 0;
                bool find_root_histroy_bool = false;
                bool find_analyze_histroy_bool = false;
                while (!(find_root_histroy_bool || find_analyze_histroy_bool))
                {
                    DateTime dtn = DateTime.Now;
                    TimeSpan ds = new TimeSpan(dtn.Hour, dtn.Minute, dtn.Second);
                    if (!(MainServer.globalCommond.Detect_Time_Start < ds && ds < MainServer.globalCommond.Detect_Time_End)) { Detect_Start = false; break; } /* 已經不再產製檔案的時間，跳出迴圈 */

                    if (find_histroy_count++ == 0) { MyCommond.WriteLog(ThisReceive, "待檔案取得"); }
                    if (!find_root_histroy_bool /*-------*/) { find_root_histroy_bool = MyCommond.CheckFile(history_fullpath); }
                    if (!find_analyze_histroy_bool /*-*/) { find_analyze_histroy_bool = MyCommond.CheckFile(analyze_fullpath); }
                    //find_histroy_count++;
                    Thread.Sleep(Time_1s);
                } /* 尋找ftproot資料夾或analyze資料夾  其中一個有前一天營運日的history檔案就跳出 */

                if (find_analyze_histroy_bool || find_root_histroy_bool) { MyCommond.WriteLog(ThisReceive, "發現資料，開始程式。"); } 
                else { MyCommond.WriteLog(ThisReceive, $@"find_analyze_histroy_bool:{find_analyze_histroy_bool,8}, find_root_histroy_bool:{find_root_histroy_bool,8}"); }

                bool Decompression_TF = false;
                bool Analyze_History_File = false;
                int lookfile_int = 0;
                int re_Check_history = 10;
                if (true)
                {
                    try
                    {
                        int rt_1 = 0;
                        MyCommond.WriteLog(ThisReceive, $@"解壓縮檔案");
                    StartDecompression:
                        rt_1++;
                        Decompression_TF = MyCommond.Cmd_Decompression(ThisReceive, history_fullpath);
                        if (MyCommond.CheckFile(analyze_fullpath))
                        {
                            MyCommond.WriteLog(ThisReceive, $@"完成解壓縮");
                        }
                        else if (rt_1 < re_Check_history)
                        {
                            MyCommond.WriteLog(ThisReceive, $@"解壓縮異常，準備重試");
                            Thread.Sleep(10 * Time_1s);
                            goto StartDecompression;
                        } /* 重試解壓縮檔案，間隔10秒，重試超過10次將跳過解壓縮 */
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $@"{ex.Message}");
                    }
                } /* 解壓縮檔案 */
                if (true)
                {
                    MyCommond.InvokeIfRequired(this, () =>
                    {
                        Thread thread1 = new Thread(() =>
                        {
                            try
                            {
                                MoveFile(Start_Dt.AddDays(1));
                            }
                            catch (Exception ex)
                            {
                                MyCommond.WriteLog(ThisReceive, $@"自動搬移錯誤!訊息：{ex.Message}");
                            }
                        });
                        thread1.IsBackground = true;
                        thread1.Start();
                    });
                } /*  搬移檔案  */
                if (true)
                {
                    try
                    {
                    Recheck_AnalyzeFile:
                        Analyze_History_File = MyCommond.CheckFile(analyze_fullpath);
                        if (!Analyze_History_File && lookfile_int++ <= re_Check_history)
                        {
                            Thread.Sleep(10 * Time_1s);
                            goto Recheck_AnalyzeFile;
                        } /* 重複確認history檔案是否完成解壓縮，最大確認次數10次，每次等待10秒 */
                        else if (lookfile_int > re_Check_history)
                        {
                            MyCommond.WriteLog(ThisReceive, $"解壓縮可能有異常，未能發現{analyze_fullpath}檔案。");
                            Detect_Start = false;
                            goto EndSub;
                        } /* 大於十次直接結束 */

                        if (Analyze_History_File)
                        {
                            MyCommond.InvokeIfRequired(this, () =>
                            {
                                MyCommond.WriteLog(ThisReceive, $@"準備啟動Mrt計算程式");
                                Thread.Sleep(10 * Time_1s);
                                MainServer.LoadTxn(new CheckBoxStatus()
                                {
                                    OperationStart /*------*/ = Start_Dt,
                                    OperationEnd /*--------*/ = Start_Dt.AddDays(1),
                                    EasyAnalyze /*---------*/ = true,
                                    TxnTypeAnalyze /*------*/ = true,
                                    IOAnalyze /*-----------*/ = true,
                                    CompareAnalyze /*------*/ = true,
                                    TrnAnalyze /*----------*/ = true,
                                    TPassAnalyze /*--------*/ = true,
                                    VolumnReportAnalyze /*-*/ = true,
                                    AddValueAnalyze /*-----------*/ = false,
                                    EveryIssuerAnalyze /*-----------*/ = false,
                                    checkbox10 /*----------*/ = false,
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
                                }, new List<ReportList>()
                            {
                                new ReportList(){methode = LrtTxnClientSwitch.Day_Volume},
                                new ReportList(){methode = LrtTxnClientSwitch.Day_Amount},
                                new ReportList(){methode = LrtTxnClientSwitch.Day_EachStation_EachTime},
                                new ReportList(){methode = LrtTxnClientSwitch.Day_ElectronicTicket},
                                new ReportList(){methode = LrtTxnClientSwitch.Day_AllRideList},
                                new ReportList(){methode = LrtTxnClientSwitch.Day_TrafficAmount},
                                new ReportList(){methode = LrtTxnClientSwitch.Day_OriginDestination},
                                new ReportList(){methode = LrtTxnClientSwitch.Day_EquipAmount},
                                new ReportList(){methode = LrtTxnClientSwitch.Month_AllRideList},
                                new ReportList(){methode = LrtTxnClientSwitch.Month_OriginDestination},
                                new ReportList(){methode = LrtTxnClientSwitch.Month_ElectronicTicket_Station},
                                new ReportList(){methode = LrtTxnClientSwitch.Month_ElectronicTicket_Day},
                                new ReportList(){methode = LrtTxnClientSwitch.Month_OwnTicketVolume_Station},
                                new ReportList(){methode = LrtTxnClientSwitch.Month_OwnTicketVolume_Day},
                                new ReportList(){methode = LrtTxnClientSwitch.Month_TrafficAmount_Station},
                                new ReportList(){methode = LrtTxnClientSwitch.Month_TrafficAmount_Day},
                                new ReportList(){methode = LrtTxnClientSwitch.Month_EquipAmount_Station},
                                new ReportList(){methode = LrtTxnClientSwitch.Month_EquipAmount_Day},
                            });
                                Thread.Sleep(10 * Time_1s);
                                timer_MrtRunStatus.Start();
                            });
                        }  /* 執行報表產製主程式 */
                        else
                        {
                            MyCommond.WriteLog(ThisReceive, "未發現檔案，結束程式。");
                        }
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $@"自動執行錯誤!訊息：{ex.Message}");
                    }
                } /*  產製報表  */

            EndSub:
                MyCommond.WriteLog(ThisReceive, $@"退出啟動用副程式。");
            });
            thread.IsBackground = true;
            thread.Start();

        }

        private void WhileEnd()
        {
            MyCommond.InvokeIfRequired(this, () => notifyIcon1.Visible = false);
            Console.WriteLine("press any key to exit the process...");
            //Console.ReadKey(false);
            MyCommond.WriteLog(ThisReceive, "結束啟動程式");
        StartEnd:
            try
            {
                System.Environment.Exit(0);
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, "重試關閉程序!");
                goto StartEnd;
            }
            
        }

        /// <summary>
        /// From UI 最小化ICon後將滑鼠移置上方出現定義的文字
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void NotifyIcon1_MouseMove(object sender, MouseEventArgs e) { MyCommond.InvokeIfRequired(this, () => { notifyIcon1.Text = $"{MouseMoveShowStatueString:N}"; }); }

        /// <summary>
        /// 在右下角彈出訊息。
        /// </summary>
        /// <param name="TitleString">訊息概要。</param>
        /// <param name="InsideString">詳細訊息。</param>
        private void NotifyShowBalloonTip(string TitleString, string InsideString)
        {
            MyCommond.InvokeIfRequired(this, () => notifyIcon1.ShowBalloonTip(1000, TitleString, InsideString, ToolTipIcon.None));
            MyCommond.WriteLog(ThisReceive, $"A1:{TitleString}, A2:{InsideString}");
        }

        /// <summary>
        /// 將A點的資料搬到B點的資料。
        /// </summary>
        /// <param name="End_dt"></param>
        /// <param name="MoreAgo"></param>
        private void MoveFile(DateTime End_dt, bool MoreAgo = false)
        {
            int MoveCount = 0;
            int Switch_TCLTxn_History = 0;
            string[] Path_0 = {
                      MyCommond.Path_FileZilla_TclTxn
                    , MyCommond.Path_FileZilla_History
            };

            MyCommond.WriteLog(ThisReceive, "開始搬移前一日或以前的檔案");
            foreach (var path in Path_0)
            {
                Switch_TCLTxn_History++;

                MyCommond.WriteLog(ThisReceive, $@"準備搬移 {MyCommond.Path_FileZilla_root + path} 中的檔案 至 {MyCommond.Path_Program + Path_MrtData_Backup + path}");
                DirectoryInfo di = new DirectoryInfo(@"C:/");
                try { di = new DirectoryInfo(MyCommond.Path_FileZilla_root + path); }
                catch (Exception ex) { MyCommond.WriteLog(ThisReceive, $@"Error! 訊息:{ex.Message}"); continue; }

                if (di.GetFiles().GetLength(0) == 0 || di.GetFiles().GetLength(0) == null) { MyCommond.WriteLog(ThisReceive, $@"未發現檔案。"); continue; }
                foreach (var fi in di.GetFiles())
                {
                    try
                    {
                        DateTime dt_file = new DateTime();
                        int mOffset = -1;
                        if (MoveFileHandleString2(End_dt, fi.Name, ref dt_file, ref mOffset))
                        {
                            MyCommond.WriteLog(ThisReceive, $@"" + 
                                $@"搬移第 {++MoveCount,4} 個檔案" +
                                $@"至資料夾:{MyCommond.Path_Program}{Path_MrtData_Backup}{path}{((Switch_TCLTxn_History == 1) ? $@"{dt_file.ToString("yyyyMMdd")}\" : "")}, " +
                                $@"檔名(Offset:{mOffset, 2})：{fi.Name}");
                            MyCommond.MoveFile($@"{MyCommond.Path_FileZilla_root}{path}\"
                                , $@"{MyCommond.Path_Program}{Path_MrtData_Backup}{path}{((Switch_TCLTxn_History == 1) ? $@"{dt_file.ToString("yyyyMMdd")}\" : "")}"
                                , fi.Name); 
                        }
                    }
                    catch (Exception ex) { MyCommond.WriteLog(ThisReceive, $@"搬移出現錯誤!訊息:{ex.Message}"); }
                }
            }
            MyCommond.WriteLog(ThisReceive, "已完成前一日或以前的檔案搬移");
        }

        /// <summary>
        /// 依檔案名稱判斷時間。
        /// </summary>
        /// <param name="End_dt">要判斷的時間。</param>
        /// <param name="FileName">檔案名稱。</param>
        /// <param name="FindDate">回傳該檔案的時間。</param>
        /// <returns>回傳是否處裡。</returns>
        private bool MoveFileHandleString(DateTime End_dt, string FileName, ref DateTime FindDate, ref int Offset)
        {
            try
            {
                try
                {
                    Offset = 7;
                    /** y */ var mYear__ = Convert.ToInt16(FileName.Substring(Offset, 4));
                    /** M */ var mMonth_ = Convert.ToInt16(FileName.Substring(Offset + 4, 2));
                    /** d */ var mDay___ = Convert.ToInt16(FileName.Substring(Offset + 4 + 2, 2));
                    /** H */ var mHour__ = Convert.ToInt16(FileName.Substring(Offset + 4 + 2 + 2, 2));
                    /** m */ var mMinute = Convert.ToInt16(FileName.Substring(Offset + 4 + 2 + 2 + 2, 2));
                    /** s */ var mSecond = Convert.ToInt16(FileName.Substring(Offset + 4 + 2 + 2 + 2 + 2, 2));
                    FindDate = new DateTime(mYear__, mMonth_, mDay___, mHour__, mMinute, mSecond);
                }
                catch
                {
                    Offset = 6;
                    /** y */ var mYear__ = Convert.ToInt16(FileName.Substring(Offset, 4));
                    /** M */ var mMonth_ = Convert.ToInt16(FileName.Substring(Offset + 4, 2));
                    /** d */ var mDay___ = Convert.ToInt16(FileName.Substring(Offset + 4 + 2, 2));
                    /** H */ var mHour__ = Convert.ToInt16(FileName.Substring(Offset + 4 + 2 + 2, 2));
                    /** m */ var mMinute = Convert.ToInt16(FileName.Substring(Offset + 4 + 2 + 2 + 2, 2));
                    /** s */ var mSecond = Convert.ToInt16(FileName.Substring(Offset + 4 + 2 + 2 + 2 + 2, 2));
                    FindDate = new DateTime(mYear__, mMonth_, mDay___, mHour__, mMinute, mSecond);
                }
            }
            catch
            {
                try
                {
                    Offset = 7;
                    /** y */ var mYear_ = Convert.ToInt16(FileName.Substring(Offset, 4));
                    /** M */ var mMonth = Convert.ToInt16(FileName.Substring(Offset + 4, 2));
                    /** d */ var mDay__ = Convert.ToInt16(FileName.Substring(Offset + 4 + 2, 2));
                    FindDate = new DateTime(mYear_, mMonth, mDay__);
                }
                catch
                {
                    Offset = 6;
                    /** y */ var mYear_ = Convert.ToInt16(FileName.Substring(Offset, 4));
                    /** M */ var mMonth = Convert.ToInt16(FileName.Substring(Offset + 4, 2));
                    /** d */ var mDay__ = Convert.ToInt16(FileName.Substring(Offset + 4 + 2, 2));
                    FindDate = new DateTime(mYear_, mMonth, mDay__);
                }
            }
            return FindDate < End_dt;
        }
        
        private bool MoveFileHandleString2(DateTime End_dt, string FileName, ref DateTime FindDate, ref int Offset)
        {
            int StringLength = FileName.Length;
            
            for (Offset = 0; Offset < StringLength; Offset++)
            {
                try
                {
                    try
                    {
                        /** y */ var mYear__ = Convert.ToInt16(FileName.Substring(Offset,  4));
                        /** M */ var mMonth_ = Convert.ToInt16(FileName.Substring(Offset + 4, 2));
                        /** d */ var mDay___ = Convert.ToInt16(FileName.Substring(Offset + 4 + 2, 2));
                        /** H */ var mHour__ = Convert.ToInt16(FileName.Substring(Offset + 4 + 2 + 2, 2));
                        /** m */ var mMinute = Convert.ToInt16(FileName.Substring(Offset + 4 + 2 + 2 + 2, 2));
                        /** s */ var mSecond = Convert.ToInt16(FileName.Substring(Offset + 4 + 2 + 2 + 2 + 2, 2));
                        FindDate = new DateTime(mYear__, mMonth_, mDay___, mHour__, mMinute, mSecond);
                    }
                    catch
                    {
                        /** y */ var mYear_ = Convert.ToInt16(FileName.Substring(Offset,  4));
                        /** M */ var mMonth = Convert.ToInt16(FileName.Substring(Offset + 4, 2));
                        /** d */ var mDay__ = Convert.ToInt16(FileName.Substring(Offset + 4 + 2, 2));
                        FindDate = new DateTime(mYear_, mMonth, mDay__);
                        End_dt = new DateTime(End_dt.Year, End_dt.Month, End_dt.Day);
                    }
                    break;
                }
                catch
                {
                    continue;
                }
            }
            return FindDate < End_dt;
        }

    }

}
