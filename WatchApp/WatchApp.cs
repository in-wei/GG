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
using UsuallyCommond;
using System.IO;
using System.Threading;
using System.Reflection;
using UsuallyCommond.MyEnum;

namespace WatchApp
{
    public partial class WatchApp : Form
    {
        #region Program
        MyCommonds MyCommond;
        string ThisReceive = "WatchApp";
        int Time_OneSecend = 1000;
        #endregion
        #region Form Control
        bool Form_Close = false;
        #endregion

        List<List_listviews> list_Listviews;
        string autoExe = "";

        public WatchApp()
        {
            InitializeComponent();

            Class_Init();       // 激活所有class
            Variable_Init();    // 初始化變數
            Other_Init();       // 其他執行
            RegisterEvents();   // 掛聽程式

        }

        private void Class_Init()
        {
            MyCommond = new MyCommonds();

            MyCommond.WriteLog(ThisReceive, "==程式啟動==");
            string ProgramFullPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName.ToString();
            string ProgramLastWriteTime = File.GetLastWriteTime(ProgramFullPath).ToString();
            MyCommond.WriteLog(ThisReceive, $"This Program Last Write Time： {ProgramLastWriteTime}");



        }

        private void Variable_Init()
        {
            if (!true)
            {
                aaaa();
            }
            else
            {
                autoExe = "Main_Console_Auto";
                list_Listviews = new List<List_listviews>()
                {
                    new List_listviews(1, "Main_Console_Auto"),
                    new List_listviews(2, "Report.Lrt.UI"),
                    new List_listviews(3, "Report.Mrt.UI"),
                };
            }
            
        }

        private void aaaa()
        {
            DirectoryInfo di = new DirectoryInfo($@"{MyCommond.Path_Program}");
            string MyName = this.GetType().Assembly.GetName().Name;
            int counts = 1;
            list_Listviews = new List<List_listviews>();
            foreach (var fi in di.GetFiles("*.exe"))
            {
                try
                {
                    if (fi.Name == MyName + ".exe") continue;
                    var ff = fi.Name.Equals("*Auto*");
                    List_listviews a = new List_listviews(counts++, fi.Name);
                    list_Listviews.Add(a);
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, ex.Message);
                }
            }
        }

        /// <summary>
        /// 其他初始化或是執行程式。
        /// </summary>
        private void Other_Init()
        {
            if (true) MyCommond.CheckApplicationAndKill("EXCEL");
            return;

            if (true) MyCommond.ScanFolderAndExport($@"C:\");
            if (true) MyCommond.ScanFolderAndExport($@"D:\");
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

        }

        /// <summary>
        /// 掛聽 From 中的物件。
        /// </summary>
        private void RegisterEvents()
        {
            //Form
            this.Load += Form_Load;
            this.FormClosing += Form1_FormClosing;

            // Menu
            this.notifyIcon1.MouseMove += NotifyIcon1_MouseMove;
            this.notifyIcon1.MouseClick += NotifyIcon1_MouseClick;

            this.toolStripMenuItem1.Click += ToolStripMenuItem1_Click;

            this.listView1.MouseClick += ListView1_MouseClick;                      // 在ListView物件上的滑鼠操作
            this.listView1.DoubleClick += ListView1_DoubleClick;                    // 在ListView物件上雙擊左鍵
            this.listView1.ItemSelectionChanged += ListView1_ItemSelectionChanged;  // 點選ListView物件

            this.contextMenuStrip1.MouseLeave += ContextMenuStrip1_MouseLeave;  // 滑鼠離開ListView物件右鍵後出現的Menu
            this.Listview_menu_Switch.Click += Listview_menu_Switch_Click;      // 掛聽ListView物件右鍵後出現的Menu中的 開啟/關閉 按鈕

        }

        #region From - Event

        private void Form_Load(object sender, EventArgs e)
        {
            ListView1_SetITem();
            SetTimer();
            Timer_EverySecond.Start();
            //Timer_Item_1.Start();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Form_Close)
            {
                GC.Collect();
                MyCommond.WriteLog(ThisReceive, "程式關閉");
                System.Environment.Exit(0);
            }
            else
            {
                e.Cancel = true;
            }
        }

        #endregion

        #region Form - Timer

        System.Windows.Forms.Timer Timer_EverySecond;
        System.Windows.Forms.Timer Timer_Item_1;
        System.Windows.Forms.Timer Timer_Item_2;
        System.Windows.Forms.Timer Timer_Item_3;
        System.Windows.Forms.Timer Timer_Item_4;
        System.Windows.Forms.Timer Timer_Item_5;

        bool OnListViewMenu = false;

        private void SetTimer()
        {
            Timer_EverySecond = new System.Windows.Forms.Timer();
            Timer_EverySecond.Interval = 1 * Time_OneSecend;
            Timer_EverySecond.Tick += EverySecond_Tick;

            try
            {
                Timer_Item_1 = new System.Windows.Forms.Timer();
                Timer_Item_1.Interval = 30 * Time_OneSecend;
                Timer_Item_1.Tick += Item1_Tick;
                list_Listviews[0].SetTimer(Timer_Item_1);

                Timer_Item_2 = new System.Windows.Forms.Timer();
                Timer_Item_2.Interval = 30 * Time_OneSecend;
                Timer_Item_2.Tick += Item2_Tick;
                list_Listviews[1].SetTimer(Timer_Item_2);

                Timer_Item_3 = new System.Windows.Forms.Timer();
                Timer_Item_3.Interval = 30 * Time_OneSecend;
                Timer_Item_3.Tick += Item3_Tick;
                list_Listviews[2].SetTimer(Timer_Item_3);

                Timer_Item_4 = new System.Windows.Forms.Timer();
                Timer_Item_4.Interval = 60 * Time_OneSecend;
                Timer_Item_4.Tick += Item4_Tick;
                list_Listviews[3].SetTimer(Timer_Item_4);

                Timer_Item_5 = new System.Windows.Forms.Timer();
                Timer_Item_5.Interval = 60 * Time_OneSecend;
                Timer_Item_5.Tick += Item5_Tick;
                list_Listviews[4].SetTimer(Timer_Item_5);
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"{ex}");
            }

            

        }

        private void EverySecond_Tick(object sender, EventArgs e)
        {
            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                MyCommond.InvokeIfRequired(this, () =>
                {
                    var dt = DateTime.Now;
                    var ds = new TimeSpan(dt.Hour,dt.Minute,dt.Second);
                    this.Text = $@"WatchApp - 現在時間:{dt:yyyy/MM/dd HH:mm:ss}";
                    if (true)
                    {
                        if (!(ds == new TimeSpan(23, 59, 59))) goto ToEnd;
                        if (MyCommond.CheckApplicationAndKill(autoExe, "")) goto ToEnd;
                        MyCommond.CheckApplicationAndKill(autoExe, "<>");
                    }
                ToEnd:
                    int zz = 0;
                });
                if (!OnListViewMenu) MyCommond.InvokeIfRequired(this, ListView1_SetITem);
                MyCommond.InvokeIfRequired(this, TimeScan);
            });
            thread.IsBackground = true;
            thread.Start();
        }

        private void TimeScan()
        {
            foreach (var item in list_Listviews)
            {
                     if ( item.TimerStatus &&  item.Switch)
                {

                } // 開關 Open  且 計時器 Start (不作動)
                else if (!item.TimerStatus &&  item.Switch)
                {
                    if (item.TimerStatus) continue;
                    item.Timer_.Start();
                    item.SetTimerStatus(true);
                } // 開關 Open  且 計時器 Stop  (要把計時器開啟)
                else if ( item.TimerStatus && !item.Switch)
                {
                    item.Timer_.Stop();
                    item.SetTimerStatus(false);
                } // 開關 Close 且 計時器 Start (要把計時器關閉)
                else if (!item.TimerStatus && !item.Switch)
                {

                } // 開關 Close 且 計時器 Stop  (不作動)
            }
        }

        private void ListView1_SetITem()
        {
            listView1.Items.Clear();
            foreach (var item in list_Listviews)
            {
                var NewStatus = (MyCommond.CheckApplication(item.ApplactionName)) ? ViewListStatus.Running : ViewListStatus.Closed;

                DateTime DatetimeNow = DateTime.Now;

                if (item.Status == ViewListStatus.Running && NewStatus == ViewListStatus.Running)
                {
                    item.SetDatetime(DateTimeSwitch.LastDatetime, DatetimeNow);
                } // 原狀態為 執行中 且 新狀態為 執行中 (在畫面     紀錄最後時間)
                else if (item.Status == ViewListStatus.Closed && NewStatus == ViewListStatus.Running)
                {
                    MyCommond.WriteLog(ThisReceive, $"程式被啟動，啟動時間{DatetimeNow}");
                    item.SetDatetime(DateTimeSwitch.RunDatetime, DatetimeNow);
                } // 原狀態為 未執行 且 新狀態為 執行中 (在畫面及Log紀錄啟動時間)
                else if (item.Status == ViewListStatus.Running && NewStatus == ViewListStatus.Closed)
                {
                    MyCommond.WriteLog(ThisReceive, $"程式已斷線，斷線時間{DatetimeNow}");
                } // 原狀態為 執行中 且 新狀態為 未執行 (在      Log紀錄斷線時間)
                else if (item.Status == ViewListStatus.Closed && NewStatus == ViewListStatus.Closed)
                {

                } // 原狀態為 未執行 且 新狀態為 未執行 (暫不動作)



                item.ChangeStatus(NewStatus);
                string[] str = new string[listView1.Columns.Count];
                string[] str2 = item.ToString().Split(',');

                int offset = -1;
                     if (str.GetLength(0) == str2.GetLength(0)) offset =  str.GetLength(0);
                else if (str.GetLength(0) <  str2.GetLength(0)) offset =  str.GetLength(0);
                else if (str.GetLength(0) >  str2.GetLength(0)) offset = str2.GetLength(0);

                for (int i = 0; i < offset; i++) str[i] = str2[i];

                ListViewItem Lwi = new ListViewItem(str);
                listView1.Items.Add(Lwi);
            }

        }

        private void Item_Process(List_listviews item)
        {
            if (!MyCommond.CheckApplication(item.ApplactionName))
            {
                MyCommond.WriteLog(ThisReceive, $"工作管理員裡未發現{item.ApplactionName}程式，準備啟動");
                System.Threading.Thread thread = new System.Threading.Thread(() =>
                {
                    MyCommond.Cmd_(ThisReceive, $@"{MyCommond.Path_Program}\{item.ApplactionName}.exe");
                    int countDelay = 0;
                    while (!MyCommond.CheckApplication(item.ApplactionName))
                    {
                        System.Threading.Thread.Sleep(10 * Time_OneSecend);
                        if (countDelay++ > 6)
                        {
                            item.ChangeSwitch(false);
                            MyCommond.WriteLog(ThisReceive, "重試6次未發現程式被啟動，將開關強制改為關閉。");
                        }
                    }
                    if (MyCommond.CheckApplication(item.ApplactionName)) item.SetDatetime(DateTimeSwitch.LastDatetime, DateTime.Now);
                });
                thread.IsBackground = true;
                thread.Start();
            }
        }

        private void Item1_Tick(object sender, EventArgs e)
        {
            List_listviews item = list_Listviews[0];
            if (true)
            {
                Item_Process(item);
            }
            else
            {
                if (!MyCommond.CheckApplication(item.ApplactionName))
                {
                    MyCommond.WriteLog(ThisReceive, $"工作管理員裡未發現{item.ApplactionName}程式，準備啟動");
                    System.Threading.Thread thread = new System.Threading.Thread(() =>
                    {
                        MyCommond.Cmd_(ThisReceive, $@"{MyCommond.Path_Program}\{item.ApplactionName}.exe");
                        int countDelay = 0;
                        while (!MyCommond.CheckApplication(item.ApplactionName))
                        {
                            System.Threading.Thread.Sleep(10 * Time_OneSecend);
                            if (countDelay++ > 6)
                            {
                                item.ChangeSwitch(false);
                                MyCommond.WriteLog(ThisReceive, "重試6次未發現程式被啟動，將開關強制改為關閉。");
                            }
                        }
                        if (MyCommond.CheckApplication(item.ApplactionName)) item.SetDatetime(DateTimeSwitch.LastDatetime, DateTime.Now);
                    });
                    thread.IsBackground = true;
                    thread.Start();
                }
            }
        }

        private void Item2_Tick(object sender, EventArgs e)
        {
            List_listviews item = list_Listviews[1];
            if (true)
            {
                Item_Process(item);
            }
            else
            {
                if (!MyCommond.CheckApplication(item.ApplactionName))
                {
                    MyCommond.WriteLog(ThisReceive, $"工作管理員裡未發現{item.ApplactionName}程式，準備啟動");
                    System.Threading.Thread thread = new System.Threading.Thread(() =>
                    {
                        MyCommond.Cmd_(ThisReceive, $@"{MyCommond.Path_Program}\{item.ApplactionName}.exe");
                        int countDelay = 0;
                        while (!MyCommond.CheckApplication(item.ApplactionName))
                        {
                            System.Threading.Thread.Sleep(10 * Time_OneSecend);
                            if (countDelay++ > 6)
                            {
                                item.ChangeSwitch(false);
                                MyCommond.WriteLog(ThisReceive, "重試6次未發現程式被啟動，將開關強制改為關閉。");
                            }
                        }
                        if (MyCommond.CheckApplication(item.ApplactionName)) item.SetDatetime(DateTimeSwitch.LastDatetime, DateTime.Now);
                    });
                    thread.IsBackground = true;
                    thread.Start();
                }
            }
        }

        private void Item3_Tick(object sender, EventArgs e)
        {
            List_listviews item = list_Listviews[2];
            if (true)
            {
                Item_Process(item);
            }
            else
            {
                if (!MyCommond.CheckApplication(item.ApplactionName))
                {
                    MyCommond.WriteLog(ThisReceive, $"工作管理員裡未發現{item.ApplactionName}程式，準備啟動");
                    System.Threading.Thread thread = new System.Threading.Thread(() =>
                    {
                        MyCommond.Cmd_(ThisReceive, $@"{MyCommond.Path_Program}\{item.ApplactionName}.exe");
                        int countDelay = 0;
                        while (!MyCommond.CheckApplication(item.ApplactionName))
                        {
                            System.Threading.Thread.Sleep(10 * Time_OneSecend);
                            if (countDelay++ > 6)
                            {
                                item.ChangeSwitch(false);
                                MyCommond.WriteLog(ThisReceive, "重試6次未發現程式被啟動，將開關強制改為關閉。");
                            }
                        }
                        if (MyCommond.CheckApplication(item.ApplactionName)) item.SetDatetime(DateTimeSwitch.LastDatetime, DateTime.Now);
                    });
                    thread.IsBackground = true;
                    thread.Start();
                }
            }
        }

        private void Item4_Tick(object sender, EventArgs e)
        {
            List_listviews item = list_Listviews[3];
            if (true)
            {
                Item_Process(item);
            }
            else
            {
                if (!MyCommond.CheckApplication(item.ApplactionName))
                {
                    MyCommond.WriteLog(ThisReceive, $"工作管理員裡未發現{item.ApplactionName}程式，準備啟動");
                    System.Threading.Thread thread = new System.Threading.Thread(() =>
                    {
                        MyCommond.Cmd_(ThisReceive, $@"{MyCommond.Path_Program}\{item.ApplactionName}.exe");
                        int countDelay = 0;
                        while (!MyCommond.CheckApplication(item.ApplactionName))
                        {
                            System.Threading.Thread.Sleep(10 * Time_OneSecend);
                            if (countDelay++ > 6)
                            {
                                item.ChangeSwitch(false);
                                MyCommond.WriteLog(ThisReceive, "重試6次未發現程式被啟動，將開關強制改為關閉。");
                            }
                        }
                        if (MyCommond.CheckApplication(item.ApplactionName)) item.SetDatetime(DateTimeSwitch.LastDatetime, DateTime.Now);
                    });
                    thread.IsBackground = true;
                    thread.Start();
                }
            }
        }

        private void Item5_Tick(object sender, EventArgs e)
        {
            List_listviews item = list_Listviews[4];
            if (true)
            {
                Item_Process(item);
            }
            else
            {
                if (!MyCommond.CheckApplication(item.ApplactionName))
                {
                    MyCommond.WriteLog(ThisReceive, $"工作管理員裡未發現{item.ApplactionName}程式，準備啟動");
                    System.Threading.Thread thread = new System.Threading.Thread(() =>
                    {
                        MyCommond.Cmd_(ThisReceive, $@"{MyCommond.Path_Program}\{item.ApplactionName}.exe");
                        int countDelay = 0;
                        while (!MyCommond.CheckApplication(item.ApplactionName))
                        {
                            System.Threading.Thread.Sleep(10 * Time_OneSecend);
                            if (countDelay++ > 6)
                            {
                                item.ChangeSwitch(false);
                                MyCommond.WriteLog(ThisReceive, "重試6次未發現程式被啟動，將開關強制改為關閉。");
                            }
                        }
                        if (MyCommond.CheckApplication(item.ApplactionName)) item.SetDatetime(DateTimeSwitch.LastDatetime, DateTime.Now);
                    });
                    thread.IsBackground = true;
                    thread.Start();
                }
            }
        }

        #endregion

        #region Form - OtherListener

        int check_listview_Index = -1;
        private void ListView1_MouseClick(object sender, MouseEventArgs e)
        {
            //Console.WriteLine($@"{listView1.TabIndex}");

            if (e.Button == MouseButtons.Left)
            {

            }
            else if (e.Button == MouseButtons.Right)
            {
                OnListViewMenu = true;
                Listview_menu_Switch.Text = (list_Listviews[check_listview_Index].Switch) ? "關閉" : "開啟";
                contextMenuStrip1.Visible = true;
                contextMenuStrip1.Show(System.Windows.Forms.Cursor.Position);
            }
            else if (e.Button == MouseButtons.Middle)
            {

            }
        }

        private void ListView1_DoubleClick(object sender, EventArgs e)
        {
            var aa = list_Listviews[check_listview_Index].ApplactionName;
            var bb = list_Listviews[check_listview_Index].Status;
            list_Listviews[check_listview_Index].ChangeSwitch(!list_Listviews[check_listview_Index].Switch);
            var cc = list_Listviews[check_listview_Index].Status;
            MyCommond.WriteLog(ThisReceive, $"{aa}原狀態{bb}更新{cc}");
        }

        private void Listview_menu_Switch_Click(object sender, EventArgs e)
        {
            var aa = list_Listviews[check_listview_Index].ApplactionName;
            var bb = list_Listviews[check_listview_Index].Status;
            list_Listviews[check_listview_Index].ChangeSwitch(!list_Listviews[check_listview_Index].Switch);
            var cc = list_Listviews[check_listview_Index].Status;
            MyCommond.WriteLog(ThisReceive, $"{aa}原狀態{bb}更新{cc}");
            OnListViewMenu = false;
        }

        private void ContextMenuStrip1_MouseLeave(object sender, EventArgs e)
        {
            contextMenuStrip1.Visible = false;
            OnListViewMenu = false;
        }

        private void ListView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            check_listview_Index = e.ItemIndex;
            Console.WriteLine($@"{check_listview_Index}");
        }

        private void ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            WhileEnd();
        }

        private void NotifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void NotifyIcon1_MouseMove(object sender, MouseEventArgs e)
        { 
            MyCommond.InvokeIfRequired(this, () => 
            { 
                notifyIcon1.Text = ThisReceive;
            });
        }



        #endregion

        #region a




        #endregion

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


    }

    public class List_listviews
    {
        public int No { get; private set; }
        public string ApplactionName { get; private set; }
        public ViewListStatus Status { get; private set; }
        public bool Switch { get; private set; }
        public DateTime RunDatetime { get; private set; }
        public DateTime LastDatetime { get; private set; }

        public System.Windows.Forms.Timer Timer_ { get; private set; }
        public bool TimerStatus { get; private set; }

        public List_listviews(int no) : this(no, "", ViewListStatus.NULL, false)
        {
        }
        public List_listviews(int no, string appname) : this(no, appname, ViewListStatus.NULL, false)
        {
        }
        public List_listviews(int no, string appname, bool switch_) : this(no, appname, ViewListStatus.NULL, switch_)
        {
        }
        public List_listviews(int no, string appname, ViewListStatus status, bool switch_)
        {
            No = no;
            ApplactionName = appname;
            Status = (new MyCommonds().CheckApplication(appname)) ? ViewListStatus.Running : ViewListStatus.Closed;
            Switch = switch_;
            RunDatetime = new DateTime();
            LastDatetime = new DateTime();
        }
        public void ChangeStatus(ViewListStatus status) => Status = status;
        public void ChangeSwitch(bool switch_) => Switch = switch_;
        public void SetTimer(System.Windows.Forms.Timer timer)
        {
            Timer_ = timer;
            TimerStatus = false;
        }
        public void SetTimerStatus(bool timeStatus) => TimerStatus = timeStatus;
        public void SetDatetime(DateTimeSwitch ss, DateTime dateTime)
        {
            if (ss == DateTimeSwitch.RunDatetime)
            {
                RunDatetime = dateTime;
            }
            else if (ss == DateTimeSwitch.LastDatetime)
            {
                LastDatetime = dateTime;
            }
        }
        public override string ToString() => $@""
            + $@"{No},"
            + $@"{ApplactionName},"
            + $@"{Status.ToDescription()},"
            + $@"{((Switch) ? "啟動" : "關閉")},"
            + $@"{((RunDatetime == new DateTime()) ? "" : RunDatetime.ToString("yyyy/MM/dd HH:mm:ss"))},"
            + $@"{((LastDatetime == new DateTime()) ? "" : LastDatetime.ToString("yyyy/MM/dd HH:mm:ss"))},";

    }

}
