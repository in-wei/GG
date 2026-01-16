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
using System.Data.SQLite;
using System.IO;
using System.Diagnostics;
using UsuallyCommond.MyEnum;
using UsuallyCommond;
using System.Collections;

namespace Test
{
    public partial class Form1 : Form
    {
        MyCommonds MyCommond = new MyCommonds();
        string ThisReceive;
        FormWindowState FromState;

        bool UseDoubleListView = true;
        bool hideForm = false;

        #region G2 Block

        string DB_Title = "Data Source=";
        string DB_Path = "./DataBase/";
        string DB_Path_Mrt = "Mrt/";
        string DB_Mrt_Title = "";
        string DB_Mrt = "Mrt_Txn_Data.db";
        string DB_Connect = "";

        string Operation_Line = "Mrt";
        string DataBaseName = "Txn_Data.db";

        string MrtTxnColumn;
        string MrtTxnColumn2;
        string[] TxnColumnSplit;
        string[] TxnTable;

        #endregion

        public Form1()
        {
            InitializeComponent();
            ThisReceive = this.Text;
            string ProgramFullPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName.ToString();
            string ProgramLastWriteTime = File.GetLastWriteTime(ProgramFullPath).ToString();
            MyCommond.WriteLog(ThisReceive, $"This Program Last Write Time： {ProgramLastWriteTime}");
            RegisterEvents();
        }

        private void RegisterEvents()
        {

            //Form
            this.Load += Form_Load;
            this.MinimumSizeChanged += Form_MinimumSizeChanged;
            this.FormClosing += Form_Close;
            this.SizeChanged += Form_SizeChanged;

            //UI
            this.button1.Click += Button1_Click;
            this.button4.Click += Button4_Click;
            this.button5.Click += Button5_Click;
            this.button6.Click += Button6_Click;
            this.DB_Create.Click += DB_Create_Click;
            this.DB_Insert.Click += DB_Insert_Click;
            this.DB_UpDate.Click += DB_UpDate_Click;
            this.DB_Delete.Click += DB_Delete_Click;
            this.DB_Read.Click += DB_Read_Click;
            this.StopAct.Click += Action_Stop_Click;
            this.Add_Title.Click += Add_Title_Click;
            this.Add_Column.Click += Add_Column_Click;
            this.MenuItem_Top.CheckedChanged += Menu_Form_Top_CheckedChanged;
            this.MenuItem_Mini.CheckedChanged += Menu_Form_Mini_CheckedChanged;

            // Menu
            this.notifyIcon1.MouseMove += NotifyIcon1_MouseMove;
            this.notifyIcon1.DoubleClick += NotifyIcon1_DoubleClick;
            this.MenuItem_Close.Click += MenuItem_Close_Click;

        }

        void test()
        {
            PerformanceCounter cpuCounter;
            PerformanceCounter ramCounter;

            //cpuCounter = new PerformanceCounter();
            //cpuCounter.CategoryName = "Processor";
            //cpuCounter.CounterName = "% Processor Time";
            //cpuCounter.InstanceName = "_Total";
            cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
            ramCounter = new PerformanceCounter("Memory", "Available MBytes");



            Console.WriteLine("電腦CPU使用率：" + cpuCounter.NextValue() + "%");
            Console.WriteLine("電腦可使用記憶體：" + ramCounter.NextValue() + "MB");
            Console.WriteLine();



            while (true)
            {
                System.Threading.Thread.Sleep(1000);
                Console.WriteLine("電腦CPU使用率：" + cpuCounter.NextValue() + " %");
                Console.WriteLine("電腦可使用記憶體：" + ramCounter.NextValue() + "MB");
                Console.WriteLine();

                if ((int)cpuCounter.NextValue() > 80)
                {
                    System.Threading.Thread.Sleep(1000 * 60);
                }
            }
        }

        static void UsingProcess(string pname)
        {
            
            using (var pro = Process.GetProcessesByName(pname)[0])
            {
                PerformanceCounter pf = new PerformanceCounter(); //性能計數器
                pf.CategoryName = "Process";
                pf.CounterName = "% Processor Time";
                pf.InstanceName = pro.ProcessName;
                pf.MachineName = ".";

                while (true)
                {
                    Console.WriteLine("RAM:" + (Convert.ToInt64(pro.WorkingSet64.ToString()) / 1024).ToString() + " CPU:" + Math.Round(pf.NextValue(), 2).ToString() + "%");
                    System.Threading.Thread.Sleep(1000);
                }
            }
        }

        private void WriteText2Command(string msg)
        {
            MyCommond.InvokeIfRequired(this, () =>
            {
                var txtIndex = txtCommand.SelectionStart;
                string T_string = txtCommand.Text;
                string space = "";
                if (T_string != "")
                {
                    var L1 = T_string.Length - 1;
                    space = T_string.Substring(L1, 1);
                    if (space != " ") space = " "; else space = "";
                }

                txtCommand.Focus();
                txtCommand.Select(txtIndex, 0);
                txtCommand.Paste($"{space}{msg} ");
            });
        }

        #region Form

        public void Form_Load(object sender, EventArgs e)
        {

            G2_UpLoad();
            SetTipUI();
            FromState = this.WindowState;
        }

        private void Add_Column_Click(object sender, EventArgs e)
        {
            if (true)
            {
                WriteText2Command(comboBox2.SelectedItem.ToString());
            }
            else
            {
                MyCommond.InvokeIfRequired(this, () =>
                {
                    string T_string;
                    string space;
                    T_string = txtCommand.Text;
                    if (T_string != "")
                    {
                        var L1 = T_string.Length - 1;
                        space = T_string.Substring(L1, 1);
                        if (space != " ") txtCommand.Text += " ";
                    }
                    txtCommand.Text += comboBox2.SelectedItem.ToString();
                    //txtCommand.Focus();
                });
            }
        }

        private void Add_Title_Click(object sender, EventArgs e)
        {
            if (true)
            {
                WriteText2Command(comboBox1.SelectedItem.ToString());
            }
            else
            {
                MyCommond.InvokeIfRequired(this, () =>
                {
                    string T_string;
                    string space;
                    T_string = txtCommand.Text;
                    if (T_string != "")
                    {
                        var L1 = T_string.Length - 1;
                        space = T_string.Substring(L1, 1);
                        if (space != " ") txtCommand.Text += " ";
                    }
                    txtCommand.Text += comboBox1.SelectedItem.ToString();
                    //txtCommand.Focus();
                });
            }
        }

        private void Form_MinimumSizeChanged(object sender, EventArgs e)
        {
            MyCommond.InvokeIfRequired(this, () =>
            {
                notifyIcon1.ShowBalloonTip(1000, "最小化執行", "程式於後台執行中...", ToolTipIcon.None);
            });
        }

        public void Form_Close(object sender, FormClosingEventArgs e)
        {
            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                MyCommond.InvokeIfRequired(this, () =>
                {
                    this.Hide();
                    //notifyIcon1.Visible = true;
                    notifyIcon1.ShowBalloonTip(1000, "最小化執行", "資料尚在匯入中，將執行至完成後關閉", ToolTipIcon.None);
                });

                Action_Stop();
                while (msql.Stop)
                {
                    if (msql.CountValue == -1)
                    {
                        FromClose("已完成匯入，自動關閉程式");
                    }
                }
            });
            thread.IsBackground = true;

            if (msql.CountValue > 0)
            {
                MyCommond.WriteLog(ThisReceive, "資料匯入中，待資料完成匯入將自動關閉");
                thread.Start();
                e.Cancel = true;
            }
            else
            {
                FromClose("程式關閉");
            }
        }

        private void FromClose(string message)
        {
            MyCommond.InvokeIfRequired(this, () =>
            {
                if (!true)
                {
                    this.ShowInTaskbar = false;
                    this.WindowState = FormWindowState.Minimized;
                    this.Hide();
                }
                else
                {
                    MyCommond.WriteLog(ThisReceive, message);
                    notifyIcon1.Visible = false;
                    GC.Collect();
                    System.Environment.Exit(0);
                }
            });
        }

        private void Form_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                if (!hideForm) return;
                MyCommond.InvokeIfRequired(this, () =>
                {
                    this.Hide();
                    notifyIcon1.ShowBalloonTip(1000, "最小化執行", "程式於後台執行中...", ToolTipIcon.None);
                });
                return;
            }
            else
            {

                if (UseDoubleListView)
                {
                    this.doubleBufferListView1.Width = this.Width - doubleBufferListView1.Location.X - 30;
                    //this.doubleBufferListView1.Height = this.Height - doubleBufferListView1.Location.Y - 50;
                }
                else
                {
                    this.listView1.Width = this.Width - listView1.Location.X - 30;
                    this.listView1.Height = this.Height - listView1.Location.Y - 50;
                }

                this.txtCommand.Width = this.Width - txtCommand.Location.X - 30;
                this.txtCommand.Height = this.Height - txtCommand.Location.Y - 50;
            }
        }

        private void NotifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            MyCommond.InvokeIfRequired(this, () =>
            {
                this.Show();
                this.ShowInTaskbar = true;
                this.WindowState = FormWindowState.Normal;
            });
        }

        private void NotifyIcon1_MouseMove(object sender, MouseEventArgs e)
        {
            if (msql.CountValue > 0)
            {
                decimal cc = ((decimal)msql.CountValue / (decimal)msql.MaxCount) * (decimal)100;
                MyCommond.InvokeIfRequired(this, () => { notifyIcon1.Text = $"{cc:N}%"; });
            }
        }

        #endregion

        #region Menu

        private void Menu_Form_Mini_CheckedChanged(object sender, EventArgs e)
        {
            hideForm = MenuItem_Mini.Checked;
        }

        private void Menu_Form_Top_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = MenuItem_Top.Checked;
        }

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
            toolTip1.SetToolTip(comboBox1, "選擇Table");
            toolTip1.SetToolTip(comboBox2, "選擇Column");

            MenuItem_Mini.Checked = true;
            MenuItem_Top.Checked = true;
        }

        #endregion

        #region G2 Block event

        DoubleBufferListView doubleBufferListView1 = new DoubleBufferListView();
        DateTime DtStart;
        DateTime DtEnd;

        MySQLite msql;

        Boolean Using_New = true;

        private void G2_UpLoad()
        {
            dateTimePicker1.Value = DateTime.Now.AddDays(-1);
            dateTimePicker2.Value = DateTime.Now;

            textBox3.Text = "1000";

            MrtTxnColumn = (new TxnTransaction()).ExportTitle();
            TxnColumnSplit = MrtTxnColumn.Split(',');
            foreach (var item in TxnColumnSplit) MrtTxnColumn2 += $"${item},";
            MrtTxnColumn2 = MrtTxnColumn2.Substring(0, MrtTxnColumn2.Length - 1);
            TxnTable = new string[]
                 {
                    "Entry", "Exit", "AddValue", "AddValueCancel", "SaleCard", "SaleCardCancel", "SaleTicket", "SaleTicketRefound", "Other",
                 };

            if (UseDoubleListView)
            {
                // 
                // doubleBufferListView1
                // 
                this.doubleBufferListView1.Font = new System.Drawing.Font("微軟雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
                this.doubleBufferListView1.FullRowSelect = true;
                this.doubleBufferListView1.HideSelection = false;
                this.doubleBufferListView1.Location = listView1.Location;
                this.doubleBufferListView1.Name = "doubleBufferListView1";
                this.doubleBufferListView1.Size = listView1.Size;
                this.doubleBufferListView1.TabIndex = 2;
                this.doubleBufferListView1.UseCompatibleStateImageBehavior = false;
                this.doubleBufferListView1.View = System.Windows.Forms.View.Details;

                this.Controls.Add(this.doubleBufferListView1);
            }

            msql = new MySQLite(Operation_Line, DataBaseName);

            System.Threading.Thread thread_0 = new System.Threading.Thread(() => { ComBoxSet(); /*ListViewSet(TxnColumnSplit);*/ });
            thread_0.IsBackground = true;
            thread_0.Start();
        }

        private void G2_StartSet()
        {
            DtStart = Convert.ToDateTime(dateTimePicker1.Value.ToString("yyyy/MM/dd 05:00:00"));
            DtEnd = Convert.ToDateTime(dateTimePicker2.Value.ToString("yyyy/MM/dd 05:00:00"));

            if (Using_New)
            {
                msql.SetDateTime(DtStart, DtEnd);
                msql.LoadLimit = Convert.ToInt16(textBox3.Text);
            }
            else
            {
                MyCommond.CheckFolder(DB_Path + DB_Path_Mrt);
                DB_Mrt_Title = $"{DtStart.ToString("yyyy.MM")}月_";
                DB_Connect = DB_Title + DB_Path + DB_Path_Mrt + DB_Mrt_Title + DB_Mrt;
            }
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("是否確定要創建新的'資料庫'?", "新增", MessageBoxButtons.OKCancel);
            if (result == DialogResult.Cancel) return;
            try
            {
                System.Threading.Thread thread = new System.Threading.Thread(B4_create);
                thread.IsBackground = true;
                thread.Start();
                //threads.Add(new Threads()
                //{
                //    Methon = "Create",
                //    Status = StatusEnum.Start,
                //    Thread = thread,
                //});
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"Error! {ex.Message}");
            }
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("是否確認'載入資料'?", "載入", MessageBoxButtons.OKCancel);
            if (result == DialogResult.Cancel) return;
            try
            {
                System.Threading.Thread thread = new System.Threading.Thread(B5_insert);
                thread.IsBackground = true;
                thread.Start();
                //threads.Add(new Threads()
                //{
                //    Methon = "Insert",
                //    Status = StatusEnum.Start,
                //    Thread = thread,
                //});
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"Error! {ex.Message}");
            }
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            try
            {
                System.Threading.Thread thread = new System.Threading.Thread(B6_read2);
                thread.IsBackground = true;
                thread.Start();
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"Error! {ex.Message}");
            }
        }

        private void B4_create()
        {
            MyCommond.WriteLog(ThisReceive, "準備創建資料庫");
            ItemEnable(false);
            G2_StartSet();
            var checkCreate = false;
            if (Using_New)
            {
                string CreatMrtTxnColumn = "";
                foreach (var item in TxnColumnSplit) CreatMrtTxnColumn += item + " text,";
                CreatMrtTxnColumn = CreatMrtTxnColumn.Substring(0, CreatMrtTxnColumn.Length - 1);

                string[] CommandMsg = new string[TxnTable.Length];
                for( int i = 0; i < TxnTable.Length; i++)
                {
                    CommandMsg[i] = $@"Create table {TxnTable[i]} ({CreatMrtTxnColumn});";
                }

                msql.ReadDt = DtStart;
                msql.ReLoad();
                checkCreate = msql.CreateDataBaseFile(CommandMsg);
            }
            else
            {
                if (MyCommond.CheckFile(DB_Path + DB_Path_Mrt + DB_Mrt_Title + DB_Mrt))
                {
                    MyCommond.WriteLog(ThisReceive, "資料庫已存在");
                    ItemEnable(true);
                    return;
                }

                string CreatMrtTxnColumn = "";
                foreach (var item in TxnColumnSplit) CreatMrtTxnColumn += item + " text,";
                CreatMrtTxnColumn = CreatMrtTxnColumn.Substring(0, CreatMrtTxnColumn.Length - 1);

                using (var connection = new SQLiteConnection(DB_Connect))
                {
                    MyCommond.WriteLog(ThisReceive, "未發現資料庫，開始創建");
                    connection.Open();
                    var command0 = connection.CreateCommand();


                    foreach (var item in TxnTable)
                    {
                        command0.CommandText = $@"Create table {item} ({CreatMrtTxnColumn});";
                        try
                        {
                            command0.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }

            }

            MyCommond.WriteLog(ThisReceive, $"資料庫創建{((checkCreate) ? "完成":"失敗")}");
            GC.Collect();
            ItemEnable(true);
            //threads.Find(x => x.Methon == "Create").Status = StatusEnum.Stop;
        }

        private void B5_insert()
        {
            ItemEnable(false);
            G2_StartSet();

            if (Using_New)
            {
                int DtMax = (DtEnd - DtStart).Days;
                int DtCount = new int();
                for (DateTime dt = DtStart; dt < DtEnd; dt = dt.AddDays(1))
                {
                    if (msql.Stop) { msql.Stop = false; Label_Msg("已中斷匯入", "Red"); ItemEnable(true); GC.Collect(); return; } // 強制停止

                    msql.ReadDt = dt;
                    msql.ReLoad();
                Insert_Create:
                    if (!msql.Check_DB())
                    {
                        string CreatMrtTxnColumn = "";
                        foreach (var item in TxnColumnSplit) CreatMrtTxnColumn += item + " text,";
                        CreatMrtTxnColumn = CreatMrtTxnColumn.Substring(0, CreatMrtTxnColumn.Length - 1);

                        string[] CommandMsg = new string[TxnTable.Length];
                        for (int i = 0; i < TxnTable.Length; i++)
                        {
                            CommandMsg[i] = $@"Create table {TxnTable[i]} ({CreatMrtTxnColumn});";
                        }

                        msql.CreateDataBaseFile(CommandMsg);
                        goto Insert_Create;
                    }
                    ItemEnable(false);
                    MyCommond.InvokeIfRequired(this, () => { label7.Text = $"{++DtCount}/{DtMax}"; });

                    string path = $"./Analyze/history{msql.ReadDt:yyyyMMdd}_nc.csv";
                    string[] sTxnAllData;

                    #region 檢查檔案

                    if (!MyCommond.CheckFile(path))
                    {
                        MyCommond.WriteLog(ThisReceive, "未發檔案");
                        continue;
                    }
                    sTxnAllData = File.ReadAllLines(path);
                    if (sTxnAllData.Count() < 1)
                    {
                        MyCommond.WriteLog(ThisReceive, "未發資料");
                        continue;
                    }

                    #endregion

                    #region 檢查是否為當前營運日的資料

                    int Jump = 1, Jump_C = 0;
                    int Check_T = 0, Check_F = 0, Check_C = 0;
                    string WarnData = "";
                    foreach (var item in sTxnAllData)
                    {
                        if (Jump_C < Jump)
                        {
                            Jump_C++;
                            continue;
                        }
                        var jtem = item.Split(',');
                        if (Convert.ToDateTime(jtem[3]) > msql.DtStart && Convert.ToDateTime(jtem[3]) < msql.DtEnd.AddDays(1))
                        {
                            Check_T++;
                        }
                        else
                        {
                            WarnData = Convert.ToDateTime(jtem[3]).ToString("yyyy/MM/dd");
                            Check_F++;
                        }
                        Check_C++;
                        if (Check_C < 1000) continue;
                        if (((float)Check_T / (float)Check_C) < 0.7)
                        {
                            var Result_Msg = MessageBox.Show($"發現營運時間非當日\n        檢查營運日期:{dt.ToString("yyyy/MM/dd")}\n    被檢查營運日期:{WarnData}\n是否確認繼續載入", "警告", MessageBoxButtons.OKCancel);
                            //var Result_Msg = MessageBox.Show($"發現營運時間非當日，是否確認繼續載入", "警告", MessageBoxButtons.OKCancel);
                            if (Result_Msg == DialogResult.OK)
                            {
                                MyCommond.WriteLog(ThisReceive, "使用者確認繼續執行");
                                break;
                            }
                            else
                            {
                                MyCommond.WriteLog(ThisReceive, "使用者取消執行");
                                ItemEnable(true);
                                return;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    #endregion

                    #region 開始寫入資料

                    System.Threading.Thread thread = new System.Threading.Thread(() =>
                    {
                    //try
                    //{

                    ThreadStart:
                        if (msql.CountValue == -1)
                        {
                            System.Threading.Thread.Sleep(1000);
                            goto ThreadStart;
                        } // 開始進行載入時再進行下一步


                        while (msql.CountValue >= 0)
                        {
                            // 迴圈計算當前進度
                            decimal cc = ((decimal)msql.CountValue / (decimal)msql.MaxCount) * (decimal)100;
                            MyCommond.InvokeIfRequired(this, () => { label3.Text = $"{cc:N}%"; });

                            System.Threading.Thread.Sleep(1000);
                        } // 當值變回 -1 時表示執行結束，跳出迴圈
                        MyCommond.InvokeIfRequired(this, () => { label6.Text = $"{msql.ReadDt:yyyy-MM-dd} 完成"; });
                        //}
                        //catch (Exception ex)
                        //{
                        //    MyCommond.WriteLog(ThisReceive, $"Error! {ex.Message}");
                        //}
                    });
                    thread.IsBackground = true;
                    thread.Start();
                    var CheckInsert = false;
                    msql.SetCommand = "Insert into {0}(" + MrtTxnColumn + ") values ( " + MrtTxnColumn2 + "); ";
                    try
                    {
                        CheckInsert = msql.Insert_(MrtTxnColumn, sTxnAllData, 1);
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"Error! {ex.Message}");
                    }

                    MyCommond.WriteLog(ThisReceive, $"載入{dt.ToString("yyyy/MM/dd")} - {((CheckInsert) ? "成功" : "失敗")}");
                    #endregion

                    MyCommond.MoveFile($"./Analyze/", $"./Analyze/Old/", $"history{msql.ReadDt:yyyyMMdd}_nc.csv");

                }
            }
            else
            {
                MyCommond.WriteLog(ThisReceive, "準備寫入資料");

                if (!MyCommond.CheckFile(DB_Path + DB_Path_Mrt + DB_Mrt_Title + DB_Mrt))
                {
                    if (true)
                    {
                        MyCommond.WriteLog(ThisReceive, "未發現資料庫");
                        B4_create();
                    }
                    else
                    {
                        MyCommond.WriteLog(ThisReceive, "未發現資料庫");
                        ItemEnable(true);
                        return;
                    }
                }

                int DtMax = (DtEnd - DtStart).Days;
                int DtCount = new int();
                //DateTime dt = Convert.ToDateTime(dateTimePicker1.Value.ToString("yyyy/MM/dd 05:00:00"));
                for (DateTime dt = DtStart; dt < DtEnd; dt = dt.AddDays(1))
                {

                    MyCommond.InvokeIfRequired(this, () => { label3.Text = $"0%"; label7.Text = $"{++DtCount}/{DtMax}"; });
                    MyCommond.WriteLog(ThisReceive, $"第{DtCount}天,營運日:{dt.ToString("yyyy/MM/dd")}");

                    string path = $"./Analyze/history{dt:yyyyMMdd}_nc.csv";
                    if (!MyCommond.CheckFile(path))
                    {
                        MyCommond.WriteLog(ThisReceive, "未發檔案");
                        continue;
                    }
                    string[] sTxnAllData = File.ReadAllLines(path);
                    if (sTxnAllData.Count() < 1)
                    {
                        MyCommond.WriteLog(ThisReceive, "未發資料");
                        continue;
                    }
                    int Jump = 1, Jump_C = 0;
                    int Check_T = 0, Check_F = 0, Check_C = 0;
                    foreach (var item in sTxnAllData)
                    {
                        if (Jump_C < Jump)
                        {
                            Jump_C++;
                            continue;
                        }
                        var jtem = item.Split(',');
                        if (Convert.ToDateTime(jtem[3]) > dt && Convert.ToDateTime(jtem[3]) < dt.AddDays(1))
                        {
                            Check_T++;
                        }
                        else
                        {
                            Check_F++;
                        }
                        Check_C++;
                        if (Check_C < 100) continue;
                        if (((float)Check_T / (float)Check_C) < 0.7)
                        {
                            var Result_Msg = MessageBox.Show("發現營運時間非當日，是否確認繼續載入", "警告", MessageBoxButtons.OKCancel);
                            if (Result_Msg == DialogResult.OK)
                            {
                                MyCommond.WriteLog(ThisReceive, "使用者確認繼續執行");
                                break;
                            }
                            else
                            {
                                MyCommond.WriteLog(ThisReceive, "使用者取消執行");
                                ItemEnable(true);
                                return;
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    using (var connection = new SQLiteConnection(DB_Connect))
                    {
                        connection.Open();
                        var command1 = connection.CreateCommand();
                        int Count = new int();

                        foreach (var item in sTxnAllData)
                        {
                            if (Count++ < 1) continue;
                            var jtem = item.Split(',');
                            string thisCommand = "Insert into ";
                            string thisTable = "";
                            switch (Convert.ToInt16(jtem[0]))
                            {
                                case 20:
                                case 151:
                                case 153: thisTable += $"{TxnTable[0]} "; break; // Entry
                                case 1:
                                case 4:
                                case 8:
                                case 13:
                                case 152:
                                case 154: thisTable += $"{TxnTable[1]} "; break; // Exit
                                case 2: thisTable += $"{TxnTable[2]} "; break; // AddValue
                                case 11: thisTable += $"{TxnTable[3]} "; break; // AddValueCancel
                                case 7: thisTable += $"{TxnTable[4]} "; break; // SaleCard
                                case 9: thisTable += $"{TxnTable[5]} "; break; // SaleCardCancel
                                case 190:
                                case 193: thisTable += $"{TxnTable[6]} "; break; // SaleTicket
                                case 191:
                                case 194: thisTable += $"{TxnTable[7]} "; break; // SaleTicketRefound
                                default: thisTable += $"{TxnTable[8]} "; break; // Other
                                                                                //default: continue;
                            }
                            thisCommand += thisTable;
                            thisCommand += $"({MrtTxnColumn}) values ({MrtTxnColumn2});";

                            command1.CommandText = thisCommand;


                            for (int i = 0; i < TxnColumnSplit.Count(); i++)
                            {
                                command1.Parameters.AddWithValue($"${TxnColumnSplit[i]}", jtem[i]);
                            }
                            command1.ExecuteNonQuery();

                            decimal cc = ((decimal)Count / (decimal)sTxnAllData.Count()) * (decimal)100;

                            MyCommond.InvokeIfRequired(this, () => { label3.Text = $"{cc:N}%"; });

                            //MyCommond.WriteLog(ThisReceive, $"{Count,8}/{sTxnAllData.Count()}, {thisTable} Add {item.ToString()}");
                        }
                    }
                }
            }

            MyCommond.WriteLog(ThisReceive, "結束寫入");
            GC.Collect();
            ItemEnable(true);
            //threads.Find(x => x.Methon == "Insert").Status = StatusEnum.Stop;
        }

        List<TxnTransaction> TxnAllData;
        /**
select
from Entry
where 
Txn_TIMESTAMP > '2023-08-14 05:00:00' and
Txn_TIMESTAMP < '2023-08-14 06:00:00'
        */
        private void B6_read()
        {
            ItemEnable(false);
            G2_StartSet();

            if (Using_New)
            {
                string cmd = "";
                if (txtCommand.Text == "")
                {
                    MyCommond.InvokeIfRequired(this, () =>
                    {
                        cmd = $@"
select *
from {comboBox1.SelectedItem.ToString()}
where 
TXN_TIMESTAMP >= '{DtStart.ToString("yyyy-MM-dd hh:mm:ss")}' and 
TXN_TIMESTAMP < '{DtEnd.ToString("yyyy-MM-dd hh:mm:ss")}'
";
                    });
                }
                else
                {
                    cmd = txtCommand.Text;
                }

                Label_Msg($"讀取中");
                msql.SetCommand = cmd;
                System.Threading.Thread thread = new System.Threading.Thread(() =>
                {
                ThreadStart:
                    if (msql.CountValue == -1)
                    {
                        System.Threading.Thread.Sleep(1000);
                        goto ThreadStart;
                    }
                    while (msql.CountValue >= 0)
                    {
                        MyCommond.InvokeIfRequired(this, () => { label6.Text = $"已撈取 {msql.CountValue:#,##0} 筆"; });
                    }
                    GC.Collect();
                });
                thread.IsBackground = true;
                thread.Start();

                try
                {
                    msql.ReadDt = DtStart;
                    msql.ReLoad();
                    var mData = msql.Read();
                    ExportSwitch(mData);
                }
                catch (Exception ex)
                {
                    Label_Msg("ERROR","Red");
                    MyCommond.WriteLog(ThisReceive, $"Error! {ex.Message}");
                }
                
                //ListView_Show(mData);
            }
            else
            {
                if (!true)
                {
                    string mTable_Name = "TxnCode";
                    string[] mTable_Title = new string[] { "Line", "Code", "Name" };
                    string str1 = "", str2 = "";

                    #region Create

                    string CreateCommand = $@"create table {mTable_Name} (Id integer";
                    foreach (var item in mTable_Title)
                    {
                        CreateCommand += $", {item} text";
                        str1 += $"{item},";
                        str2 += $"${item},";
                    }
                    CreateCommand += ", primary key(Id autoincrement));";

                    msql.SetCommand = CreateCommand;
                    msql.CreateDataBaseFile(null);

                    #endregion

                    #region Insert

                    //foreach (var s in mTable_Title) { str1 += $"{s},"; str2 += $"${s},"; }
                    str1 = str1.Substring(0, str1.Length - 1);
                    str2 = str2.Substring(0, str2.Length - 1);
                    msql.SetCommand = $@"insert into {mTable_Name} ({str1}) values ({str2});";

                    string[] mData_0 = new string[]
                    {
                    "Mrt,20,Entry", "Mrt,151,Entry",
                    "Mrt,153,Entry", "Mrt,1,Exit", "Mrt,4,Exit", "Mrt,8,Exit", "Mrt,13,Exit", "Mrt,152,Exit", "Mrt,154,Exit" ,
                    "Mrt,2,AddValue", "Mrt,11,AddValueCancel", "Mrt,7,SaleCard", "Mrt,9,SaleCardCancel",
                    "Mrt,190,SaleTicket", "Mrt,193,SaleTicket",
                    "Mrt,191,SaleTicketRefound", "Mrt,194,SaleTicketRefound"
                    };
                    System.Threading.Thread thread = new System.Threading.Thread(() =>
                    {
                    ThreadStart:
                        if (msql.CountValue == -1)
                        {
                            System.Threading.Thread.Sleep(1000);
                            goto ThreadStart;
                        }
                        while (msql.CountValue >= 0)
                        {
                            decimal cc = ((decimal)msql.CountValue / (decimal)msql.MaxCount) * (decimal)100;
                            MyCommond.InvokeIfRequired(this, () => { label3.Text = $"{cc:N}%"; });

                            System.Threading.Thread.Sleep(1000);
                        }
                    });
                    thread.IsBackground = true;
                    thread.Start();
                    //msql.Insert(str1, mData_0, 0);

                    #endregion

                    #region Read

                    msql.SetCommand = $"select * from {mTable_Name}";
                    var mData = msql.Read();

                    List<TableNum> list = new List<TableNum>();
                    foreach (var item in mData)
                    {
                        //TableNum tr = MyCommond.T2<TableNum>(ThisReceive, (new TableNum()).ExportTitle(), item);
                        //list.Add(tr);
                    }

                    #endregion

                    //msql.test_1();
                    ItemEnable(true);
                    return;
                } //test

                if (!MyCommond.CheckFile(DB_Path + DB_Path_Mrt + DB_Mrt_Title + DB_Mrt))
                {
                    MyCommond.WriteLog(ThisReceive, "未發現資料庫");
                    Label_Msg("No DB", "Red");
                    ItemEnable(true);
                    return;
                }
                try
                {
                    //List<ListViewItem> vs = new List<ListViewItem>();
                    using (var connection = new SQLiteConnection(DB_Connect))
                    {
                        connection.Open();
                        var command = connection.CreateCommand();
                        string cmd = "";
                        if (txtCommand.Text == "")
                        {
                            MyCommond.InvokeIfRequired(this, () =>
                            {
                                cmd = $@"
select *
from {comboBox1.SelectedItem.ToString()}
where 
TXN_TIMESTAMP >= '{DtStart.ToString("yyyy-MM-dd hh:mm:ss")}' and 
TXN_TIMESTAMP < '{DtEnd.ToString("yyyy-MM-dd hh:mm:ss")}'
";
                            });
                        }
                        else
                        {
                            cmd = txtCommand.Text;
                        }
                        command.CommandText = cmd;

                        using (var reader = command.ExecuteReader())
                        {
                            int Count = new int();
                            TxnAllData = new List<TxnTransaction>();
                            while (reader.Read())
                            {
                                var ColumnCount = reader.FieldCount;
                                if (!(ColumnCount > 0)) continue;

                                string[] toS = new string[ColumnCount];
                                for (int i = 0; i < ColumnCount; i++) { toS[i] = reader.GetString(i); }

                                Label_Msg($"已讀取{++Count:#,##0}筆");

                                //TxnAllData.Add(MyCommond.T2<TxnTransaction>(ThisReceive,(new TxnTransaction()).ExportTitle(), toS));
                            }
                        }
                    }
                    ListView_Show<TxnTransaction>(TxnAllData);
                }
                catch (Exception ex)
                {
                    Label_Msg($"ERROR", "Red");
                    MessageBox.Show(ex.Message);
                }
            }

            MyCommond.WriteLog(ThisReceive, "結束讀取");
            GC.Collect();
            ItemEnable(true);
            //threads.Find(x => x.Methon == "Read").Status = StatusEnum.Stop;
        }
        
        private void ExportSwitch(List<string[]> Data)
        {
            if (Data == null || Data.Count < 1)
            {
                Label_Msg("No Data");
                return;
            }
            int Count = new int();
            int ViewCountMax = 0;
            MyCommond.InvokeIfRequired(this, () =>
            {
                ViewCountMax = Convert.ToInt16(textBox3.Text);
            });

            string Export_FileName = MyCommond.CheckAndReName(MyCommond.Path_Report, $"{msql.DtStart.ToString("yyyyMMdd")}_{msql.DB_Line}_Export", "csv");
            string Export_Path = $"{MyCommond.Path_Report}{Export_FileName}";
            if (radioButton2.Checked)
            {
                MyCommond.ExportData(Export_Path, Encoding.UTF8, (new TxnTransaction()).ExportTitle());
            }

            if (UseDoubleListView)
            {
                if (Count == 0) { MyCommond.InvokeIfRequired(this, () => doubleBufferListView1.Items.Clear()); Label_Msg("No Data"); }
            }
            else
            {
                if (Count == 0) { MyCommond.InvokeIfRequired(this, () => listView1.Items.Clear()); Label_Msg("No Data"); }
            }

            foreach (var item in Data)
            {
                int MaxCount = ViewCountMax;
                if (msql.Stop) { msql.Stop = false; GC.Collect(); return; } // 強制停止

                if (radioButton1.Checked)
                {
                    if (Count > MaxCount)
                    {
                        Label_Msg($"Over {MaxCount:#,##0} / {Data.Count():#,##0}", "Red");
                        break;
                    } // 輸出限制筆數

                    if (UseDoubleListView)
                    {
                        MyCommond.InvokeIfRequired(this, () => doubleBufferListView1.Items.Add(new ListViewItem(item)));
                    }
                    else
                    {
                        MyCommond.InvokeIfRequired(this, () => listView1.Items.Add(new ListViewItem(item)));
                    }
                }
                else if (radioButton2.Checked)
                {
                    string ss = "";
                    foreach (var jtem in item) ss += jtem + ",";
                    ss = ss.Substring(0, ss.Length - 1);
                    MyCommond.ExportData(Export_Path, Encoding.UTF8, ss);
                }

                Label_Msg($"{++Count:#,##0} / {Data.Count():#,##0}");
            }

        }

        private void B6_read2()
        {
            ItemEnable(false);
            G2_StartSet();
            int CountValue = -1;
            string cmd = "";

            

            MyCommond.InvokeIfRequired(this, () =>
            {
                if (txtCommand.Text == "")
                {
                    cmd = $@"
select *
from {comboBox1.SelectedItem.ToString()}
where 
TXN_TIMESTAMP >= '{DtStart.ToString("yyyy-MM-dd hh:mm:ss")}' and 
TXN_TIMESTAMP < '{DtEnd.ToString("yyyy-MM-dd hh:mm:ss")}'
";
                }
                else
                {
                    cmd = txtCommand.Text;
                }
            });

            Label_Msg($"讀取中");
            msql.ReadDt = DtStart;
            msql.ReLoad();
            msql.Locol_Open();
            msql.DB_command.CommandText = cmd;

            try
            {
                System.Threading.Thread thread = new System.Threading.Thread(() =>
                {
                ThreadStart:
                    if (CountValue == -1)
                    {
                        System.Threading.Thread.Sleep(1000);
                        goto ThreadStart;
                    }
                    while (CountValue >= 0)
                    {
                        MyCommond.InvokeIfRequired(this, () => { label6.Text = $"已撈取 {CountValue:#,##0} 筆"; });
                    }
                    GC.Collect();
                });
                thread.IsBackground = true;
                thread.Start();

                using (var reader = msql.DB_command.ExecuteReader())
                {
                    CountValue = 0;
                    bool ExportTitle_B = true;
                    string[] ReadTitle = null;
                    List<string[]> ReadData = new List<string[]>();
                    while (reader.Read())
                    {
                        if (msql.Stop) { msql.Stop = false; break; } // 強制停止
                        if (CountValue > 3000000) { break; } // 強制停止

                        var ColumnCount = reader.FieldCount;
                        if (!(ColumnCount > 0)) continue;

                        if (ExportTitle_B)
                        {
                            ReadTitle = new string[ColumnCount];
                            for (int i = 0; i < ColumnCount; i++) ReadTitle[i] = $"{reader.GetName(i)}";

                            ExportTitle_B = false;
                        }

                        CountValue++;
                        string[] ArrayString = new string[ColumnCount];
                        for (int i = 0; i < ColumnCount; i++)
                        {
                            switch (reader.GetFieldType(i).Name)
                            {
                                case "String"  : ArrayString[i] = Convert.ToString(reader.GetString(i)  );    break;
                                case "DateTime": ArrayString[i] = Convert.ToString(reader.GetDateTime(i));    break;
                                case "Int16"   : ArrayString[i] = Convert.ToString(reader.GetInt16(i)   );    break;
                                case "Int32"   : ArrayString[i] = Convert.ToString(reader.GetInt32(i)   );    break;
                                case "Int64"   : ArrayString[i] = Convert.ToString(reader.GetInt64(i)   );    break;
                                case "Byte"    : ArrayString[i] = Convert.ToString(reader.GetByte(i)    );    break;
                                case "Char"    : ArrayString[i] = Convert.ToString(reader.GetChar(i)    );    break;
                                case "Boolean" : ArrayString[i] = Convert.ToString(reader.GetBoolean(i) );    break;
                                case "Double"  : ArrayString[i] = Convert.ToString(reader.GetDouble(i)  );    break;
                                case "Float"   : ArrayString[i] = Convert.ToString(reader.GetFloat(i)   );    break;
                                case "Guid"    : ArrayString[i] = Convert.ToString(reader.GetGuid(i)    );    break;
                                case "Object"  : ArrayString[i] = Convert.ToString(reader.GetValue(i)   );    break;
                                default: MyCommond.WriteLog(ThisReceive, reader.GetFieldType(i).Name); break;
                            }
                        }
                        ReadData.Add(ArrayString);
                    }
                    CountValue = -1;

                    ExportSwitch2(ReadTitle, ReadData);
                }
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"Error! {ex.Message}");
            }
            finally
            {
                msql.Locol_Close();
            }

            MyCommond.WriteLog(ThisReceive, "結束讀取");
            GC.Collect();
            ItemEnable(true);
        }

        private void ExportSwitch2(string[] Title, List<string[]> Data)
        {
            if (radioButton1.Checked)
            {
                Export2Form(Title, Data);
            }
            else if (radioButton2.Checked)
            {
                Export2Csv(Title, Data);
            }
        }

        private void Export2Form(string[] Title, List<string[]> Data)
        {

            Label_Msg("No Data");
            if (Data == null || Data.Count < 1) { return; }
            int Count = new int();
            int ViewCountMax = 0;
            MyCommond.InvokeIfRequired(this, () => { ViewCountMax = Convert.ToInt16(textBox3.Text); });

            ListViewSet(Title);

            foreach (var item in Data)
            {
                int MaxCount = ViewCountMax;
                if (msql.Stop) { msql.Stop = false; GC.Collect(); return; } // 強制停止
                if (Count > MaxCount) { Label_Msg($"Over {MaxCount:#,##0} / {Data.Count():#,##0}", "Red"); break; } // 輸出限制筆數

                if (UseDoubleListView) { MyCommond.InvokeIfRequired(this, () => doubleBufferListView1.Items.Add(new ListViewItem(item))); }
                else { MyCommond.InvokeIfRequired(this, () => listView1.Items.Add(new ListViewItem(item))); }

                Label_Msg($"{++Count:#,##0} / {Data.Count():#,##0}");
            }
        }

        private void Export2Csv(string[] Title, List<string[]> Data)
        {
            Label_Msg("No Data");
            if (Data == null || Data.Count < 1) { return; }
            int Count = new int();
            int ViewCountMax = 0;
            MyCommond.InvokeIfRequired(this, () => { ViewCountMax = Convert.ToInt16(textBox3.Text); });

            string Export_FileName = MyCommond.CheckAndReName(MyCommond.Path_Report, $"{DateTime.Now:yyyy-MM-dd}_SearchData", "csv");
            string Export_Path = $"{MyCommond.Path_Report}{Export_FileName}";

            if (UseDoubleListView) { if (Count == 0) { MyCommond.InvokeIfRequired(this, () => doubleBufferListView1.Items.Clear()); } }
            else { if (Count == 0) { MyCommond.InvokeIfRequired(this, () => listView1.Items.Clear()); } }

            string ExportTitle = "";
            for (int i = 0; i < Title.Length; i++) ExportTitle += $"{Title[i]},";
            ExportTitle = ExportTitle.Substring(0, ExportTitle.Length - 1);
            MyCommond.ExportData(Export_Path, Encoding.UTF8, ExportTitle);

            foreach (var item in Data)
            {
                if (msql.Stop) { msql.Stop = false; GC.Collect(); return; } // 強制停止
                string ExportString = "";
                for (int i = 0; i < item.Length; i++) ExportString += $"{item[i]},";
                ExportString = ExportString.Substring(0, ExportString.Length - 1);

                MyCommond.ExportData(Export_Path, Encoding.UTF8, ExportString);

                Label_Msg($"{++Count:#,##0} / {Data.Count():#,##0}");
            }
        }

        private void ListView_Show(List<string[]> Data) 
        {
            int Count = new int();
            int ViewCountMax = 0;
            MyCommond.InvokeIfRequired(this, () =>
            {
                ViewCountMax = Convert.ToInt16(textBox3.Text);
            });

            foreach (var item in Data)
            {
                if (UseDoubleListView)
                {
                    if (Count == 0) { MyCommond.InvokeIfRequired(this, () => doubleBufferListView1.Items.Clear()); Label_Msg("Msg"); }
                }
                else
                {
                    if (Count == 0) { MyCommond.InvokeIfRequired(this, () => listView1.Items.Clear()); Label_Msg("Msg"); }
                }

                int MaxCount = ViewCountMax; if (Count > MaxCount) { Label_Msg($"Over {MaxCount:#,##0} / {Data.Count():#,##0}", "Red"); break; }

                if (UseDoubleListView)
                {
                    MyCommond.InvokeIfRequired(this, () => doubleBufferListView1.Items.Add(new ListViewItem(item)));
                }
                else
                {
                    MyCommond.InvokeIfRequired(this, () => listView1.Items.Add(new ListViewItem(item)));
                }
                Label_Msg($"{++Count:#,##0} / {Data.Count():#,##0}");
            }

        }

        private void ListView_Show<TResult>(List<TResult> Data) where TResult :new()
        {
            int Count = new int();
            foreach (var item in Data)
            {
                if (UseDoubleListView)
                {
                    if (Count == 0) { MyCommond.InvokeIfRequired(this, () => doubleBufferListView1.Items.Clear()); Label_Msg("Msg"); }
                }
                else
                {
                    if (Count == 0) { MyCommond.InvokeIfRequired(this, () => listView1.Items.Clear()); Label_Msg("Msg"); }
                }

                if (Count > msql.LoadLimit) { Label_Msg($"Over {msql.LoadLimit:#,##0} / {Data.Count():#,##0}", "Red"); break; }

                string[] toS = item.ToString().Split(',');
                if (UseDoubleListView)
                {
                    MyCommond.InvokeIfRequired(this, () => doubleBufferListView1.Items.Add(new ListViewItem(toS)));
                }
                else
                {
                    MyCommond.InvokeIfRequired(this, () => listView1.Items.Add(new ListViewItem(toS)));
                }
                Label_Msg($"{++Count:#,##0} / {Data.Count():#,##0}");
            }
            
        }

        Color defaultColor = Color.Black;
        private void Label_Msg(string msg, string color = "Black")
        {
            MyCommond.InvokeIfRequired(this, () => {
                if (color == "Red") label6.ForeColor = Color.Red;
                else label6.ForeColor = Color.Black;
                label6.Text = msg;
            });
        }

        private void ComBoxSet()
        {
            MyCommond.InvokeIfRequired(this, () => 
            {
                comboBox1.Items.Clear();
                for (int i = 0; i < TxnTable.Length; i++)
                {
                    comboBox1.Items.Add(TxnTable[i]);
                }
                comboBox1.SelectedIndex = 0;

                string[] Txn_Column = (new TxnTransaction()).ExportTitle().Split(',');
                comboBox2.Items.Clear();
                for(int i = 0; i < Txn_Column.Length; i++)
                {
                    comboBox2.Items.Add(Txn_Column[i]);
                }
                comboBox2.SelectedIndex = 0;
            });
        }

        private void ListViewSet(string[] Title)
        {
            MyCommond.InvokeIfRequired(this, () =>
            {
                if (UseDoubleListView) { doubleBufferListView1.Columns.Clear(); doubleBufferListView1.Items.Clear(); }
                else { listView1.Columns.Clear(); listView1.Items.Clear(); }

                foreach (var item in Title)
                {
                    if (UseDoubleListView)
                    {
                        doubleBufferListView1.Columns.Add(item, (item.Length * 10), HorizontalAlignment.Left);
                    }
                    else
                    {
                        listView1.Columns.Add(item, (item.Length * 10), HorizontalAlignment.Left);
                    }
                }

                //if (UseDoubleListView) doubleBufferListView1.Items.Add(new ListViewItem());
                //else listView1.Items.Add(new ListViewItem());
            });
        }

        private void ItemEnable(bool TF)
        {
            MyCommond.InvokeIfRequired(this, () => 
            { 
                groupBox2.Enabled = TF; 
                if(TF) StopAct.Enabled = TF;
            });
        }

        #endregion

        #region G1 Block event
        //UsuallyCommond.MySQLite sq = new UsuallyCommond.MySQLite(DateTime.Now,"Mrt");
        MySQLite sq = new MySQLite();
        private void DB_Create_Click(object sender, EventArgs e)
        {
            CreateDatabaseFile();
            //            sq.CreateDataBaseFile(@"
            //CREATE TABLE users (
            //    id INTEGER, 
            //    user_name TEXT NOT NULL UNIQUE, 
            //    user_name2 TEXT, 
            //    user_name3 TEXT, 
            //    PRIMARY KEY(id AUTOINCREMENT));");
        }

        private void DB_Insert_Click(object sender, EventArgs e)
        {
            Insert(textBox2.Text);
        }

        private void DB_UpDate_Click(object sender, EventArgs e)
        {
            Update(Convert.ToInt16(textBox1.Text), textBox2.Text);
        }

        private void DB_Delete_Click(object sender, EventArgs e)
        {
            Delete(Convert.ToInt16(textBox1.Text));
        }

        private void DB_Read_Click(object sender, EventArgs e)
        {
            Read();
        }

        /// <summary>
        /// 資料庫連接字串
        /// </summary>
        //private string _connectionString = "Data Source=db.db;";
        private string _connectionString = "Data Source=db.db;";

        /// <summary>
        /// 建立資料庫檔案
        /// </summary>
        private void CreateDatabaseFile()
        {
            Debug.WriteLine("CreateDatabaseFile");
            if (File.Exists("db.db")) return;
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText =
                    @"CREATE TABLE users (
            	        id	INTEGER,
            	        user_name	TEXT NOT NULL UNIQUE,
            	        user_name2	TEXT ,
            	        user_name3	TEXT ,
                        PRIMARY KEY(id AUTOINCREMENT)
                    );";
                command.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 新增
        /// </summary>
        /// <param name="userName"></param>
        private void Insert(string userName)
        {
            Debug.WriteLine("Insert");
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText =
                    @"  INSERT INTO users (user_name)
                        values ($userName);
                        select last_insert_rowid();";
                command.Parameters.AddWithValue("$userName", userName);
                int id = Convert.ToInt32((object)command.ExecuteScalar());
                Debug.WriteLine($"\tid = {id}, userName = {userName}");
            }
        }

        /// <summary>
        /// 修改
        /// </summary>
        /// <param name="id"></param>
        /// <param name="userName"></param>
        private void Update(int id, string userName)
        {
            Debug.WriteLine("Update");
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText =
                    @"  UPDATE users
                        SET user_name= $userName
                        WHERE id = $id;";
                command.Parameters.AddWithValue("$userName", userName);
                command.Parameters.AddWithValue("id", id);
                command.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="id"></param>
        private void Delete(int id)
        {
            Debug.WriteLine("Delete");
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                command.CommandText =
                    @"  delete from users where id = $id;";
                command.Parameters.AddWithValue("id", id);
                command.ExecuteNonQuery();
            }
        }

        /// <summary>
        /// 查詢、讀取
        /// </summary>
        private void Read()
        {
            Debug.WriteLine("Read");
            using (var connection = new SQLiteConnection(_connectionString))
            {
                connection.Open();

                var command = connection.CreateCommand();
                command.CommandText = @"  select * from users";
                
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var C1 = reader.StepCount;
                        var C2 = reader.FieldCount;
                        var C3 = reader.VisibleFieldCount;
                        var C4 = reader.GetType();
                        var C5 = C4.GetProperties()[0].PropertyType.Name;
                        var C6 = C4.GetProperties()[1].PropertyType.Name;
                        var C7 = C4.GetProperties()[2].PropertyType.Name;
                        var C8 = C4.GetProperties()[3].PropertyType.Name;
                        var C9 = C4.GetProperties()[4].PropertyType.Name;


                        var ColumnCount = reader.FieldCount;
                        var toS = "";
                        for (int i = 0; i < ColumnCount; i++)
                        {
                            var a1 = reader.GetFieldType(i).Name;
                            try
                            {
                                switch (reader.GetFieldType(i).Name)
                                {
                                    case "String":      /*if (reader.GetString(i)     == "")   continue;*/ toS += $"{Convert.ToString(reader.GetString(i))},";     break;
                                    case "DateTime":    /*if (reader.GetDateTime(i)   == null) continue;*/ toS += $"{Convert.ToString(reader.GetDateTime(i))},";   break;
                                    case "Int16":       /*if (reader.GetInt16(i)      == null) continue;*/ toS += $"{Convert.ToString(reader.GetInt16(i))},";      break;
                                    case "Int32":       /*if (reader.GetInt32(i)      == null) continue;*/ toS += $"{Convert.ToString(reader.GetInt32(i))},";      break;
                                    case "Int64":       /*if (reader.GetInt64(i)      == null) continue;*/ toS += $"{Convert.ToString(reader.GetInt64(i))},";      break;
                                    case "Byte":        /*if (reader.GetByte(i)       == null) continue;*/ toS += $"{Convert.ToString(reader.GetByte(i))},";       break;
                                    case "Char":        /*if (reader.GetChar(i)       == null) continue;*/ toS += $"{Convert.ToString(reader.GetChar(i))},";       break;
                                    case "Boolean":     /*if (reader.GetBoolean(i)    == null) continue;*/ toS += $"{Convert.ToString(reader.GetBoolean(i))},";    break;
                                    case "Double":      /*if (reader.GetDouble(i)     == null) continue;*/ toS += $"{Convert.ToString(reader.GetDouble(i))},";     break;
                                    case "Float":       /*if (reader.GetFloat(i)      == null) continue;*/ toS += $"{Convert.ToString(reader.GetFloat(i))},";      break;
                                    case "Guid":        /*if (reader.GetGuid(i)       == null) continue;*/ toS += $"{Convert.ToString(reader.GetGuid(i))},";       break;
                                    case "Object":      /*if (reader.GetValue(i)      == null) continue;*/ toS += $"{Convert.ToString(reader.GetValue(i))},";      break;
                                    default: MyCommond.WriteLog(ThisReceive, $"{reader.GetFieldType(i).Name}"); break;
                                }



                            }
                            catch (Exception ex)
                            {

                            }
                            
                        }


                        var id = reader.GetInt32(0);
                        var name = reader.GetString(1);
                        
                        Debug.WriteLine($"\t{C1} id = {id}, name = {name} => toS = {toS}");
                    }
                }
            }
        }

        #endregion

        private void Button1_Click(object sender, EventArgs e)
        {
            System.Threading.Thread t = new System.Threading.Thread(() =>
            {
                testCatch_4_3.Main(null);
            });
            t.IsBackground = true;
            t.Start();
            return;
            new testCatch_2().GetWhatMac();
            return;
            new testCatch_32().GetComputerInfo();
            return;
            string FILE_NAME = "test.txt";

            if (File.Exists(FILE_NAME))
            {
                Console.WriteLine($"{FILE_NAME} already exists!");
                return;
            }

            using (FileStream fs = new FileStream(FILE_NAME, FileMode.CreateNew))
            {
                using (BinaryWriter w = new BinaryWriter(fs))
                {
                    for (int i = 0; i < 11; i++)
                    {
                        w.Write(i);
                    }
                }
            }

            using (FileStream fs = new FileStream(FILE_NAME, FileMode.Open, FileAccess.Read))
            {
                using (BinaryReader r = new BinaryReader(fs))
                {
                    for (int i = 0; i < 11; i++)
                    {
                        Console.WriteLine(r.ReadInt32());
                    }
                }
            }

            return;
            DirSearch(@"C:\Users\1012M0188\Desktop\環狀資料庫\Log\2023.11.07\");

            return;
            System.Threading.Thread c = new System.Threading.Thread(() =>
            {
                var c_a = System.IO.Directory.GetCurrentDirectory();
                var c_b = System.IO.Directory.GetFiles(c_a);

                foreach (var item in c_b)
                {
                    string path = item;
                    var p = path.Split('.');
                    if (p[p.Length - 1] != "exe") continue;
                    // Get a list of all the processes currently running on the system
                    Process[] processes = Process.GetProcesses();

                    // Iterate through the list of processes
                    foreach (Process process in processes)
                    {
                        // Check if the process path matches the one we're looking for
                        if (process.MainModule.FileName == path)
                        {
                            // If it does, print the process ID to the console
                            Console.WriteLine("Process ID: {0}", process.Id);
                        }
                    }
                }



            });
            c.IsBackground = true;
            c.Start();

            return;
            System.Threading.Thread b = new System.Threading.Thread(() => UsingProcess("OUTLOOK"));
            b.IsBackground = true;
            b.Start();

            return;
            System.Threading.Thread a = new System.Threading.Thread(() => test());
            a.IsBackground = true;
            a.Start();
        }

        private void Action_Stop_Click(object sender, EventArgs e)
        {
            //if (!msql.Stop) return;
            StopAct.Enabled = false;
            Action_Stop();
        }

        private void Action_Stop()
        {
            Label_Msg("等待停止中", "Red");
            msql.Stop = true;
        }

        #region a

        /// <summary>
        /// 查詢資料夾目錄下所有檔案名稱(無法找沒有資料夾的資料夾..)
        /// https://devil0827.pixnet.net/blog/post/327334731-c%23-%E6%9F%A5%E8%A9%A2%E8%B3%87%E6%96%99%E5%A4%BE%E7%9B%AE%E9%8C%84%E4%B8%8B%E6%89%80%E6%9C%89%E6%AA%94%E6%A1%88%E6%B8%85%E5%96%AE
        /// </summary>
        /// <param name="sDir"></param>
        public void DirSearch(string sDir)
        {
            try
            {
                //先找出所有目錄 
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    //先針對目前目路的檔案做處理 
                    foreach (string f in Directory.GetFiles(d))
                    {
                        Console.WriteLine(f);
                    }
                    //此目錄處理完再針對每個子目錄做處理 
                    DirSearch(d);
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
        }

        /* Form 查詢資料夾內容清單
        String[] FileCollection;

        String FilePath = "F:\\Uploads";

        FileInfo theFileInfo;


        FileCollection = Directory.GetFiles(FilePath, "*.txt");

        for(int i = 0 ; i < FileCollection.Length ; i++)

        {

            theFileInfo = new FileInfo(FileCollection[i]);

            Response.Write(theFileInfo.Name.ToString());

        }
        */

        #endregion

        #region 取得在是窗外點選的檔案或是資料夾的路徑

        /// <summary>
        /// 取得在是窗外點選的檔案或是資料夾的路徑。
        /// https://kknews.cc/zh-tw/code/n22y5xq.html
        /// </summary>
        private void ssssss()
        {
            string filename;
            ArrayList selected = new ArrayList();
            foreach (SHDocVw.InternetExplorer window in new SHDocVw.ShellWindows())
            {
                filename = Path.GetFileNameWithoutExtension(window.FullName).ToLower();
                if (filename.ToLowerInvariant() == "explorer" || filename.ToLowerInvariant() == "資源管理器")
                {
                    Shell32.FolderItems items = ((Shell32.IShellFolderViewDual2)window.Document).SelectedItems();
                    foreach (Shell32.FolderItem item in items)
                    {
                        //selected.Add(item.Path);
                        MessageBox.Show(item.Path);
                    }
                }
            }
        }

        #endregion

        public List<string> FindFile(string sSourcePath)
        {
            List<string> list = new List<string>();

            //在指定目錄及子目錄下查找文件,在list中列出子目錄及文件
            DirectoryInfo Dir = new DirectoryInfo(sSourcePath);
            DirectoryInfo[] DirSub = Dir.GetDirectories();
            if (DirSub.Length <= 0)
            {
                foreach (FileInfo f in Dir.GetFiles("*.*", SearchOption.TopDirectoryOnly)) //查找文件
                {
                    //listBox1.Items.Add(Dir+f.ToString()); //listBox1中填加文件名
                    list.Add(Dir + @"\" + f.ToString());
                }
            }

            int t = 1;
            foreach (DirectoryInfo d in DirSub)//查找子目錄 
            {
                FindFile(Dir + @"\" + d.ToString());
                list.Add(Dir + @"\" + d.ToString());
                if (t == 1)
                {
                    foreach (FileInfo f in Dir.GetFiles("*.*", SearchOption.TopDirectoryOnly)) //查找文件
                    {
                        list.Add(Dir + @"\" + f.ToString());
                    }
                    t = t + 1;
                }
            }
            return list;
        }

        public List<string> FindFile2(string sSourcePath)
        {
            List<String> list = new List<string>();
            //遍歷文件夾
            DirectoryInfo theFolder = new DirectoryInfo(sSourcePath);
            FileInfo[] thefileInfo = theFolder.GetFiles("*.*", SearchOption.TopDirectoryOnly);
            foreach (FileInfo NextFile in thefileInfo)  //遍歷文件
                list.Add(NextFile.FullName);

            //遍歷子文件夾
            DirectoryInfo[] dirInfo = theFolder.GetDirectories();
            foreach (DirectoryInfo NextFolder in dirInfo)
            {
                //list.Add(NextFolder.ToString());
                FileInfo[] fileInfo = NextFolder.GetFiles("*.*", SearchOption.AllDirectories);
                foreach (FileInfo NextFile in fileInfo)  //遍歷文件
                    list.Add(NextFile.FullName);
            }
            return list;
        }

    }

    public class TxnTransaction
    {
        #region Properties

        public string CARD_TXN_TYPE_ID { get; set; }
        public string CARD_TXN_SUBTYPE_ID { get; set; }
        public string DEV_ID { get; set; }
        public string TXN_TIMESTAMP { get; set; }
        public string CARD_PHYSICAL_ID { get; set; }
        public string ISSUER_ID { get; set; }
        public string CARD_TXN_SEQ_NO { get; set; }
        public string TXN_AMT { get; set; }
        public string ELECTRONIC_VALUE { get; set; }
        public string SVCE_LOC_ID { get; set; }
        public string PROCESSING_DATE { get; set; }
        public string BUSINESS_DATE { get; set; }
        public string ENTRY_LOC_ID { get; set; }
        public string XFER_CODE { get; set; }
        public string XFER_DISC { get; set; }
        public string PERSONAL_DISC { get; set; }
        public string PENALTY { get; set; }
        public string LOYALTY_COUNTER { get; set; }
        public string LOYALTY_POINTS { get; set; }
        public string FARE_PRODUCT_TYPE_ID { get; set; }
        public string ENTRY_DATETIME { get; set; }
        public string AREA_CODE { get; set; }
        public string TICKET_SUBTYPE_ID { get; set; }
        public string XFER_DISC_BUSTOMRT { get; set; }
        public string USER_PROFILE { get; set; }
        public string FIRST_UTILISATION_DATE { get; set; }
        public string UP_UTILISATION_DATE { get; set; }
        public string LAST_UTILISATION_DATE { get; set; }

        #endregion

        #region Methods

        public override string ToString() => $"" +
                $"{CARD_TXN_TYPE_ID}," +
                $"{CARD_TXN_SUBTYPE_ID}," +
                $"{DEV_ID}," +
                $"{TXN_TIMESTAMP}," +
                $"{CARD_PHYSICAL_ID}," +
                $"{ISSUER_ID}," +
                $"{CARD_TXN_SEQ_NO}," +
                $"{TXN_AMT}," +
                $"{ELECTRONIC_VALUE}," +
                $"{SVCE_LOC_ID}," +
                $"{PROCESSING_DATE}," +
                $"{BUSINESS_DATE}," +
                $"{ENTRY_LOC_ID}," +
                $"{XFER_CODE}," +
                $"{XFER_DISC}," +
                $"{PERSONAL_DISC}," +
                $"{PENALTY}," +
                $"{LOYALTY_COUNTER}," +
                $"{LOYALTY_POINTS}," +
                $"{FARE_PRODUCT_TYPE_ID}," +
                $"{ENTRY_DATETIME}," +
                $"{AREA_CODE}," +
                $"{TICKET_SUBTYPE_ID}," +
                $"{XFER_DISC_BUSTOMRT}," +
                $"{USER_PROFILE}," +
                $"{FIRST_UTILISATION_DATE}," +
                $"{UP_UTILISATION_DATE}," +
                $"{LAST_UTILISATION_DATE}";

        public string ExportTitle() => $"CARD_TXN_TYPE_ID," +
            $"CARD_TXN_SUBTYPE_ID," +
            $"DEV_ID,TXN_TIMESTAMP," +
            $"CARD_PHYSICAL_ID," +
            $"ISSUER_ID," +
            $"CARD_TXN_SEQ_NO," +
            $"TXN_AMT," +
            $"ELECTRONIC_VALUE," +
            $"SVCE_LOC_ID," +
            $"PROCESSING_DATE," +
            $"BUSINESS_DATE," +
            $"ENTRY_LOC_ID," +
            $"XFER_CODE," +
            $"XFER_DISC," +
            $"PERSONAL_DISC," +
            $"PENALTY," +
            $"LOYALTY_COUNTER," +
            $"LOYALTY_POINTS," +
            $"FARE_PRODUCT_TYPE_ID," +
            $"ENTRY_DATETIME," +
            $"AREA_CODE," +
            $"TICKET_SUBTYPE_ID," +
            $"XFER_DISC_BUSTOMRT," +
            $"USER_PROFILE," +
            $"FIRST_UTILISATION_DATE," +
            $"UP_UTILISATION_DATE," +
            $"LAST_UTILISATION_DATE";

        #endregion
    }

    public class DoubleBufferListView : ListView
    {
        public DoubleBufferListView()
        {
            SetStyle(ControlStyles.DoubleBuffer |
              ControlStyles.OptimizedDoubleBuffer |
              ControlStyles.AllPaintingInWmPaint, true);
            UpdateStyles();
        }
    }

    internal class TableNum
    {
        public string Id { get; set; }
        public string Line { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }

        public string ExportTitle()
        {
            return "Id,Line,Code,Name";
        }
    }

    /*
    //https://topic.alibabacloud.com/tc/a/c--web-page-automatic-logon-for-network-programming-i-use-the-webbrower-control-to-simulate-logon_1_31_31852121.html
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    //新添加命名空間
    using System.Net;
    using System.IO;
    namespace HttpWeb
    {    
        public partial class Form1 : Form    
        {        
            public Form1()        
            {            
                InitializeComponent();        
            }        
            //雙擊"瀏覽"添加Click事件
            private void button1_Click(object sender, EventArgs e)       
            {           
                //擷取輸入的URL
                string url = textBox1.Text;            
                //string url = "http://mail.163.com/";            
                //建立http連結            
                //HttpWebRequest對象執行個體:該類用於擷取和操作HTTP請求 var可改成HttpWebRequest
                var request = (HttpWebRequest)WebRequest.Create(url);   
                //建立WebRequest對象            
                //HttpWebResponse對象執行個體:該類用於擷取和操作HTTP應答 var可改成HttpWebResponse
                var response = (HttpWebResponse)request.GetResponse();  
                //GetResponse:擷取回覆            
                //構造資料流對象執行個體
                Stream stream = response.GetResponseStream();  
                //GetResponseStream:擷取應答流
                StreamReader sr = new StreamReader(stream);    
                //從位元組流中讀取字元            
                //從流當前位置讀取到末尾並顯示在WebBrower控制項中
                string content = sr.ReadToEnd();            
                webBrowser1.DocumentText = content;        
            }        
            //web瀏覽器控制項中承載的文檔全部載入後發生
            private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)        
            {            
                //定義html元素 通過Name擷取控制項值            
                //HtmlElement tbUserid = webBrowser1.Document.All["userName"];            
                //HtmlElement tbPasswd = webBrowser1.Document.All["password"];                        
                //定義html元素 通過ID擷取控制項值 (使用者名稱 密碼 登入按鈕)
                HtmlElement tbUserid = webBrowser1.Document.GetElementById("idInput");            
                HtmlElement tbPasswd = webBrowser1.Document.GetElementById("pwdInput");            
                HtmlElement btnSubmit = webBrowser1.Document.GetElementById("loginBtn");            
                //三個元素其一為空白返回 載入後才執行賦值 否則會出現為null值的崩潰錯誤
                if (tbUserid == null || tbPasswd == null || btnSubmit == null)            
                {                
                    return;            
                }                        
                //設定元素value屬性值 (使用者名稱 密碼值)
                tbUserid.SetAttribute("value", "Eastmount");            
                tbPasswd.SetAttribute("value", "Eastmount");                        
                //執行元素的方法：如click submit
                btnSubmit.InvokeMember("click");             
            }    
        }
    }

    //擷取網頁元素 (使用者名稱 密碼 登入按鈕)
    HtmlElement tbUserid = webBrowser1.Document.GetElementById("idInput");
    HtmlElement tbPasswd = webBrowser1.Document.GetElementById("pwdInput");
    HtmlElement btnSubmit = webBrowser1.Document.GetElementById("loginBtn");
    //用相應方法為元素賦值
    tbUserid.SetAttribute("value", "Eastmount");
    tbPasswd.SetAttribute("value", "Eastmount");
    btnSubmit.InvokeMember("click");

    */
    /*
     //https://stackoverflow.com/questions/45802507/c-sharp-web-api-response-handling-using-list
    
    [HttpPost]
    public List<ValueStory> UserValueStories([FromBody] ValueStory valuestory)
    //public void UserValueStories([FromBody] ValueStory Id)
    {
        if (valuestory.Id == "" || valuestory.Id == null)
        {
            //what code to add to change status code to 400 and to display error message?
        }

        //what code to put if the id is not valid, what status code and what message?


        var valueStoryName = (from vs in db.ValueStories
                              where vs.Id == valuestory.Id
                              select vs).ToList();


        List<ValueStory> vs1 = new List<ValueStory>();
        foreach (var v in valueStoryName)
        {
            vs1.Add(new ValueStory()
            {
                Id = v.Id,
                ValueStoryName = v.ValueStoryName,
                Organization = v.Organization,
                Industry = v.Industry,
                Location = v.Location,
                AnnualRevenue = v.AnnualRevenue,
                CreatedDate = v.CreatedDate,
                ModifiedDate = v.ModifiedDate,
                MutualActionPlan = v.MutualActionPlan,
                Currency = v.Currency,
                VSId = v.VSId
            });
        }
        return vs1.ToList();

    }
    public class ValueStory
    {
        public string Id { get; set; }

        public string ValueStoryName { get; set; }
        public string Organization { get; set; }
        public string Industry { get; set; }
        public string Location { get; set; }
        public string AnnualRevenue { get; set; }
        public string CreatedDate { get; set; }
        public string ModifiedDate { get; set; }
        public string MutualActionPlan { get; set; }
        public string Currency { get; set; }
        public string VSId { get; set; }
    }

     */

}
